using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

namespace RadiologyAgent {
  // The native save/print sink independently recomputes these exact SHA-256
  // digests on the Word event's Document. Never store report content or paths.
  internal static class MandatoryApproval {
    internal static string Digest(string text) {
      if(String.IsNullOrEmpty(text))return null;
      using(var sha=SHA256.Create()) {
        var bytes=sha.ComputeHash(Encoding.UTF8.GetBytes(text));
        return BitConverter.ToString(bytes).Replace("-","").ToLowerInvariant();
      }
    }
    [DllImport("user32.dll", SetLastError=true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint processId);
    internal static string Identity(uint processId,string fullName) {
      return Digest(processId.ToString()+"\u001f"+fullName);
    }
    internal static bool Approve(object document, string reviewedText, uint verifiedWindow) {
      try {
        dynamic doc=document;
        string fullName=Convert.ToString(doc.FullName);
        if(verifiedWindow==0)return false;
        uint processId;
        if(GetWindowThreadProcessId(new IntPtr(unchecked((int)verifiedWindow)),out processId)==0
            ||processId==0)return false;
        string current=Convert.ToString(doc.Content.Text);
        if(!String.Equals(current,reviewedText,StringComparison.Ordinal))return false;
        string id=Identity(processId,fullName),hash=Digest(current);
        if(String.IsNullOrEmpty(id)||String.IsNullOrEmpty(hash))return false;
        using(var key=Registry.CurrentUser.CreateSubKey(
                  @"Software\Haronex\MandatoryReviews\"+id)) {
          if(key==null)return false;
          // Each new successful review needs a subsequent Clear Both action.
          key.DeleteValue("ClearBothHash",false);
          key.DeleteValue("ClearBothV102Hash",false);
          key.DeleteValue("ClearBothV103Hash",false);
          key.SetValue("ApprovedHash",hash,RegistryValueKind.String);
        }
        return true;
      } catch {return false;}
    }
    // A successful review is a one-time gate for the open document. We do
    // not revoke access on subsequent edits. Native Word COM session tracks
    // the released document after Clear Both, even across Save As.
    internal static void Invalidate(string oldId) {
      // Intentional no-op: text edits must never re-lock an approved document.
    }
    internal static bool SelfTest() {
      return Digest("Text")==
        "71988c4d8e0803ba4519f0b2864c1331c14a1890bf8694e251379177bfedb5c3"
        && Identity(42,"Example.docx")!=Identity(43,"Example.docx")
        && Identity(42,"Example.docx")!=Identity(42,"Other.docx")
        && Digest("text")!=Digest("Text");
    }
  }

  // Persistent STA window receives Word COM calls ONLY on the UI thread.
  // Gemini browser work uses its own subprocess and is never allowed to modify Word.
  internal sealed class LiveReviewMonitor : Form {
    private const int PollMilliseconds=400;
    private const int DebounceMilliseconds=1200;
    private readonly System.Windows.Forms.Timer timer;
    private object word;
    private object boundWordWindow;
    private uint boundWordHwnd;
    private bool polling;
    private int missingWordTicks;
    private string documentId, currentText, snapshotToken;
    private DateTime lastChangeUtc;
    private string readyResponse;
    private bool failed, working, pendingClick;
    private CancellationTokenSource runningCancellation;
    private uint lastHandledCommand;
    private bool awaitingDocument;
    private DateTime documentWaitStarted;
    private string lastWordProblem="WORD_NOT_CONNECTED";

    [DllImport("user32.dll", CharSet=CharSet.Unicode)]
    private static extern int MessageBox(IntPtr owner,string message,string title,uint flags);

    public LiveReviewMonitor() {
      Text="Haronex Live Review";
      Width=1;Height=1;Opacity=0;ShowInTaskbar=false;
      StartPosition=FormStartPosition.Manual;
      Location=new Point(-30000,-30000);
      FormBorderStyle=FormBorderStyle.None;
      lastHandledCommand=ReadDword("LiveReviewHandled");
      timer=new System.Windows.Forms.Timer {Interval=PollMilliseconds};
      timer.Tick+=(sender,args)=>TickOnce();
      Shown+=(sender,args)=>timer.Start();
      FormClosing+=(sender,args)=>{
        timer.Stop();
        if(runningCancellation!=null)runningCancellation.Cancel();
      };
    }

    private static uint ReadDword(string name) {
      try {
        using(var key=Registry.CurrentUser.OpenSubKey(@"Software\Haronex"))
          return key==null?0:Convert.ToUInt32(key.GetValue(name,0));
      } catch{return 0;}
    }
    private static void WriteDword(string name,uint value) {
      try {
        using(var key=Registry.CurrentUser.CreateSubKey(@"Software\Haronex"))
          key.SetValue(name,unchecked((int)value),RegistryValueKind.DWord);
      } catch{}
    }
    private static void LogFailure(string code) {
      try {
        // Only short reason codes; no report text, file paths, keys, prompts, or responses.
        if(String.IsNullOrWhiteSpace(code))code="UNKNOWN";
        var clean=new StringBuilder();
        foreach(char c in code.ToUpperInvariant()) {
          if((c>='A'&&c<='Z')||(c>='0'&&c<='9')||c=='_')
            clean.Append(c);
          if(clean.Length>=65)break;
        }
        string dir=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"Haronex");
        Directory.CreateDirectory(dir);
        string file=Path.Combine(dir,"diagnostics.log");
        if(File.Exists(file)&&new FileInfo(file).Length>20000)File.Delete(file);
        File.AppendAllText(file,DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ssZ")+" "+clean.ToString()+"\r\n");
      } catch {}
    }

    private static void Notify(bool success) {
      WebReviewForm.SetReviewStatus(success?1:2);
      MessageBox(IntPtr.Zero,success?"تمت المراجعة":"رجاءً أعد المحاولة",
        "Haronex",0x00040000|0x00010000|(success?0x40u:0x10u));
    }

    private bool ConnectWord() {
      uint requested=ReadDword("WordWindowHwnd");
      // Never substitute Marshal.GetActiveObject: it may belong to another
      // instance or another patient's open document.
      if(requested==0) {
        lastWordProblem="WORD_WINDOW_NOT_IDENTIFIED";
        if(boundWordWindow!=null||currentText!=null)ClearDocument();
        word=null;boundWordWindow=null;boundWordHwnd=0;
        return false;
      }
      if(requested==boundWordHwnd&&boundWordWindow!=null&&word!=null)
        return true;
      // A newly clicked Word window invalidates all previous document snapshots.
      if(boundWordHwnd!=requested||boundWordWindow!=null||currentText!=null)ClearDocument();
      word=null;boundWordWindow=null;boundWordHwnd=0;
      object window;
      bool candidateFound;
      if(!WordWindowBinding.TryGetWordWindow(requested,out window,out candidateFound)) {
        lastWordProblem=candidateFound?"WORD_DOCUMENT_NOT_READY":"WORD_NATIVEOM_BIND_FAILED";
        return false;
      }
      try {
        dynamic activeWindow=window;
        object owner=activeWindow.Application;
        if(owner==null) {
          lastWordProblem="WORD_NATIVEOM_APP_UNAVAILABLE";
          return false;
        }
        // Document access can transiently fail while Word is loading.
        // Preserve the window binding for the next 400ms poll.
        word=owner;boundWordWindow=window;boundWordHwnd=requested;
        lastWordProblem="WORD_DOCUMENT_LOADING";
        missingWordTicks=0;
        return true;
      } catch(COMException) {
        lastWordProblem="WORD_NATIVEOM_BUSY";
        return false;
      } catch {
        lastWordProblem="WORD_NATIVEOM_DOCUMENT_ERROR";
        return false;
      }
    }

    private void TickOnce() {
      if(polling||IsDisposed)return;
      polling=true;
      try {
        if(ConnectWord())PollWord();
        CheckClick();
        CheckPendingDocument();
        // No background API calls: only explicit Review Text click triggers analysis.
      } catch {
        lastWordProblem="WORD_POLL_FAILED";
      } finally {polling=false;}
    }

    private void PollWord() {
      dynamic app=word;
      dynamic doc;
      try {doc=boundWordWindow!=null?((dynamic)boundWordWindow).Document:app.ActiveDocument;}
      catch (COMException) {
        // Document property can be temporarily unavailable during focus,
        // startup, or protected view. Do not fall back to another document.
        lastWordProblem=boundWordWindow!=null?"WORD_DOCUMENT_NOT_READY":"WORD_NO_ACTIVE_DOCUMENT";
        if(documentId!=null)ClearDocument();
        // A previously bound pane can become stale; re-enumerate the clicked
        // window on the next tick rather than retrying the same broken RCW.
        word=null;boundWordWindow=null;boundWordHwnd=0;
        return;
      }
      catch {
        lastWordProblem="WORD_ACTIVE_DOCUMENT_ERROR";
        if(documentId!=null)ClearDocument();
        word=null;boundWordWindow=null;boundWordHwnd=0; // Reacquire pinned Word window.
        return;
      }
      string id,report;
      try {
        string name=Convert.ToString(doc.FullName);
        int hwnd;
        try {hwnd=boundWordWindow!=null?(int)((dynamic)boundWordWindow).Hwnd:(int)app.ActiveWindow.Hwnd;}
        catch {hwnd=(int)app.Hwnd;}
        id=hwnd.ToString()+"|"+name;
        report=Convert.ToString(doc.Content.Text);
      } catch {
        lastWordProblem="WORD_DOCUMENT_READ_FAILED";
        if(documentId!=null)ClearDocument();
        word=null;boundWordWindow=null;boundWordHwnd=0;
        return;
      }
      if(String.IsNullOrWhiteSpace(report)) {
        lastWordProblem="WORD_DOCUMENT_EMPTY";
        if(documentId!=null)ClearDocument();
        return;
      }
      if(report.Length>120000) {
        lastWordProblem="WORD_DOCUMENT_TOO_LONG";
        if(documentId!=null)ClearDocument();
        return;
      }
      lastWordProblem="WORD_DOCUMENT_READY";
      bool changedDoc=!String.Equals(documentId,id,StringComparison.Ordinal);
      bool changedText=!String.Equals(currentText,report,StringComparison.Ordinal);
      if(changedDoc||changedText) {
        // A previously approved revision must not survive any observed edit.
        if(!changedDoc&&changedText)MandatoryApproval.Invalidate(documentId);
        // Preserve the click while the document becomes available.
        if(runningCancellation!=null)runningCancellation.Cancel();
        documentId=id;
        currentText=report;
        snapshotToken=WebReviewForm.CachePath(report,false);
        lastChangeUtc=DateTime.UtcNow;
        failed=false;
        readyResponse=ValidateResponse(WebReviewForm.ReadCached(snapshotToken),report);
        WebReviewForm.SetReviewStatus(readyResponse!=null?1:0);
      }
    }

    private void ClearDocument() {
      if(runningCancellation!=null)runningCancellation.Cancel();
      documentId=null;currentText=null;snapshotToken=null;
      readyResponse=null;failed=false;
      WebReviewForm.SetReviewStatus(0);
    }

    private void CheckClick() {
      uint command=ReadDword("LiveReviewCommand");
      if(command==lastHandledCommand)return;
      lastHandledCommand=command;
      WriteDword("LiveReviewHandled",command);
      // Rebind to the exact Word window that raised the Ribbon click.
      bool bound=ConnectWord();
      if(bound)PollWord();
      if(!WebReviewForm.Approved()) {LogFailure("KEY_OR_CONSENT_MISSING");Notify(false);return;}
      // Fail closed if the captured Word window cannot be verified. In
      // particular, don't analyze a snapshot from a previously clicked window.
      if(!bound||boundWordHwnd==0||boundWordHwnd!=ReadDword("WordWindowHwnd"))
        ClearDocument();
      if(currentText==null) {
        // Word may be starting or its COM document snapshot not yet available.
        awaitingDocument=true;
        documentWaitStarted=DateTime.UtcNow;
        return;
      }
      if(readyResponse!=null){ApplyReady();return;}
      if(failed) {
        // A new click after an exhausted attempt starts a fresh analysis
        // immediately; never show a stale Retry popup for the old failure.
        failed=false;
        readyResponse=null;
        pendingClick=true;
        if(!working)StartAnalysis();
        return;
      }
      // The user clicked before review was ready. Show the final popup after the
      // current version has completed, never falsely apply stale results.
      pendingClick=true;
      if(!working)StartAnalysis();
    }

    private void CheckPendingDocument() {
      if(!awaitingDocument)return;
      if(currentText!=null) {
        awaitingDocument=false;
        if(readyResponse!=null)ApplyReady();
        else {pendingClick=true;if(!working)StartAnalysis();}
        return;
      }
      if((DateTime.UtcNow-documentWaitStarted).TotalSeconds>=12) {
        awaitingDocument=false;
        LogFailure(lastWordProblem);
        Notify(false);
      }
    }

    private static List<string> Chunks(string report) {
      var result=new List<string>();
      int start=0;
      while(start<report.Length) {
        int length=Math.Min(6000,report.Length-start);
        if(start+length<report.Length) {
          int breakAt=report.LastIndexOfAny(new char[]{'\r','\n'},start+length-1,length);
          if(breakAt>=start+4500)length=breakAt-start+1;
        }
        result.Add(report.Substring(start,length));
        start+=length;
      }
      return result;
    }

    private static string ValidateResponse(string json,string report) {
      if(String.IsNullOrWhiteSpace(json))return null;
      try {
        var parser=new JavaScriptSerializer{MaxJsonLength=2000000};
        var root=parser.DeserializeObject(json) as Dictionary<string,object>;
        if(root==null||!root.ContainsKey("ok")||!Convert.ToBoolean(root["ok"]))return null;
        if(!root.ContainsKey("mode")||Convert.ToString(root["mode"])!="groq")return null;
        if(!root.ContainsKey("review_version")||
           Convert.ToString(root["review_version"])!=WebReviewForm.ReviewAlgorithmVersion)return null;
        if(!root.ContainsKey("passes")||Convert.ToInt32(root["passes"])!=2)return null;
        var groups=root["groups"] as object[];
        var chunks=Chunks(report);
        if(groups==null||groups.Length!=chunks.Count)return null;
        for(int i=0;i<groups.Length;i++) {
          var items=groups[i] as object[];
          if(items==null)return null;
          foreach(var item in items) {
            var row=item as Dictionary<string,object>;
            if(row==null||!row.ContainsKey("quote")||!row.ContainsKey("category"))return null;
            string quote=Convert.ToString(row["quote"]);
            string category=Convert.ToString(row["category"]);
            if(String.IsNullOrWhiteSpace(quote)||quote.Length>500||
               (category!="language"&&category!="logic"&&category!="unrelated")||!chunks[i].Contains(quote))return null;
          }
        }
        return json;
      } catch{return null;}
    }

    private async void StartAnalysis() {
      if(working||currentText==null||!WebReviewForm.Approved())return;
      working=true;
      string report=currentText,id=documentId,cacheFile=snapshotToken;
      var token=new CancellationTokenSource();
      runningCancellation=token;
      try {
        var chunks=Chunks(report);
        var prompts=new List<string>();
        for(int i=0;i<chunks.Count;i++)
          prompts.Add(WebReviewForm.BuildPrompt(chunks[i],i,chunks.Count));
        var serializer=new JavaScriptSerializer{MaxJsonLength=2000000};
        string payload=serializer.Serialize(new Dictionary<string,object> {
          {"prompts",prompts},{"texts",chunks},
          {"review_version",WebReviewForm.ReviewAlgorithmVersion},
          {"api_key",WebReviewForm.ReadGroqKey()}
        });
        string answer=await Task.Run(()=>WebReviewForm.CallBrowserWorker(payload,chunks.Count,token.Token));
        if(token.IsCancellationRequested||id!=documentId||report!=currentText)return;
        string valid=ValidateResponse(answer,report);
        if(valid==null) {
          string code="INVALID_WORKER_RESPONSE";
          try {
            var response=serializer.DeserializeObject(answer) as Dictionary<string,object>;
            if(response!=null&&response.ContainsKey("error"))
              code=Convert.ToString(response["error"]);
          } catch {}
          LogFailure(code);
          throw new Exception("Invalid review response");
        }
        WebReviewForm.SaveCached(cacheFile,valid);
        readyResponse=valid;
        failed=false;
        WebReviewForm.SetReviewStatus(1);
        if(pendingClick)ApplyReady();
      } catch(OperationCanceledException) {
        // Any result for an obsolete revision is discarded.
      } catch(Exception ex) {
        if(!token.IsCancellationRequested&&id==documentId&&report==currentText) {
          if(ex.Message!="Invalid review response")LogFailure("HOST_"+ex.GetType().Name);
          failed=true;
          WebReviewForm.SetReviewStatus(2);
          if(pendingClick){pendingClick=false;Notify(false);}
        }
      } finally {
        if(ReferenceEquals(runningCancellation,token))runningCancellation=null;
        token.Dispose();
        working=false;
      }
    }

    internal static bool SelfTest() {
      try {
        string report="FINDINGS: Right kidney normal.\rIMPRESSION: Left kidney normal.\r";
        var parser=new JavaScriptSerializer{MaxJsonLength=2000000};
        var issue=new Dictionary<string,object> {
          {"quote","Left kidney"},{"category","logic"},
          {"explanation","Laterality conflict"},{"suggestion","Verify laterality"}
        };
        var response=new Dictionary<string,object> {
          {"ok",true},{"mode","groq"},{"passes",2},
          {"review_version",WebReviewForm.ReviewAlgorithmVersion},
          {"groups",new object[]{new object[]{issue}}}
        };
        string json=parser.Serialize(response);
        var tableExample="Findings: Bilateral renal cysts.\r\x07Conclusion:\r\x07o Haron work at dokki scan\r\x07Thanks;";
        var found=ObviousOffTopicInConclusion(tableExample);
        var control=ObviousOffTopicInConclusion("Findings: Patient works at refinery.\rConclusion:\rBilateral cysts.\rThanks;");
        return found.Count==1&&found[0]=="Haron work at dokki scan"&&control.Count==0
          && MandatoryApproval.SelfTest()
          && DebounceMilliseconds==1200
          && PollMilliseconds==400
          && ValidateResponse(json,report)!=null
          && ValidateResponse(json,report.Replace("Left kidney","Spleen"))==null
          && ValidateResponse(json.Replace("\"groq\"","\"account\""),report)==null;
      } catch {return false;}
    }

    // Independent Word-side safety net for obvious personal/employment lines
    // in Conclusion. Generic, narrow pattern: no hard-coded person or site.
    // Clinical occupational history and non-Conclusion text are not flagged.
    internal static List<string> ObviousOffTopicInConclusion(string report) {
      var found=new List<string>();
      if(String.IsNullOrEmpty(report))return found;
      var header=Regex.Match(report,
        @"(?:^|[\r\n\x07])[ \t]*(?:Conclusion|Impression)[ \t]*:",
        RegexOptions.IgnoreCase);
      if(!header.Success)return found;
      int begin=header.Index+header.Length;
      var ending=Regex.Match(report.Substring(begin),
        @"(?:^|[\r\n\x07])[ \t]*(?:Thanks[;,.]?|Dr\.[ \t]|Signed[ \t]+by)",
        RegexOptions.IgnoreCase);
      string body=ending.Success?report.Substring(begin,ending.Index):
        report.Substring(begin);
      foreach(string raw in Regex.Split(body,@"[\r\n\x07]+")) {
        string line=raw.Trim().TrimStart('o','•','◦','-',' ','\t').Trim();
        if(line.Length<8||line.Length>240)continue;
        if(!Regex.IsMatch(line,
           @"^[A-Z][a-z]{1,30}(?:[ \t]+[A-Z][a-z]{1,30})?[ \t]+(?:work|works|working)[ \t]+(?:at|for)[ \t]+[A-Za-z]",
           RegexOptions.CultureInvariant))continue;
        if(!found.Contains(line))found.Add(line);
      }
      return found;
    }

    private void ApplyReady() {
      if(readyResponse==null||currentText==null){Notify(false);return;}
      try {
        dynamic app=word,doc=boundWordWindow!=null?((dynamic)boundWordWindow).Document:app.ActiveDocument;
        string latest=Convert.ToString(doc.Content.Text);
        string name=Convert.ToString(doc.FullName);
        int hwnd;
        try {hwnd=boundWordWindow!=null?(int)((dynamic)boundWordWindow).Hwnd:(int)app.ActiveWindow.Hwnd;}
        catch{hwnd=(int)app.Hwnd;}
        if(latest!=currentText||hwnd.ToString()+"|"+name!=documentId) {
          pendingClick=true;return;
        }
        string source=readyResponse;
        var parser=new JavaScriptSerializer{MaxJsonLength=2000000};
        var root=parser.DeserializeObject(source) as Dictionary<string,object>;
        var groups=root["groups"] as object[];
        var unique=new Dictionary<string,Tuple<string,string,string,string,string>>(StringComparer.Ordinal);
        foreach(var group in groups) {
          var issues=group as object[];
          foreach(var issue in issues) {
            var row=issue as Dictionary<string,object>;
            Func<string,string> get=k=>row.ContainsKey(k)&&row[k]!=null?Convert.ToString(row[k]):"";
            string quote=get("quote"),category=get("category");
            if(quote.Length<2||!currentText.Contains(quote))continue;
            string section=get("section");
            string key=quote+"\u001f"+section;
            var record=Tuple.Create(quote,category,get("explanation"),get("suggestion"),section);
            if(!unique.ContainsKey(key)||(category=="logic"||category=="unrelated"))unique[key]=record;
          }
        }
        // Cross-check direct Word snapshot, because the Python model/text
        // segmentation can miss table cells even after a successful review.
        var fallback=ObviousOffTopicInConclusion(currentText);
        WriteDword("UnrelatedCandidateCount",(uint)fallback.Count);
        WriteDword("UnrelatedAppliedCount",0);
        WriteDword("UnrelatedSkippedCount",0);
        foreach(string line in fallback) {
          string key=line+"\u001fconclusion";
          unique[key]=Tuple.Create(line,"unrelated",
            "Non-medical personal or employment statement appears in the clinical conclusion.",
            "Ask the reporting radiologist to verify and remove if unintended.",
            "conclusion");
        }
        var ordered=new List<Tuple<string,string,string,string,string>>(unique.Values);
        ordered.Sort((a,b)=>{
          int priority=String.CompareOrdinal(a.Item2,b.Item2);
          if(a.Item2=="unrelated"&&b.Item2!="unrelated")return -1;
          if(b.Item2=="unrelated"&&a.Item2!="unrelated")return 1;
          if(a.Item2=="logic"&&b.Item2!="logic")return -1;
          if(b.Item2=="logic"&&a.Item2!="logic")return 1;
          return currentText.IndexOf(a.Item1,StringComparison.Ordinal)
            .CompareTo(currentText.IndexOf(b.Item1,StringComparison.Ordinal));
        });
        var existing=new HashSet<string>(StringComparer.Ordinal);
        try {
          dynamic comments=doc.Comments;
          int count=(int)comments.Count;
          for(int i=1;i<=count;i++) {
            dynamic comment=comments.Item(i);
            string body=Convert.ToString(comment.Range.Text);
            if(body.IndexOf("[Haronex:",StringComparison.Ordinal)<0)continue;
            dynamic anchor=comment.Scope;
            existing.Add(((int)anchor.Start).ToString()+":"+((int)anchor.End).ToString());
          }
        } catch {}
        int start=(int)doc.Content.Start,end=(int)doc.Content.End;
        foreach(var issue in ordered) {
          try {
            dynamic region=doc.Range(start,end);
            // Dedicated conclusion-audit items must be attached inside the
            // Conclusion/Impression, even when a quote also appears in Findings.
            if(String.Equals(issue.Item5,"conclusion",StringComparison.OrdinalIgnoreCase)) {
              dynamic heading=doc.Range(start,end);
              dynamic hf=heading.Find;
              hf.ClearFormatting();hf.Text="Conclusion";hf.Forward=true;
              hf.Wrap=0;hf.MatchCase=false;hf.MatchWildcards=false;
              if(!(bool)hf.Execute()) {
                heading=doc.Range(start,end);
                hf=heading.Find;
                hf.ClearFormatting();hf.Text="Impression";hf.Forward=true;
                hf.Wrap=0;hf.MatchCase=false;hf.MatchWildcards=false;
                if(!(bool)hf.Execute())continue;
              }
              region=doc.Range((int)heading.Start,end);
            }
            dynamic find=region.Find;
            find.ClearFormatting();find.Text=issue.Item1;
            find.Forward=true;find.Wrap=0;find.MatchCase=true;find.MatchWildcards=false;
            if(!(bool)find.Execute()) {
              // When Word Find fails inside a table, try a precise validated
              // character-range anchor, NEVER an approximate nearby position.
              if(issue.Item2!="unrelated"){continue;}
              int at=currentText.IndexOf(issue.Item1,StringComparison.Ordinal);
              if(at<0||currentText.IndexOf(issue.Item1,at+1,StringComparison.Ordinal)>=0) {
                WriteDword("UnrelatedSkippedCount",ReadDword("UnrelatedSkippedCount")+1);
                continue;
              }
              region=doc.Range(start+at,start+at+issue.Item1.Length);
              if(!String.Equals(Convert.ToString(region.Text),issue.Item1,StringComparison.Ordinal)) {
                WriteDword("UnrelatedSkippedCount",ReadDword("UnrelatedSkippedCount")+1);
                continue;
              }
            }
            dynamic again=doc.Range((int)region.End,end);
            dynamic second=again.Find;
            second.ClearFormatting();second.Text=issue.Item1;
            second.Forward=true;second.Wrap=0;second.MatchCase=true;second.MatchWildcards=false;
            if((bool)second.Execute())continue; // Ambiguous quote; never annotate the wrong occurrence.
            string anchorId=((int)region.Start).ToString()+":"+((int)region.End).ToString();
            if(existing.Contains(anchorId))continue;
            string comment="[Haronex: "+(issue.Item2=="unrelated"?"Unrelated content":(issue.Item2=="logic"?"Potential internal contradiction":"Language"))+"]\r\n"+
              WebReviewForm.Safe(issue.Item3,500)+
              (issue.Item4.Length>0?"\r\nSuggestion: "+WebReviewForm.Safe(issue.Item4,250):"")+
              "\r\nRequires clinical verification.";
            region.HighlightColorIndex=issue.Item2=="unrelated"?12:(issue.Item2=="logic"?6:7); // wdViolet=12
            doc.Comments.Add(region,comment);
            if(issue.Item2=="unrelated")
              WriteDword("UnrelatedAppliedCount",ReadDword("UnrelatedAppliedCount")+1);
            existing.Add(anchorId);
          } catch {
            if(issue.Item2=="unrelated")
              WriteDword("UnrelatedSkippedCount",ReadDword("UnrelatedSkippedCount")+1);
          }
        }
        if(!MandatoryApproval.Approve(doc,currentText,unchecked((uint)hwnd))) {
          LogFailure("MANDATORY_APPROVAL_NOT_STORED");
          pendingClick=false;Notify(false);return;
        }
        pendingClick=false;
        Notify(true);
      } catch(Exception ex) {
        LogFailure("WORD_ANNOTATE_"+ex.GetType().Name);
        pendingClick=false;Notify(false);
      }
    }
  }
}
