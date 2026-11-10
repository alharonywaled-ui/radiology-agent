using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

namespace RadiologyAgent {
  // Persistent STA window receives Word COM calls ONLY on the UI thread.
  // Gemini browser work uses its own subprocess and is never allowed to modify Word.
  internal sealed class LiveReviewMonitor : Form {
    private const int PollMilliseconds=400;
    private const int DebounceMilliseconds=1200;
    private readonly System.Windows.Forms.Timer timer;
    private object word;
    private bool polling;
    private int missingWordTicks;
    private string documentId, currentText, snapshotToken;
    private DateTime lastChangeUtc;
    private string readyResponse;
    private bool failed, working, pendingClick;
    private CancellationTokenSource runningCancellation;
    private uint lastHandledCommand;

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
    private static void Notify(bool success) {
      WebReviewForm.SetReviewStatus(success?1:2);
      MessageBox(IntPtr.Zero,success?"تمت المراجعة":"رجاءً أعد المحاولة",
        "Haronex",0x00040000|0x00010000|(success?0x40u:0x10u));
    }

    private bool ConnectWord() {
      if(word!=null)return true;
      try {
        word=Marshal.GetActiveObject("Word.Application");
        missingWordTicks=0;
        return true;
      } catch {
        // Word may not yet be present in the ROT at add-in OnConnection.
        if(++missingWordTicks>90)Close();
        return false;
      }
    }

    private void TickOnce() {
      if(polling||IsDisposed)return;
      polling=true;
      try {
        if(!ConnectWord())return;
        PollWord();
        CheckClick();
        // No background API calls: only explicit Review Text click triggers analysis.
      } catch {
        // Word may temporarily be in a modal operation; retry without leaking report text.
      } finally {polling=false;}
    }

    private void PollWord() {
      dynamic app=word;
      dynamic doc;
      try {doc=app.ActiveDocument;}
      catch {
        if(documentId!=null)ClearDocument();
        return;
      }
      string id,report;
      try {
        string name=Convert.ToString(doc.FullName);
        int hwnd;
        try {hwnd=(int)app.ActiveWindow.Hwnd;}
        catch {hwnd=(int)app.Hwnd;}
        id=hwnd.ToString()+"|"+name;
        report=Convert.ToString(doc.Content.Text);
      } catch {return;}
      if(String.IsNullOrWhiteSpace(report)||report.Length>120000) {
        if(documentId!=null)ClearDocument();
        return;
      }
      bool changedDoc=!String.Equals(documentId,id,StringComparison.Ordinal);
      if(changedDoc||!String.Equals(currentText,report,StringComparison.Ordinal)) {
        if(changedDoc)pendingClick=false;
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
      readyResponse=null;failed=false;pendingClick=false;
      WebReviewForm.SetReviewStatus(0);
    }

    private void CheckClick() {
      uint command=ReadDword("LiveReviewCommand");
      if(command==lastHandledCommand)return;
      lastHandledCommand=command;
      WriteDword("LiveReviewHandled",command);
      if(currentText==null) {Notify(false);return;}
      if(readyResponse!=null){ApplyReady();return;}
      if(failed) {
        Notify(false);
        failed=false;
        lastChangeUtc=DateTime.UtcNow; // next click can reuse the new analysis
        return;
      }
      // The user clicked before review was ready. Show the final popup after the
      // current version has completed, never falsely apply stale results.
      pendingClick=true;
      if(!working)StartAnalysis();
    }

    private static List<string> Chunks(string report) {
      var result=new List<string>();
      int start=0;
      while(start<report.Length) {
        int length=Math.Min(12000,report.Length-start);
        if(start+length<report.Length) {
          int breakAt=report.LastIndexOf('\n',start+length-1,length);
          if(breakAt>=start+6000)length=breakAt-start+1;
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
        if(!root.ContainsKey("passes")||Convert.ToInt32(root["passes"])!=1)return null;
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
               (category!="language"&&category!="logic")||!chunks[i].Contains(quote))return null;
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
        if(valid==null)throw new Exception("Incomplete review");
        WebReviewForm.SaveCached(cacheFile,valid);
        readyResponse=valid;
        failed=false;
        WebReviewForm.SetReviewStatus(1);
        if(pendingClick)ApplyReady();
      } catch(OperationCanceledException) {
        // Any result for an obsolete revision is discarded.
      } catch {
        if(!token.IsCancellationRequested&&id==documentId&&report==currentText) {
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
          {"ok",true},{"mode","groq"},{"passes",1},
          {"review_version",WebReviewForm.ReviewAlgorithmVersion},
          {"groups",new object[]{new object[]{issue}}}
        };
        string json=parser.Serialize(response);
        return DebounceMilliseconds==1200
          && ValidateResponse(json,report)!=null
          && ValidateResponse(json,report.Replace("Left kidney","Spleen"))==null
          && ValidateResponse(json.Replace("\"groq\"","\"account\""),report)==null;
      } catch {return false;}
    }

    private void ApplyReady() {
      if(readyResponse==null||currentText==null){Notify(false);return;}
      try {
        dynamic app=word,doc=app.ActiveDocument;
        string latest=Convert.ToString(doc.Content.Text);
        string name=Convert.ToString(doc.FullName);
        int hwnd;
        try {hwnd=(int)app.ActiveWindow.Hwnd;}
        catch{hwnd=(int)app.Hwnd;}
        if(latest!=currentText||hwnd.ToString()+"|"+name!=documentId) {
          pendingClick=true;return;
        }
        string source=readyResponse;
        var parser=new JavaScriptSerializer{MaxJsonLength=2000000};
        var root=parser.DeserializeObject(source) as Dictionary<string,object>;
        var groups=root["groups"] as object[];
        var unique=new Dictionary<string,Tuple<string,string,string,string>>(StringComparer.Ordinal);
        foreach(var group in groups) {
          var issues=group as object[];
          foreach(var issue in issues) {
            var row=issue as Dictionary<string,object>;
            Func<string,string> get=k=>row.ContainsKey(k)&&row[k]!=null?Convert.ToString(row[k]):"";
            string quote=get("quote"),category=get("category");
            if(quote.Length<2||!currentText.Contains(quote))continue;
            string key=quote;
            var record=Tuple.Create(quote,category,get("explanation"),get("suggestion"));
            if(!unique.ContainsKey(key)||category=="logic")unique[key]=record;
          }
        }
        var ordered=new List<Tuple<string,string,string,string>>(unique.Values);
        ordered.Sort((a,b)=>{
          int priority=String.CompareOrdinal(a.Item2,b.Item2);
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
            dynamic find=region.Find;
            find.ClearFormatting();find.Text=issue.Item1;
            find.Forward=true;find.Wrap=0;find.MatchCase=true;find.MatchWildcards=false;
            if(!(bool)find.Execute())continue;
            dynamic again=doc.Range((int)region.End,end);
            dynamic second=again.Find;
            second.ClearFormatting();second.Text=issue.Item1;
            second.Forward=true;second.Wrap=0;second.MatchCase=true;second.MatchWildcards=false;
            if((bool)second.Execute())continue; // Ambiguous quote; never annotate the wrong occurrence.
            string anchorId=((int)region.Start).ToString()+":"+((int)region.End).ToString();
            if(existing.Contains(anchorId))continue;
            string comment="[Haronex: "+(issue.Item2=="logic"?"Potential internal contradiction":"Language")+"]\r\n"+
              WebReviewForm.Safe(issue.Item3,500)+
              (issue.Item4.Length>0?"\r\nSuggestion: "+WebReviewForm.Safe(issue.Item4,250):"")+
              "\r\nRequires clinical verification.";
            region.HighlightColorIndex=issue.Item2=="logic"?6:7;
            doc.Comments.Add(region,comment);
            existing.Add(anchorId);
          } catch {}
        }
        pendingClick=false;
        Notify(true);
      } catch {
        pendingClick=false;Notify(false);
      }
    }
  }
}
