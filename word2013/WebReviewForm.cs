using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

namespace RadiologyAgent {
  internal sealed class WebReviewForm : Form {
    private readonly object word;
    private readonly Label status;
    private bool running;
    internal const string ReviewAlgorithmVersion="haronex-groq-qwen38-v9-4";
    public WebReviewForm(object application) {
      word=application;
      Text="Haronex — Word 2013"; Width=1; Height=1;
      ShowInTaskbar=false; Opacity=0;
      StartPosition=FormStartPosition.Manual; Location=new Point(-30000,-30000);
      FormBorderStyle=FormBorderStyle.None;
      MaximizeBox=false;MinimizeBox=false;
      var title=new Label{Text="Haronex",Font=new Font("Segoe UI",18,FontStyle.Bold),
        Location=new Point(20,18),Size=new Size(550,45)};
      status=new Label{Text="Ready. Starting review…",
        Location=new Point(20,80),Size=new Size(550,100),Font=new Font("Segoe UI",10)};
      Controls.Add(title);Controls.Add(status);
      Shown+=async (sender,args)=>await RunOnce();
      FormClosing+=(sender,args)=>{if(running){args.Cancel=true;}};
    }
    internal static string Safe(string text,int max) {
      return String.IsNullOrEmpty(text)?"":text.Substring(0,Math.Min(text.Length,max));
    }
    internal static string BuildPrompt(string text,int index,int total) {
      return "You are Haronex, an English radiology report proofreader. "
        +"PRIORITY ONE (LOGIC): compare findings against impression, laterality, positive/negative claims, "
        +"lesion counts, dates and measurements, and only identify conflicts explicitly supported by this text. "
        +"Do not claim a contradiction where there is merely uncertainty or different context. "
        +"PRIORITY TWO: verify medical terminology and abbreviations are internally consistent. "
        +"PRIORITY THREE: check every sentence for spelling, grammar and punctuation. "
        +"PRIORITY FOUR: re-scan each section for remaining independent issues. "
        +"Never invent a medical fact, diagnosis, or finding. All logic findings require clinical verification. "
        +"Return ONLY a single strict JSON object with schema "
        +"{\"issues\":[{\"quote\":\"exact short substring from report\","
        +"\"category\":\"logic or language\",\"explanation\":\"brief reason without identifiers\","
        +"\"suggestion\":\"optional proposed wording\"}]}. "
        +"Each quote must match an exact character sequence in the text. "
        +"Use category logic for explicit internal contradictions, language otherwise. "
        +"Deduplicate findings. Empty issues array ONLY if no well-supported findings after all checks. "
        +"Never repeat patient identifiers in explanations or suggestions. "
        +"PART "+(index+1)+" OF "+total+":\n"+text;
    }
    internal static bool Approved() {
      using(var key=Registry.CurrentUser.OpenSubKey(@"Software\Haronex")) {
        return key!=null&&Convert.ToString(key.GetValue("WebTransferAllowed",0))=="1" && File.Exists(GroqKeyPath());
      }
    }
    internal static string GroqKeyPath() {
      return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "RadiologyAgent","groq.key.dpapi");
    }
    internal static string ReadGroqKey() {
      string path=GroqKeyPath();
      if(!File.Exists(path))throw new Exception("Groq API key is missing. Reinstall Haronex.");
      byte[] encrypted=Convert.FromBase64String(File.ReadAllText(path).Trim());
      byte[] clear=ProtectedData.Unprotect(encrypted,null,DataProtectionScope.CurrentUser);
      try {
        string key=Encoding.UTF8.GetString(clear);
        if(!key.StartsWith("gsk_",StringComparison.Ordinal)||key.Length<20)
          throw new Exception("Groq API key format is invalid.");
        return key;
      } finally {Array.Clear(clear,0,clear.Length);}
    }
    internal static string CachePath(string report,bool selection) {
      string identity=ReviewAlgorithmVersion+"|"+(selection?"selection":"document")+"|"+report;
      byte[] hash;
      using(var sha=SHA256.Create())hash=sha.ComputeHash(Encoding.UTF8.GetBytes(identity));
      string token=BitConverter.ToString(hash).Replace("-","").ToLowerInvariant();
      string folder=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Haronex","ReviewCache");
      return Path.Combine(folder,token+".dpapi");
    }
    internal static string ReadCached(string path) {
      try {
        if(!File.Exists(path))return null;
        if(DateTime.UtcNow-File.GetLastWriteTimeUtc(path)>TimeSpan.FromDays(30))return null;
        byte[] encrypted=File.ReadAllBytes(path);
        byte[] clear=ProtectedData.Unprotect(encrypted,null,DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(clear);
      } catch { return null; }
    }
    internal static void SaveCached(string path,string response) {
      try {
        string directory=Path.GetDirectoryName(path);
        Directory.CreateDirectory(directory);
        byte[] encrypted=ProtectedData.Protect(Encoding.UTF8.GetBytes(response),null,DataProtectionScope.CurrentUser);
        string temp=path+".tmp";
        File.WriteAllBytes(temp,encrypted);
        if(File.Exists(path))File.Delete(path);
        File.Move(temp,path);
        // Retain only the 30 newest encrypted results.
        var entries=new DirectoryInfo(directory).GetFiles("*.dpapi");
        Array.Sort(entries,(a,b)=>b.LastWriteTimeUtc.CompareTo(a.LastWriteTimeUtc));
        for(int i=30;i<entries.Length;i++)try{entries[i].Delete();}catch{}
      } catch { /* Cache write failures cannot block reviewing. */ }
    }
    internal static void SetReviewStatus(int code) {
      try {
        using(var key=Registry.CurrentUser.CreateSubKey(@"Software\Haronex"))
          key.SetValue("ReviewResult",code,RegistryValueKind.DWord);
      } catch { /* Never fail reviewing due to UI status. */ }
    }
    private async Task RunOnce() {
      if(running)return;
      running=true;
      SetReviewStatus(0);
      try {
        if(!Approved())throw new Exception("External web transfer is not authorized. Reinstall Haronex and approve Web Mode data transfer during setup.");
        dynamic app=word,doc=app.ActiveDocument;
        int selectionStart=(int)app.Selection.Start,selectionEnd=(int)app.Selection.End;
        bool hasSelection=selectionEnd>selectionStart;
        int rangeStart=hasSelection?selectionStart:(int)doc.Content.Start;
        int rangeEnd=hasSelection?selectionEnd:(int)doc.Content.End;
        string report=hasSelection?(string)app.Selection.Text:(string)doc.Content.Text;
        if(String.IsNullOrWhiteSpace(report))throw new Exception("No text to review.");
        if(report.Length>120000)throw new Exception("Report exceeds 120,000 characters. Select a smaller section.");
        var chunks=new List<string>();
        int pos=0;
        while(pos<report.Length) {
          int length=Math.Min(6000,report.Length-pos);
          if(pos+length<report.Length) {
            int boundary=report.LastIndexOfAny(new char[]{'\r','\n'},pos+length-1,length);
            if(boundary>=pos+4500)length=boundary-pos+1;
          }
          chunks.Add(report.Substring(pos,length));
          pos+=length;
        }
        var prompts=new List<string>();
        for(int i=0;i<chunks.Count;i++)prompts.Add(BuildPrompt(chunks[i],i,chunks.Count));
        var serializer=new JavaScriptSerializer{MaxJsonLength=2000000};
        var payload=serializer.Serialize(new Dictionary<string,object>{{"prompts",prompts},{"texts",chunks},{"review_version",ReviewAlgorithmVersion},{"api_key",ReadGroqKey()}});
        string cacheFile=CachePath(report,hasSelection);
        string workerResult=ReadCached(cacheFile);
        bool cached=workerResult!=null;
        if(!cached) {
          status.Text="Groq Qwen API reviewing "+report.Length+" characters in "+chunks.Count+" part(s).";
          workerResult=await Task.Run(()=>CallBrowserWorker(payload,chunks.Count));
        } else {
          status.Text="Reusing encrypted previous review for unchanged text…";
        }
        var root=serializer.DeserializeObject(workerResult) as Dictionary<string,object>;
        if(root==null||!root.ContainsKey("ok")||!Convert.ToBoolean(root["ok"])) {
          string error=root!=null&&root.ContainsKey("error")?Convert.ToString(root["error"]):"No browser response.";
          throw new Exception("Web review did not complete: "+Safe(error,330));
        }
        // Cache only completed, validated responses (never error pages or partial replies).
        if(!root.ContainsKey("mode")||Convert.ToString(root["mode"])!="groq"||
           !root.ContainsKey("review_version")||Convert.ToString(root["review_version"])!=ReviewAlgorithmVersion)
          throw new Exception("Unexpected non-guest browser result; no Word changes applied.");
        if(!root.ContainsKey("passes")||Convert.ToInt32(root["passes"])!=1)
          throw new Exception("Incomplete review stages; no Word changes applied.");
        var groups=root["groups"] as object[];
        if(groups==null||groups.Length!=chunks.Count)throw new Exception("Incomplete web response: no Word changes applied.");
        var valid=new List<Tuple<string,string,string,string>>();
        int received=0,rejected=0;
        for(int i=0;i<groups.Length;i++) {
          var rows=groups[i] as object[];
          if(rows==null)throw new Exception("Invalid response group: no Word changes applied.");
          foreach(var entry in rows) {
            received++;
            var row=entry as Dictionary<string,object>;
            if(row==null){rejected++;continue;}
            Func<string,string> field=k=>row.ContainsKey(k)&&row[k]!=null?Convert.ToString(row[k]):"";
            var quote=field("quote");
            var category=field("category").Trim().ToLowerInvariant();
            if(quote.Length<2||quote.Length>500||!chunks[i].Contains(quote)||
               (category!="language"&&category!="logic")) {rejected++;continue;}
            valid.Add(Tuple.Create(quote,category,field("explanation"),field("suggestion")));
          }
        }
        if(!cached)SaveCached(cacheFile,workerResult);
        // Apply only after all chunks have been checked. Word's COM objects are used on the UI STA thread.
        status.Text="Adding "+valid.Count+" Haronex Word annotations…";
        int applied=0,skipped=0,duplicates=0;
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
        } catch { /* Duplicates can still be checked against current-session changes. */ }
        foreach(var issue in valid) {
          try {
            dynamic region=doc.Range(rangeStart,rangeEnd);
            dynamic find=region.Find;
            find.ClearFormatting();find.Text=issue.Item1;
            find.Forward=true;find.Wrap=0;find.MatchCase=true;find.MatchWildcards=false;
            if(!(bool)find.Execute()){skipped++;continue;}
            dynamic again=doc.Range((int)region.End,rangeEnd);
            dynamic second=again.Find;
            second.ClearFormatting();second.Text=issue.Item1;
            second.Forward=true;second.Wrap=0;second.MatchCase=true;second.MatchWildcards=false;
            if((bool)second.Execute()){skipped++;continue;}
            string anchorId=((int)region.Start).ToString()+":"+((int)region.End).ToString();
            if(existing.Contains(anchorId)){duplicates++;continue;}
            string comment="[Haronex: "+(issue.Item2=="logic"?"Potential internal contradiction":"Language")+"]\r\n"+
              Safe(issue.Item3,500)+(issue.Item4.Length>0?"\r\nSuggestion: "+Safe(issue.Item4,250):"")+
              "\r\nRequires clinical verification.";
            region.HighlightColorIndex=issue.Item2=="logic"?6:7;
            doc.Comments.Add(region,comment);
            existing.Add(anchorId);
            applied++;
          } catch {skipped++;}
        }
        SetReviewStatus(1);
        MessageBox.Show("تمت المراجعة", "Haronex", MessageBoxButtons.OK, MessageBoxIcon.Information);
      }
      catch(Exception ex) {
        SetReviewStatus(2);
        // Do not expose service or report details in an end-user popup.
        MessageBox.Show("رجاءً أعد المحاولة", "Haronex", MessageBoxButtons.OK, MessageBoxIcon.Error);
      }
      finally {running=false; BeginInvoke((Action)(()=>Close()));}
    }
    private static void KillWorkerTree(Process process) {
      try {
        using(var kill=Process.Start(new ProcessStartInfo("taskkill.exe",
            "/PID "+process.Id+" /T /F") {
          CreateNoWindow=true,UseShellExecute=false,
          WindowStyle=ProcessWindowStyle.Hidden
        })) {
          if(kill!=null)kill.WaitForExit(3000);
        }
      } catch { try {process.Kill();} catch {} }
    }
    internal static string CallBrowserWorker(string payload,int count) {
      return CallBrowserWorker(payload,count,CancellationToken.None);
    }
    internal static string CallBrowserWorker(string payload,int count,CancellationToken cancellation) {
      string executable=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"HaronexGroqWorker.exe");
      if(!File.Exists(executable))throw new Exception("Web worker missing. Reinstall Haronex.");
      var info=new ProcessStartInfo(executable){
        UseShellExecute=false,CreateNoWindow=true,WindowStyle=ProcessWindowStyle.Hidden,
        RedirectStandardInput=true,RedirectStandardOutput=true,
        StandardOutputEncoding=new UTF8Encoding(false)
      };
      using(var process=Process.Start(info)) {
        if(process==null)throw new Exception("Unable to start Haronex browser worker.");
        var bytes=new UTF8Encoding(false).GetBytes(payload);
        process.StandardInput.BaseStream.Write(bytes,0,bytes.Length);
        process.StandardInput.BaseStream.Flush();
        process.StandardInput.Close();
        // Read output asynchronously so an unresponsive browser cannot block timeout handling.
        var outputTask=process.StandardOutput.ReadToEndAsync();
        var watch=Stopwatch.StartNew();
        // Four transport attempts (45 seconds each) plus 2/5/10 second backoff.
        // Longer documents may have multiple chunks. Keep a finite upper bound.
        int maxMs=(int)Math.Min(1800000L,Math.Max(215000L,(long)count*215000L));
        while(!process.WaitForExit(250)) {
          if(cancellation.IsCancellationRequested) {
            KillWorkerTree(process);
            cancellation.ThrowIfCancellationRequested();
          }
          if(watch.ElapsedMilliseconds>maxMs) {
            KillWorkerTree(process);
            throw new Exception("Browser review timed out.");
          }
        }
        cancellation.ThrowIfCancellationRequested();
        string output=outputTask.GetAwaiter().GetResult();
        if(String.IsNullOrWhiteSpace(output))throw new Exception("Browser worker returned no result.");
        return output.Trim();
      }
    }
  }
}
