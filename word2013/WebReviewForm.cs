using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using Microsoft.Win32;

namespace RadiologyAgent {
  internal sealed class WebReviewForm : Form {
    private readonly object word;
    private readonly Label status;
    private bool running;
    private const string ReviewAlgorithmVersion="haronex-web-review-v3.2-cache1";
    public WebReviewForm(object application) {
      word=application;
      Text="Haronex — Word 2013"; Width=610; Height=250;
      StartPosition=FormStartPosition.CenterScreen;
      FormBorderStyle=FormBorderStyle.FixedDialog;
      MaximizeBox=false;MinimizeBox=false;
      var title=new Label{Text="Haronex",Font=new Font("Segoe UI",18,FontStyle.Bold),
        Location=new Point(20,18),Size=new Size(550,45)};
      status=new Label{Text="Ready. Starting review…",
        Location=new Point(20,80),Size=new Size(550,100),Font=new Font("Segoe UI",10)};
      Controls.Add(title);Controls.Add(status);
      Shown+=async (sender,args)=>await RunOnce();
      FormClosing+=(sender,args)=>{if(running){args.Cancel=true;}};
    }
    private static string Safe(string text,int max) {
      return String.IsNullOrEmpty(text)?"":text.Substring(0,Math.Min(text.Length,max));
    }
    private static string BuildPrompt(string text,int index,int total) {
      return "You are Haronex, a cautious English radiology report proofreader. "+
        "Identify spelling, grammar, terminology consistency mistakes and ONLY clear internal contradictions. "+
        "Never invent medical findings, diagnoses, treatment, or measurements. "+
        "Return ONE strict JSON object, no Markdown, no introduction. "+
        "Schema: {\"issues\":[{\"quote\":\"exact substring from text\",\"category\":\"language or logic\",\"explanation\":\"brief explanation, no patient identifiers\",\"suggestion\":\"optional proposed wording\"}]}. "+
        "Use only categories language and logic. Keep each quote exactly as in the report. "+
        "Do not repeat patient identifiers. If no reliable issues return {\"issues\":[]}. "+
        "Clinical verification is mandatory. Part "+(index+1)+"/"+total+":\n"+text;
    }
    private static bool Approved() {
      using(var key=Registry.CurrentUser.OpenSubKey(@"Software\Haronex")) {
        return key!=null&&Convert.ToString(key.GetValue("WebTransferAllowed",0))=="1";
      }
    }
    private static string CachePath(string report,bool selection) {
      string identity=ReviewAlgorithmVersion+"|"+(selection?"selection":"document")+"|"+report;
      byte[] hash;
      using(var sha=SHA256.Create())hash=sha.ComputeHash(Encoding.UTF8.GetBytes(identity));
      string token=BitConverter.ToString(hash).Replace("-","").ToLowerInvariant();
      string folder=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Haronex","ReviewCache");
      return Path.Combine(folder,token+".dpapi");
    }
    private static string ReadCached(string path) {
      try {
        if(!File.Exists(path))return null;
        if(DateTime.UtcNow-File.GetLastWriteTimeUtc(path)>TimeSpan.FromDays(30))return null;
        byte[] encrypted=File.ReadAllBytes(path);
        byte[] clear=ProtectedData.Unprotect(encrypted,null,DataProtectionScope.CurrentUser);
        return Encoding.UTF8.GetString(clear);
      } catch { return null; }
    }
    private static void SaveCached(string path,string response) {
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
    private async Task RunOnce() {
      if(running)return;
      running=true;
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
          int length=Math.Min(12000,report.Length-pos);
          if(pos+length<report.Length) {
            int boundary=report.LastIndexOf('\n',pos+length-1,length);
            if(boundary>=pos+6000)length=boundary-pos+1;
          }
          chunks.Add(report.Substring(pos,length));
          pos+=length;
        }
        var prompts=new List<string>();
        for(int i=0;i<chunks.Count;i++)prompts.Add(BuildPrompt(chunks[i],i,chunks.Count));
        var serializer=new JavaScriptSerializer{MaxJsonLength=2000000};
        var payload=serializer.Serialize(new Dictionary<string,object>{{"prompts",prompts}});
        string cacheFile=CachePath(report,hasSelection);
        string workerResult=ReadCached(cacheFile);
        bool cached=workerResult!=null;
        if(!cached) {
          status.Text="Reviewing "+report.Length+" characters in "+chunks.Count+
            " part(s) with Haronex Web. You may need to sign in once in Edge.";
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
        int applied=0,skipped=0;
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
            string comment="[Haronex: "+(issue.Item2=="logic"?"Potential internal contradiction":"Language")+"]\r\n"+
              Safe(issue.Item3,500)+(issue.Item4.Length>0?"\r\nSuggestion: "+Safe(issue.Item4,250):"")+
              "\r\nRequires clinical verification.";
            region.HighlightColorIndex=issue.Item2=="logic"?6:7;
            doc.Comments.Add(region,comment);
            applied++;
          } catch {skipped++;}
        }
        status.Text="Haronex completed"+(cached?" (same saved review)":"")+
          ". "+report.Length+" chars, "+received+" returned, "+applied+
          " applied, "+rejected+" rejected, "+skipped+" skipped."+
          (received==0?" No issues returned; NOT clinical clearance.":"");
      }
      catch(Exception ex) {status.Text="Haronex: "+Safe(ex.Message,420);}
      finally {running=false;}
    }
    private static string CallBrowserWorker(string payload,int count) {
      string executable=Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"HaronexWebWorker.exe");
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
        var output=process.StandardOutput.ReadToEnd();
        if(!process.WaitForExit(Math.Max(240000,count*210000))) {
          try{process.Kill();}catch{}
          throw new Exception("Browser review timed out.");
        }
        if(String.IsNullOrWhiteSpace(output))throw new Exception("Browser worker returned no result.");
        return output.Trim();
      }
    }
  }
}
