using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace RadiologyAgent {
  public enum ext_ConnectMode { ext_cm_AfterStartup=0, ext_cm_Startup=1, ext_cm_External=2, ext_cm_CommandLine=3, ext_cm_Solution=4, ext_cm_UISetup=5 }
  public enum ext_DisconnectMode { ext_dm_HostShutdown=0, ext_dm_UserClosed=1, ext_dm_UISetupComplete=2, ext_dm_SolutionClosed=3 }

  [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
  public interface IDTExtensibility2 {
    [DispId(1)] void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object application, [In] ext_ConnectMode mode, [In, MarshalAs(UnmanagedType.IDispatch)] object addin, [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(2)] void OnDisconnection([In] ext_DisconnectMode mode, [In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(3)] void OnAddInsUpdate([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(4)] void OnStartupComplete([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
    [DispId(5)] void OnBeginShutdown([In, Out, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType=VarEnum.VT_VARIANT)] ref Array custom);
  }
  [ComImport, Guid("000C0396-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
  public interface IRibbonExtensibility { [return: MarshalAs(UnmanagedType.BStr)] string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string ribbonID); }
  [ComImport, Guid("000C0395-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
  public interface IRibbonControl { string Id { get; } object Context { get; } string Tag { get; } }

  [ComVisible(true), Guid("D1FBEFD5-3B43-4516-9B7B-CDEB47AB0FA1"), ProgId("RadiologyAgent.WordAddin"), ClassInterface(ClassInterfaceType.None)]
  public class WordAddin : IDTExtensibility2, IRibbonExtensibility {
    private object application;
    public void OnConnection(object app, ext_ConnectMode mode, object addin, ref Array custom) { application=app; }
    public void OnDisconnection(ext_DisconnectMode mode, ref Array custom) { application=null; }
    public void OnAddInsUpdate(ref Array custom) {}
    public void OnStartupComplete(ref Array custom) {}
    public void OnBeginShutdown(ref Array custom) { application=null; }

    public string GetCustomUI(string ribbonID) {
      return "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
      "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
      "<ribbon><tabs><tab id=\"RadiologyTab\" label=\"Radiology Agent\">" +
      "<group id=\"RadiologyGroup\" label=\"Report Review\">" +
      "<button id=\"ReviewText\" label=\"Review Text\" size=\"large\" imageMso=\"ReviewAcceptChange\" onAction=\"ReviewTextClick\"/>" +
      "</group></tab></tabs></ribbon></customUI>";
    }
    public void ReviewTextClick(IRibbonControl control) {
      try {
        if (application==null) { MessageBox.Show("Word is not connected."); return; }
        using(var form=new ReviewForm(application)) { form.ShowDialog(); }
      } catch(Exception ex) { MessageBox.Show("Radiology Agent: "+ex.Message, "Radiology Agent", MessageBoxButtons.OK, MessageBoxIcon.Error); }
    }
  }

  internal sealed class Issue {
    public string Quote, Category, Explanation, Suggestion;
  }

  internal static class Gemini {
    public static List<Issue> Review(string report) {
      string key=ReadKey();
      var prompt="You are a cautious English radiology-report proofreader. Identify grammar errors and clear INTERNAL contradictions only. Do not invent medical findings. Do not disclose patient identifiers in explanations. Return at most 25 issues in strict JSON of the form {\"issues\":[{\"quote\":\"an exact phrase copied from the report\",\"category\":\"language or logic\",\"explanation\":\"brief explanation\",\"suggestion\":\"optional suggested wording\"}]}. Every quote must match the report exactly. Do not rewrite the report. Clinical verification remains essential.\nREPORT:\n"+report;
      var serializer=new JavaScriptSerializer { MaxJsonLength=300000 };
      var payload=new Dictionary<string,object> {
        {"contents",new object[]{new Dictionary<string,object>{{"parts",new object[]{new Dictionary<string,object>{{"text",prompt}}}}}}},
        {"generationConfig",new Dictionary<string,object>{{"responseMimeType","application/json"},{"temperature",0.1}}}
      };
      ServicePointManager.SecurityProtocol=(SecurityProtocolType)3072;
      var request=(HttpWebRequest)WebRequest.Create("https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent");
      request.Method="POST";request.ContentType="application/json; charset=utf-8";request.Timeout=85000;request.ReadWriteTimeout=85000;request.Headers["x-goog-api-key"]=key;
      byte[] bytes=Encoding.UTF8.GetBytes(serializer.Serialize(payload));
      string response;
      try {
        using(var stream=request.GetRequestStream()) stream.Write(bytes,0,bytes.Length);
        using(var result=(HttpWebResponse)request.GetResponse())
        using(var reader=new StreamReader(result.GetResponseStream())) response=reader.ReadToEnd();
      }
      catch(WebException ex) {
        var http=ex.Response as HttpWebResponse;
        if(http==null)throw new Exception("Gemini connection failed ("+ex.Status+"). Check TLS, proxy, firewall and internet access.");
        string status="HTTP "+(int)http.StatusCode;
        string googleStatus="";
        string reason="";
        try {
          using(var reader=new StreamReader(http.GetResponseStream())) {
            string body=reader.ReadToEnd();
            var envelope=serializer.DeserializeObject(body) as Dictionary<string,object>;
            if(envelope!=null&&envelope.ContainsKey("error")){
              var error=envelope["error"] as Dictionary<string,object>;
              if(error!=null&&error.ContainsKey("status"))googleStatus=Convert.ToString(error["status"]);
              if(error!=null&&error.ContainsKey("details")){
                var details=error["details"] as object[];
                if(details!=null)foreach(var entry in details){
                  var d=entry as Dictionary<string,object>;
                  if(d!=null&&d.ContainsKey("reason")){reason=Convert.ToString(d["reason"]);break;}
                }
              }
            }
          }
        } catch { /* Do not expose remote error bodies, patient data, or credentials. */ }
        string detail=(googleStatus.Length>0?" "+googleStatus:"")+(reason.Length>0?" ["+reason+"]":"");
        string advice;
        switch((int)http.StatusCode) {
          case 400: advice="Request or key rejected. Review API configuration.";break;
          case 401: advice="Authentication rejected. Rotate the exposed key and enter a new Google AI Studio key.";break;
          case 402: advice="Billing credit exhausted. Check your Google AI Studio billing.";break;
          case 403: advice="Project/key lacks permission or billing/region prerequisites.";break;
          case 404: advice="Model unavailable for this key/API version. Check available models.";break;
          case 429: advice="Quota or rate limit reached. Check your Gemini usage limits.";break;
          case 500: case 502: case 503: case 504: advice="Google service issue or timeout. Try again later.";break;
          default: advice="Request failed. Verify key, model and network access.";break;
        }
        throw new Exception("Gemini "+status+detail+". "+advice);
      }
      var root=serializer.DeserializeObject(response) as Dictionary<string,object>;
      var candidates=(object[])root["candidates"];
      var candidate=(Dictionary<string,object>)candidates[0];
      var content=(Dictionary<string,object>)candidate["content"];
      var parts=(object[])content["parts"];
      string json=Convert.ToString(((Dictionary<string,object>)parts[0])["text"]);
      var parsed=serializer.DeserializeObject(json) as Dictionary<string,object>;
      var list=new List<Issue>();
      if(parsed==null||!parsed.ContainsKey("issues"))return list;
      object[] issues=parsed["issues"] as object[];if(issues==null)return list;
      foreach(object item in issues) {
        if(list.Count>=25)break;
        var row=item as Dictionary<string,object>;if(row==null)continue;
        Func<string,string> get=name=>row.ContainsKey(name) ? Convert.ToString(row[name]) : "";
        string quote=get("quote"), category=get("category");
        if(quote.Length==0||quote.Length>300||!report.Contains(quote)||(category!="language"&&category!="logic"))continue;
        list.Add(new Issue{Quote=quote,Category=category,Explanation=get("explanation"),Suggestion=get("suggestion")});
      }
      return list;
    }
    private static string ReadKey() {
      string path=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"RadiologyAgent","key.dpapi");
      if(!File.Exists(path))throw new Exception("Gemini key is missing. Reinstall Radiology Agent and enter your key.");
      byte[] raw=Convert.FromBase64String(File.ReadAllText(path).Trim());
      string key=Encoding.UTF8.GetString(ProtectedData.Unprotect(raw,null,DataProtectionScope.CurrentUser));
      if(key.Length<12)throw new Exception("Invalid saved Gemini key.");
      return key;
    }
  }

  internal sealed class ReviewForm : Form {
    private readonly object application;
    private readonly Button review;
    private readonly Button test;
    private readonly Button changeKey;
    private readonly Label status;
    public ReviewForm(object app) {
      application=app;
      Text="Radiology Agent — Word 2013";Width=530;Height=340;
      StartPosition=FormStartPosition.CenterScreen;FormBorderStyle=FormBorderStyle.FixedDialog;MaximizeBox=false;MinimizeBox=false;
      var title=new Label{Text="Radiology Agent",Font=new Font("Segoe UI",17,FontStyle.Bold),Location=new Point(20,18),Size=new Size(470,42)};
      var sub=new Label{Text="Review selected text or entire document. AI suggestions require clinical verification.",Location=new Point(20,65),Size=new Size(470,45)};
      review=new Button{Text="Review By Agent",Location=new Point(20,120),Size=new Size(470,43),BackColor=Color.FromArgb(20,96,208),ForeColor=Color.White,FlatStyle=FlatStyle.Flat,Font=new Font("Segoe UI",10,FontStyle.Bold)};
      test=new Button{Text="Test Connection (synthetic text only)",Location=new Point(20,175),Size=new Size(300,35)};
      changeKey=new Button{Text="Change API Key",Location=new Point(330,175),Size=new Size(160,35)};
      status=new Label{Text="Ready. Do not send identifiable patient data without organizational approval.",Location=new Point(20,220),Size=new Size(470,70)};
      Controls.Add(title);Controls.Add(sub);Controls.Add(review);Controls.Add(test);Controls.Add(changeKey);Controls.Add(status);
      review.Click+=async (sender,args)=>await RunReview();
      test.Click+=async (sender,args)=>await RunTest();
      changeKey.Click+=(sender,args)=>ChangeKey();
    }
    private async Task RunTest() {
      test.Enabled=false;
      status.Text="Testing API with synthetic text only…";
      try {
        await Task.Run(()=>Gemini.Review("The lungs are clear. No pleural effusion is seen."));
        status.Text="Gemini connection succeeded. No patient report was sent.";
      } catch(Exception ex) {status.Text="Connection test: "+ex.Message;}
      finally {test.Enabled=true;}
    }
    private void ChangeKey() {
      using(var dialog=new Form()){
        dialog.Text="Replace Gemini API Key";dialog.Width=470;dialog.Height=190;
        dialog.StartPosition=FormStartPosition.CenterParent;
        dialog.FormBorderStyle=FormBorderStyle.FixedDialog;
        dialog.MaximizeBox=false;dialog.MinimizeBox=false;
        var instruction=new Label{Text="Enter a NEW Google AI Studio key. The old key will be replaced using Windows DPAPI.",Location=new Point(15,12),Size=new Size(430,38)};
        var field=new TextBox{Location=new Point(15,58),Width=425,UseSystemPasswordChar=true};
        var save=new Button{Text="Save new key",Location=new Point(280,95),Width=160,DialogResult=DialogResult.OK};
        dialog.Controls.Add(instruction);dialog.Controls.Add(field);dialog.Controls.Add(save);
        dialog.AcceptButton=save;
        if(dialog.ShowDialog(this)!=DialogResult.OK)return;
        string key=field.Text.Trim();
        if(key.Length<30||!(key.StartsWith("AQ.",StringComparison.Ordinal)||key.StartsWith("AIza",StringComparison.Ordinal))){
          status.Text="Key not saved: unrecognized key format.";return;
        }
        foreach(char character in key)
          if(!(Char.IsLetterOrDigit(character)||character=='_'||character=='-'||character=='.')){
            status.Text="Key not saved: invalid characters.";return;
          }
        string dir=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"RadiologyAgent");
        Directory.CreateDirectory(dir);
        string target=Path.Combine(dir,"key.dpapi");
        byte[] protectedKey=ProtectedData.Protect(Encoding.UTF8.GetBytes(key),null,DataProtectionScope.CurrentUser);
        File.WriteAllText(target,Convert.ToBase64String(protectedKey),Encoding.ASCII);
        field.Clear();key=null;
        status.Text="New key saved securely. Run Test Connection before reviewing documents.";
      }
    }
    private async Task RunReview() {
      review.Enabled=false;
      try {
        dynamic word=application;
        dynamic doc=word.ActiveDocument;
        string selected=(string)word.Selection.Text;
        string full=(string)doc.Content.Text;
        string report=String.IsNullOrWhiteSpace(selected)?full:selected;
        if(String.IsNullOrWhiteSpace(report))throw new Exception("The document is empty.");
        if(report.Length>18000)throw new Exception("Select a section of 18,000 characters or fewer.");
        status.Text="Reviewing with Gemini…";
        var issues=await Task.Run(()=>Gemini.Review(report));
        status.Text="Adding Word annotations…";
        int added=0,skipped=0;
        int selectedStart=(int)word.Selection.Start;
        int selectedEnd=(int)word.Selection.End;
        bool useSelection=!String.IsNullOrWhiteSpace(selected);
        foreach(var issue in issues) {
          dynamic search=doc.Range(useSelection?selectedStart:0,useSelection?selectedEnd:(int)doc.Content.End);
          dynamic find=search.Find;
          find.ClearFormatting();
          find.Text=issue.Quote;find.Forward=true;find.Wrap=0;find.MatchCase=true;
          bool found=(bool)find.Execute();
          if(!found){skipped++;continue;}
          // Skip ambiguous duplicate phrases rather than annotate the wrong occurrence.
          int firstEnd=(int)search.End;
          dynamic second=doc.Range(firstEnd,useSelection?selectedEnd:(int)doc.Content.End);
          dynamic secondFind=second.Find;
          secondFind.Text=issue.Quote;secondFind.Forward=true;secondFind.Wrap=0;secondFind.MatchCase=true;
          if((bool)secondFind.Execute()){skipped++;continue;}
          search.HighlightColorIndex=(issue.Category=="logic"?6:7); // Word 2013: red=6, yellow=7
          string comment="[AI suggestion: "+(issue.Category=="logic"?"Potential internal contradiction":"Language")+"]\r\n"+
             Truncate(issue.Explanation,550)+
             (String.IsNullOrWhiteSpace(issue.Suggestion)?"":"\r\nSuggestion: "+Truncate(issue.Suggestion,250))+
             "\r\nHuman clinical verification required.";
          doc.Comments.Add(search,comment);
          added++;
        }
        status.Text="Finished: "+added+" comments, "+skipped+" skipped.";
      } catch(Exception ex) { status.Text="Error: "+ex.Message; }
      finally { review.Enabled=true; }
    }
    private static string Truncate(string text,int n){return text==null?"":text.Substring(0,Math.Min(text.Length,n));}
  }
}
