using System;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace RadiologyAgent {
  // Web Mode deliberately does not automate login, bypass web limits, or call an API.
  // The user explicitly submits the prompt and copies the response in their own browser.
  internal sealed class WebReviewForm : Form {
    private readonly object word;
    private readonly Button action;
    private readonly Label status;
    private string documentText;
    private int rangeStart, rangeEnd;
    private bool selectedRange;
    private readonly List<string> chunks=new List<string>();
    private int chunkIndex;
    private int received, applied, skipped, rejected;
    private enum Phase { Initial, Import }
    private Phase phase=Phase.Initial;
    public WebReviewForm(object application) {
      word=application;
      Text="Haronex — Word 2013"; Width=610; Height=320;
      StartPosition=FormStartPosition.CenterScreen;
      FormBorderStyle=FormBorderStyle.FixedDialog; MaximizeBox=false; MinimizeBox=false;
      var title=new Label{Text="Haronex",Font=new Font("Segoe UI",18,FontStyle.Bold),
        Location=new Point(20,18),Size=new Size(550,45)};
      var help=new Label{Text="Web review: paste the prepared prompt in your browser, then copy the complete JSON answer.",
        Location=new Point(20,65),Size=new Size(550,50)};
      action=new Button{Text="Review By Agent",Location=new Point(20,124),Size=new Size(550,44),
        BackColor=Color.FromArgb(20,96,208),ForeColor=Color.White,
        FlatStyle=FlatStyle.Flat,Font=new Font("Segoe UI",10,FontStyle.Bold)};
      status=new Label{Text="Ready. Only send authorized, de-identified reports to the web service.",
        Location=new Point(20,185),Size=new Size(550,82)};
      Controls.Add(title);Controls.Add(help);Controls.Add(action);Controls.Add(status);
      action.Click+=(s,e)=>RunStep();
    }
    private void RunStep() {
      try {if(phase==Phase.Initial)Prepare();else Import();}
      catch(Exception ex) {status.Text="Haronex: "+ex.Message;}
    }
    private void Prepare() {
      dynamic app=word;
      dynamic doc=app.ActiveDocument;
      int a=(int)app.Selection.Start,b=(int)app.Selection.End;
      selectedRange=b>a;
      rangeStart=selectedRange?a:(int)doc.Content.Start;
      rangeEnd=selectedRange?b:(int)doc.Content.End;
      documentText=selectedRange?(string)app.Selection.Text:(string)doc.Content.Text;
      if(String.IsNullOrWhiteSpace(documentText))throw new Exception("Document/selection has no text.");
      if(documentText.Length>120000)throw new Exception("Review limited to 120,000 characters. Select a smaller part.");
      var consent=MessageBox.Show(this,
        "The prepared prompt may contain names, patient IDs and clinical data.\n"+
        "Only continue if you are authorized to send this text to the web service "+
        "and have removed personal identifiers. The full prompt will be copied to the clipboard.\n\nContinue?",
        "Haronex — Protect patient privacy",MessageBoxButtons.YesNo,MessageBoxIcon.Warning);
      if(consent!=DialogResult.Yes)return;
      chunks.Clear();chunkIndex=0;received=0;applied=0;skipped=0;rejected=0;
      int pos=0;
      while(pos<documentText.Length) {
        int length=Math.Min(12000,documentText.Length-pos);
        if(pos+length<documentText.Length) {
          int newline=documentText.LastIndexOf('\n',pos+length-1,length);
          if(newline>=pos+6000)length=newline-pos+1;
        }
        chunks.Add(documentText.Substring(pos,length));pos+=length;
      }
      if(!String.IsNullOrEmpty(documentText))SendCurrentPrompt();
      phase=Phase.Import;action.Text="Import Result From Clipboard";
      try {Process.Start(new ProcessStartInfo("https://gemini.google.com/app"){UseShellExecute=true});}
      catch {status.Text+=" Open https://gemini.google.com/app manually."; }
    }
    private string BuildPrompt(string text,int index,int count) {
      return "You are Haronex, a careful English radiology report proofreading assistant. "+
        "Check grammar, spelling, medical terminology consistency, and ONLY clear internal contradictions. "+
        "Never invent findings, measurements, or diagnoses. "+
        "Return a SINGLE raw JSON object, with NO markdown fences or introduction. "+
        "Schema: {\"issues\":[{\"quote\":\"an EXACT short substring of this chunk\",\"category\":\"language or logic\",\"explanation\":\"brief explanation without patient identifiers\",\"suggestion\":\"optional corrected wording\"}]}. "+
        "Use only categories language and logic. Keep quotes exact, including punctuation and spacing. "+
        "If there are no reliable issues, return {\"issues\":[]}. "+
        "Do not repeat patient identifiers in explanations or suggestions. "+
        "Do not rewrite the report. Clinical verification is mandatory. "+
        "CHUNK "+(index+1)+" OF "+count+"; TEXT:\n"+text;
    }
    private void SendCurrentPrompt() {
      Clipboard.SetText(BuildPrompt(chunks[chunkIndex],chunkIndex,chunks.Count));
      status.Text="Copied prompt "+(chunkIndex+1)+"/"+chunks.Count+
        " to clipboard. Paste and send in Gemini Web, copy the FULL JSON reply, "+
        "then click Import Result. Do not copy the report into this chat.";
    }
    private void Import() {
      if(chunkIndex>=chunks.Count)throw new Exception("No chunk is pending.");
      if(!Clipboard.ContainsText())throw new Exception("Clipboard contains no text. Copy the web response first.");
      string answer=Clipboard.GetText();
      // The prompt itself is not a response. Avoid processing it accidentally.
      if(answer.Contains("CHUNK "+(chunkIndex+1)+" OF "+chunks.Count+"; TEXT:"))
        throw new Exception("Clipboard still contains the prompt. Copy the web response first.");
      var parsed=ExtractObject(answer);
      var serializer=new JavaScriptSerializer{MaxJsonLength=500000};
      var root=serializer.DeserializeObject(parsed) as Dictionary<string,object>;
      if(root==null||!root.ContainsKey("issues"))throw new Exception("No issues array in copied response. Copy the full JSON object.");
      var rows=root["issues"] as object[];
      if(rows==null)throw new Exception("Invalid issues list. Copy a valid JSON response.");
      int localReceived=rows.Length,localRejected=0,localSkipped=0,localApplied=0;
      // Validate every row before applying any document changes.
      var valid=new List<Tuple<string,string,string,string>>();
      foreach(var entry in rows) {
        var row=entry as Dictionary<string,object>;
        if(row==null){localRejected++;continue;}
        Func<string,string> field=name=>row.ContainsKey(name)&&row[name]!=null?Convert.ToString(row[name]):"";
        var quote=field("quote");
        var category=field("category").Trim().ToLowerInvariant();
        if(quote.Length<2||quote.Length>500||(category!="language"&&category!="logic")||
           !chunks[chunkIndex].Contains(quote)){localRejected++;continue;}
        valid.Add(Tuple.Create(quote,category,field("explanation"),field("suggestion")));
      }
      dynamic doc=((dynamic)word).ActiveDocument;
      foreach(var issue in valid) {
        try {
          dynamic region=doc.Range(rangeStart,rangeEnd);
          dynamic find=region.Find;
          find.ClearFormatting();find.Text=issue.Item1;find.Forward=true;find.Wrap=0;
          find.MatchCase=true;find.MatchWildcards=false;
          if(!(bool)find.Execute()){localSkipped++;continue;}
          int end=(int)region.End;
          dynamic again=doc.Range(end,rangeEnd);
          dynamic nextFind=again.Find;
          nextFind.ClearFormatting();nextFind.Text=issue.Item1;nextFind.Forward=true;
          nextFind.Wrap=0;nextFind.MatchCase=true;nextFind.MatchWildcards=false;
          if((bool)nextFind.Execute()){localSkipped++;continue;}
          region.HighlightColorIndex=issue.Item2=="logic"?6:7;
          string comment="[Haronex — "+(issue.Item2=="logic"?"Potential internal contradiction":"Language")+"]\r\n"+
            Short(issue.Item3,500)+(issue.Item4.Length>0?"\r\nSuggestion: "+Short(issue.Item4,250):"")+
            "\r\nRequires clinical verification.";
          doc.Comments.Add(region,comment);localApplied++;
        } catch {localSkipped++; }
      }
      received+=localReceived;rejected+=localRejected;skipped+=localSkipped;applied+=localApplied;
      chunkIndex++;
      if(chunkIndex<chunks.Count) {
        SendCurrentPrompt();
        status.Text+=" Previous chunk: "+localReceived+" received, "+localApplied+" applied, "+
          localRejected+" rejected, "+localSkipped+" skipped.";
      } else {
        phase=Phase.Initial;action.Text="Review By Agent";
        status.Text="Haronex finished. Sent "+documentText.Length+" chars across "+chunks.Count+
          " chunks; received "+received+" issues, applied "+applied+", rejected "+rejected+
          ", skipped "+skipped+". "+(received==0?"No issues returned; NOT a clinical clearance.":"Clinical verification required.");
      }
    }
    private static string ExtractObject(string answer) {
      // Extract first balanced JSON object; handles quoted braces and Markdown fences.
      int start=answer.IndexOf('{');
      if(start<0)throw new Exception("Copied text has no JSON object.");
      bool quoted=false,escape=false;int depth=0;
      for(int i=start;i<answer.Length;i++){
        char c=answer[i];
        if(quoted) {
          if(escape)escape=false;
          else if(c=='\\')escape=true;
          else if(c=='"')quoted=false;
        } else {
          if(c=='"')quoted=true;
          else if(c=='{')depth++;
          else if(c=='}'&&--depth==0)return answer.Substring(start,i-start+1);
        }
      }
      throw new Exception("Incomplete JSON response. Copy the entire response.");
    }
    private static string Short(string text,int max) {return String.IsNullOrEmpty(text)?"":text.Substring(0,Math.Min(text.Length,max));}
  }
}
