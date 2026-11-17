using System;
using System.Drawing;
using System.Windows.Forms;

namespace RadiologyAgent {
  // Word 2013 does not support buttons inside native comment balloons.
  // This non-modal popup appears near the selected Haronex comment.
  internal sealed class SuggestionPopup : Form {
    private readonly RichTextBox original;
    private readonly RichTextBox suggestion;
    private readonly CheckBox physicianConfirmation;
    private readonly Button apply;
    private readonly bool clinical;
    private readonly Func<string,bool> onApply;

    internal SuggestionPopup(string category,string quote,string advice,
                             Func<string,bool> applyAction) {
      clinical=String.Equals(category,"logic",StringComparison.OrdinalIgnoreCase)
            ||String.Equals(category,"unrelated",StringComparison.OrdinalIgnoreCase);
      onApply=applyAction;
      Text="Haronex - Apply Suggestion";
      Font=new Font("Segoe UI",9.5f);
      Size=new Size(435,394);
      MinimumSize=new Size(395,360);
      FormBorderStyle=FormBorderStyle.SizableToolWindow;
      StartPosition=FormStartPosition.Manual;
      ShowInTaskbar=false;
      TopMost=true;
      var head=new Label{Text="Haronex - "+category.ToUpperInvariant(),
                         Font=new Font("Segoe UI",11,FontStyle.Bold),
                         Location=new Point(14,12),Size=new Size(395,24)};
      var originalTitle=new Label{Text="Original text",Location=new Point(14,44),Size=new Size(380,20)};
      original=new RichTextBox{Location=new Point(14,67),Size=new Size(393,69),
                               ReadOnly=true,Text=quote,ScrollBars=RichTextBoxScrollBars.Vertical,
                               BackColor=Color.FromArgb(250,247,247)};
      var suggestedTitle=new Label{Text="Suggestion (bold) - edit replacement if needed",
                                   Location=new Point(14,145),Size=new Size(400,20)};
      suggestion=new RichTextBox{Location=new Point(14,168),Size=new Size(393,87),
                                 Font=new Font("Segoe UI",10,FontStyle.Bold),
                                 ScrollBars=RichTextBoxScrollBars.Vertical};
      // Explanatory instructions are never inserted into clinical reports.
      // The user must explicitly type a replacement for those comments.
      bool literal=IsLiteralReplacement(advice);
      suggestion.Text=literal?advice:"";
      if(!literal) {
        suggestedTitle.Text="Enter corrected report text (advice is not replacement)";
        suggestion.BackColor=Color.FromArgb(255,253,236);
      }
      physicianConfirmation=new CheckBox {
        Text="I am the reporting physician / authorized clinician and approve this clinical change",
        Location=new Point(14,263),Size=new Size(402,39),Visible=clinical,
        AutoSize=false
      };
      apply=new Button{Text="Apply Suggestion",Location=new Point(14,308),
                       Size=new Size(210,33),BackColor=Color.FromArgb(22,132,64),
                       ForeColor=Color.White,FlatStyle=FlatStyle.Flat};
      var cancel=new Button{Text="Cancel",Location=new Point(236,308),Size=new Size(171,33)};
      Controls.Add(head);Controls.Add(originalTitle);Controls.Add(original);
      Controls.Add(suggestedTitle);Controls.Add(suggestion);
      Controls.Add(physicianConfirmation);Controls.Add(apply);Controls.Add(cancel);
      cancel.Click+=(s,e)=>Close();
      apply.Click+=(s,e)=> {
        string replacement=suggestion.Text;
        if(String.IsNullOrWhiteSpace(replacement)||
           replacement==original.Text) {
          MessageBox.Show(this,"Enter a different, complete replacement text.",
                          "Haronex",MessageBoxButtons.OK,MessageBoxIcon.Information);
          return;
        }
        if(clinical&&!physicianConfirmation.Checked) {
          MessageBox.Show(this,"Clinical changes require explicit physician confirmation.",
                          "Haronex",MessageBoxButtons.OK,MessageBoxIcon.Warning);
          return;
        }
        if(replacement.Length>2000||replacement.IndexOf('\0')>=0) {
          MessageBox.Show(this,"Invalid replacement length or control characters.",
                          "Haronex",MessageBoxButtons.OK,MessageBoxIcon.Warning);
          return;
        }
        if(onApply(replacement))Close();
        else MessageBox.Show(this,
          "The original text, comment, or Word document has changed. Review it again; nothing was replaced.",
          "Haronex",MessageBoxButtons.OK,MessageBoxIcon.Warning);
      };
      suggestion.TextChanged+=(s,e)=>UpdateButton();
      physicianConfirmation.CheckedChanged+=(s,e)=>UpdateButton();
      Resize+=(s,e)=>{
        int width=ClientSize.Width-28;
        original.Width=width;suggestion.Width=width;
        apply.Width=Math.Max(140,(width-12)/2);
        cancel.Left=apply.Right+12;
        cancel.Width=Math.Max(100,width-apply.Width-12);
      };
      UpdateButton();
    }

    // A suggestion is prefilled only if it looks like clean report wording
    // rather than AI advice, a discussion, or an uncertain clinical instruction.
    internal static bool IsLiteralReplacement(string advice) {
      if(String.IsNullOrWhiteSpace(advice)||advice.Length>500)return false;
      string trimmed=advice.Trim();
      string low=trimmed.ToLowerInvariant();
      if(low.StartsWith("suggestion:"))return false;
      string[] instructions={"consider ","verify ","confirm ","review ","ask ",
        "check ","clarify ","correlate ","ensure ","possibly ","perhaps ",
        "if ","may ","could ","should ","requires ","either ","add supporting",
        "reconcile ","depending ","it is "};
      foreach(string term in instructions)if(low.StartsWith(term))return false;
      if(trimmed.IndexOf('\r')>=0||trimmed.IndexOf('\n')>=0)return false;
      return true;
    }

    private void UpdateButton() {
      apply.Enabled=!String.IsNullOrWhiteSpace(suggestion.Text)
          &&!String.Equals(suggestion.Text,original.Text,StringComparison.Ordinal)
          &&(!clinical||physicianConfirmation.Checked);
    }

    internal void PositionNearCursor() {
      var screen=Screen.FromPoint(Cursor.Position).WorkingArea;
      int x=Cursor.Position.X+20,y=Cursor.Position.Y+16;
      if(x+Width>screen.Right)x=Math.Max(screen.Left,Cursor.Position.X-Width-16);
      if(y+Height>screen.Bottom)y=Math.Max(screen.Top,screen.Bottom-Height);
      Location=new Point(x,y);
    }
  }
}
