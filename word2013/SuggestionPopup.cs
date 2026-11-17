using System;
using System.Drawing;
using System.Windows.Forms;

namespace RadiologyAgent {
  // Modern, keyboard-accessible nonmodal popup for Word 2013. It displays
  // the precise replacement scope, never merely the smaller AI quote.
  internal sealed class SuggestionPopup : Form {
    private readonly RichTextBox original;
    private readonly RichTextBox suggestion;
    private readonly Button apply;
    private readonly Func<string,bool> onApply;
    private readonly string initialOriginal;
    private readonly Label state;

    internal SuggestionPopup(string category,string quote,string proposed,
                             int relatedCount,Func<string,bool> applyAction) {
      onApply=applyAction;
      initialOriginal=quote??"";
      Text="Haronex | Suggested Correction";
      Font=new Font("Segoe UI",10f);
      BackColor=Color.FromArgb(245,247,251);
      ForeColor=Color.FromArgb(31,42,59);
      ClientSize=new Size(564,444);
      MinimumSize=new Size(500,420);
      FormBorderStyle=FormBorderStyle.Sizable;
      StartPosition=FormStartPosition.Manual;
      ShowInTaskbar=false;
      TopMost=true;
      MaximizeBox=false;
      AutoScaleMode=AutoScaleMode.Dpi;

      var header=new Panel{Dock=DockStyle.Top,Height=78,
         BackColor=Color.FromArgb(24,54,103)};
      var title=new Label{Text="HARONEX  /  SUGGESTED CORRECTION",
         Dock=DockStyle.Top,Height=35,Padding=new Padding(22,12,0,0),
         ForeColor=Color.White,Font=new Font("Segoe UI",11.5f,FontStyle.Bold)};
      var sub=new Label{Text=relatedCount>0?
         ("One change · "+relatedCount+" linked note(s) will be checked and resolved"):
         "One change · independent comments remain open",
         Dock=DockStyle.Top,Height=30,Padding=new Padding(22,1,0,0),
         ForeColor=Color.FromArgb(208,224,249),Font=new Font("Segoe UI",9.2f)};
      header.Controls.Add(sub);header.Controls.Add(title);
      Controls.Add(header);

      var footer=new Panel{Dock=DockStyle.Bottom,Height=75,
          BackColor=Color.White,Padding=new Padding(20,13,20,15)};
      apply=new Button{Text="✓  Apply Suggestion",Dock=DockStyle.Left,
         Width=285,FlatStyle=FlatStyle.Flat,Cursor=Cursors.Hand,
         ForeColor=Color.White,BackColor=Color.FromArgb(23,132,81),
         Font=new Font("Segoe UI",10,FontStyle.Bold)};
      apply.FlatAppearance.BorderSize=0;
      var cancel=new Button{Text="Cancel",Dock=DockStyle.Right,
         Width=112,FlatStyle=FlatStyle.Flat,Cursor=Cursors.Hand,
         BackColor=Color.FromArgb(239,242,247)};
      cancel.FlatAppearance.BorderColor=Color.FromArgb(211,220,230);
      footer.Controls.Add(apply);footer.Controls.Add(cancel);
      Controls.Add(footer);

      var body=new Panel{Dock=DockStyle.Fill,Padding=new Padding(20,11,20,8)};
      Controls.Add(body);
      body.BringToFront();
      var badge=new Label{Text=category.ToUpperInvariant()+"   •   ONE-CLICK APPROVAL",
         Dock=DockStyle.Top,Height=29,ForeColor=Color.FromArgb(46,93,153),
         Font=new Font("Segoe UI",9,FontStyle.Bold)};
      var originalLabel=new Label{Text="CURRENT REPORT TEXT  ·  replacement scope",
         Dock=DockStyle.Top,Height=28,
         Font=new Font("Segoe UI",9,FontStyle.Bold)};
      original=new RichTextBox{Dock=DockStyle.Top,Height=87,ReadOnly=true,
         BorderStyle=BorderStyle.FixedSingle,BackColor=Color.FromArgb(251,240,239),
         ForeColor=Color.FromArgb(84,40,43),
         Font=new Font("Segoe UI",10.2f),ScrollBars=RichTextBoxScrollBars.Vertical,
         Text=initialOriginal};
      var suggestedLabel=new Label{Text="SUGGESTED REPLACEMENT  ·  editable",
         Dock=DockStyle.Top,Height=35,Padding=new Padding(0,12,0,0),
         Font=new Font("Segoe UI",9,FontStyle.Bold)};
      suggestion=new RichTextBox{Dock=DockStyle.Fill,ReadOnly=false,
         BorderStyle=BorderStyle.FixedSingle,BackColor=Color.FromArgb(239,248,246),
         ForeColor=Color.FromArgb(20,91,73),
         Font=new Font("Segoe UI",10.2f,FontStyle.Bold),
         ScrollBars=RichTextBoxScrollBars.Vertical};
      bool literal=IsLiteralReplacement(proposed);
      suggestion.Text=literal?proposed.Trim():"";
      state=new Label{Dock=DockStyle.Bottom,Height=39,
         ForeColor=Color.FromArgb(104,116,132),Font=new Font("Segoe UI",8.5f),
         Text=literal?"Only the text shown above will be replaced.":
          "No safe literal suggestion. Type the exact replacement to enable Apply."};
      body.Controls.Add(suggestion);
      body.Controls.Add(suggestedLabel);
      body.Controls.Add(original);
      body.Controls.Add(originalLabel);
      body.Controls.Add(badge);
      body.Controls.Add(state);
      state.BringToFront();

      cancel.Click+=(sender,args)=>Close();
      apply.Click+=(sender,args)=>{
        string replacement=suggestion.Text.Trim();
        if(String.IsNullOrWhiteSpace(replacement)||replacement==initialOriginal)return;
        if(replacement.Length>2000||replacement.IndexOf('\0')>=0||
           replacement.IndexOf('\r')>=0||replacement.IndexOf('\n')>=0) {
          MessageBox.Show(this,"Use one complete replacement without paragraph breaks.",
            "Haronex",MessageBoxButtons.OK,MessageBoxIcon.Warning);return;
        }
        apply.Enabled=false;
        try {
          if(onApply(replacement))Close();
          else MessageBox.Show(this,
            "Word text or comment changed. Nothing was safely replaced; review again.",
            "Haronex",MessageBoxButtons.OK,MessageBoxIcon.Warning);
        } finally {if(!IsDisposed)UpdateButton();}
      };
      suggestion.TextChanged+=(sender,args)=>UpdateButton();
      AcceptButton=apply;
      CancelButton=cancel;
      UpdateButton();
    }

    internal static bool IsLiteralReplacement(string advice) {
      if(String.IsNullOrWhiteSpace(advice)||advice.Length>500)return false;
      string low=advice.Trim().ToLowerInvariant();
      if(low.StartsWith("suggestion:"))return false;
      string[] instructions={"consider ","verify ","confirm ","review ","ask ",
        "check ","clarify ","correlate ","ensure ","possibly ","perhaps ",
        "if ","may ","could ","should ","requires ","either ","add supporting",
        "reconcile ","depending ","it is ","suggest ","recommend ","the intended "};
      foreach(string term in instructions)if(low.StartsWith(term))return false;
      if(advice.IndexOf('\r')>=0||advice.IndexOf('\n')>=0||
         advice.IndexOf('\x07')>=0)return false;
      return true;
    }

    private void UpdateButton() {
      apply.Enabled=!String.IsNullOrWhiteSpace(suggestion.Text)&&
        !String.Equals(suggestion.Text.Trim(),initialOriginal.Trim(),
                      StringComparison.Ordinal);
    }
    internal void PositionNearCursor() {
      var screen=Screen.FromPoint(Cursor.Position).WorkingArea;
      int x=Cursor.Position.X+20,y=Cursor.Position.Y+16;
      if(x+Width>screen.Right)x=Math.Max(screen.Left,Cursor.Position.X-Width-12);
      if(y+Height>screen.Bottom)y=Math.Max(screen.Top,screen.Bottom-Height);
      Location=new Point(x,y);
    }
  }
}
