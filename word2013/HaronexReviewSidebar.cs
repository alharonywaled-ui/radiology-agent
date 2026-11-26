using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace RadiologyAgent {
  // The Word 2013 add-in is a native COM shim; this .NET x86 companion offers
  // a fixed, Word-window-anchored sidebar without rewriting that COM shim.
  internal sealed class HaronexReviewSidebar : Form {
    internal sealed class Issue {
      internal int Start, End;
      internal string Quote, Body, Category, Why, Suggestion;
      internal bool Resolved;
      internal string Id {get{return Start+":"+End+"|"+Body;}}
    }
    private readonly FlowLayoutPanel cards;
    private readonly Label summary;
    private readonly Label status;
    private readonly Label demographics;
    private readonly Dictionary<string,string> issueFeedback=new Dictionary<string,string>(StringComparer.Ordinal);
    private readonly Dictionary<string,Button> filters=
      new Dictionary<string,Button>(StringComparer.Ordinal);
    private readonly Action<Issue,string> apply;
    private readonly Action<Issue> ignore;
    private readonly Action<Issue> navigate;
    private string filter="All";
    private List<Issue> issues=new List<Issue>();
    private string activeId;
    private bool rebuilding;
    internal event EventHandler Dismissed;

    private static readonly Color Navy=Color.FromArgb(20,63,130);
    private static readonly Color Ink=Color.FromArgb(36,48,64);
    private static readonly Color Muted=Color.FromArgb(104,119,139);
    private static readonly Color Border=Color.FromArgb(218,227,237);
    private static readonly Color Back=Color.FromArgb(244,247,251);
    private static readonly Color Green=Color.FromArgb(26,135,82);

    internal HaronexReviewSidebar(Action<Issue,string> onApply,
         Action<Issue> onIgnore,Action<Issue> onNavigate) {
      apply=onApply;ignore=onIgnore;navigate=onNavigate;
      Text="Haronex Review Panel";
      Font=new Font("Segoe UI",9f);
      BackColor=Back;
      ForeColor=Ink;
      FormBorderStyle=FormBorderStyle.None;
      StartPosition=FormStartPosition.Manual;
      ShowInTaskbar=false;
      TopMost=true;
      Width=446;Height=730;
      MinimumSize=new Size(375,350);
      DoubleBuffered=true;

      var header=new Panel{Dock=DockStyle.Top,Height=49,BackColor=Navy};
      var brand=new Label{Text="✦  Haronex Review Panel",ForeColor=Color.White,
        Font=new Font("Segoe UI",12,FontStyle.Bold),Dock=DockStyle.Fill,
        Padding=new Padding(15,12,0,0)};
      var close=new Button{Text="×",Dock=DockStyle.Right,Width=44,
        Font=new Font("Segoe UI",17),ForeColor=Color.White,BackColor=Navy,
        FlatStyle=FlatStyle.Flat,Cursor=Cursors.Hand};
      close.FlatAppearance.BorderSize=0;
      close.Click+=(s,e)=>{
        Hide();
        var dismissed=Dismissed;
        if(dismissed!=null)dismissed(this,EventArgs.Empty);
      };
      header.Controls.Add(brand);header.Controls.Add(close);
      Controls.Add(header);

      var top=new Panel{Dock=DockStyle.Top,Height=130,BackColor=Color.White,
          Padding=new Padding(8,8,8,4)};
      var tabs=new FlowLayoutPanel{Dock=DockStyle.Top,Height=38,
         FlowDirection=FlowDirection.LeftToRight,WrapContents=false,AutoScroll=false};
      foreach(string category in new[]{"All","Language","Logic","Unrelated","Resolved"}) {
        string c=category;
        var button=new Button{Text=c,Tag=c,AutoSize=false,
          Height=29,Width=c=="Unrelated"?91:(c=="Language"||c=="Resolved"?83:58),
          Font=new Font("Segoe UI",8.5f),Cursor=Cursors.Hand,
          Margin=new Padding(1,1,2,1),FlatStyle=FlatStyle.Flat};
        button.FlatAppearance.BorderSize=1;
        button.Click+=(sender,args)=>{filter=c;RefreshCards();};
        filters.Add(c,button);tabs.Controls.Add(button);
      }
      summary=new Label{Dock=DockStyle.Fill,Text="Review Text to load suggestions",
        ForeColor=Ink,Font=new Font("Segoe UI",9,FontStyle.Bold),
        Padding=new Padding(8,8,0,0)};
      demographics=new Label{Dock=DockStyle.Top,Height=43,
        Text="Patient context: not reviewed",ForeColor=Muted,Font=new Font("Segoe UI",8.5f),
        Padding=new Padding(8,3,0,0)};
      top.Controls.Add(summary);top.Controls.Add(demographics);top.Controls.Add(tabs);
      Controls.Add(top);

      var footer=new Panel{Dock=DockStyle.Bottom,Height=31,
        BackColor=Color.White,Padding=new Padding(14,6,6,4)};
      status=new Label{Dock=DockStyle.Fill,ForeColor=Muted,
        Text="Select a card to locate its text in Word.",
        Font=new Font("Segoe UI",8.3f)};
      footer.Controls.Add(status);Controls.Add(footer);
      cards=new FlowLayoutPanel{Dock=DockStyle.Fill,AutoScroll=true,
        WrapContents=false,FlowDirection=FlowDirection.TopDown,
        BackColor=Back,Padding=new Padding(8,8,8,12)};
      // Resize only when actual available width changes. Don't rebuild on
      // every Word COM polling tick or transient layout scroll event.
      cards.SizeChanged+=(s,e)=>ResizeCards();
      Controls.Add(cards);cards.BringToFront();
    }

    internal void UpdatePatientContext(string context,bool headerDetected,bool reviewed) {
      string sex="Not detected",age="Not detected";
      foreach(string part in (context??"").Split(';')) {
        if(part.StartsWith("sex="))sex=part.Substring(4)=="male"?"Male":"Female";
        if(part.StartsWith("age="))age=part.Substring(4);
      }
      string source=headerDetected?"Header":String.IsNullOrEmpty(context)?"No readable demographics":"Report";
      string state=String.IsNullOrEmpty(context)?"Unavailable":reviewed?"Checked":"Detected; review pending";
      demographics.Text="Patient context — "+source+" | "+state+"\nGender: "+sex+"   Age: "+age;
      demographics.ForeColor=String.IsNullOrEmpty(context)?Muted:Ink;
    }
    internal void UpdateIssues(IList<Issue> updates) {
      issues=updates==null?new List<Issue>():new List<Issue>(updates);
      RefreshCards();
    }

    private void RefreshCards() {
      if(rebuilding)return;
      rebuilding=true;
      try {
        cards.SuspendLayout();
        foreach(Control c in cards.Controls.Cast<Control>().ToArray())c.Dispose();
        cards.Controls.Clear();
        int all=issues.Count(x=>!x.Resolved);
        int resolved=issues.Count(x=>x.Resolved);
        summary.Text=all+" open suggestion"+(all==1?"":"s")+"  ·  "+resolved+" resolved";
        foreach(var pair in filters) {
          var b=pair.Value;
          int count=pair.Key=="All"?all:
            pair.Key=="Resolved"?resolved:
            issues.Count(x=>!x.Resolved&&x.Category==pair.Key);
          b.Text=pair.Key+" ("+count+")";
          b.BackColor=pair.Key==filter?Navy:Color.FromArgb(238,243,248);
          b.ForeColor=pair.Key==filter?Color.White:Ink;
          b.FlatAppearance.BorderColor=pair.Key==filter?Navy:Border;
        }
        var shown=issues.Where(x=>filter=="Resolved"?x.Resolved:
          !x.Resolved&&(filter=="All"||x.Category==filter)).ToList();
        if(shown.Count==0) {
          cards.Controls.Add(new Label{Text="No "+filter.ToLowerInvariant()+
            " suggestions to display.",ForeColor=Muted,AutoSize=false,
            Height=70,Width=385,TextAlign=ContentAlignment.MiddleCenter});
        } else {
          int number=0;
          foreach(Issue item in shown)cards.Controls.Add(CreateCard(item,++number));
        }
      } finally {cards.ResumeLayout(true);rebuilding=false;}
    }
    private static string Short(string value,int limit) {
      if(String.IsNullOrEmpty(value))return "";
      string clean=value.Replace('\r',' ').Replace('\n',' ').Trim();
      return clean.Length<=limit?clean:clean.Substring(0,limit-1)+"…";
    }
    private static int TextHeight(string value,int width,Font font,int min,int max) {
      int needed=TextRenderer.MeasureText(value??"",font,
        new Size(Math.Max(125,width),3000),
        TextFormatFlags.WordBreak|TextFormatFlags.TextBoxControl).Height+10;
      return Math.Min(max,Math.Max(min,needed));
    }
    private Control CreateCard(Issue issue,int index) {
      int contentWidth=Math.Max(305,cards.ClientSize.Width-42);
      var box=new Panel{Width=contentWidth,Height=260,BackColor=Color.White,
         BorderStyle=BorderStyle.FixedSingle,Margin=new Padding(1,1,0,9),
         Tag=issue.Id};
      int x=13,y=9,width=contentWidth-26;
      Color accent=issue.Category=="Logic"?Color.FromArgb(145,80,193):
        issue.Category=="Unrelated"?Color.FromArgb(226,127,51):Navy;
      var type=new Label{Text=issue.Resolved?"✓ RESOLVED":issue.Category.ToUpperInvariant(),
        Location=new Point(x,y),Size=new Size(125,22),BackColor=issue.Resolved?
          Color.FromArgb(220,242,229):Color.FromArgb(231,239,252),
        ForeColor=issue.Resolved?Green:accent,
        Font=new Font("Segoe UI",8.5f,FontStyle.Bold),
        TextAlign=ContentAlignment.MiddleCenter};
      box.Controls.Add(type);
      var count=new Label{Text="#"+index,Location=new Point(contentWidth-49,y+2),
        Size=new Size(35,20),ForeColor=Muted,TextAlign=ContentAlignment.MiddleRight};
      box.Controls.Add(count);
      y+=30;
      var quote=AddLine(box,"Issue",issue.Quote,x,ref y,width,Color.FromArgb(248,250,253),false);
      var why=AddLine(box,"Why",issue.Why,x,ref y,width,Color.White,false);
      var proposed=AddLine(box,"Suggestion",issue.Suggestion,x,ref y,width,
        Color.FromArgb(229,247,236),true);
      if(issue.Resolved) {
        var badge=new Label{Text="✓  Resolved",Location=new Point(x,y+5),
          Size=new Size(110,25),BackColor=Color.FromArgb(227,247,233),
          ForeColor=Green,Font=new Font("Segoe UI",9,FontStyle.Bold),
          TextAlign=ContentAlignment.MiddleCenter};
        box.Controls.Add(badge);y+=41;
      } else {
        var applyButton=new Button{Text="Apply",Location=new Point(x,y+5),
          Size=new Size(116,29),FlatStyle=FlatStyle.Flat,
          BackColor=Navy,ForeColor=Color.White,Cursor=Cursors.Hand,
          Font=new Font("Segoe UI",9,FontStyle.Bold)};
        applyButton.FlatAppearance.BorderSize=0;
        applyButton.Enabled=!String.IsNullOrWhiteSpace(issue.Suggestion)||
          !String.IsNullOrWhiteSpace(proposed.Text);
        applyButton.Click+=(s,e)=>apply(issue,proposed.Text);
        proposed.TextChanged+=(s,e)=>applyButton.Enabled=
          !String.IsNullOrWhiteSpace(proposed.Text);
        var ignoreButton=new Button{Text="Ignore",Location=new Point(x+125,y+5),
          Size=new Size(97,29),BackColor=Color.FromArgb(243,245,248),
          FlatStyle=FlatStyle.Flat,Cursor=Cursors.Hand};
        ignoreButton.FlatAppearance.BorderColor=Border;
        ignoreButton.Click+=(s,e)=>ignore(issue);
        box.Controls.Add(applyButton);box.Controls.Add(ignoreButton);y+=45;
      }
      string feedback;
      if(issueFeedback.TryGetValue(issue.Id,out feedback)) {
        var hint=new Label{
          Text=feedback,AutoSize=false,
          Location=new Point(x,y+2),Size=new Size(width-4,46),
          BackColor=Color.FromArgb(255,245,221),
          ForeColor=Color.FromArgb(128,74,16),
          Font=new Font("Segoe UI",8.5f),
          TextAlign=ContentAlignment.MiddleLeft
        };
        box.Controls.Add(hint);y+=54;
      }
      var clicker=new Label{Location=new Point(contentWidth-128,35),
        Size=new Size(108,18),Text="Locate in Word",
        ForeColor=Color.FromArgb(37,104,189),Cursor=Cursors.Hand,
        TextAlign=ContentAlignment.MiddleRight,Font=new Font("Segoe UI",8)};
      clicker.Click+=(s,e)=>navigate(issue);
      box.Controls.Add(clicker);
      box.Height=y+7;
      return box;
    }
    private TextBox AddLine(Panel parent,string caption,string content,
           int x,ref int y,int width,Color background,bool bold) {
      var name=new Label{Text=caption+":",Location=new Point(x,y),
        Size=new Size(90,19),ForeColor=Ink,
        Font=new Font("Segoe UI",9,FontStyle.Bold)};
      parent.Controls.Add(name);
      y+=21;
      string value=content??"";
      int height;
      using(var measureFont=new Font("Segoe UI",9,
          bold?FontStyle.Bold:FontStyle.Regular))
        height=TextHeight(value,width-4,measureFont,34,117);
      var text=new TextBox{Text=value,Location=new Point(x,y),
        Size=new Size(width-3,height),Multiline=true,ReadOnly=!bold,
        BorderStyle=BorderStyle.None,ScrollBars=ScrollBars.Vertical,
        BackColor=background,ForeColor=Ink,
        Font=new Font("Segoe UI",9,bold?FontStyle.Bold:FontStyle.Regular)};
      parent.Controls.Add(text);
      y+=height+7;
      return text;
    }
    private void ResizeCards() {
      if(rebuilding||cards.Controls.Count==0)return;
      int width=Math.Max(285,cards.ClientSize.Width-42);
      foreach(Control item in cards.Controls) {
        if(item is Panel&&Math.Abs(item.Width-width)>8) {
          RefreshCards();return;
        }
      }
    }
    internal void UpdateStatus(string message) {
      status.Text=message;
    }
    internal void ShowIssueFeedback(Issue issue,string message) {
      if(issue==null)return;
      if(String.IsNullOrWhiteSpace(message))issueFeedback.Remove(issue.Id);
      else issueFeedback[issue.Id]=message;
      RefreshCards();
    }
    internal void SetWindowPlacement(Rectangle wordRectangle) {
      if(wordRectangle.Width<750||wordRectangle.Height<300)return;
      int paneWidth=Math.Min(480,Math.Max(380,wordRectangle.Width/3));
      int headerOffset=Math.Min(170,Math.Max(108,wordRectangle.Height/6));
      var target=new Rectangle(wordRectangle.Right-paneWidth-8,
        wordRectangle.Top+headerOffset,paneWidth,
        Math.Max(260,wordRectangle.Height-headerOffset-30));
      if(Bounds!=target)Bounds=target;
    }
  }
}
