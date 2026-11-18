using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace RadiologyAgent {
  /// <summary>
  /// A non-persistent, blocking frosted screenshot overlay, *not* a Word
  /// document element and never a Win32 process/thread freeze. It intercepts
  /// the user's mouse/keyboard while async analysis continues on its worker.
  /// All bitmaps are held only in memory and disposed when review ends.
  /// </summary>
  internal sealed class HaronexReviewLoadingOverlay : Form {
    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect { public int Left,Top,Right,Bottom; }
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hwnd,out NativeRect rect);
    [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr hwnd);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr hwnd,uint flags);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
    [DllImport("user32.dll")] private static extern bool SetWindowPos(
      IntPtr hwnd,IntPtr after,int x,int y,int cx,int cy,uint flags);
    private static readonly IntPtr HwndTopmost=new IntPtr(-1);
    private const uint SWP_NOACTIVATE=0x0010;

    private readonly IntPtr wordHwnd;
    private readonly Stopwatch clock=Stopwatch.StartNew();
    private readonly Timer ticker;
    private readonly Button cancel;
    private readonly Action cancelRequested;
    private Bitmap blurredSnapshot;
    private int spinnerPhase;
    private string message="Reviewing report…";
    private bool cancelled;
    private bool positioning;
    private static readonly Color Navy=Color.FromArgb(21,62,125);
    private static readonly Color Blue=Color.FromArgb(34,117,216);
    private static readonly Color Ink=Color.FromArgb(33,53,81);

    internal HaronexReviewLoadingOverlay(IntPtr ownerHwnd,Action onCancel) {
      wordHwnd=ownerHwnd;
      cancelRequested=onCancel;
      Text="Haronex · Reviewing";
      FormBorderStyle=FormBorderStyle.None;
      ShowInTaskbar=false;
      StartPosition=FormStartPosition.Manual;
      TopMost=true;
      BackColor=Color.FromArgb(237,244,252);
      DoubleBuffered=true;
      KeyPreview=true;
      MinimumSize=new Size(250,150);
      // Prevent users from editing an incomplete report during AI review.
      // Escape is a safe stop, NEVER a hidden process freeze.
      KeyDown+=(s,e)=>{
        if(e.KeyCode==Keys.Escape){e.Handled=true;RequestCancellation();}
        else {e.Handled=true;e.SuppressKeyPress=true;}
      };
      cancel=new Button{
        Text="Cancel Review",Width=146,Height=39,
        FlatStyle=FlatStyle.Flat,
        BackColor=Color.FromArgb(242,245,250),
        ForeColor=Navy,
        Font=new Font("Segoe UI",10,FontStyle.Bold),
        Cursor=Cursors.Hand
      };
      cancel.FlatAppearance.BorderSize=1;
      cancel.FlatAppearance.BorderColor=Color.FromArgb(200,213,231);
      cancel.Click+=(s,e)=>RequestCancellation();
      Controls.Add(cancel);
      ticker=new Timer{Interval=100};
      ticker.Tick+=(s,e)=>{
        if(IsDisposed)return;
        spinnerPhase=(spinnerPhase+1)%12;
        if(!UpdatePlacement()) {End();return;}
        if(Visible)Invalidate(new Rectangle(
          Math.Max(0,ClientSize.Width/2-198),
          Math.Max(0,ClientSize.Height/2-148),396,330));
      };
      Shown+=(s,e)=>ticker.Start();
      FormClosed+=(s,e)=>{
        ticker.Stop();ticker.Dispose();
        if(blurredSnapshot!=null){blurredSnapshot.Dispose();blurredSnapshot=null;}
      };
      Resize+=(s,e)=>PositionControls();
    }

    internal static bool SelfTest() {
      return FormatElapsed(TimeSpan.Zero)=="00:00"
        && FormatElapsed(TimeSpan.FromSeconds(71))=="01:11"
        && FormatElapsed(TimeSpan.FromSeconds(3605))=="60:05";
    }
    internal static string FormatElapsed(TimeSpan elapsed) {
      int total=Math.Max(0,(int)elapsed.TotalSeconds);
      return (total/60).ToString("00")+":"+(total%60).ToString("00");
    }

    internal bool Begin() {
      if(wordHwnd==IntPtr.Zero||!IsWindow(wordHwnd)||IsIconic(wordHwnd))return false;
      NativeRect r;
      if(!GetWindowRect(wordHwnd,out r))return false;
      Rectangle area=Rectangle.FromLTRB(r.Left,r.Top,r.Right,r.Bottom);
      if(area.Width<300||area.Height<200)return false;
      // Capture only once BEFORE we display our overlay. No image is written
      // to disk, sent to the API, or stored in logs.
      try {blurredSnapshot=CaptureSoftBlur(area);}catch{blurredSnapshot=null;}
      Bounds=area;
      Show();
      BringToFront();
      return true;
    }

    // A fast, soft blur by downsample + smooth upscale, with a translucent
    // wash. Much cheaper than running a full Gaussian filter every frame.
    private static Bitmap CaptureSoftBlur(Rectangle area) {
      using(var screenshot=new Bitmap(area.Width,area.Height)) {
        using(var g=Graphics.FromImage(screenshot))
          g.CopyFromScreen(area.Location,Point.Empty,area.Size,CopyPixelOperation.SourceCopy);
        int smallWidth=Math.Max(32,area.Width/19);
        int smallHeight=Math.Max(24,area.Height/19);
        using(var reduced=new Bitmap(smallWidth,smallHeight)) {
          using(var g=Graphics.FromImage(reduced)) {
            g.InterpolationMode=InterpolationMode.HighQualityBilinear;
            g.DrawImage(screenshot,0,0,smallWidth,smallHeight);
          }
          var result=new Bitmap(area.Width,area.Height);
          using(var g=Graphics.FromImage(result)) {
            g.InterpolationMode=InterpolationMode.HighQualityBicubic;
            g.DrawImage(reduced,0,0,area.Width,area.Height);
            using(var wash=new SolidBrush(Color.FromArgb(145,247,250,255)))
              g.FillRectangle(wash,0,0,area.Width,area.Height);
          }
          return result;
        }
      }
    }

    private bool UpdatePlacement() {
      if(wordHwnd==IntPtr.Zero||!IsWindow(wordHwnd)||IsIconic(wordHwnd))
        return false;
      IntPtr active=GetForegroundWindow();
      uint activePid=0,wordPid=0;
      if(active!=IntPtr.Zero)GetWindowThreadProcessId(active,out activePid);
      GetWindowThreadProcessId(wordHwnd,out wordPid);
      // Don't cover another application when the user Alt-Tabs away from Word.
      // Reappear as soon as the pinned Word window regains the foreground.
      IntPtr activeRoot=active==IntPtr.Zero?IntPtr.Zero:GetAncestor(active,2);
      bool otherWordWindow=activePid==wordPid&&activeRoot!=wordHwnd;
      if(otherWordWindow||
         (activePid!=wordPid&&activePid!=(uint)Process.GetCurrentProcess().Id)) {
        if(Visible)Hide();
        return true;
      }
      NativeRect r;
      if(!GetWindowRect(wordHwnd,out r))return false;
      Rectangle target=Rectangle.FromLTRB(r.Left,r.Top,r.Right,r.Bottom);
      if(target.Width<200||target.Height<150)return false;
      if(Bounds!=target) {
        Bounds=target;
        // The captured bitmap stretches on resize. No second screen capture.
      }
      if(!Visible&&!cancelled)Show();
      return true;
    }

    internal void SetMessage(string value) {
      if(!String.IsNullOrWhiteSpace(value))message=value;
      Invalidate();
    }
    internal void End() {
      if(IsDisposed)return;
      ticker.Stop();Close();
    }

    private void RequestCancellation() {
      if(cancelled)return;
      cancelled=true;
      cancel.Enabled=false;
      message="Cancelling review…";
      Invalidate();
      if(cancelRequested!=null)cancelRequested();
      End();
    }

    private void PositionControls() {
      if(positioning)return;
      positioning=true;
      try {
        cancel.Left=Math.Max(5,(ClientSize.Width-cancel.Width)/2);
        cancel.Top=Math.Min(ClientSize.Height-45,
          Math.Max(100,ClientSize.Height/2+104));
      } finally {positioning=false;}
    }

    protected override void OnPaint(PaintEventArgs e) {
      base.OnPaint(e);
      Graphics g=e.Graphics;
      g.SmoothingMode=SmoothingMode.AntiAlias;
      if(blurredSnapshot!=null)
        g.DrawImage(blurredSnapshot,ClientRectangle);
      else using(var fill=new SolidBrush(Color.FromArgb(241,247,253)))
        g.FillRectangle(fill,ClientRectangle);
      using(var tint=new SolidBrush(Color.FromArgb(63,246,250,255)))
        g.FillRectangle(tint,ClientRectangle);
      int cx=ClientSize.Width/2,cy=ClientSize.Height/2;
      Rectangle card=new Rectangle(cx-195,cy-145,390,325);
      using(var shadow=new SolidBrush(Color.FromArgb(45,37,63,106)))
        g.FillEllipse(shadow,cx-174,cy-119,348,282);
      using(var white=new SolidBrush(Color.FromArgb(251,255,255,255)))
        g.FillRectangle(white,card);
      using(var outline=new Pen(Color.FromArgb(209,221,237),1))
        g.DrawRectangle(outline,card);

      // 12-segment animated spinner is custom drawn; no media assets, browser,
      // GIF files, or repeatedly decoding images on the UI thread.
      for(int i=0;i<12;i++) {
        double angle=(Math.PI*2*i/12)-Math.PI/2;
        double inner=41,outer=57;
        int strength=(i-spinnerPhase+12)%12;
        int alpha=Math.Max(52,255-strength*17);
        using(var pen=new Pen(Color.FromArgb(alpha,25,102,202),6)) {
          pen.StartCap=LineCap.Round;pen.EndCap=LineCap.Round;
          g.DrawLine(pen,
            cx+(float)(Math.Cos(angle)*inner),
            cy-42+(float)(Math.Sin(angle)*inner),
            cx+(float)(Math.Cos(angle)*outer),
            cy-42+(float)(Math.Sin(angle)*outer));
        }
      }
      DrawCentered(g,"HARONEX",new Font("Segoe UI",12,FontStyle.Bold),
        Navy,new Rectangle(card.Left,card.Top+20,card.Width,24));
      DrawCentered(g,message,new Font("Segoe UI",12,FontStyle.Bold),
        Ink,new Rectangle(card.Left+10,cy+28,card.Width-20,33));
      DrawCentered(g,FormatElapsed(clock.Elapsed),
        new Font("Consolas",23,FontStyle.Bold),Blue,
        new Rectangle(card.Left,cy+58,card.Width,40));
      DrawCentered(g,"Word is temporarily locked while the report is reviewed",
        new Font("Segoe UI",8.5f),Color.FromArgb(95,108,127),
        new Rectangle(card.Left+5,cy+100,card.Width-10,24));
    }
    private static void DrawCentered(Graphics g,string value,Font font,
                                     Color color,Rectangle area) {
      using(font) {
        TextRenderer.DrawText(g,value,font,area,color,
          TextFormatFlags.HorizontalCenter|TextFormatFlags.VerticalCenter|
          TextFormatFlags.EndEllipsis|TextFormatFlags.SingleLine);
      }
    }
  }
}
