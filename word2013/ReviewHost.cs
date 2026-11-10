using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

namespace RadiologyAgent {
  internal static class ReviewHost {
    private static void Failed() {
      try {
        using(var key=Registry.CurrentUser.CreateSubKey(@"Software\Haronex"))
          key.SetValue("ReviewResult",2,RegistryValueKind.DWord);
      } catch { }
      MessageBox.Show("رجاءً أعد المحاولة","Haronex",MessageBoxButtons.OK,MessageBoxIcon.Error);
    }
    [STAThread]
    private static void Main() {
      Application.EnableVisualStyles();
      Application.SetCompatibleTextRenderingDefault(false);
      using(var gate=new Mutex(false,@"Local\HaronexWordReviewHost")) {
        bool acquired=false;
        try {
          acquired=gate.WaitOne(0,false);
          if(!acquired){Failed();return;}
          object word=Marshal.GetActiveObject("Word.Application");
          using(var form=new WebReviewForm(word))form.ShowDialog();
        } catch(AbandonedMutexException) {
          // Another reviewer died; allow this review to take over on next click.
          Failed();
        } catch(Exception) {
          Failed();
        } finally {
          if(acquired)try{gate.ReleaseMutex();}catch{}
        }
      }
    }
  }
}
