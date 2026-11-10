using System;
using System.Threading;
using System.Windows.Forms;

namespace RadiologyAgent {
  internal static class ReviewHost {
    [STAThread]
    private static void Main(string[] args) {
      if(args!=null&&args.Length>0&&args[0]=="--self-test") {
        Environment.Exit(LiveReviewMonitor.SelfTest()?0:7);
        return;
      }
      Application.EnableVisualStyles();
      Application.SetCompatibleTextRenderingDefault(false);
      // Exactly one persistent monitor per interactive Windows user.
      using(var gate=new Mutex(false,@"Local\HaronexLiveReviewV6")) {
        bool owner=false;
        try {
          try {owner=gate.WaitOne(0,false);}
          catch(AbandonedMutexException){owner=true;}
          if(!owner)return; // Duplicate startup is expected from Word's COM connection callbacks.
          Application.Run(new LiveReviewMonitor());
        } catch {
          // Automatic background startup must not interrupt Word.
          // The Ribbon stays available and a new monitor can start on its next click.
        } finally {
          if(owner)try{gate.ReleaseMutex();}catch{}
        }
      }
    }
  }
}
