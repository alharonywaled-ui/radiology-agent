using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace RadiologyAgent {
  internal static class ReviewHost {
    [STAThread]
    private static void Main() {
      Application.EnableVisualStyles();
      Application.SetCompatibleTextRenderingDefault(false);
      try {
        object word=Marshal.GetActiveObject("Word.Application");
        using(var form=new WebReviewForm(word)) form.ShowDialog();
      }
      catch(COMException ex) {
        MessageBox.Show("Unable to connect to Microsoft Word. Ensure Word 2013 is open and a document is active.\r\n"+ex.Message,
          "Radiology Agent",MessageBoxButtons.OK,MessageBoxIcon.Warning);
      }
      catch(Exception ex) {
        MessageBox.Show(ex.Message,"Radiology Agent",MessageBoxButtons.OK,MessageBoxIcon.Error);
      }
    }
  }
}