using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Security.Principal;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Windows.Forms;
using Microsoft.Win32;

namespace RadiologyAgent {
  internal static class KeySetup {
    internal static string KeyFile {
      get {return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                               "RadiologyAgent", "groq.key.dpapi");}
    }
    private static bool Valid(string key) {
      if(String.IsNullOrWhiteSpace(key)||!key.StartsWith("gsk_",StringComparison.Ordinal)
         ||key.Length<20||key.Length>256)return false;
      foreach(char ch in key)if(!Char.IsLetterOrDigit(ch)&&ch!='_'&&ch!='-')return false;
      return true;
    }
    // Authorization is based on the joined domain's actual SID and token group,
    // not a user name, password in the installer, or a local group with the same name.
    private static int CheckDomain(bool requireAdmin) {
      try {
        using (var domain=Domain.GetComputerDomain()) {
          if (!String.Equals(domain.Name,"dscan.corp",StringComparison.OrdinalIgnoreCase))return 10;
          if (!requireAdmin)return 0;
          using (var entry=domain.GetDirectoryEntry())
          using (var identity=WindowsIdentity.GetCurrent()) {
            var bytes=entry.Properties["objectSid"].Value as byte[];
            if(bytes==null)return 11;
            var domainSid=new SecurityIdentifier(bytes,0);
            var admins=new SecurityIdentifier(domainSid.Value+"-512");
            return new WindowsPrincipal(identity).IsInRole(admins)?0:12;
          }
        }
      } catch {return 13;} // No domain/DC access: do not authorize removal.
    }
    private static void StoreKey(string key) {
      Directory.CreateDirectory(Path.GetDirectoryName(KeyFile));
      byte[] bytes=ProtectedData.Protect(Encoding.UTF8.GetBytes(key),null,DataProtectionScope.CurrentUser);
      string temp=KeyFile+".tmp";
      File.WriteAllText(temp,Convert.ToBase64String(bytes),Encoding.ASCII);
      if(File.Exists(KeyFile))File.Replace(temp,KeyFile,null);
      else File.Move(temp,KeyFile);
    }
    private static int MigrateLegacy() {
      // Run in each user's own session, never as another user's administrator.
      // Only remove known Haronex per-user COM registration if the installed
      // machine-wide registration is present. Keep keys, reports and old files.
      const string clsid="{7B918AD1-1C42-4B73-A4BB-9FC83743DD6E}";
      const string prog="RadiologyAgent.Word2013.Safe";
      using(var lm=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry32))
      using(var machine=lm.OpenSubKey(@"Software\Haronex\Enterprise"))
      using(var server=lm.OpenSubKey(@"Software\Classes\CLSID\"+clsid+@"\InprocServer32")) {
        if(machine==null||server==null)return 20;
        string expected=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),"Haronex","RadiologyAgent.NativeShim.dll");
        if(String.IsNullOrEmpty(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)))
          expected=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),"Haronex","RadiologyAgent.NativeShim.dll");
        if(!String.Equals(Convert.ToString(server.GetValue(null)),expected,StringComparison.OrdinalIgnoreCase)||!File.Exists(expected))return 21;
      }
      using(var cu=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,RegistryView.Registry32)) {
        cu.DeleteSubKeyTree(@"Software\Classes\CLSID\"+clsid,false);
        cu.DeleteSubKeyTree(@"Software\Classes\"+prog,false);
        cu.DeleteSubKeyTree(@"Software\Microsoft\Office\Word\Addins\"+prog,false);
        cu.DeleteSubKeyTree(@"Software\Microsoft\Office\15.0\Word\Addins\"+prog,false);
        cu.DeleteSubKeyTree(@"Software\Microsoft\Office\Word\Addins\RadiologyAgent.WordAddin",false);
        cu.DeleteSubKeyTree(@"Software\Microsoft\Office\15.0\Word\Addins\RadiologyAgent.WordAddin",false);
      }
      return 0;
    }
    private static int Configure() {
      Application.EnableVisualStyles();
      using(var form=new Form {Text="Haronex — User Configuration",Width=560,Height=340,
          StartPosition=FormStartPosition.CenterScreen,FormBorderStyle=FormBorderStyle.FixedDialog,
          MaximizeBox=false,MinimizeBox=false}) {
        var note=new Label {Left=18,Top=15,Width=510,Height=65,
          Text="Configure Haronex for this Windows account. Reports are sent to Groq only when authorized. Existing encrypted API key can be kept by leaving the field blank."};
        var label=new Label {Left=18,Top=85,Width=140,Text="Groq API key:"};
        var key=new TextBox {Left=18,Top=108,Width=510,UseSystemPasswordChar=true};
        var consent=new CheckBox {Left=18,Top=145,Width=510,Height=38,
          Text="I am authorized by my organization to send medical report text to Groq."};
        var mandatory=new CheckBox {Left=18,Top=187,Width=510,Text="Enable Mandatory Review"};
        var recovery=new CheckBox {Left=18,Top=216,Width=510,Text="Disable AutoRecover (unsaved work may be lost)"};
        using(var settings=Registry.CurrentUser.CreateSubKey(@"Software\Haronex")) {
          consent.Checked=Convert.ToInt32(settings.GetValue("WebTransferAllowed",0))==1;
          mandatory.Checked=Convert.ToInt32(settings.GetValue("MandatoryReviewEnabled",0))==1;
          recovery.Checked=Convert.ToInt32(settings.GetValue("DisableAutoRecover",0))==1;
          var save=new Button {Left=345,Top=255,Width=85,Text="Save"};
          var cancel=new Button {Left=440,Top=255,Width=85,Text="Cancel",DialogResult=DialogResult.Cancel};
          save.Click+=(sender,e)=>{
            try {
              string value=key.Text.Trim();
              if(consent.Checked) {
                if(value.Length!=0) {if(!Valid(value))throw new InvalidOperationException("Invalid Groq API key.");}
                else if(!File.Exists(KeyFile)||!Valid(Encoding.UTF8.GetString(ProtectedData.Unprotect(
                  Convert.FromBase64String(File.ReadAllText(KeyFile)),null,DataProtectionScope.CurrentUser))))
                  throw new InvalidOperationException("Enter a valid Groq API key.");
              }
              if(value.Length!=0) {if(!Valid(value))throw new InvalidOperationException("Invalid Groq API key.");StoreKey(value);}
              int migrated=MigrateLegacy();
              if(migrated!=0)throw new InvalidOperationException("Machine installation not detected. Contact IT.");
              settings.SetValue("WebTransferAllowed",consent.Checked?1:0,RegistryValueKind.DWord);
              settings.SetValue("MandatoryReviewEnabled",mandatory.Checked?1:0,RegistryValueKind.DWord);
              settings.SetValue("DisableAutoRecover",recovery.Checked?1:0,RegistryValueKind.DWord);
              settings.SetValue("LiveAutoReviewAllowed",0,RegistryValueKind.DWord);
              form.DialogResult=DialogResult.OK;form.Close();
            }catch(Exception ex){MessageBox.Show(form,ex is InvalidOperationException?ex.Message:"Unable to save Haronex configuration.","Haronex",MessageBoxButtons.OK,MessageBoxIcon.Warning);}
          };
          form.Controls.AddRange(new Control[]{note,label,key,consent,mandatory,recovery,save,cancel});
          form.AcceptButton=save;form.CancelButton=cancel;
          return form.ShowDialog()==DialogResult.OK?0:2;
        }
      }
    }
    [STAThread]
    public static int Main(string[] args) {
      try {
        if(args.Length==1&&args[0]=="--check-domain")return CheckDomain(false);
        if(args.Length==1&&args[0]=="--authorize-uninstall")return CheckDomain(true);
        if(args.Length==1&&args[0]=="--migrate-legacy")return MigrateLegacy();
        if(args.Length==1&&args[0]=="--configure")return Configure();
        if(args.Length>0&&args[0]=="--check") {
          if(!File.Exists(KeyFile))return 2;
          var saved=Encoding.UTF8.GetString(ProtectedData.Unprotect(
            Convert.FromBase64String(File.ReadAllText(KeyFile)),null,DataProtectionScope.CurrentUser));
          return Valid(saved)?0:3;
        }
        // Installer writes one short-lived local staging file, removed in finally.
        if(args.Length!=2||args[0]!="--import")return 2;
        string incoming=Path.GetFullPath(args[1]);
        try {
          var key=File.ReadAllText(incoming,Encoding.UTF8).Trim();
          if(!Valid(key))return 3;
          Directory.CreateDirectory(Path.GetDirectoryName(KeyFile));
          byte[] bytes=ProtectedData.Protect(Encoding.UTF8.GetBytes(key),null,DataProtectionScope.CurrentUser);
          string temp=KeyFile+".tmp";
          File.WriteAllText(temp,Convert.ToBase64String(bytes),Encoding.ASCII);
          if(File.Exists(KeyFile))File.Delete(KeyFile);
          File.Move(temp,KeyFile);
          return 0;
        } finally {try {File.Delete(incoming);}catch{}}
      } catch {return 1;}
    }
  }
}