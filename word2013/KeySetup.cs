using System;
using System.IO;
using System.Diagnostics;
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
    private const string AddinClsid="{7B918AD1-1C42-4B73-A4BB-9FC83743DD6E}";
    private const string AddinProg="RadiologyAgent.Word2013.Safe";
    private static string EnterpriseRoot() {
      string pf=Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
      if(String.IsNullOrEmpty(pf))pf=Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
      return Path.Combine(pf,"Haronex");
    }
    private static bool EnterpriseReady() {
      string root=EnterpriseRoot();
      using(var lm=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry32))
      using(var machine=lm.OpenSubKey(@"Software\Haronex\Enterprise"))
      using(var server=lm.OpenSubKey(@"Software\Classes\CLSID\"+AddinClsid+@"\InprocServer32"))
      using(var addin=lm.OpenSubKey(@"Software\Microsoft\Office\Word\Addins\"+AddinProg)) {
        return machine!=null&&server!=null&&addin!=null&&
          String.Equals(Convert.ToString(machine.GetValue("InstallPath")),root,StringComparison.OrdinalIgnoreCase)&&
          String.Equals(Convert.ToString(server.GetValue(null)),Path.Combine(root,"RadiologyAgent.NativeShim.dll"),StringComparison.OrdinalIgnoreCase)&&
          Convert.ToInt32(addin.GetValue("LoadBehavior",0))==3&&InstallHealth.CheckEssential(root)==null;
      }
    }
    private static void RemoveOldClasses(RegistryKey classes) {
      if(classes==null)return;
      classes.DeleteSubKeyTree(@"CLSID\"+AddinClsid,false);
      classes.DeleteSubKeyTree(AddinProg,false);
      classes.DeleteSubKeyTree(@"CLSID\{D1FBEFD5-3B43-4516-9B7B-CDEB47AB0FA1}",false);
      classes.DeleteSubKeyTree("RadiologyAgent.WordAddin",false);
    }
    private static void RemoveOldUserRegistration(RegistryKey user) {
      using(var classes=user.OpenSubKey(@"Software\Classes",true))RemoveOldClasses(classes);
      foreach(string prog in new[]{AddinProg,"RadiologyAgent.WordAddin"}) {
        user.DeleteSubKeyTree(@"Software\Microsoft\Office\Word\Addins\"+prog,false);
        user.DeleteSubKeyTree(@"Software\Microsoft\Office\15.0\Word\Addins\"+prog,false);
      }
    }
    private static int MigrateLegacy() {
      if(!EnterpriseReady())return 20;
      // Never mutate the registration of a running Word session.
      int session=Process.GetCurrentProcess().SessionId;
      foreach(var proc in Process.GetProcessesByName("WINWORD"))using(proc)
        if(proc.SessionId==session)return 31;
      using(var cu=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,RegistryView.Registry32))
        RemoveOldUserRegistration(cu);
      return 0;
    }
    private static int MigrateLoadedUsers() {
      if(!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))return 12;
      if(!EnterpriseReady())return 20;
      foreach(var proc in Process.GetProcessesByName("WINWORD"))using(proc)return 31;
      // Only loaded real user hives, only Haronex's known keys. Do not load
      // offline profile hives or touch keys, reports, consent or other add-ins.
      using(var users=RegistryKey.OpenBaseKey(RegistryHive.Users,RegistryView.Registry32)) {
        foreach(string sid in users.GetSubKeyNames()) {
          if(!sid.StartsWith("S-1-5-21-",StringComparison.OrdinalIgnoreCase)||
             sid.EndsWith("_Classes",StringComparison.OrdinalIgnoreCase))continue;
          using(var user=users.OpenSubKey(sid,true))if(user!=null)RemoveOldUserRegistration(user);
          using(var classes=users.OpenSubKey(sid+"_Classes",true))RemoveOldClasses(classes);
        }
      }
      return 0;
    }
    private static int PrepareInstall(bool machine) {
      int session=Process.GetCurrentProcess().SessionId;
      foreach(var proc in Process.GetProcessesByName("WINWORD"))using(proc)
        if(machine||proc.SessionId==session)return 31;
      // Stop only Haronex processes, never Word or unrelated processes.
      foreach(string name in new[]{"RadiologyAgent.ReviewHost","HaronexGroqWorker"})
        foreach(var proc in Process.GetProcessesByName(name))using(proc) {
          if(!machine&&proc.SessionId!=session)continue;
          if(!proc.HasExited){proc.Kill();if(!proc.WaitForExit(8000))return 32;}
        }
      return 0;
    }
    private static int VerifyInstall() {
      string code=InstallHealth.Verify(AppDomain.CurrentDomain.BaseDirectory);
      Console.WriteLine(code??"INSTALL_INTEGRITY_OK");
      return code==null?0:40;
    }
    private static bool MachineOption(string name,out bool enabled) {
      enabled=false;
      using(var lm=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry32))
      using(var policy=lm.OpenSubKey(@"Software\Haronex\Enterprise")) {
        if(policy==null||policy.GetValue(name)==null||policy.GetValueKind(name)!=RegistryValueKind.DWord)return false;
        enabled=Convert.ToInt32(policy.GetValue(name))==1;
        return true;
      }
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
        var mandatory=new CheckBox {Left=18,Top=187,Width=510,Text="Enable Mandatory Review (successful Review Text)"};
        var recovery=new CheckBox {Left=18,Top=216,Width=510,Text="Disable AutoRecover (unsaved work may be lost)"};
        using(var settings=Registry.CurrentUser.CreateSubKey(@"Software\Haronex")) {
          consent.Checked=Convert.ToInt32(settings.GetValue("WebTransferAllowed",0))==1;
          mandatory.Checked=Convert.ToInt32(settings.GetValue("MandatoryReviewEnabled",0))==1;
          recovery.Checked=Convert.ToInt32(settings.GetValue("DisableAutoRecover",0))==1;
          bool machineMandatory,machineRecovery;
          if(MachineOption("MandatoryReviewEnabled",out machineMandatory)) {
            mandatory.Checked=machineMandatory;mandatory.Enabled=false;
            mandatory.Text="Enable Mandatory Review (managed by IT)";
          }
          if(MachineOption("DisableAutoRecover",out machineRecovery)) {
            recovery.Checked=machineRecovery;recovery.Enabled=false;
            recovery.Text="Disable AutoRecover (managed by IT; unsaved work may be lost)";
          }
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
              if(EnterpriseReady()) {
                int migrated=MigrateLegacy();
                if(migrated!=0)throw new InvalidOperationException("Close Word before migrating Haronex. Contact IT if the problem continues.");
              }
              settings.SetValue("WebTransferAllowed",consent.Checked?1:0,RegistryValueKind.DWord);
              if(mandatory.Enabled)settings.SetValue("MandatoryReviewEnabled",mandatory.Checked?1:0,RegistryValueKind.DWord);
              if(recovery.Enabled)settings.SetValue("DisableAutoRecover",recovery.Checked?1:0,RegistryValueKind.DWord);
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
        if(args.Length==1&&args[0]=="--migrate-loaded-users")return MigrateLoadedUsers();
        if(args.Length==1&&args[0]=="--prepare-machine-install")return PrepareInstall(true);
        if(args.Length==1&&args[0]=="--prepare-user-install")return PrepareInstall(false);
        if(args.Length==1&&args[0]=="--verify-install")return VerifyInstall();
        if(args.Length==1&&args[0]=="--configure")return Configure();
        if(args.Length>0&&args[0]=="--check") {
          if(!File.Exists(KeyFile))return 2;
          var saved=Encoding.UTF8.GetString(ProtectedData.Unprotect(
            Convert.FromBase64String(File.ReadAllText(KeyFile)),null,DataProtectionScope.CurrentUser));
          return Valid(saved)?0:3;
        }
        // Installer writes one short-lived local staging file, removed in finally.
        if(args.Length!=2||(args[0]!="--import"&&args[0]!="--import-if-missing"))return 2;
        bool preserveExisting=args[0]=="--import-if-missing";
        string incoming=Path.GetFullPath(args[1]);
        try {
          if(preserveExisting&&File.Exists(KeyFile))return 5;
          var key=File.ReadAllText(incoming,Encoding.UTF8).Trim();
          if(!Valid(key))return 3;
          Directory.CreateDirectory(Path.GetDirectoryName(KeyFile));
          byte[] bytes=ProtectedData.Protect(Encoding.UTF8.GetBytes(key),null,DataProtectionScope.CurrentUser);
          string temp=KeyFile+"."+Guid.NewGuid().ToString("N")+".tmp";
          try {
            File.WriteAllText(temp,Convert.ToBase64String(bytes),Encoding.ASCII);
            // File.Move fails instead of replacing a file that appeared concurrently.
            if(preserveExisting) {
              if(File.Exists(KeyFile))return 5;
              try {File.Move(temp,KeyFile);}catch(IOException){if(File.Exists(KeyFile))return 5;throw;}
            } else if(File.Exists(KeyFile))File.Replace(temp,KeyFile,null);
            else File.Move(temp,KeyFile);
            return 0;
          }finally{try{File.Delete(temp);}catch{}}
        } finally {try {File.Delete(incoming);}catch{}}
      } catch {return 1;}
    }
  }
}