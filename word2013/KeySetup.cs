using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
namespace RadiologyAgent {
  internal static class KeySetup {
    private static bool IsPlausibleGeminiKey(string key) {
      if(String.IsNullOrEmpty(key)||!key.StartsWith("AIza",StringComparison.Ordinal)||key.Length<30||key.Length>128) return false;
      foreach(char ch in key)
        if(!Char.IsLetterOrDigit(ch) && ch!='_' && ch!='-') return false;
      return true;
    }
    private static int Main() {
      string dir=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"RadiologyAgent");
      string incoming=Path.Combine(dir,"pending.key");
      string encrypted=Path.Combine(dir,"key.dpapi");
      try {
        if (Environment.GetCommandLineArgs().Length>1 && Environment.GetCommandLineArgs()[1]=="--check") {
          if (!File.Exists(encrypted)) return 2;
          string stored=Encoding.UTF8.GetString(ProtectedData.Unprotect(
            Convert.FromBase64String(File.ReadAllText(encrypted).Trim()),null,DataProtectionScope.CurrentUser));
          return IsPlausibleGeminiKey(stored)?0:3;
        }
        if(!File.Exists(incoming))return File.Exists(encrypted)?0:2;
        string key=File.ReadAllText(incoming,Encoding.UTF8).Trim();
        if(!IsPlausibleGeminiKey(key))throw new Exception("Invalid Gemini key format");
        byte[] raw=ProtectedData.Protect(Encoding.UTF8.GetBytes(key),null,DataProtectionScope.CurrentUser);
        Directory.CreateDirectory(dir);
        File.WriteAllText(encrypted,Convert.ToBase64String(raw),Encoding.ASCII);
        key=null;
        return 0;
      } catch(Exception ex) {
        File.WriteAllText(Path.Combine(dir,"keysetup-error.log"),ex.GetType().Name,Encoding.UTF8);
        return 1;
      } finally {
        try { if(File.Exists(incoming))File.Delete(incoming); } catch {}
      }
    }
  }
}