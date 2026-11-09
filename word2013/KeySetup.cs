using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
namespace RadiologyAgent {
  internal static class KeySetup {
    private static int Main() {
      string dir=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"RadiologyAgent");
      string incoming=Path.Combine(dir,"pending.key");
      string encrypted=Path.Combine(dir,"key.dpapi");
      try {
        if(!File.Exists(incoming))return File.Exists(encrypted)?0:2;
        string key=File.ReadAllText(incoming,Encoding.UTF8).Trim();
        if(key.Length<12||key.IndexOfAny(new [] {'\r','\n',' '})>=0)throw new Exception("Invalid Gemini key");
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