using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

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
    public static int Main(string[] args) {
      try {
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