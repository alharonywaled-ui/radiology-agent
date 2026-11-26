using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;

namespace RadiologyAgent {
  internal sealed class HaronexHostException : Exception {
    internal readonly string Code;
    internal HaronexHostException(string code) : base(code) { Code=code; }
    internal HaronexHostException(string code, Exception inner) : base(code,inner) { Code=code; }
  }
  internal static class InstallHealth {
    internal const string Manifest="HaronexFiles.sha256";
    private static readonly string[] Essential={
      "RadiologyAgent.NativeShim.dll","RadiologyAgent.ReviewHost.exe",
      "HaronexGroqWorker.exe","RadiologyAgent.KeySetup.exe",Manifest
    };
    internal static string CheckEssential(string root) {
      if(!File.Exists(Path.Combine(root,"RadiologyAgent.ReviewHost.exe")))return "HOST_REVIEW_HOST_MISSING";
      if(!File.Exists(Path.Combine(root,"RadiologyAgent.NativeShim.dll")))return "HOST_NATIVE_ADDIN_MISSING";
      if(!File.Exists(Path.Combine(root,"HaronexGroqWorker.exe")))return "HOST_WORKER_MISSING";
      if(!Directory.Exists(Path.Combine(root,"_HaronexRuntime")))return "HOST_WORKER_RUNTIME_MISSING";
      if(!File.Exists(Path.Combine(root,"RadiologyAgent.KeySetup.exe")))return "HOST_KEY_SETUP_MISSING";
      if(!File.Exists(Path.Combine(root,Manifest)))return "HOST_INSTALL_MANIFEST_MISSING";
      return null;
    }
    // The manifest is produced by CI from the exact packaged binaries.
    // No network access; compare hashes only during setup/explicit diagnostics.
    internal static string Verify(string root) {
      string check=CheckEssential(root);if(check!=null)return check;
      try {
        root=Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar)+Path.DirectorySeparatorChar;
        var seen=new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int runtimeCount=0;
        foreach(string line in File.ReadAllLines(Path.Combine(root,Manifest))) {
          if(String.IsNullOrWhiteSpace(line))continue;
          int split=line.IndexOf('|');
          if(split!=64||line.Length<=65)return "HOST_INSTALL_MANIFEST_INVALID";
          string hash=line.Substring(0,64),rel=line.Substring(65).Replace('/',Path.DirectorySeparatorChar);
          if(Path.IsPathRooted(rel)||!seen.Add(rel))return "HOST_INSTALL_MANIFEST_INVALID";
          string file=Path.GetFullPath(Path.Combine(root,rel));
          if(!file.StartsWith(root,StringComparison.OrdinalIgnoreCase))return "HOST_INSTALL_MANIFEST_INVALID";
          if(!File.Exists(file))return "HOST_INSTALL_FILE_MISSING";
          if(rel.StartsWith("_HaronexRuntime"+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))runtimeCount++;
          using(var input=File.OpenRead(file))
          using(var sha=SHA256.Create()) {
            string actual=BitConverter.ToString(sha.ComputeHash(input)).Replace("-","");
            if(!String.Equals(actual,hash,StringComparison.OrdinalIgnoreCase))return "HOST_INSTALL_HASH_MISMATCH";
          }
        }
        foreach(string rel in Essential)if(rel!=Manifest&&!seen.Contains(rel))return "HOST_INSTALL_MANIFEST_INVALID";
        return runtimeCount>0?null:"HOST_INSTALL_MANIFEST_INVALID";
      } catch (UnauthorizedAccessException) {return "HOST_INSTALL_ACCESS_DENIED";}
        catch (IOException) {return "HOST_INSTALL_READ_FAILED";}
        catch {return "HOST_INSTALL_MANIFEST_INVALID";}
    }
  }
}
