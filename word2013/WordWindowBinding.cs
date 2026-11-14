using System;
using System.Runtime.InteropServices;
using System.Text;

namespace RadiologyAgent {
  // Word 2013 native document object model, bound to the exact clicked Word window.
  // Microsoft's AccessibleObjectFromWindow/OBJID_NATIVEOM for _WwG returns Word.Window.
  internal static class WordWindowBinding {
    private const uint NativeOm=0xFFFFFFF0;
    private static readonly Guid DispatchId=new Guid("00020400-0000-0000-C000-000000000046");

    [DllImport("user32.dll", CharSet=CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hwnd,StringBuilder name,int max);
    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hwnd);
    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr parent,EnumWindowsProc callback,IntPtr arg);
    private delegate bool EnumWindowsProc(IntPtr hwnd,IntPtr arg);
    [DllImport("oleacc.dll", PreserveSig=true)]
    private static extern int AccessibleObjectFromWindow(IntPtr hwnd,uint objectId,
      ref Guid iid,[MarshalAs(UnmanagedType.Interface)] out object result);

    internal static bool TryGetWordWindow(uint wordRoot,out object nativeWordWindow,out bool foundNativeCandidate) {
      nativeWordWindow=null;
      foundNativeCandidate=false;
      if(wordRoot==0)return false;
      IntPtr root=new IntPtr(unchecked((int)wordRoot));
      if(!IsWindow(root))return false;
      object found=null;
      bool sawCandidate=false;
      EnumWindowsProc callback=(hwnd,arg)=>{
        var name=new StringBuilder(64);
        GetClassName(hwnd,name,name.Capacity);
        if(!String.Equals(name.ToString(),"_WwG",StringComparison.Ordinal))return true;
        object candidate=null;
        Guid iid=DispatchId;
        try {
          int hr=AccessibleObjectFromWindow(hwnd,NativeOm,ref iid,out candidate);
          if(hr>=0&&candidate!=null) {
            sawCandidate=true;
            // NativeOM can expose a transitioning or inactive pane.
            // Verify Document before selecting an object from this exact root.
            try {
              dynamic win=candidate;
              object doc=win.Document;
              if(doc!=null) {
                found=candidate;
                return false;
              }
            } catch(COMException) {
              // Try other panes of this same Word window, then poll again.
            } catch {
              // Invalid object. Never fall back to another Word instance.
            }
          }
        } catch (COMException) {}
        return true;
      };
      EnumChildWindows(root,callback,IntPtr.Zero);
      GC.KeepAlive(callback);
      foundNativeCandidate=sawCandidate;
      nativeWordWindow=found;
      return nativeWordWindow!=null;
    }
  }
}