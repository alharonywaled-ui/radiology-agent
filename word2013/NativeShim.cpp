#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <shellapi.h>
#include <string>
#include <new>
#include <cwchar>

static const wchar_t* PROGID = L"RadiologyAgent.Word2013.Safe";
static const wchar_t* CLSID_TEXT = L"{7B918AD1-1C42-4B73-A4BB-9FC83743DD6E}";
static const GUID CLSID_RA = {0x7B918AD1,0x1C42,0x4B73,{0xA4,0xBB,0x9F,0xC8,0x37,0x43,0xDD,0x6E}};
static const GUID IID_EXT = {0xB65AD801,0xABAF,0x11D0,{0xBB,0x8B,0x00,0xA0,0xC9,0x0F,0x27,0x44}};
static const GUID IID_RIBBON = {0x000C0396,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static HMODULE g_instance = nullptr;
static LONG g_objects = 0;
static LONG g_locks = 0;

struct IExtensibility2 : IDispatch {
  virtual HRESULT STDMETHODCALLTYPE OnConnection(IDispatch*, LONG, IDispatch*, SAFEARRAY**) = 0;
  virtual HRESULT STDMETHODCALLTYPE OnDisconnection(LONG, SAFEARRAY**) = 0;
  virtual HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**) = 0;
  virtual HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**) = 0;
  virtual HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**) = 0;
};
struct IRibbonExtensibilityNative : IDispatch {
  virtual HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR, BSTR*) = 0;
};

static HRESULT RegistryValue(HKEY root, const std::wstring& key, const wchar_t* name, const wchar_t* value) {
  HKEY h = nullptr;
  LONG rc = RegCreateKeyExW(root,key.c_str(),0,nullptr,REG_OPTION_NON_VOLATILE,KEY_SET_VALUE,nullptr,&h,nullptr);
  if(rc!=ERROR_SUCCESS)return HRESULT_FROM_WIN32(rc);
  DWORD size=static_cast<DWORD>((wcslen(value)+1)*sizeof(wchar_t));
  rc=RegSetValueExW(h,name,0,REG_SZ,reinterpret_cast<const BYTE*>(value),size);
  RegCloseKey(h);
  return HRESULT_FROM_WIN32(rc);
}
static HRESULT RegistryDword(HKEY root, const std::wstring& key, const wchar_t* name, DWORD value) {
  HKEY h=nullptr;
  LONG rc=RegCreateKeyExW(root,key.c_str(),0,nullptr,REG_OPTION_NON_VOLATILE,KEY_SET_VALUE,nullptr,&h,nullptr);
  if(rc!=ERROR_SUCCESS)return HRESULT_FROM_WIN32(rc);
  rc=RegSetValueExW(h,name,0,REG_DWORD,reinterpret_cast<BYTE*>(&value),sizeof(value));
  RegCloseKey(h);
  return HRESULT_FROM_WIN32(rc);
}
static std::wstring ModuleDirectory() {
  wchar_t path[MAX_PATH]={0};
  GetModuleFileNameW(g_instance,path,MAX_PATH);
  wchar_t* slash=wcsrchr(path,L'\\');
  if(slash)*(slash+1)=L'\0';
  return std::wstring(path);
}
static void LaunchReview() {
  std::wstring exe=ModuleDirectory()+L"RadiologyAgent.ReviewHost.exe";
  if(GetFileAttributesW(exe.c_str())==INVALID_FILE_ATTRIBUTES) {
    MessageBoxW(nullptr,L"Radiology Agent ReviewHost.exe was not found. Reinstall the application.",L"Radiology Agent",MB_OK|MB_ICONERROR);
    return;
  }
  HINSTANCE outcome=ShellExecuteW(nullptr,L"open",exe.c_str(),nullptr,nullptr,SW_SHOWNORMAL);
  if(reinterpret_cast<INT_PTR>(outcome)<=32)
    MessageBoxW(nullptr,L"Unable to start the report review window.",L"Radiology Agent",MB_OK|MB_ICONERROR);
}
class Agent : public IExtensibility2, public IRibbonExtensibilityNative {
  LONG refs_;
public:
  Agent():refs_(1){InterlockedIncrement(&g_objects);}
  ~Agent(){InterlockedDecrement(&g_objects);}
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid,void** ppv) override {
    if(!ppv)return E_POINTER;
    *ppv=nullptr;
    if(IsEqualIID(riid,IID_IUnknown)||IsEqualIID(riid,IID_IDispatch)||IsEqualIID(riid,IID_EXT))
      *ppv=static_cast<IExtensibility2*>(this);
    else if(IsEqualIID(riid,IID_RIBBON)) *ppv=static_cast<IRibbonExtensibilityNative*>(this);
    if(!*ppv)return E_NOINTERFACE;
    AddRef();return S_OK;
  }
  ULONG STDMETHODCALLTYPE AddRef() override{return static_cast<ULONG>(InterlockedIncrement(&refs_));}
  ULONG STDMETHODCALLTYPE Release() override {LONG r=InterlockedDecrement(&refs_);if(r==0)delete this;return static_cast<ULONG>(r);}
  HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* count) override{if(!count)return E_POINTER;*count=0;return S_OK;}
  HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT,LCID,ITypeInfo**) override{return E_NOTIMPL;}
  HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID,LPOLESTR* names,UINT count,LCID,DISPID* ids) override {
    if(!names||!ids)return E_POINTER;
    for(UINT i=0;i<count;i++){
      const wchar_t* n=names[i];
      if(!_wcsicmp(n,L"OnConnection")) ids[i]=1;
      else if(!_wcsicmp(n,L"OnDisconnection")) ids[i]=2;
      else if(!_wcsicmp(n,L"OnAddInsUpdate")) ids[i]=3;
      else if(!_wcsicmp(n,L"OnStartupComplete")) ids[i]=4;
      else if(!_wcsicmp(n,L"OnBeginShutdown")) ids[i]=5;
      else if(!_wcsicmp(n,L"ReviewTextClick")) ids[i]=100;
      else return DISP_E_UNKNOWNNAME;
    }
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE Invoke(DISPID id,REFIID,LCID,WORD,DISPPARAMS*,VARIANT* result,EXCEPINFO*,UINT*) override {
    if(result)VariantInit(result);
    if(id==100){LaunchReview();return S_OK;}
    if(id>=1&&id<=5)return S_OK;
    return DISP_E_MEMBERNOTFOUND;
  }
  HRESULT STDMETHODCALLTYPE OnConnection(IDispatch*,LONG,IDispatch*,SAFEARRAY**) override{return S_OK;}
  HRESULT STDMETHODCALLTYPE OnDisconnection(LONG,SAFEARRAY**) override{return S_OK;}
  HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**) override{return S_OK;}
  HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**) override{return S_OK;}
  HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**) override{return S_OK;}
  HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR,BSTR* xml) override {
    if(!xml)return E_POINTER;
    static const wchar_t* ui =
      LR"(<?xml version="1.0" encoding="utf-8"?>)"
      LR"(<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">)"
      LR"(<ribbon><tabs><tab id="RadiologyAgentTab" label="Radiology Agent">)"
      LR"(<group id="RadiologyAgentGroup" label="Report Review">)"
      LR"(<button id="RadiologyAgentReview" label="Review Text" size="large" imageMso="ReviewAcceptChange" onAction="ReviewTextClick"/>)"
      LR"(</group></tab></tabs></ribbon></customUI>)";
    *xml=SysAllocString(ui);
    return *xml?S_OK:E_OUTOFMEMORY;
  }
};
class Factory : public IClassFactory {
  LONG refs_;
public:
  Factory():refs_(1){InterlockedIncrement(&g_objects);}
  ~Factory(){InterlockedDecrement(&g_objects);}
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid,void** ppv) override {
    if(!ppv)return E_POINTER;*ppv=nullptr;
    if(IsEqualIID(iid,IID_IUnknown)||IsEqualIID(iid,IID_IClassFactory))*ppv=static_cast<IClassFactory*>(this);
    else return E_NOINTERFACE;
    AddRef();return S_OK;
  }
  ULONG STDMETHODCALLTYPE AddRef() override{return static_cast<ULONG>(InterlockedIncrement(&refs_));}
  ULONG STDMETHODCALLTYPE Release() override{LONG r=InterlockedDecrement(&refs_);if(!r)delete this;return static_cast<ULONG>(r);}
  HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer,REFIID iid,void** out) override {
    if(outer)return CLASS_E_NOAGGREGATION;
    if(!out)return E_POINTER;
    Agent* obj=new(std::nothrow) Agent();
    if(!obj)return E_OUTOFMEMORY;
    HRESULT hr=obj->QueryInterface(iid,out);
    obj->Release();
    return hr;
  }
  HRESULT STDMETHODCALLTYPE LockServer(BOOL lock) override {
    if(lock)InterlockedIncrement(&g_locks);else InterlockedDecrement(&g_locks);
    return S_OK;
  }
};
extern "C" BOOL WINAPI DllMain(HINSTANCE h,DWORD reason,LPVOID) {
  if(reason==DLL_PROCESS_ATTACH){g_instance=h;DisableThreadLibraryCalls(h);}
  return TRUE;
}
extern "C" __declspec(dllexport) HRESULT STDAPICALLTYPE RADllCanUnloadNow() {
  return (g_objects==0&&g_locks==0)?S_OK:S_FALSE;
}
extern "C" __declspec(dllexport) HRESULT STDAPICALLTYPE RADllGetClassObject(REFCLSID clsid,REFIID iid,void** out) {
  if(!IsEqualGUID(clsid,CLSID_RA))return CLASS_E_CLASSNOTAVAILABLE;
  Factory* factory=new(std::nothrow) Factory();
  if(!factory)return E_OUTOFMEMORY;
  HRESULT hr=factory->QueryInterface(iid,out);
  factory->Release();
  return hr;
}
extern "C" __declspec(dllexport) HRESULT STDAPICALLTYPE DllRegisterServer() {
  wchar_t dll[MAX_PATH]={0};
  if(!GetModuleFileNameW(g_instance,dll,MAX_PATH))return HRESULT_FROM_WIN32(GetLastError());
  const std::wstring base=std::wstring(L"Software\\Classes\\CLSID\\")+CLSID_TEXT;
  HRESULT hr=RegistryValue(HKEY_CURRENT_USER,std::wstring(L"Software\\Classes\\")+PROGID,nullptr,L"Radiology Agent for Word 2013");
  if(FAILED(hr))return hr;
  hr=RegistryValue(HKEY_CURRENT_USER,std::wstring(L"Software\\Classes\\")+PROGID+L"\\CLSID",nullptr,CLSID_TEXT);if(FAILED(hr))return hr;
  hr=RegistryValue(HKEY_CURRENT_USER,base,nullptr,PROGID);if(FAILED(hr))return hr;
  hr=RegistryValue(HKEY_CURRENT_USER,base+L"\\InprocServer32",nullptr,dll);if(FAILED(hr))return hr;
  hr=RegistryValue(HKEY_CURRENT_USER,base+L"\\InprocServer32",L"ThreadingModel",L"Apartment");if(FAILED(hr))return hr;
  hr=RegistryValue(HKEY_CURRENT_USER,base+L"\\ProgID",nullptr,PROGID);if(FAILED(hr))return hr;
  const std::wstring word=std::wstring(L"Software\\Microsoft\\Office\\Word\\Addins\\")+PROGID;
  hr=RegistryValue(HKEY_CURRENT_USER,word,L"FriendlyName",L"Radiology Agent");if(FAILED(hr))return hr;
  hr=RegistryValue(HKEY_CURRENT_USER,word,L"Description",L"AI report proofreading. All suggestions require clinical verification.");if(FAILED(hr))return hr;
  return RegistryDword(HKEY_CURRENT_USER,word,L"LoadBehavior",3);
}
extern "C" __declspec(dllexport) HRESULT STDAPICALLTYPE DllUnregisterServer() {
  RegDeleteTreeW(HKEY_CURRENT_USER,(std::wstring(L"Software\\Classes\\CLSID\\")+CLSID_TEXT).c_str());
  RegDeleteTreeW(HKEY_CURRENT_USER,(std::wstring(L"Software\\Classes\\")+PROGID).c_str());
  RegDeleteTreeW(HKEY_CURRENT_USER,(std::wstring(L"Software\\Microsoft\\Office\\Word\\Addins\\")+PROGID).c_str());
  return S_OK;
}
