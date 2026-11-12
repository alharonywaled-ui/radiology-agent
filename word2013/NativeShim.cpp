#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <ocidl.h>
#define DllRegisterServer OlectlDllRegisterServer
#define DllUnregisterServer OlectlDllUnregisterServer
#include <olectl.h>
#undef DllRegisterServer
#undef DllUnregisterServer
#include <shellapi.h>
#include <string>
#include <new>
#include <cwchar>
#include "WordLocalCleanup.h"

static const wchar_t* PROGID = L"RadiologyAgent.Word2013.Safe";
static const wchar_t* CLSID_TEXT = L"{7B918AD1-1C42-4B73-A4BB-9FC83743DD6E}";
static const GUID CLSID_RA = {0x7B918AD1,0x1C42,0x4B73,{0xA4,0xBB,0x9F,0xC8,0x37,0x43,0xDD,0x6E}};
static const GUID IID_EXT = {0xB65AD801,0xABAF,0x11D0,{0xBB,0x8B,0x00,0xA0,0xC9,0x0F,0x27,0x44}};
static const GUID IID_RIBBON = {0x000C0396,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static HMODULE g_instance = nullptr;
static LONG g_objects = 0;
static LONG g_locks = 0;
static IDispatch* g_ribbonUi = nullptr;
static UINT_PTR g_ribbonTimer = 0;
static DWORD g_lastReviewResult = 0xFFFFFFFF;
static DWORD ReadReviewResult() {
  DWORD value=0, size=sizeof(value);
  HKEY key=nullptr;
  if(RegOpenKeyExW(HKEY_CURRENT_USER,L"Software\\Haronex",0,KEY_QUERY_VALUE,&key)==ERROR_SUCCESS) {
    if(RegQueryValueExW(key,L"ReviewResult",nullptr,nullptr,reinterpret_cast<BYTE*>(&value),&size)!=ERROR_SUCCESS)value=0;
    RegCloseKey(key);
  }
  return value;
}
static void RefreshRibbonButton() {
  if(!g_ribbonUi)return;
  LPOLESTR method=const_cast<LPOLESTR>(L"InvalidateControl");
  DISPID dispid=0;
  if(FAILED(g_ribbonUi->GetIDsOfNames(IID_NULL,&method,1,LOCALE_USER_DEFAULT,&dispid)))return;
  VARIANTARG arg;VariantInit(&arg);
  arg.vt=VT_BSTR;arg.bstrVal=SysAllocString(L"RadiologyAgentReview");
  if(!arg.bstrVal)return;
  DISPPARAMS params={&arg,nullptr,1,0};
  g_ribbonUi->Invoke(dispid,IID_NULL,LOCALE_USER_DEFAULT,DISPATCH_METHOD,&params,nullptr,nullptr,nullptr);
  VariantClear(&arg);
}
static VOID CALLBACK RibbonStateTimer(HWND,UINT,UINT_PTR,DWORD) {
  DWORD state=ReadReviewResult();
  if(state!=g_lastReviewResult) {
    g_lastReviewResult=state;
    RefreshRibbonButton();
  }
}
static HRESULT CreateReviewIcon(DWORD state,VARIANT* out) {
  if(!out)return E_POINTER;
  HDC screen=GetDC(nullptr);
  if(!screen)return E_FAIL;
  HDC dc=CreateCompatibleDC(screen);
  HBITMAP bmp=CreateCompatibleBitmap(screen,32,32);
  ReleaseDC(nullptr,screen);
  if(!dc||!bmp){if(dc)DeleteDC(dc);if(bmp)DeleteObject(bmp);return E_OUTOFMEMORY;}
  HGDIOBJ old=SelectObject(dc,bmp);
  RECT area={0,0,32,32};
  HBRUSH white=CreateSolidBrush(RGB(255,255,255));
  FillRect(dc,&area,white);DeleteObject(white);
  COLORREF color=state==1?RGB(24,160,75):(state==2?RGB(210,45,45):RGB(30,100,200));
  HBRUSH brush=CreateSolidBrush(color);
  HGDIOBJ oldBrush=SelectObject(dc,brush);
  HPEN pen=CreatePen(PS_SOLID,1,color);
  HGDIOBJ oldPen=SelectObject(dc,pen);
  Ellipse(dc,2,2,30,30);
  SelectObject(dc,oldPen);DeleteObject(pen);
  SelectObject(dc,oldBrush);DeleteObject(brush);
  HPEN markPen=CreatePen(PS_SOLID,3,RGB(255,255,255));
  oldPen=SelectObject(dc,markPen);
  if(state==2){
    MoveToEx(dc,10,10,nullptr);LineTo(dc,22,22);
    MoveToEx(dc,22,10,nullptr);LineTo(dc,10,22);
  }else{
    MoveToEx(dc,9,16,nullptr);LineTo(dc,14,22);
    MoveToEx(dc,14,22,nullptr);LineTo(dc,24,10);
  }
  SelectObject(dc,oldPen);DeleteObject(markPen);
  SelectObject(dc,old);DeleteDC(dc);
  PICTDESC desc={};desc.cbSizeofstruct=sizeof(desc);desc.picType=PICTYPE_BITMAP;
  desc.bmp.hbitmap=bmp;desc.bmp.hpal=nullptr;
  IPictureDisp* picture=nullptr;
  HRESULT hr=OleCreatePictureIndirect(&desc,IID_IPictureDisp,TRUE,reinterpret_cast<void**>(&picture));
  if(FAILED(hr)){DeleteObject(bmp);return hr;}
  VariantInit(out);out->vt=VT_DISPATCH;out->pdispVal=picture;
  return S_OK;
}


static HRESULT CreateCleanupIcon(bool comments, VARIANT* out) {
  if(!out)return E_POINTER;
  HDC screen=GetDC(nullptr);
  if(!screen)return E_FAIL;
  HDC dc=CreateCompatibleDC(screen);
  HBITMAP bmp=CreateCompatibleBitmap(screen,32,32);
  ReleaseDC(nullptr,screen);
  if(!dc||!bmp){if(dc)DeleteDC(dc);if(bmp)DeleteObject(bmp);return E_OUTOFMEMORY;}
  HGDIOBJ oldBmp=SelectObject(dc,bmp);
  RECT all={0,0,32,32};HBRUSH bg=CreateSolidBrush(RGB(255,255,255));
  FillRect(dc,&all,bg);DeleteObject(bg);
  COLORREF accent=comments?RGB(207,55,63):RGB(240,181,36);
  HBRUSH brush=CreateSolidBrush(accent);
  HGDIOBJ oldBrush=SelectObject(dc,brush);
  HPEN outline=CreatePen(PS_SOLID,1,comments?RGB(161,35,41):RGB(173,114,12));
  HGDIOBJ oldPen=SelectObject(dc,outline);
  RoundRect(dc,3,4,29,27,7,7);
  SelectObject(dc,oldPen);DeleteObject(outline);
  SelectObject(dc,oldBrush);DeleteObject(brush);
  HPEN stroke=CreatePen(PS_SOLID,3,comments?RGB(255,255,255):RGB(55,55,55));
  oldPen=SelectObject(dc,stroke);
  if(comments){
    MoveToEx(dc,11,11,nullptr);LineTo(dc,21,20);
    MoveToEx(dc,21,11,nullptr);LineTo(dc,11,20);
  }else{
    MoveToEx(dc,8,18,nullptr);LineTo(dc,24,18);
    MoveToEx(dc,8,22,nullptr);LineTo(dc,24,22);
    MoveToEx(dc,10,10,nullptr);LineTo(dc,22,24);
  }
  SelectObject(dc,oldPen);DeleteObject(stroke);
  SelectObject(dc,oldBmp);DeleteDC(dc);
  PICTDESC desc={};desc.cbSizeofstruct=sizeof(desc);desc.picType=PICTYPE_BITMAP;
  desc.bmp.hbitmap=bmp;desc.bmp.hpal=nullptr;
  IPictureDisp* picture=nullptr;
  HRESULT hr=OleCreatePictureIndirect(&desc,IID_IPictureDisp,TRUE,
      reinterpret_cast<void**>(&picture));
  if(FAILED(hr)){DeleteObject(bmp);return hr;}
  VariantInit(out);out->vt=VT_DISPATCH;out->pdispVal=picture;
  return S_OK;
}


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
static void StartLiveMonitor() {
  std::wstring exe=ModuleDirectory()+L"RadiologyAgent.ReviewHost.exe";
  if(GetFileAttributesW(exe.c_str())==INVALID_FILE_ATTRIBUTES)return;
  // The C# monitor enforces a per-Windows-user singleton and talks to Word on its STA.
  ShellExecuteW(nullptr,L"open",exe.c_str(),L"--monitor",nullptr,SW_HIDE);
}
static void RequestReview() {
  const wchar_t* key=L"Software\\Haronex";
  HKEY h=nullptr;
  if(RegCreateKeyExW(HKEY_CURRENT_USER,key,0,nullptr,REG_OPTION_NON_VOLATILE,
                     KEY_QUERY_VALUE|KEY_SET_VALUE,nullptr,&h,nullptr)!=ERROR_SUCCESS)
    return;
  // Record the exact Word window that invoked the ribbon action. The monitor
  // must not blindly bind to a different Word.Application instance from ROT.
  HWND clicked=GetForegroundWindow();
  HWND root=clicked?GetAncestor(clicked,GA_ROOT):nullptr;
  DWORD windowHandle=static_cast<DWORD>(reinterpret_cast<ULONG_PTR>(root));
  RegSetValueExW(h,L"WordWindowHwnd",0,REG_DWORD,
    reinterpret_cast<const BYTE*>(&windowHandle),sizeof(windowHandle));
  DWORD seq=0,size=sizeof(seq),type=REG_DWORD;
  if(RegQueryValueExW(h,L"LiveReviewCommand",nullptr,&type,
                      reinterpret_cast<BYTE*>(&seq),&size)!=ERROR_SUCCESS)seq=0;
  ++seq;
  RegSetValueExW(h,L"LiveReviewCommand",0,REG_DWORD,
                 reinterpret_cast<const BYTE*>(&seq),sizeof(seq));
  RegCloseKey(h);
  StartLiveMonitor();
}

class Agent : public IExtensibility2, public IRibbonExtensibilityNative {
  LONG refs_;
  IDispatch* wordApplication_=nullptr;
public:
  Agent():refs_(1){InterlockedIncrement(&g_objects);}
  ~Agent(){if(wordApplication_)wordApplication_->Release();InterlockedDecrement(&g_objects);}
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
      else if(!_wcsicmp(n,L"RibbonLoaded")) ids[i]=101;
      else if(!_wcsicmp(n,L"ReviewIcon")) ids[i]=102;
      else if(!_wcsicmp(n,L"DeleteAllCommentsClick")) ids[i]=103;
      else if(!_wcsicmp(n,L"ClearRedYellowClick")) ids[i]=104;
      else if(!_wcsicmp(n,L"DeleteCommentsIcon")) ids[i]=105;
      else if(!_wcsicmp(n,L"ClearHighlightsIcon")) ids[i]=106;
      else return DISP_E_UNKNOWNNAME;
    }
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE Invoke(DISPID id,REFIID,LCID,WORD,DISPPARAMS* params,VARIANT* result,EXCEPINFO*,UINT*) override {
    if(result)VariantInit(result);
    if(id==100){RequestReview();return S_OK;}
    if(id==101) {
      if(params&&params->cArgs>=1) {
        VARIANTARG& arg=params->rgvarg[0];
        IDispatch* ui=arg.vt==VT_DISPATCH?arg.pdispVal:nullptr;
        if(ui) {
          if(g_ribbonUi)g_ribbonUi->Release();
          g_ribbonUi=ui;g_ribbonUi->AddRef();
          g_lastReviewResult=ReadReviewResult();
          if(!g_ribbonTimer)g_ribbonTimer=SetTimer(nullptr,0,500,RibbonStateTimer);
        }
      }
      return S_OK;
    }
    if(id==102)return CreateReviewIcon(ReadReviewResult(),result);
    if(id==103){LocalCleanup::DeleteAllComments(wordApplication_);return S_OK;}
    if(id==104){LocalCleanup::RemoveRedYellowHighlights(wordApplication_);return S_OK;}
    if(id==105)return CreateCleanupIcon(true,result);
    if(id==106)return CreateCleanupIcon(false,result);
    if(id>=1&&id<=5)return S_OK;
    return DISP_E_MEMBERNOTFOUND;
  }
  HRESULT STDMETHODCALLTYPE OnConnection(IDispatch* app,LONG,IDispatch*,SAFEARRAY**) override{
    if(wordApplication_)wordApplication_->Release();
    wordApplication_=app;
    if(wordApplication_)wordApplication_->AddRef();
    StartLiveMonitor();return S_OK;
  }
  HRESULT STDMETHODCALLTYPE OnDisconnection(LONG,SAFEARRAY**) override{
    if(wordApplication_){wordApplication_->Release();wordApplication_=nullptr;}
    if(g_ribbonTimer){KillTimer(nullptr,g_ribbonTimer);g_ribbonTimer=0;}
    if(g_ribbonUi){g_ribbonUi->Release();g_ribbonUi=nullptr;}
    return S_OK;
  }
  HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**) override{return S_OK;}
  HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**) override{StartLiveMonitor();return S_OK;}
  HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**) override{return S_OK;}
  HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR,BSTR* xml) override {
    if(!xml)return E_POINTER;
    static const wchar_t* ui =
      LR"(<?xml version="1.0" encoding="utf-8"?>)"
      LR"(<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="RibbonLoaded">)"
      LR"(<ribbon><tabs><tab id="RadiologyAgentTab" label="Radiology Agent">)"
      LR"(<group id="RadiologyAgentGroup" label="Report Review">)"
      LR"(<button id="RadiologyAgentReview" label="Review Text" size="large" getImage="ReviewIcon" onAction="ReviewTextClick"/>)"
      LR"(<button id="RadiologyAgentDeleteComments" label="Clear Comments" size="large" getImage="DeleteCommentsIcon" onAction="DeleteAllCommentsClick" screentip="Remove all comments" supertip="Delete every Word comment in the active document locally. No API call." />)"
      LR"(<button id="RadiologyAgentClearHighlights" label="Clear Highlights" size="large" getImage="ClearHighlightsIcon" onAction="ClearRedYellowClick" screentip="Remove red and yellow highlights" supertip="Remove red and yellow text highlighting from the active Word document, preserving other colors. Runs locally." />)"
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
