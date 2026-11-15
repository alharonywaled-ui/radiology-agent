#pragma once
// Haronex mandatory Word review gate: native ApplicationEvents4 save/print
// interception. No report text or patient identity is written to disk.
#include <windows.h>
#include <ole2.h>
#include <ocidl.h>
#include <wincrypt.h>
#include <string>

namespace HaronexMandatory {
static const IID WordEvents4 =
  {0x00020A01,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};

inline void HookStatus(DWORD value) {
  HKEY key=nullptr;
  if(RegCreateKeyExW(HKEY_CURRENT_USER,L"Software\\Haronex",0,nullptr,0,
                     KEY_SET_VALUE,nullptr,&key,nullptr)==ERROR_SUCCESS) {
    RegSetValueExW(key,L"MandatoryHookStatus",0,REG_DWORD,
                   reinterpret_cast<const BYTE*>(&value),sizeof(value));
    RegCloseKey(key);
  }
}
inline bool Prop(IDispatch* obj,const wchar_t* field,VARIANT* value) {
  if(!obj||!value)return false;
  VariantInit(value);
  LPOLESTR name=const_cast<LPOLESTR>(field);
  DISPID id=0;
  if(FAILED(obj->GetIDsOfNames(IID_NULL,&name,1,LOCALE_USER_DEFAULT,&id)))return false;
  DISPPARAMS args={nullptr,nullptr,0,0};
  return SUCCEEDED(obj->Invoke(id,IID_NULL,LOCALE_USER_DEFAULT,
                        DISPATCH_PROPERTYGET,&args,value,nullptr,nullptr));
}
inline bool AsDispatch(VARIANT* item,IDispatch** object) {
  *object=nullptr;
  if(item->vt==VT_DISPATCH&&item->pdispVal) {
    *object=item->pdispVal;(*object)->AddRef();return true;
  }
  if(item->vt==VT_UNKNOWN&&item->punkVal)
    return SUCCEEDED(item->punkVal->QueryInterface(IID_IDispatch,
                  reinterpret_cast<void**>(object)));
  return false;
}
inline bool ReadString(IDispatch* obj,const wchar_t* property,std::wstring& output) {
  VARIANT v; if(!Prop(obj,property,&v))return false;
  bool ok=v.vt==VT_BSTR&&v.bstrVal!=nullptr;
  if(ok)output.assign(v.bstrVal,SysStringLen(v.bstrVal));
  VariantClear(&v);return ok;
}
inline bool ReadDispatch(IDispatch* obj,const wchar_t* property,IDispatch** dest) {
  VARIANT v;if(!Prop(obj,property,&v))return false;
  bool ok=AsDispatch(&v,dest);
  VariantClear(&v);return ok;
}
inline bool ReadWindowHwnd(IDispatch* doc,DWORD& hwnd) {
  IDispatch* window=nullptr;
  if(!ReadDispatch(doc,L"ActiveWindow",&window))return false;
  VARIANT v;bool ok=Prop(window,L"Hwnd",&v);
  window->Release();
  if(!ok)return false;
  VARIANT num;VariantInit(&num);
  ok=SUCCEEDED(VariantChangeType(&num,&v,0,VT_I4));
  if(ok)hwnd=static_cast<DWORD>(num.lVal);
  VariantClear(&num);VariantClear(&v);
  return ok&&hwnd!=0;
}
inline bool ReadDocument(IDispatch* doc,std::wstring& name,
                         std::wstring& text,DWORD& hwnd) {
  if(!ReadString(doc,L"FullName",name))return false;
  if(!ReadWindowHwnd(doc,hwnd))return false;
  IDispatch* range=nullptr;
  if(!ReadDispatch(doc,L"Content",&range))return false;
  bool ok=ReadString(range,L"Text",text);
  range->Release();return ok;
}
inline std::wstring Hash(const std::wstring& source) {
  if(source.size()>250000)return L""; // fail closed on unexpected input sizes
  int n=WideCharToMultiByte(CP_UTF8,WC_ERR_INVALID_CHARS,source.c_str(),
                            static_cast<int>(source.size()),nullptr,0,nullptr,nullptr);
  if(n<=0)return L"";
  std::string bytes(static_cast<size_t>(n),'\0');
  if(WideCharToMultiByte(CP_UTF8,WC_ERR_INVALID_CHARS,source.c_str(),
                        static_cast<int>(source.size()),&bytes[0],n,nullptr,nullptr)!=n)
    return L"";
  HCRYPTPROV provider=0;HCRYPTHASH hash=0;
  BYTE digest[32]={0};DWORD length=sizeof(digest);
  if(!CryptAcquireContextW(&provider,nullptr,nullptr,PROV_RSA_AES,CRYPT_VERIFYCONTEXT))
    return L"";
  bool ok=CryptCreateHash(provider,CALG_SHA_256,0,0,&hash)!=0;
  if(ok)ok=CryptHashData(hash,reinterpret_cast<const BYTE*>(bytes.data()),
                        static_cast<DWORD>(bytes.size()),0)!=0;
  if(ok)ok=CryptGetHashParam(hash,HP_HASHVAL,digest,&length,0)!=0&&length==32;
  if(hash)CryptDestroyHash(hash);
  CryptReleaseContext(provider,0);
  if(!ok)return L"";
  static const wchar_t digits[]=L"0123456789abcdef";
  std::wstring out;out.reserve(64);
  for(BYTE x:digest){out+=digits[x>>4];out+=digits[x&15];}
  return out;
}
inline bool Approved(IDispatch* doc) {
  std::wstring name,text;
  DWORD hwnd=0;
  if(!ReadDocument(doc,name,text,hwnd)||text.empty())return false;
  std::wstring identity=std::to_wstring(hwnd)+L"\x001f"+name;
  std::wstring keyDigest=Hash(identity),textDigest=Hash(text);
  if(keyDigest.empty()||textDigest.empty())return false;
  std::wstring path=L"Software\\Haronex\\MandatoryReviews\\"+keyDigest;
  HKEY key=nullptr;
  if(RegOpenKeyExW(HKEY_CURRENT_USER,path.c_str(),0,KEY_QUERY_VALUE,&key)!=ERROR_SUCCESS)
    return false;
  wchar_t result[80]={0};DWORD type=0,size=sizeof(result);
  LONG status=RegQueryValueExW(key,L"ApprovedHash",nullptr,&type,
                              reinterpret_cast<BYTE*>(result),&size);
  RegCloseKey(key);
  return status==ERROR_SUCCESS&&type==REG_SZ&&textDigest==result;
}
inline void SetCancel(VARIANTARG& arg) {
  if(arg.vt==(VT_BYREF|VT_BOOL)&&arg.pboolVal)
    *arg.pboolVal=VARIANT_TRUE;
  else if(arg.vt==(VT_BYREF|VT_VARIANT)&&arg.pvarVal) {
    VARIANT* target=arg.pvarVal;
    if(target->vt==(VT_BYREF|VT_BOOL)&&target->pboolVal)
      *target->pboolVal=VARIANT_TRUE;
    else {VariantClear(target);target->vt=VT_BOOL;target->boolVal=VARIANT_TRUE;}
  }
}
inline void ShowBlocked() {
  MessageBoxW(nullptr,
      L"Haronex: Review Text must finish successfully before Save or Print.\n"
      L"Any change to the report text requires a new review.\n"
      L"Please click Review Text in the Radiology Agent ribbon.",
      L"Haronex Mandatory Review",MB_OK|MB_ICONWARNING|MB_TOPMOST);
}

// The COM event sink lives inside Word (same thread as Word's save/print event).
class Gate : public IDispatch {
  LONG refs_=1;
  IConnectionPoint* point_=nullptr;
  DWORD cookie_=0;
  DISPID save_=-1,print_=-1;
public:
  ~Gate(){Detach();}
  void Detach(){
    if(point_){if(cookie_)point_->Unadvise(cookie_);point_->Release();}
    point_=nullptr;cookie_=0;HookStatus(0);
  }
  bool Attach(IDispatch* app) {
    Detach();
    if(!app)return false;
    ITypeInfo* appInfo=nullptr;ITypeLib* typeLib=nullptr;
    UINT index=0;ITypeInfo* events=nullptr;
    bool types=SUCCEEDED(app->GetTypeInfo(0,LOCALE_USER_DEFAULT,&appInfo))&&appInfo;
    if(types)types=SUCCEEDED(appInfo->GetContainingTypeLib(&typeLib,&index))&&typeLib;
    if(types)types=SUCCEEDED(typeLib->GetTypeInfoOfGuid(WordEvents4,&events))&&events;
    if(types) {
      LPOLESTR saveName=const_cast<LPOLESTR>(L"DocumentBeforeSave");
      LPOLESTR printName=const_cast<LPOLESTR>(L"DocumentBeforePrint");
      types=SUCCEEDED(events->GetIDsOfNames(&saveName,1,&save_))&&
            SUCCEEDED(events->GetIDsOfNames(&printName,1,&print_));
    }
    if(events)events->Release();
    if(typeLib)typeLib->Release();
    if(appInfo)appInfo->Release();
    if(!types)return false;
    IConnectionPointContainer* container=nullptr;
    if(FAILED(app->QueryInterface(IID_IConnectionPointContainer,
                    reinterpret_cast<void**>(&container))))return false;
    HRESULT hr=container->FindConnectionPoint(WordEvents4,&point_);
    container->Release();
    if(FAILED(hr)||!point_){point_=nullptr;return false;}
    hr=point_->Advise(static_cast<IUnknown*>(this),&cookie_);
    if(FAILED(hr)){point_->Release();point_=nullptr;return false;}
    HookStatus(1);
    return true;
  }
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID id,void** dest) override {
    if(!dest)return E_POINTER;*dest=nullptr;
    if(IsEqualIID(id,IID_IUnknown)||IsEqualIID(id,IID_IDispatch)||
       IsEqualIID(id,WordEvents4)){*dest=static_cast<IDispatch*>(this);AddRef();return S_OK;}
    return E_NOINTERFACE;
  }
  ULONG STDMETHODCALLTYPE AddRef() override {return static_cast<ULONG>(InterlockedIncrement(&refs_));}
  ULONG STDMETHODCALLTYPE Release() override {
    LONG n=InterlockedDecrement(&refs_);
    if(!n)delete this;
    return static_cast<ULONG>(n);
  }
  HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* n) override {
    if(!n)return E_POINTER;*n=0;return S_OK;
  }
  HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT,LCID,ITypeInfo**) override {return E_NOTIMPL;}
  HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID,LPOLESTR*,UINT,LCID,DISPID*) override {
    return DISP_E_UNKNOWNNAME;
  }
  HRESULT STDMETHODCALLTYPE Invoke(DISPID id,REFIID,LCID,WORD,DISPPARAMS* args,
                                   VARIANT*,EXCEPINFO*,UINT*) override {
    const bool saving=id==save_,printing=id==print_;
    if(!saving&&!printing)return S_OK;
    UINT needed=saving?3:2;
    if(!args||args->cArgs<needed)return S_OK;
    VARIANTARG& cancel=args->rgvarg[0];
    VARIANTARG& document=args->rgvarg[saving?2:1];
    IDispatch* doc=nullptr;
    bool pass=AsDispatch(&document,&doc)&&Approved(doc);
    if(doc)doc->Release();
    if(!pass){SetCancel(cancel);ShowBlocked();}
    return S_OK;
  }
};
} // namespace HaronexMandatory
