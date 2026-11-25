#pragma once
// Haronex mandatory Word review gate: native ApplicationEvents4 save/print
// interception. No report text or patient identity is written to disk.
#include <windows.h>
#include <ole2.h>
#include <ocidl.h>
#include <wincrypt.h>
#include <string>
#include <vector>

namespace HaronexMandatory {
static const IID WordEvents4 =
  {0x00020A01,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};

// Enterprise settings are written by elevated Setup in the 32-bit machine
// registry. When a policy value exists (including 0), it takes precedence
// over the older per-user preference. Ordinary installs retain HKCU behavior.
inline bool ReadPolicyDword(HKEY root,const wchar_t* path,const wchar_t* name,DWORD& value) {
  HKEY key=nullptr; DWORD size=sizeof(value),type=0;
  if(RegOpenKeyExW(root,path,0,KEY_QUERY_VALUE|KEY_WOW64_32KEY,&key)!=ERROR_SUCCESS)return false;
  LONG status=RegQueryValueExW(key,name,nullptr,&type,reinterpret_cast<BYTE*>(&value),&size);
  RegCloseKey(key);
  return status==ERROR_SUCCESS&&type==REG_DWORD&&size==sizeof(value);
}
inline bool SettingEnabled(const wchar_t* name) {
  DWORD value=0;
  if(ReadPolicyDword(HKEY_LOCAL_MACHINE,L"Software\\Haronex\\Enterprise",name,value))return value==1;
  return ReadPolicyDword(HKEY_CURRENT_USER,L"Software\\Haronex",name,value)&&value==1;
}
inline void DiagnosticValue(const wchar_t* name,DWORD value) {
  HKEY key=nullptr;
  if(RegCreateKeyExW(HKEY_CURRENT_USER,L"Software\\Haronex",0,nullptr,0,
                     KEY_SET_VALUE,nullptr,&key,nullptr)==ERROR_SUCCESS) {
    RegSetValueExW(key,name,0,REG_DWORD,
      reinterpret_cast<const BYTE*>(&value),sizeof(value));
    RegCloseKey(key);
  }
}
// Only numeric event metadata, never report content or patient information.
inline void BeginDiagnostics() {
  DiagnosticValue(L"SaveEventCount",0);
  DiagnosticValue(L"SaveAllowedCount",0);
  DiagnosticValue(L"SaveBlockedCount",0);
  DiagnosticValue(L"SaveEmergencyCount",0);
  DiagnosticValue(L"PrintAllowedCount",0);
  DiagnosticValue(L"PrintBlockedCount",0);
  DiagnosticValue(L"ClearBothUnlockCount",0);
  DiagnosticValue(L"SuccessfulReviewUnlockCount",0);
  DiagnosticValue(L"MalformedEventCount",0);
  DiagnosticValue(L"MissingDocumentCount",0);
  DiagnosticValue(L"PrintEventCount",0);
  DiagnosticValue(L"OtherEventCount",0);
  DiagnosticValue(L"LastEventDispId",0);
  DiagnosticValue(L"LastEventArgCount",0);
  DiagnosticValue(L"LastCancelArgumentType",0);
  DiagnosticValue(L"LastArg0Type",0);
  DiagnosticValue(L"LastArg1Type",0);
  DiagnosticValue(L"LastArg2Type",0);
  DiagnosticValue(L"SaveArgCount",0);
  DiagnosticValue(L"SaveArg0Type",0);
  DiagnosticValue(L"SaveArg1Type",0);
  DiagnosticValue(L"SaveArg2Type",0);
  DiagnosticValue(L"SaveEventDispId",0);
  DiagnosticValue(L"PrintEventDispId",0);
}
inline void IncrementDiagnostic(const wchar_t* name) {
  HKEY key=nullptr;
  if(RegCreateKeyExW(HKEY_CURRENT_USER,L"Software\\Haronex",0,nullptr,0,
                     KEY_QUERY_VALUE|KEY_SET_VALUE,nullptr,&key,nullptr)!=ERROR_SUCCESS)return;
  DWORD count=0,sz=sizeof(count),type=0;
  RegQueryValueExW(key,name,nullptr,&type,reinterpret_cast<BYTE*>(&count),&sz);
  if(count<0xFFFFFFFE)++count;
  RegSetValueExW(key,name,0,REG_DWORD,
            reinterpret_cast<const BYTE*>(&count),sizeof(count));
  RegCloseKey(key);
}
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
  if(item->vt==(VT_BYREF|VT_DISPATCH)&&item->ppdispVal&&*item->ppdispVal) {
    *object=*item->ppdispVal;(*object)->AddRef();return true;
  }
  if(item->vt==(VT_BYREF|VT_VARIANT)&&item->pvarVal)
    return AsDispatch(item->pvarVal,object);
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
// Native creates a unique request token for the precise open COM document.
// The managed host returns a receipt only after a successful, verified review.
// Neither an old report hash nor Clear Both grants permission.
inline bool DocumentRequestIdentity(IDispatch* doc,std::wstring& identity) {
  std::wstring name;
  if(!ReadString(doc,L"FullName",name)||name.empty())return false;
  identity=Hash(std::to_wstring(GetCurrentProcessId())+L"\x001f"+name);
  return !identity.empty();
}
inline bool WriteReviewString(HKEY key,const wchar_t* name,const std::wstring& value) {
  return RegSetValueExW(key,name,0,REG_SZ,reinterpret_cast<const BYTE*>(value.c_str()),
    static_cast<DWORD>((value.size()+1)*sizeof(wchar_t)))==ERROR_SUCCESS;
}
inline bool ReviewReceiptMatches(const std::wstring& token,const std::wstring& identity) {
  HKEY key=nullptr;
  if(token.empty()||identity.empty()||RegOpenKeyExW(HKEY_CURRENT_USER,
      L"Software\\Haronex\\ReviewReceipts",0,KEY_QUERY_VALUE,&key)!=ERROR_SUCCESS)return false;
  wchar_t value[80]={0};DWORD type=0,size=sizeof(value);
  LONG status=RegQueryValueExW(key,token.c_str(),nullptr,&type,
    reinterpret_cast<BYTE*>(value),&size);
  RegCloseKey(key);
  return status==ERROR_SUCCESS&&type==REG_SZ&&size<=sizeof(value)&&value[79]==0&&identity==value;
}
inline void DeleteReviewReceipt(const std::wstring& token) {
  HKEY key=nullptr;
  if(RegOpenKeyExW(HKEY_CURRENT_USER,L"Software\\Haronex\\ReviewReceipts",0,
      KEY_SET_VALUE,&key)==ERROR_SUCCESS) {
    RegDeleteValueW(key,token.c_str());RegCloseKey(key);
  }
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
// A successful review unlocks the current open document; cleanup is optional.
// The COM event sink lives inside Word (same thread as Word's save/print event).
class Gate : public IDispatch {
  LONG refs_=1;
  IConnectionPoint* point_=nullptr;
  DWORD cookie_=0;
  DISPID save_=-1,print_=-1;
  // A retained COM identity is unique to this open document. No filename
  // fallback: a different/reopened document at the same path must be reviewed.
  struct PendingReview { IUnknown* doc; std::wstring token; std::wstring identity; };
  std::vector<IUnknown*> released_;
  std::vector<PendingReview> pending_;
  static IUnknown* Canonical(IDispatch* doc) {
    IUnknown* identity=nullptr;
    if(!doc||FAILED(doc->QueryInterface(IID_IUnknown,reinterpret_cast<void**>(&identity))))return nullptr;
    return identity;
  }
 public:
  bool IsReleased(IDispatch* doc) const {
    IUnknown* identity=Canonical(doc);if(!identity)return false;
    bool yes=false;
    for(auto* item:released_)if(item==identity){yes=true;break;}
    identity->Release();return yes;
  }
  bool RememberReleasedDocument(IDispatch* doc) {
    if(IsReleased(doc))return true;
    IUnknown* identity=Canonical(doc);if(!identity)return false;
    released_.push_back(identity);return true;
  }
  bool TryReleaseReviewedDocument(IDispatch* doc) {
    if(IsReleased(doc))return true;
    IUnknown* identity=Canonical(doc);if(!identity)return false;
    bool released=false;
    for(auto it=pending_.begin();it!=pending_.end();++it) {
      if(it->doc!=identity)continue;
      if(ReviewReceiptMatches(it->token,it->identity)&&RememberReleasedDocument(doc)) {
        DeleteReviewReceipt(it->token);
        it->doc->Release();pending_.erase(it);
        IncrementDiagnostic(L"SuccessfulReviewUnlockCount");released=true;
      }
      break;
    }
    identity->Release();return released;
  }
  bool BeginReview(IDispatch* doc,DWORD command,HKEY request) {
    if(!request||!command)return false;
    // Preserve a previously completed review if a second click occurs.
    TryReleaseReviewedDocument(doc);
    IUnknown* identity=Canonical(doc);if(!identity)return false;
    std::wstring binding;
    GUID guid={};wchar_t token[40]={0};
    if(!DocumentRequestIdentity(doc,binding)||FAILED(CoCreateGuid(&guid))||
       !StringFromGUID2(guid,token,40)) {identity->Release();return false;}
    if(!WriteReviewString(request,L"ReviewRequestToken",token)||
       !WriteReviewString(request,L"ReviewRequestIdentity",binding)||
       RegSetValueExW(request,L"ReviewRequestCommand",0,REG_DWORD,
         reinterpret_cast<const BYTE*>(&command),sizeof(command))!=ERROR_SUCCESS) {
      identity->Release();return false;
    }
    for(auto it=pending_.begin();it!=pending_.end();) {
      if(it->doc==identity) {
        DeleteReviewReceipt(it->token);it->doc->Release();it=pending_.erase(it);
      } else ++it;
    }
    pending_.push_back({identity,token,binding});
    return true;
  }
  ~Gate(){Detach();}
  void Detach(){
    for(auto* item:released_)if(item)item->Release();
    released_.clear();
    for(auto& item:pending_){DeleteReviewReceipt(item.token);if(item.doc)item.doc->Release();}
    pending_.clear();
    if(point_){if(cookie_)point_->Unadvise(cookie_);point_->Release();}
    point_=nullptr;cookie_=0;HookStatus(0);
  }
  bool Attach(IDispatch* app) {
    Detach();
    BeginDiagnostics();
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
    DiagnosticValue(L"SaveEventDispId",static_cast<DWORD>(save_));
    DiagnosticValue(L"PrintEventDispId",static_cast<DWORD>(print_));
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
  void ConfigureEventIdsForSmokeTest(DISPID saveEvent,DISPID printEvent){
    if(!point_){save_=saveEvent;print_=printEvent;}
  }
  // The safety policy can be tested without a Word installation.
  // Native Save, Save As and Print require a successful review for this
  // open document when enabled. Clear Both never grants permission.
  static bool ShouldCancel(bool unlocked,bool,bool) {
    return EnforcementEnabled()&&!unlocked;
  }
  static bool EnforcementEnabled(){return SettingEnabled(L"MandatoryReviewEnabled");}
  static bool ValidCancel(VARIANTARG& arg) {
    if(arg.vt==(VT_BYREF|VT_BOOL)&&arg.pboolVal)return true;
    if(arg.vt==(VT_BYREF|VT_VARIANT)&&arg.pvarVal) {
      VARIANT* nested=arg.pvarVal;
      return nested->vt==(VT_BYREF|VT_BOOL)&&nested->pboolVal;
    }
    return false;
  }
  // The actual Word 2013 Save event is [Document, SaveAsUI, Cancel]
  // on the user's 32-bit COM sink: VT_DISPATCH, 0x400B, 0x400B.
  // The two BYREF BOOLs are NOT ambiguous: the final one is Cancel.
  // Also support a conventional reversed COM argument vector, explicitly.
  static bool IdentifyEventArguments(DISPPARAMS* args,UINT expected,
                                     VARIANTARG** cancel,IDispatch** doc) {
    if(!cancel||!doc)return false;
    *cancel=nullptr;*doc=nullptr;
    if(!args||!args->rgvarg||args->cArgs!=expected)return false;
    VARIANTARG* v=args->rgvarg;
    const bool forward=AsDispatch(&v[0],doc);
    if(forward) {
      const bool valid=expected==3 ?
          (ValidCancel(v[1])&&ValidCancel(v[2])) :
          ValidCancel(v[1]);
      if(valid) {*cancel=&v[expected-1];return true;}
      (*doc)->Release();*doc=nullptr;
      return false;
    }
    // Reversed arguments: [Cancel, SaveAsUI, Document] on Save, or
    // [Cancel, Document] on Print. Treat any other shape as malformed.
    if(!AsDispatch(&v[expected-1],doc))return false;
    const bool valid=expected==3 ?
        (ValidCancel(v[0])&&ValidCancel(v[1])) :
        ValidCancel(v[0]);
    if(valid) {*cancel=&v[0];return true;}
    (*doc)->Release();*doc=nullptr;
    return false;
  }
  HRESULT STDMETHODCALLTYPE Invoke(DISPID dispid,REFIID,LCID,WORD,
                                   DISPPARAMS* args,VARIANT*,EXCEPINFO*,UINT*) override {
    DiagnosticValue(L"LastEventDispId",static_cast<DWORD>(dispid));
    DiagnosticValue(L"LastEventArgCount",args?args->cArgs:0);
    DiagnosticValue(L"LastCancelArgumentType",
      args&&args->rgvarg&&args->cArgs?args->rgvarg[0].vt:0);
    // Record only COM argument types, not any contents of the report.
    DiagnosticValue(L"LastArg0Type",args&&args->rgvarg&&args->cArgs>0?args->rgvarg[0].vt:0);
    DiagnosticValue(L"LastArg1Type",args&&args->rgvarg&&args->cArgs>1?args->rgvarg[1].vt:0);
    DiagnosticValue(L"LastArg2Type",args&&args->rgvarg&&args->cArgs>2?args->rgvarg[2].vt:0);
    const bool saving=dispid==save_,printing=dispid==print_;
    if((saving||printing)&&!EnforcementEnabled()) {
      IncrementDiagnostic(saving?L"SaveAllowedCount":L"PrintAllowedCount");
      return S_OK;
    }
    if(saving) {
      DiagnosticValue(L"SaveArgCount",args?args->cArgs:0);
      DiagnosticValue(L"SaveArg0Type",args&&args->rgvarg&&args->cArgs>0?args->rgvarg[0].vt:0);
      DiagnosticValue(L"SaveArg1Type",args&&args->rgvarg&&args->cArgs>1?args->rgvarg[1].vt:0);
      DiagnosticValue(L"SaveArg2Type",args&&args->rgvarg&&args->cArgs>2?args->rgvarg[2].vt:0);
      IncrementDiagnostic(L"SaveEventCount");
    }
    else if(printing)IncrementDiagnostic(L"PrintEventCount");
    else {IncrementDiagnostic(L"OtherEventCount");return S_OK;}
    VARIANTARG* cancel=nullptr;
    IDispatch* document=nullptr;
    if(!IdentifyEventArguments(args,saving?3:2,&cancel,&document)) {
      IncrementDiagnostic(L"MalformedEventCount");
      return S_OK;
    }
    bool unlocked=TryReleaseReviewedDocument(document);
    document->Release();
    if(unlocked) {
      IncrementDiagnostic(saving?L"SaveAllowedCount":L"PrintAllowedCount");
      return S_OK;
    }
    if(saving) {
      SetCancel(*cancel);
      IncrementDiagnostic(L"SaveBlockedCount");
      MessageBoxW(nullptr,
        L"Saving this report has been locked. Run Review Text in the "
        L"Radiology Agent tab and wait for successful completion. "
        L"Clear Both is optional.",
        L"Haronex - Save locked",MB_OK|MB_ICONINFORMATION);
    } else {
      SetCancel(*cancel);
      IncrementDiagnostic(L"PrintBlockedCount");
      MessageBoxW(nullptr,
        L"Printing remains locked until Review Text completes successfully "
        L"on the Radiology Agent tab. Clear Both is optional.",
        L"Haronex - Print Locked",MB_OK|MB_ICONINFORMATION);
    }
    return S_OK;
  }
};
} // namespace HaronexMandatory
