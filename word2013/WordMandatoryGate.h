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
// Approval scope is the exact Word document/window + exact document text.
// A Clear Both acceptance is explicitly distinguished from a Groq review.
inline bool ApprovalIdentity(IDispatch* doc,std::wstring& registryPath,
                            std::wstring& textDigest) {
  // Document.ActiveWindow.Hwnd is unreliable and caused false Save/Print
  // rejections. Bind the initial review to the Word process and FullName,
  // validating the exact current text on the Review/Clear Both handshake.
  std::wstring name,text;
  if(!ReadString(doc,L"FullName",name)||name.empty())return false;
  IDispatch* range=nullptr;
  if(!ReadDispatch(doc,L"Content",&range))return false;
  bool read=ReadString(range,L"Text",text);
  range->Release();
  if(!read||text.empty())return false;
  std::wstring identity=std::to_wstring(GetCurrentProcessId())+L"\x001f"+name;
  std::wstring keyDigest=Hash(identity);
  textDigest=Hash(text);
  if(keyDigest.empty()||textDigest.empty())return false;
  registryPath=L"Software\\Haronex\\MandatoryReviews\\"+keyDigest;
  return true;
}
inline bool MatchingApproval(HKEY key,const wchar_t* valueName,
                             const std::wstring& textDigest) {
  wchar_t result[80]={0};
  DWORD type=0,size=sizeof(result);
  LONG status=RegQueryValueExW(key,valueName,nullptr,&type,
                               reinterpret_cast<BYTE*>(result),&size);
  return status==ERROR_SUCCESS&&type==REG_SZ&&textDigest==result;
}
// A review alone is not the final release: the user also acknowledges the
// displayed suggestions by pressing Clear Both. No suggested edit is required.
inline bool ApprovalReady(HKEY key,const std::wstring& digest) {
  return MatchingApproval(key,L"ApprovedHash",digest)&&
         MatchingApproval(key,L"ClearBothV103Hash",digest);
}
inline bool HasReviewedHash(IDispatch* doc) {
  std::wstring path,digest;
  if(!ApprovalIdentity(doc,path,digest))return false;
  HKEY key=nullptr;
  if(RegOpenKeyExW(HKEY_CURRENT_USER,path.c_str(),0,KEY_QUERY_VALUE,&key)!=ERROR_SUCCESS)
    return false;
  bool reviewed=MatchingApproval(key,L"ApprovedHash",digest);
  RegCloseKey(key);
  return reviewed;
}
inline bool Approved(IDispatch* doc) {
  std::wstring path,digest;
  if(!ApprovalIdentity(doc,path,digest))return false;
  HKEY key=nullptr;
  if(RegOpenKeyExW(HKEY_CURRENT_USER,path.c_str(),0,KEY_QUERY_VALUE,&key)!=ERROR_SUCCESS)
    return false;
  // BOTH gates must match the current revision. An old V10.1 ClearBothHash
  // alone can no longer bypass Review Text.
  bool pass=ApprovalReady(key,digest);
  RegCloseKey(key);
  return pass;
}
inline bool AcceptClearBoth(IDispatch* doc) {
  std::wstring path,digest;
  if(!ApprovalIdentity(doc,path,digest))return false;
  HKEY key=nullptr;
  if(RegOpenKeyExW(HKEY_CURRENT_USER,path.c_str(),0,KEY_QUERY_VALUE|KEY_SET_VALUE,
                   &key)!=ERROR_SUCCESS)return false;
  // Never grant access unless a successful Review Text approval for this
  // exact content has been saved. Cleanup is not itself an AI review.
  if(!MatchingApproval(key,L"ApprovedHash",digest)) {
    RegCloseKey(key);
    return false;
  }
  LONG status=RegSetValueExW(key,L"ClearBothV103Hash",0,REG_SZ,
      reinterpret_cast<const BYTE*>(digest.c_str()),
      static_cast<DWORD>((digest.size()+1)*sizeof(wchar_t)));
  RegCloseKey(key);
  return status==ERROR_SUCCESS;
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
      L"Haronex: Complete Review Text, then click Clear Both before Save or Print.\n"
      L"Corrections are optional. After Review Text and Clear Both, editing is allowed.\n"
      L"Once unlocked, this open document can be saved and printed normally.",
      L"Haronex Mandatory Review",MB_OK|MB_ICONWARNING|MB_TOPMOST);
}

// The COM event sink lives inside Word (same thread as Word's save/print event).
class Gate : public IDispatch {
  LONG refs_=1;
  IConnectionPoint* point_=nullptr;
  DWORD cookie_=0;
  DISPID save_=-1,print_=-1;
  // Canonical COM IUnknown identity is stable across document edits and
  // Save As. These are in-memory, per-Word-instance approvals; never persisted
  // or shared with another open document.
  std::vector<IUnknown*> released_;
  std::vector<std::wstring> releasedNames_;
public:
  bool IsReleased(IDispatch* doc) const {
    if(!doc)return false;
    IUnknown* identity=nullptr;
    if(FAILED(doc->QueryInterface(IID_IUnknown,
              reinterpret_cast<void**>(&identity)))||!identity)return false;
    bool yes=false;
    for(auto* item:released_)if(item==identity){yes=true;break;}
    identity->Release();
    if(!yes) {
      std::wstring name;
      if(ReadString(doc,L"FullName",name)&&!name.empty()) {
        for(const auto& item:releasedNames_)
          if(_wcsicmp(item.c_str(),name.c_str())==0){yes=true;break;}
      }
    }
    return yes;
  }
  bool RememberReleasedDocument(IDispatch* doc) {
    if(!doc)return false;
    if(IsReleased(doc))return true;
    IUnknown* identity=nullptr;
    if(FAILED(doc->QueryInterface(IID_IUnknown,
              reinterpret_cast<void**>(&identity)))||!identity)return false;
    released_.push_back(identity);
    std::wstring name;
    if(ReadString(doc,L"FullName",name)&&!name.empty())
      releasedNames_.push_back(name);
    return true;
  }
  bool ReleaseAfterClearBoth(IDispatch* doc) {
    // Clear Both is the explicit unlock. Review Text is available, but never
    // a prerequisite that could trap the user without Save/Print.
    if(!doc)return false;
    bool released=RememberReleasedDocument(doc);
    if(released)IncrementDiagnostic(L"ClearBothUnlockCount");
    return released;
  }
  ~Gate(){Detach();}
  void Detach(){
    for(auto* item:released_)if(item)item->Release();
    released_.clear();
    releasedNames_.clear();
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
  // The Save dialog always offers a human-operated emergency escape.
  static bool ShouldCancel(bool unlocked,bool isSave,bool emergencySave) {
    return !unlocked && (!isSave || !emergencySave);
  }
  static bool EnforcementEnabled(){return true;}
  static bool ValidCancel(VARIANTARG& arg) {
    if(arg.vt==(VT_BYREF|VT_BOOL)&&arg.pboolVal)return true;
    if(arg.vt==(VT_BYREF|VT_VARIANT)&&arg.pvarVal) {
      VARIANT* nested=arg.pvarVal;
      return nested->vt==(VT_BYREF|VT_BOOL)&&nested->pboolVal;
    }
    return false;
  }
  // Word's actual COM event on the pilot PC passed VT_DISPATCH at argument
  // zero; the prior implementation falsely assumed reverse argument order.
  // Match arguments by COM VARTYPE instead, with exactly one Document and one
  // writable byref BOOL (Cancel). Ambiguous events fail open.
  static bool IdentifyEventArguments(DISPPARAMS* args,UINT expected,
                                     VARIANTARG** cancel,IDispatch** doc) {
    if(!cancel||!doc)return false;
    *cancel=nullptr;*doc=nullptr;
    if(!args||!args->rgvarg||args->cArgs!=expected)return false;
    for(UINT i=0;i<args->cArgs;++i) {
      VARIANTARG& arg=args->rgvarg[i];
      if(ValidCancel(arg)) {
        if(*cancel){if(*doc){(*doc)->Release();*doc=nullptr;}*cancel=nullptr;return false;}
        *cancel=&arg;
        continue;
      }
      // Only accept an actual dispatch for the Word Document, never infer
      // the identity from an unrelated flag or other event data.
      IDispatch* candidate=nullptr;
      if(AsDispatch(&arg,&candidate)) {
        if(*doc){candidate->Release();(*doc)->Release();*doc=nullptr;*cancel=nullptr;return false;}
        *doc=candidate;
      }
    }
    if(!*cancel||!*doc) {
      if(*doc){(*doc)->Release();*doc=nullptr;}
      *cancel=nullptr;
      return false;
    }
    return true;
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
    if(saving)IncrementDiagnostic(L"SaveEventCount");
    else if(printing)IncrementDiagnostic(L"PrintEventCount");
    else {IncrementDiagnostic(L"OtherEventCount");return S_OK;}
    VARIANTARG* cancel=nullptr;
    IDispatch* document=nullptr;
    if(!IdentifyEventArguments(args,saving?3:2,&cancel,&document)) {
      IncrementDiagnostic(L"MalformedEventCount");
      return S_OK;
    }
    bool unlocked=IsReleased(document);
    document->Release();
    if(unlocked) {
      IncrementDiagnostic(saving?L"SaveAllowedCount":L"PrintAllowedCount");
      return S_OK;
    }
    if(saving) {
      // Yes always preserves an otherwise unsaved report, even when Clear
      // Both was not clicked. No means go back and click Clear Both.
      const int choice=MessageBoxW(nullptr,
        L"Clear Both has not been clicked for this report.\n\n"
        L"YES: Save now to protect the report.\n"
        L"NO: Return to Word and click Clear Both.",
        L"Haronex - Save requires Clear Both",
        MB_YESNO|MB_ICONWARNING|MB_DEFBUTTON2|MB_TOPMOST);
      if(ShouldCancel(false,true,choice==IDYES)) {
        SetCancel(*cancel);
        IncrementDiagnostic(L"SaveBlockedCount");
      } else IncrementDiagnostic(L"SaveEmergencyCount");
    } else {
      SetCancel(*cancel);
      IncrementDiagnostic(L"PrintBlockedCount");
      MessageBoxW(nullptr,
        L"Printing is locked until you click Clear Both in the Radiology Agent tab.",
        L"Haronex - Print requires Clear Both",MB_OK|MB_ICONINFORMATION);
    }
    return S_OK;
  }
};
} // namespace HaronexMandatory
