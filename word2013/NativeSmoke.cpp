#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <cwchar>
#include <cstdio>
#include "WordMandatoryGate.h"

struct IRibbon : IDispatch {
 virtual HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR,BSTR*)=0;
};
// A fake COM document exercises the native per-document release lifetime.
// No Word install or patient report is required on the CI runner.
class FakeDoc : public IDispatch {
  LONG references=1;
public:
  HRESULT STDMETHODCALLTYPE QueryInterface(REFIID iid,void** out) override {
    if(!out)return E_POINTER;*out=nullptr;
    if(IsEqualIID(iid,IID_IUnknown)||IsEqualIID(iid,IID_IDispatch)) {
      *out=static_cast<IDispatch*>(this);AddRef();return S_OK;
    }
    return E_NOINTERFACE;
  }
  ULONG STDMETHODCALLTYPE AddRef() override {return static_cast<ULONG>(InterlockedIncrement(&references));}
  ULONG STDMETHODCALLTYPE Release() override {
    auto count=InterlockedDecrement(&references);if(!count)delete this;
    return static_cast<ULONG>(count);
  }
  HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* count) override {
    if(!count)return E_POINTER;*count=0;return S_OK;
  }
  HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT,LCID,ITypeInfo**) override {return E_NOTIMPL;}
  HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID,LPOLESTR*,UINT,LCID,DISPID*) override {
    return DISP_E_UNKNOWNNAME;
  }
  HRESULT STDMETHODCALLTYPE Invoke(DISPID,REFIID,LCID,WORD,DISPPARAMS*,VARIANT*,EXCEPINFO*,UINT*) override {
    return DISP_E_MEMBERNOTFOUND;
  }
};
static const GUID clsid={0x7B918AD1,0x1C42,0x4B73,{0xA4,0xBB,0x9F,0xC8,0x37,0x43,0xDD,0x6E}};
static const GUID iid={0x000C0396,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
static bool SetMandatoryOption(DWORD enabled) {
  HKEY key=nullptr;
  if(RegCreateKeyExW(HKEY_CURRENT_USER,L"Software\\Haronex",0,nullptr,0,
      KEY_SET_VALUE,nullptr,&key,nullptr)!=ERROR_SUCCESS)return false;
  LONG status=RegSetValueExW(key,L"MandatoryReviewEnabled",0,REG_DWORD,
      reinterpret_cast<const BYTE*>(&enabled),sizeof(enabled));
  RegCloseKey(key);return status==ERROR_SUCCESS;
}
int main(){
  HRESULT hr=CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED);
  if(FAILED(hr))return 10;
  IRibbon* ribbon=nullptr;
  hr=CoCreateInstance(clsid,nullptr,CLSCTX_INPROC_SERVER,iid,(void**)&ribbon);
  if(FAILED(hr)){wprintf(L"COM_ACTIVATION_FAILED 0x%08x\n",static_cast<unsigned>(hr));CoUninitialize();return 20;}
  IDispatch* dispatch=nullptr;
  hr=ribbon->QueryInterface(IID_IDispatch,reinterpret_cast<void**>(&dispatch));
  if(FAILED(hr)||!dispatch){wprintf(L"RIBBON_IDISPATCH_FAILED 0x%08x\n",static_cast<unsigned>(hr));ribbon->Release();CoUninitialize();return 25;}
  dispatch->Release();
  BSTR id=SysAllocString(L"Microsoft.Word.Document");
  BSTR xml=nullptr;
  hr=ribbon->GetCustomUI(id,&xml);
  SysFreeString(id);
  if(FAILED(hr)||!xml||!wcsstr(xml,L"Radiology Agent")||!wcsstr(xml,L"ReviewTextClick")||
     !wcsstr(xml,L"DeleteAllCommentsClick")||!wcsstr(xml,L"ClearRedYellowClick")||
     !wcsstr(xml,L"Clear Comments")||!wcsstr(xml,L"Clear Highlights")||
     !wcsstr(xml,L"ClearBothClick")||!wcsstr(xml,L"Clear Both")||
     !wcsstr(xml,L"getImage=\"SidebarIcon\"")||
     !wcsstr(xml,L"ToggleSidebarClick")) {
    wprintf(L"RIBBON_CHECK_FAILED 0x%08x\n",static_cast<unsigned>(hr));
    if(xml)SysFreeString(xml);
    ribbon->Release();CoUninitialize();return 30;
  }
  // Verify the Office Ribbon custom icon callback is registered and produces
  // an actual OLE image, not a blank unsupported imageMso placeholder.
  LPOLESTR callback=const_cast<LPOLESTR>(L"SidebarIcon");
  DISPID iconId=0;
  if(FAILED(ribbon->GetIDsOfNames(IID_NULL,&callback,1,LOCALE_USER_DEFAULT,
                                  &iconId))) {
    wprintf(L"SIDEBAR_ICON_CALLBACK_MISSING\n");
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 31;
  }
  DISPPARAMS args={nullptr,nullptr,0,0};
  VARIANT picture;VariantInit(&picture);
  hr=ribbon->Invoke(iconId,IID_NULL,LOCALE_USER_DEFAULT,DISPATCH_METHOD,
                    &args,&picture,nullptr,nullptr);
  if(FAILED(hr)||picture.vt!=VT_DISPATCH||!picture.pdispVal) {
    wprintf(L"SIDEBAR_ICON_IMAGE_FAILED 0x%08x\n",static_cast<unsigned>(hr));
    VariantClear(&picture);SysFreeString(xml);
    ribbon->Release();CoUninitialize();return 32;
  }
  VariantClear(&picture);
  if(HaronexMandatory::Hash(L"Text")!=L"71988c4d8e0803ba4519f0b2864c1331c14a1890bf8694e251379177bfedb5c3") {
    wprintf(L"MANDATORY_SHA256_INCOMPATIBLE\n");
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 35;
  }
  // V10.3 requires initial review + Clear Both, but session release
  // survives all later text edits. An older V10.2 marker is not accepted.
  HKEY tempKey=nullptr;
  const wchar_t* testPath=L"Software\\Haronex\\MandatoryReviews\\NativeSmokeV10_3";
  if(RegCreateKeyExW(HKEY_CURRENT_USER,testPath,0,nullptr,0,
                     KEY_SET_VALUE|KEY_QUERY_VALUE,nullptr,&tempKey,nullptr)!=ERROR_SUCCESS) {
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 36;
  }
  std::wstring reviewedHash=HaronexMandatory::Hash(L"reviewed sample");
  std::wstring editedHash=HaronexMandatory::Hash(L"edited sample");
  auto setHash=[&](const wchar_t* name,const std::wstring& value) -> bool {
    return RegSetValueExW(tempKey,name,0,REG_SZ,
             reinterpret_cast<const BYTE*>(value.c_str()),
             static_cast<DWORD>((value.size()+1)*sizeof(wchar_t)))==ERROR_SUCCESS;
  };
  bool correct=!HaronexMandatory::ApprovalReady(tempKey,reviewedHash);
  correct=correct&&setHash(L"ApprovedHash",reviewedHash);
  correct=correct&&!HaronexMandatory::ApprovalReady(tempKey,reviewedHash);
  correct=correct&&setHash(L"ClearBothV102Hash",reviewedHash);
  correct=correct&&!HaronexMandatory::ApprovalReady(tempKey,reviewedHash);
  correct=correct&&setHash(L"ClearBothV103Hash",reviewedHash);
  correct=correct&&HaronexMandatory::ApprovalReady(tempKey,reviewedHash);
  correct=correct&&!HaronexMandatory::ApprovalReady(tempKey,editedHash);
  RegCloseKey(tempKey);
  RegDeleteKeyW(HKEY_CURRENT_USER,testPath);
  if(!correct) {
    wprintf(L"MANDATORY_V103_HANDSHAKE_TEST_FAILED\n");
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 37;
  }
  HaronexMandatory::Gate* session=new HaronexMandatory::Gate();
  FakeDoc* reviewed=new FakeDoc();
  FakeDoc* another=new FakeDoc();
  // No review yet => no release, and release is scoped to one COM document.
  bool allowed=!session->IsReleased(reviewed)&&!session->IsReleased(another);
  allowed=allowed&&session->RememberReleasedDocument(reviewed);
  allowed=allowed&&session->IsReleased(reviewed)&&!session->IsReleased(another);
  // A second action on the same document does not produce a second entry.
  allowed=allowed&&session->RememberReleasedDocument(reviewed);
  session->Detach();
  allowed=allowed&&!session->IsReleased(reviewed);
  reviewed->Release();another->Release();session->Release();
  if(!allowed) {
    wprintf(L"MANDATORY_SESSION_RELEASE_TEST_FAILED\n");
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 38;
  }
  // V10.9 strict policy: no Save Now bypass before Clear Both; after release
  // Save, Save As and Print remain enabled.
  bool policy=SetMandatoryOption(1) &&
    HaronexMandatory::Gate::EnforcementEnabled() &&
    HaronexMandatory::Gate::ShouldCancel(false,true,false) &&
    HaronexMandatory::Gate::ShouldCancel(false,false,false) &&
    HaronexMandatory::Gate::ShouldCancel(false,true,true) &&
    !HaronexMandatory::Gate::ShouldCancel(true,true,false) &&
    !HaronexMandatory::Gate::ShouldCancel(true,false,false);
  if(!policy) {
    wprintf(L"V109_STRICT_CLEAR_BOTH_POLICY_FAILED\n");
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 39;
  }
  HaronexMandatory::Gate gate;
  gate.ConfigureEventIdsForSmokeTest(101,102);
  FakeDoc* releasedForTest=new FakeDoc();
  FakeDoc* other=new FakeDoc();
  bool separated=gate.ReleaseAfterClearBoth(releasedForTest) &&
    gate.IsReleased(releasedForTest) && !gate.IsReleased(other);
  VARIANT_BOOL cancel=VARIANT_FALSE;
  VARIANTARG argv[3];for(int i=0;i<3;++i)VariantInit(&argv[i]);
  argv[0].vt=VT_BYREF|VT_BOOL;argv[0].pboolVal=&cancel;
  VARIANT_BOOL saveAsUI=VARIANT_FALSE;
  argv[1].vt=VT_BYREF|VT_BOOL;argv[1].pboolVal=&saveAsUI;
  argv[2].vt=VT_DISPATCH;argv[2].pdispVal=releasedForTest;
  DISPPARAMS saveArgs={argv,nullptr,3,0};
  bool saved=SUCCEEDED(gate.Invoke(101,IID_NULL,0,DISPATCH_METHOD,
                      &saveArgs,nullptr,nullptr,nullptr))&&cancel==VARIANT_FALSE;
  argv[1].vt=VT_DISPATCH;argv[1].pdispVal=releasedForTest;
  DISPPARAMS printArgs={argv,nullptr,2,0};
  bool printed=SUCCEEDED(gate.Invoke(102,IID_NULL,0,DISPATCH_METHOD,
                      &printArgs,nullptr,nullptr,nullptr))&&cancel==VARIANT_FALSE;
  // Observed live Word 2013: [VT_DISPATCH, 16395, 16395].
  // Exercise forward order as well as reverse order before releasing.
  VARIANT_BOOL alternateCancel=VARIANT_FALSE;
  VARIANTARG forwardSave[3];for(int i=0;i<3;++i)VariantInit(&forwardSave[i]);
  forwardSave[0].vt=VT_DISPATCH;forwardSave[0].pdispVal=releasedForTest;
  forwardSave[1].vt=VT_BYREF|VT_BOOL;forwardSave[1].pboolVal=&saveAsUI;
  forwardSave[2].vt=VT_BYREF|VT_BOOL;forwardSave[2].pboolVal=&alternateCancel;
  DISPPARAMS fSave={forwardSave,nullptr,3,0};
  VARIANTARG* foundCancel=nullptr;
  IDispatch* foundDoc=nullptr;
  bool identified=HaronexMandatory::Gate::IdentifyEventArguments(
      &fSave,3,&foundCancel,&foundDoc)&&foundCancel==&forwardSave[2]&&
      foundDoc!=nullptr;
  if(foundDoc){foundDoc->Release();foundDoc=nullptr;}
  bool correctCancelOnly=identified && foundCancel==&forwardSave[2] &&
    saveAsUI==VARIANT_FALSE;
  if(correctCancelOnly)HaronexMandatory::SetCancel(*foundCancel);
  correctCancelOnly=correctCancelOnly && alternateCancel==VARIANT_TRUE &&
    saveAsUI==VARIANT_FALSE;
  alternateCancel=VARIANT_FALSE;
  bool forwardSaved=identified &&
      SUCCEEDED(gate.Invoke(101,IID_NULL,0,DISPATCH_METHOD,
                           &fSave,nullptr,nullptr,nullptr)) &&
      alternateCancel==VARIANT_FALSE;
  VARIANTARG forwardPrint[2];for(int i=0;i<2;++i)VariantInit(&forwardPrint[i]);
  forwardPrint[0].vt=VT_DISPATCH;forwardPrint[0].pdispVal=releasedForTest;
  forwardPrint[1].vt=VT_BYREF|VT_BOOL;forwardPrint[1].pboolVal=&alternateCancel;
  DISPPARAMS fPrint={forwardPrint,nullptr,2,0};
  bool forwardPrinted=SUCCEEDED(gate.Invoke(102,IID_NULL,0,DISPATCH_METHOD,
                            &fPrint,nullptr,nullptr,nullptr))&&
                            alternateCancel==VARIANT_FALSE;
  // Unknown event argument must fail open (never cancel the report).
  VARIANTARG malformedArgs[3];for(int i=0;i<3;++i)VariantInit(&malformedArgs[i]);
  malformedArgs[0].vt=VT_DISPATCH;malformedArgs[0].pdispVal=releasedForTest;
  malformedArgs[1].vt=VT_BYREF|VT_BOOL;malformedArgs[1].pboolVal=&saveAsUI;
  malformedArgs[2].vt=VT_I4;malformedArgs[2].lVal=42;
  DISPPARAMS invalidArgs={malformedArgs,nullptr,3,0};
  bool invalidFailsOpen=SUCCEEDED(gate.Invoke(101,IID_NULL,0,DISPATCH_METHOD,
                              &invalidArgs,nullptr,nullptr,nullptr));
  DISPPARAMS malformed={nullptr,nullptr,0,0};
  bool safe=SUCCEEDED(gate.Invoke(101,IID_NULL,0,DISPATCH_METHOD,
                      &malformed,nullptr,nullptr,nullptr))&&cancel==VARIANT_FALSE;
  // Installer option OFF allows Save/Save As and Print. When ON, the
  // shared policy blocks both operations until the document is released.
  FakeDoc* lockedForTest=new FakeDoc();
  VARIANT_BOOL disabledSaveCancel=VARIANT_FALSE,disabledPrintCancel=VARIANT_FALSE;
  VARIANTARG disabledSave[3];for(int i=0;i<3;++i)VariantInit(&disabledSave[i]);
  disabledSave[0].vt=VT_DISPATCH;disabledSave[0].pdispVal=lockedForTest;
  disabledSave[1].vt=VT_BYREF|VT_BOOL;disabledSave[1].pboolVal=&saveAsUI;
  disabledSave[2].vt=VT_BYREF|VT_BOOL;disabledSave[2].pboolVal=&disabledSaveCancel;
  DISPPARAMS disabledSaveArgs={disabledSave,nullptr,3,0};
  VARIANTARG disabledPrint[2];for(int i=0;i<2;++i)VariantInit(&disabledPrint[i]);
  disabledPrint[0].vt=VT_DISPATCH;disabledPrint[0].pdispVal=lockedForTest;
  disabledPrint[1].vt=VT_BYREF|VT_BOOL;disabledPrint[1].pboolVal=&disabledPrintCancel;
  DISPPARAMS disabledPrintArgs={disabledPrint,nullptr,2,0};
  bool optionalOff=SetMandatoryOption(0)&&!HaronexMandatory::Gate::EnforcementEnabled()&&
    !HaronexMandatory::Gate::ShouldCancel(false,true,false)&&
    !HaronexMandatory::Gate::ShouldCancel(false,false,true)&&
    SUCCEEDED(gate.Invoke(101,IID_NULL,0,DISPATCH_METHOD,&disabledSaveArgs,nullptr,nullptr,nullptr))&&
    disabledSaveCancel==VARIANT_FALSE&&
    SUCCEEDED(gate.Invoke(102,IID_NULL,0,DISPATCH_METHOD,&disabledPrintArgs,nullptr,nullptr,nullptr))&&
    disabledPrintCancel==VARIANT_FALSE;
  bool optionalOn=SetMandatoryOption(1)&&HaronexMandatory::Gate::EnforcementEnabled()&&
    HaronexMandatory::Gate::ShouldCancel(false,true,false)&&
    HaronexMandatory::Gate::ShouldCancel(false,false,true)&&
    !HaronexMandatory::Gate::ShouldCancel(true,true,false);
  SetMandatoryOption(0);
  lockedForTest->Release();
  gate.Detach();releasedForTest->Release();other->Release();
  if(!separated||!saved||!printed||!safe||!forwardSaved||!forwardPrinted||!invalidFailsOpen||!correctCancelOnly||!optionalOff||!optionalOn) {
    wprintf(L"V1010_WORD_EVENT_RELEASE_FAILED\n");
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 40;
  }
  wprintf(L"NATIVE_COM_AND_RIBBON_AND_MANDATORY_HASH_OK\n");
  SysFreeString(xml);ribbon->Release();CoUninitialize();return 0;
}
