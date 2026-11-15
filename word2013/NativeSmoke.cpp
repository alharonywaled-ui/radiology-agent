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
     !wcsstr(xml,L"ClearBothClick")||!wcsstr(xml,L"Clear Both")) {
    wprintf(L"RIBBON_CHECK_FAILED 0x%08x\n",static_cast<unsigned>(hr));
    if(xml)SysFreeString(xml);
    ribbon->Release();CoUninitialize();return 30;
  }
  if(HaronexMandatory::Hash(L"Text")!=L"71988c4d8e0803ba4519f0b2864c1331c14a1890bf8694e251379177bfedb5c3") {
    wprintf(L"MANDATORY_SHA256_INCOMPATIBLE\n");
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 35;
  }
  // Gate requires BOTH completed Review Text and Clear Both for same text.
  HKEY tempKey=nullptr;
  const wchar_t* testPath=L"Software\\Haronex\\MandatoryReviews\\NativeSmokeV10_2";
  if(RegCreateKeyExW(HKEY_CURRENT_USER,testPath,0,nullptr,0,
                     KEY_SET_VALUE|KEY_QUERY_VALUE,nullptr,&tempKey,nullptr)!=ERROR_SUCCESS) {
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 36;
  }
  std::wstring reviewHash=HaronexMandatory::Hash(L"original unchanged report");
  std::wstring editedHash=HaronexMandatory::Hash(L"edited report");
  auto setHash=[&](const wchar_t* name,const std::wstring& hash) -> bool {
    return RegSetValueExW(tempKey,name,0,REG_SZ,
             reinterpret_cast<const BYTE*>(hash.c_str()),
             static_cast<DWORD>((hash.size()+1)*sizeof(wchar_t)))==ERROR_SUCCESS;
  };
  bool valid=!HaronexMandatory::ApprovalReady(tempKey,reviewHash);
  valid=valid&&setHash(L"ClearBothHash",reviewHash);
  valid=valid&&!HaronexMandatory::ApprovalReady(tempKey,reviewHash);
  valid=valid&&setHash(L"ApprovedHash",reviewHash);
  valid=valid&&!HaronexMandatory::ApprovalReady(tempKey,reviewHash);
  valid=valid&&setHash(L"ClearBothV102Hash",reviewHash);
  valid=valid&&HaronexMandatory::ApprovalReady(tempKey,reviewHash);
  valid=valid&&!HaronexMandatory::ApprovalReady(tempKey,editedHash);
  valid=valid&&setHash(L"ApprovedHash",editedHash);
  valid=valid&&!HaronexMandatory::ApprovalReady(tempKey,editedHash);
  RegCloseKey(tempKey);
  RegDeleteKeyW(HKEY_CURRENT_USER,testPath);
  if(!valid) {
    wprintf(L"MANDATORY_REVIEW_THEN_CLEAR_BOTH_TEST_FAILED\n");
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
  wprintf(L"NATIVE_COM_AND_RIBBON_AND_MANDATORY_HASH_OK\n");
  SysFreeString(xml);ribbon->Release();CoUninitialize();return 0;
}
