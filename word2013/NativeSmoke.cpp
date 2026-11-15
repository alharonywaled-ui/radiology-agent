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
  // Native COM gate must accept Clear Both only when text hash matches.
  HKEY tempKey=nullptr;
  const wchar_t* testPath=L"Software\\Haronex\\MandatoryReviews\\NativeSmokeV10_1";
  if(RegCreateKeyExW(HKEY_CURRENT_USER,testPath,0,nullptr,0,
                     KEY_SET_VALUE|KEY_QUERY_VALUE,nullptr,&tempKey,nullptr)!=ERROR_SUCCESS) {
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 36;
  }
  std::wstring approval=HaronexMandatory::Hash(L"unchanged test report");
  LONG written=RegSetValueExW(tempKey,L"ClearBothHash",0,REG_SZ,
           reinterpret_cast<const BYTE*>(approval.c_str()),
           static_cast<DWORD>((approval.size()+1)*sizeof(wchar_t)));
  bool matches=written==ERROR_SUCCESS &&
      HaronexMandatory::MatchingApproval(tempKey,L"ClearBothHash",approval) &&
      !HaronexMandatory::MatchingApproval(tempKey,L"ClearBothHash",
                       HaronexMandatory::Hash(L"changed test report")) &&
      !HaronexMandatory::MatchingApproval(tempKey,L"ApprovedHash",approval);
  RegCloseKey(tempKey);
  RegDeleteKeyW(HKEY_CURRENT_USER,testPath);
  if(!matches) {
    wprintf(L"MANDATORY_CLEAR_BOTH_HASH_TEST_FAILED\n");
    SysFreeString(xml);ribbon->Release();CoUninitialize();return 37;
  }
  wprintf(L"NATIVE_COM_AND_RIBBON_AND_MANDATORY_HASH_OK\n");
  SysFreeString(xml);ribbon->Release();CoUninitialize();return 0;
}
