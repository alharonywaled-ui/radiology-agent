#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <ole2.h>
#include <oleauto.h>
#include <cwchar>
#include <cstdio>

struct IRibbon : IUnknown {
 virtual HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR,BSTR*)=0;
};
static const GUID clsid={0xD1FBEFD5,0x3B43,0x4516,{0x9B,0x7B,0xCD,0xEB,0x47,0xAB,0x0F,0xA1}};
static const GUID iid={0x000C0396,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}};
int main(){
  HRESULT hr=CoInitializeEx(nullptr,COINIT_APARTMENTTHREADED);
  if(FAILED(hr))return 10;
  IRibbon* ribbon=nullptr;
  hr=CoCreateInstance(clsid,nullptr,CLSCTX_INPROC_SERVER,iid,(void**)&ribbon);
  if(FAILED(hr)){wprintf(L"COM_ACTIVATION_FAILED 0x%08x\n",static_cast<unsigned>(hr));CoUninitialize();return 20;}
  BSTR xml=nullptr;
  hr=ribbon->GetCustomUI(SysAllocString(L"Microsoft.Word.Document"),&xml);
  if(FAILED(hr)||!xml||!wcsstr(xml,L"Radiology Agent")||!wcsstr(xml,L"ReviewTextClick")) {
    wprintf(L"RIBBON_CHECK_FAILED 0x%08x\n",static_cast<unsigned>(hr));
    if(xml)SysFreeString(xml);
    ribbon->Release();CoUninitialize();return 30;
  }
  wprintf(L"NATIVE_COM_AND_RIBBON_OK\n");
  SysFreeString(xml);ribbon->Release();CoUninitialize();return 0;
}
