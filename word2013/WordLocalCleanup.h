#pragma once
// These commands run entirely in Word's in-process COM add-in (the Ribbon STA).
// No network, model, worker, clipboard, file scanning, or other Word instance.
#include <windows.h>
#include <oleauto.h>
#include <string>

namespace LocalCleanup {
struct Obj {
  IDispatch* p=nullptr;
  Obj()=default;
  explicit Obj(IDispatch* q):p(q){}
  ~Obj(){if(p)p->Release();}
  Obj(const Obj&)=delete;
  Obj& operator=(const Obj&)=delete;
  Obj(Obj&& o) noexcept:p(o.p){o.p=nullptr;}
  Obj& operator=(Obj&& o) noexcept {if(this!=&o){if(p)p->Release();p=o.p;o.p=nullptr;}return *this;}
  explicit operator bool() const{return p!=nullptr;}
};

inline HRESULT Member(IDispatch* d,const wchar_t* name,DISPID& id) {
  if(!d)return E_POINTER;
  LPOLESTR n=const_cast<LPOLESTR>(name);
  return d->GetIDsOfNames(IID_NULL,&n,1,LOCALE_USER_DEFAULT,&id);
}
inline HRESULT Get(IDispatch* d,const wchar_t* name,VARIANT* value) {
  if(!value)return E_POINTER;
  VariantInit(value);
  DISPID id;
  HRESULT hr=Member(d,name,id);
  if(FAILED(hr))return hr;
  DISPPARAMS params={nullptr,nullptr,0,0};
  return d->Invoke(id,IID_NULL,LOCALE_USER_DEFAULT,DISPATCH_PROPERTYGET,
                   &params,value,nullptr,nullptr);
}
inline Obj ConvertToObject(VARIANT& value) {
  IDispatch* p=nullptr;
  if(value.vt==VT_DISPATCH&&value.pdispVal) {
    p=value.pdispVal;
    p->AddRef();
  } else if(value.vt==VT_UNKNOWN&&value.punkVal) {
    value.punkVal->QueryInterface(IID_IDispatch,reinterpret_cast<void**>(&p));
  }
  return Obj(p);
}
inline Obj PropertyObject(IDispatch* d,const wchar_t* name) {
  VARIANT v;VariantInit(&v);
  if(FAILED(Get(d,name,&v)))return Obj();
  Obj obj=ConvertToObject(v);
  VariantClear(&v);
  return obj;
}
inline bool PropertyNumber(IDispatch* d,const wchar_t* name,long& result) {
  VARIANT v;VariantInit(&v);
  if(FAILED(Get(d,name,&v)))return false;
  VARIANT n;VariantInit(&n);
  HRESULT hr=VariantChangeType(&n,&v,0,VT_I4);
  if(SUCCEEDED(hr))result=n.lVal;
  VariantClear(&n);VariantClear(&v);
  return SUCCEEDED(hr);
}
inline HRESULT Index(IDispatch* d,long index,VARIANT* value) {
  if(!value)return E_POINTER;
  VariantInit(value);
  DISPID id;
  HRESULT hr=Member(d,L"Item",id);
  if(FAILED(hr))return hr;
  VARIANT arg;VariantInit(&arg);arg.vt=VT_I4;arg.lVal=index;
  DISPPARAMS params={&arg,nullptr,1,0};
  return d->Invoke(id,IID_NULL,LOCALE_USER_DEFAULT,DISPATCH_METHOD|DISPATCH_PROPERTYGET,
                   &params,value,nullptr,nullptr);
}
inline Obj IndexedObject(IDispatch* d,long index) {
  VARIANT v;VariantInit(&v);
  if(FAILED(Index(d,index,&v)))return Obj();
  Obj obj=ConvertToObject(v);
  VariantClear(&v);
  return obj;
}
inline HRESULT Method(IDispatch* d,const wchar_t* name) {
  DISPID id;
  HRESULT hr=Member(d,name,id);
  if(FAILED(hr))return hr;
  DISPPARAMS params={nullptr,nullptr,0,0};
  return d->Invoke(id,IID_NULL,LOCALE_USER_DEFAULT,DISPATCH_METHOD,&params,
                   nullptr,nullptr,nullptr);
}
inline HRESULT SetNumber(IDispatch* d,const wchar_t* name,long value) {
  DISPID id;
  HRESULT hr=Member(d,name,id);
  if(FAILED(hr))return hr;
  VARIANT arg;VariantInit(&arg);arg.vt=VT_I4;arg.lVal=value;
  DISPID named=DISPID_PROPERTYPUT;
  DISPPARAMS params={&arg,&named,1,1};
  return d->Invoke(id,IID_NULL,LOCALE_USER_DEFAULT,DISPATCH_PROPERTYPUT,&params,
                   nullptr,nullptr,nullptr);
}
inline void Alert(const wchar_t* message,UINT flags=MB_OK|MB_ICONINFORMATION) {
  MessageBoxW(nullptr,message,L"Haronex - Local Word Tools",flags|MB_TOPMOST);
}
inline Obj ActiveDocument(IDispatch* word) {
  return PropertyObject(word,L"ActiveDocument");
}
// Removing *all* comments is explicitly requested. Work backwards so indices do not shift.
inline void DeleteAllComments(IDispatch* word) {
  Obj doc=ActiveDocument(word);
  if(!doc) {Alert(L"Open a Word report before removing comments.",MB_OK|MB_ICONWARNING);return;}
  Obj comments=PropertyObject(doc.p,L"Comments");
  long count=0;
  if(!comments||!PropertyNumber(comments.p,L"Count",count)) {
    Alert(L"Unable to access comments in this document.",MB_OK|MB_ICONWARNING);return;
  }
  if(count==0){Alert(L"No comments were found in this report.");return;}
  if(MessageBoxW(nullptr,L"Delete ALL comments from the current Word document?\nThis includes comments not created by Haronex.",L"Haronex - Delete Comments",
      MB_YESNO|MB_ICONWARNING|MB_DEFBUTTON2|MB_TOPMOST)!=IDYES)return;
  long deleted=0;
  for(long i=count;i>=1;--i) {
    Obj item=IndexedObject(comments.p,i);
    if(item&&SUCCEEDED(Method(item.p,L"Delete")))++deleted;
  }
  wchar_t buffer[160];
  wsprintfW(buffer,L"Deleted %ld of %ld comments from the current document.",deleted,count);
  Alert(buffer,deleted==count?MB_OK|MB_ICONINFORMATION:MB_OK|MB_ICONWARNING);
}
// Word WdColorIndex: wdNoHighlight=0, wdRed=6, wdYellow=7.
// Examine individual character ranges; this preserves other highlight colors,
// and does not remove font colors or shading. Covers main text, headers/footers,
// text boxes and linked story ranges through Word's StoryRanges collection.
inline void RemoveRedYellowHighlights(IDispatch* word) {
  Obj doc=ActiveDocument(word);
  if(!doc){Alert(L"Open a Word report before removing highlights.",MB_OK|MB_ICONWARNING);return;}
  if(MessageBoxW(nullptr,L"Remove ALL red and yellow text highlights in the current document?\n"
      L"Highlights manually applied by other people will also be removed.\n"
      L"Other colors, font colors and text will be preserved.",
      L"Haronex - Clear Highlights",MB_YESNO|MB_ICONWARNING|MB_DEFBUTTON2|MB_TOPMOST)!=IDYES)return;
  Obj stories=PropertyObject(doc.p,L"StoryRanges");
  if(!stories){Alert(L"Unable to access document content.",MB_OK|MB_ICONWARNING);return;}
  long changed=0;
  bool errors=false;
  // Word story IDs are 1..17, with some unavailable in a given document.
  for(long type=1;type<=17;++type) {
    Obj story=IndexedObject(stories.p,type);
    int linked=0;
    while(story&&linked++<200) {
      Obj characters=PropertyObject(story.p,L"Characters");
      long length=0;
      if(characters&&PropertyNumber(characters.p,L"Count",length)) {
        for(long i=1;i<=length;++i) {
          Obj ch=IndexedObject(characters.p,i);
          if(!ch){errors=true;continue;}
          long color=-1;
          if(!PropertyNumber(ch.p,L"HighlightColorIndex",color)){errors=true;continue;}
          if(color==6||color==7) {
            if(SUCCEEDED(SetNumber(ch.p,L"HighlightColorIndex",0)))++changed;
            else errors=true;
          }
        }
      } else errors=true;
      Obj next=PropertyObject(story.p,L"NextStoryRange");
      story=static_cast<Obj&&>(next);
    }
    if(linked>=200)errors=true;
  }
  wchar_t buffer[200];
  wsprintfW(buffer,L"Cleared red/yellow highlights from %ld characters.%s",
           changed,errors?L"\nSome protected or unavailable ranges could not be processed.":L"");
  Alert(buffer,errors?MB_OK|MB_ICONWARNING:MB_OK|MB_ICONINFORMATION);
}
} // namespace LocalCleanup
