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
inline HRESULT SetText(IDispatch* d,const wchar_t* name,const wchar_t* value) {
  DISPID id;
  HRESULT hr=Member(d,name,id);
  if(FAILED(hr))return hr;
  VARIANT arg;VariantInit(&arg);arg.vt=VT_BSTR;arg.bstrVal=SysAllocString(value);
  if(!arg.bstrVal)return E_OUTOFMEMORY;
  DISPID named=DISPID_PROPERTYPUT;
  DISPPARAMS params={&arg,&named,1,1};
  hr=d->Invoke(id,IID_NULL,LOCALE_USER_DEFAULT,DISPATCH_PROPERTYPUT,
               &params,nullptr,nullptr,nullptr);
  VariantClear(&arg);
  return hr;
}
inline bool ExecuteFind(IDispatch* find) {
  DISPID id;
  if(FAILED(Member(find,L"Execute",id)))return false;
  DISPPARAMS params={nullptr,nullptr,0,0};
  VARIANT result;VariantInit(&result);
  HRESULT hr=find->Invoke(id,IID_NULL,LOCALE_USER_DEFAULT,DISPATCH_METHOD,
                           &params,&result,nullptr,nullptr);
  VARIANT answer;VariantInit(&answer);
  bool found=SUCCEEDED(hr)&&SUCCEEDED(VariantChangeType(&answer,&result,0,VT_BOOL))
             &&answer.boolVal==VARIANT_TRUE;
  VariantClear(&answer);VariantClear(&result);
  return found;
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
  if(!doc){Alert(L"Open a Word report before removing comments.",MB_OK|MB_ICONWARNING);return;}
  Obj comments=PropertyObject(doc.p,L"Comments");
  long count=0;
  if(!comments||!PropertyNumber(comments.p,L"Count",count)){
    Alert(L"Unable to access comments in this document.",MB_OK|MB_ICONWARNING);return;
  }
  if(count==0){Alert(L"No comments were found in this report.");return;}
  if(MessageBoxW(nullptr,L"Delete ALL comments from the current Word document?\\nThis includes comments not created by Haronex.",
      L"Haronex - Delete Comments",MB_YESNO|MB_ICONWARNING|MB_DEFBUTTON2|MB_TOPMOST)!=IDYES)return;
  // Word 2013 supports document-level batch deletion. This is much faster than
  // deleting each COM comment; keep a bounded fallback for unusual documents.
  HRESULT hr=Method(doc.p,L"DeleteAllComments");
  if(SUCCEEDED(hr)){
    Alert(L"All comments were removed from the current Word document.");
    return;
  }
  const DWORD started=GetTickCount();
  const DWORD budgetMs=4000;
  long removed=0,failed=0;
  for(long i=count;i>=1;--i){
    if(static_cast<DWORD>(GetTickCount()-started)>=budgetMs)break;
    Obj item=IndexedObject(comments.p,i);
    if(item&&SUCCEEDED(Method(item.p,L"Delete")))++removed;
    else ++failed;
  }
  wchar_t buffer[256];
  wsprintfW(buffer,L"Removed %ld of %ld comments.\\n%s",removed,count,
    removed==count?L"Done.":L"Stopped after the safe time limit or an inaccessible comment. Click again to continue.");
  Alert(buffer,removed==count&&failed==0?MB_OK|MB_ICONINFORMATION:MB_OK|MB_ICONWARNING);
}
// Use Word's formatting-aware Find instead of iterating over every character.
// Finds only highlighted segments, then modifies red/yellow portions; other
// highlight colors are unchanged. All processing is local to Word.
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
  bool errors=false,limitReached=false;
  long scanned=0;
  // Hard limit prevents runaway loops on corrupted/complex Word documents.
  static const long MaxMatches=1500;
  static const DWORD MaxElapsedMs=4000;
  const DWORD cleanupStarted=GetTickCount();
  // Word story type IDs 1..17; unavailable types are skipped.
  for(long type=1;type<=17&&!limitReached;++type) {
    Obj story=IndexedObject(stories.p,type);
    int linked=0;
    while(story&&linked++<100&&!limitReached) {
      if(static_cast<DWORD>(GetTickCount()-cleanupStarted)>=MaxElapsedMs){limitReached=true;break;}
      long start=0,stop=0;
      if(!PropertyNumber(story.p,L"Start",start)||
         !PropertyNumber(story.p,L"End",stop)||stop<=start) {
        Obj next=PropertyObject(story.p,L"NextStoryRange");
        story=static_cast<Obj&&>(next);
        continue;
      }
      Obj search=PropertyObject(story.p,L"Duplicate");
      if(!search){errors=true;break;}
      long cursor=start;
      while(cursor<stop&&!limitReached) {
        if(static_cast<DWORD>(GetTickCount()-cleanupStarted)>=MaxElapsedMs){limitReached=true;break;}
        if(++scanned>MaxMatches){limitReached=true;break;}
        if(FAILED(SetNumber(search.p,L"Start",cursor))||
           FAILED(SetNumber(search.p,L"End",stop))) {
          errors=true;break;
        }
        Obj find=PropertyObject(search.p,L"Find");
        if(!find){errors=true;break;}
        // A blank find text plus formatting finds highlighted segments.
        if(FAILED(Method(find.p,L"ClearFormatting"))||
           FAILED(SetText(find.p,L"Text",L""))||
           FAILED(SetNumber(find.p,L"Format",1))||
           FAILED(SetNumber(find.p,L"Highlight",1))||
           FAILED(SetNumber(find.p,L"Forward",1))||
           FAILED(SetNumber(find.p,L"Wrap",0))) {
          errors=true;break;
        }
        if(!ExecuteFind(find.p))break;
        long hitStart=-1,hitEnd=-1,color=-1;
        if(!PropertyNumber(search.p,L"Start",hitStart)||
           !PropertyNumber(search.p,L"End",hitEnd)||
           hitEnd<=cursor||hitEnd<=hitStart) {
          errors=true;break;
        }
        if(PropertyNumber(search.p,L"HighlightColorIndex",color)) {
          if(color==6||color==7) {
            if(SUCCEEDED(SetNumber(search.p,L"HighlightColorIndex",0)))++changed;
            else errors=true;
          } else if(color!=0&&color!=1&&color!=2&&color!=3&&
                    color!=4&&color!=5&&color!=8&&color!=9&&color!=10&&
                    color!=11&&color!=12&&color!=13&&color!=14&&color!=15&&
                    color!=16) {
            // Mixed colors (wdUndefined). Inspect only this highlighted hit.
            // Bound expensive work to prevent Word UI lockups.
            Obj chars=PropertyObject(search.p,L"Characters");
            long count=0;
            if(chars&&PropertyNumber(chars.p,L"Count",count)) {
              if(count>1000){limitReached=true;break;}
              for(long c=1;c<=count;++c) {
                if(static_cast<DWORD>(GetTickCount()-cleanupStarted)>=MaxElapsedMs){limitReached=true;break;}
                Obj item=IndexedObject(chars.p,c);
                long itemColor=-1;
                if(item&&PropertyNumber(item.p,L"HighlightColorIndex",itemColor)&&
                   (itemColor==6||itemColor==7)) {
                  if(SUCCEEDED(SetNumber(item.p,L"HighlightColorIndex",0)))++changed;
                  else errors=true;
                }
              }
            } else errors=true;
          }
        } else errors=true;
        cursor=hitEnd;
      }
      Obj next=PropertyObject(story.p,L"NextStoryRange");
      story=static_cast<Obj&&>(next);
    }
    if(linked>=100)limitReached=true;
  }
  wchar_t buffer[256];
  if(limitReached) {
    wsprintfW(buffer,L"Stopped safely after %ld highlighted segments.\n"
      L"%ld red/yellow highlights cleared.\nSave the report and retry if needed.",
      scanned,changed);
    Alert(buffer,MB_OK|MB_ICONWARNING);
  } else {
    wsprintfW(buffer,L"Cleared %ld red/yellow highlighted segments.%s",
      changed,errors?L"\nSome protected ranges could not be processed.":L"");
    Alert(buffer,errors?MB_OK|MB_ICONWARNING:MB_OK|MB_ICONINFORMATION);
  }
}
} // namespace LocalCleanup
