/*******************************************************************************
 * COMvenience : COM/ATL a bit more convenient
 * Copyright 2013 Salsita Software http://www.salsitasoft.com/
 * Author: Arne Seib <kontakt@seiberspace.de>
 *
 * File: comvbrowser.h
 *
 * CCom(QI)Ptr derivations for WebBrowser control related interfaces.
 * Classes in this file:
 ******************************************************************************/

#pragma once

#include <comvptrs.h>
#include <ExDispid.h>

/*==============================================================================
 * namespace
 */
namespace COMvenience
{

#define IMPLEMENT_OBJECT_GETTER(_name, _name2, _type) \
  _Check_return_ CComPtr<_type> get##_name(_Out_opt_ HRESULT * aResultPtr = NULL) { \
    HRESULT hr = E_INVALIDARG; \
    CComPtr<_type> ret; \
    hr = (*this)->get_##_name2(&ret); \
    ATLASSERT(SUCCEEDED(hr)); \
    if (aResultPtr) { \
      (*aResultPtr) = hr; \
    } \
    return ret; \
  }

#define IMPLEMENT_OBJECT_GETTER2(_name, _name2, _type) \
  _Check_return_ CComPtr<_type> get##_name(_Out_opt_ HRESULT * aResultPtr = NULL) { \
    HRESULT hr = E_INVALIDARG; \
    CComQIPtr<_type> ret; \
    CComPtr<IDispatch> disp; \
    hr = (*this)->get_##_name2(&disp); \
    ATLASSERT(SUCCEEDED(hr)); \
    if (SUCCEEDED(hr) && disp) { \
      ret = disp; \
      if (!ret) { \
        hr = E_NOINTERFACE; \
      } \
    } \
    if (aResultPtr) { \
      (*aResultPtr) = hr; \
    } \
    return ret; \
  }

/*==============================================================================
 * class WebBrowser2Ptr
 *
 * specialization for IWebBrowser2
 */
class WebBrowser2Ptr : public DispatchPtr< IWebBrowser2 >
{
public:
  WebBrowser2Ptr() throw() : DispatchPtr< IWebBrowser2 >() {
  }

  WebBrowser2Ptr(_In_opt_ IWebBrowser2* lp) throw() :
      DispatchPtr< IWebBrowser2 >(lp) {
  }

  WebBrowser2Ptr(_In_ const WebBrowser2Ptr & lp) throw() :
      DispatchPtr< IWebBrowser2 >(lp.p) {
  }

  WebBrowser2Ptr(_In_opt_ IUnknown* lp) throw() :
      DispatchPtr< IWebBrowser2 >(lp) {
  }

  IWebBrowser2* operator=(_In_ const WebBrowser2Ptr & lp) throw() {
    if(*this != lp) {
      return static_cast<IWebBrowser2*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
    }
    return *this;
  }

  IWebBrowser2* operator=(_In_opt_ IUnknown* lp) throw() {
    if(*this != lp) {
      return static_cast<IWebBrowser2*>(AtlComQIPtrAssign((IUnknown**)&p, lp,
        __uuidof(IWebBrowser2)));
    }
    return *this;
  }

  _Check_return_ CComPtr<IHTMLDocument2> getDocument(
      _Out_opt_ HRESULT * aResultPtr = NULL) {
    HRESULT hr = E_INVALIDARG;
    CComPtr<IHTMLDocument2> document;
    ATLASSERT(p);
    if (p) {
      CComPtr<IDispatch> dispatch;
      hr = p->get_Document(&dispatch);
      if (SUCCEEDED(hr)) {
        hr = dispatch.QueryInterface(&document.p);
      }
    }
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }
    return document;
  }

  _Check_return_ CComPtr<IHTMLWindow2> getWindow(
      _Out_opt_ HRESULT * aResultPtr = NULL) {
    HRESULT hr = E_INVALIDARG;
    CComPtr<IHTMLWindow2> window;
    CComQIPtr<IHTMLDocument2> document(getDocument(&hr));
    if (document) {
      hr = document->get_parentWindow(&window);
    }
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }
    return window;
  }

  _Check_return_ DispatchExPtr getScriptDispatchEx(
      _Out_opt_ HRESULT * aResultPtr = NULL) {
    return DispatchExPtr(getWindow(aResultPtr));
  }

  READYSTATE getReadyState(_Out_opt_ HRESULT * aResultPtr = NULL) {
    READYSTATE ret = READYSTATE_UNINITIALIZED;
    HRESULT hr = (*this)->get_ReadyState(&ret);
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }
    return ret;
  }
};

/*==============================================================================
 * class HTMLDocument2Ptr
 *
 * specialization for IHTMLDocument2
 */
class HTMLDocument2Ptr : public DispatchPtr< IHTMLDocument2 >
{
public:
  HTMLDocument2Ptr() throw() : DispatchPtr< IHTMLDocument2 >() {
  }

  HTMLDocument2Ptr(_In_opt_ IHTMLDocument2* lp) throw() :
    DispatchPtr< IHTMLDocument2 >(lp) {
  }

  HTMLDocument2Ptr(_In_ const HTMLDocument2Ptr & lp) throw() :
    DispatchPtr< IHTMLDocument2 >(lp.p) {
  }

  HTMLDocument2Ptr(_In_opt_ IUnknown* lp) throw() :
    DispatchPtr< IHTMLDocument2 >(lp) {
  }

  IHTMLDocument2* operator=(_In_ const HTMLDocument2Ptr & lp) throw() {
    if(*this != lp) {
      return static_cast<IHTMLDocument2*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
    }
    return *this;
  }

  IHTMLDocument2* operator=(_In_opt_ IUnknown* lp) throw() {
    if(*this != lp) {
      return static_cast<IHTMLDocument2*>(AtlComQIPtrAssign((IUnknown**)&p,
          lp, __uuidof(IHTMLDocument2)));
    }
    return *this;
  }

  IMPLEMENT_OBJECT_GETTER(Body, body, IHTMLElement)
  IMPLEMENT_OBJECT_GETTER(Forms, forms, IHTMLElementCollection)
  IMPLEMENT_OBJECT_GETTER(Frames, frames, IHTMLFramesCollection2)
  IMPLEMENT_OBJECT_GETTER(Images, images, IHTMLElementCollection)
  IMPLEMENT_OBJECT_GETTER(Links, links, IHTMLElementCollection)
  IMPLEMENT_OBJECT_GETTER(Location, location, IHTMLLocation)
  IMPLEMENT_OBJECT_GETTER(ParentWindow, parentWindow, IHTMLWindow2)
  IMPLEMENT_OBJECT_GETTER(Plugins, plugins, IHTMLElementCollection)
  IMPLEMENT_OBJECT_GETTER(Scripts, scripts, IHTMLElementCollection)
  IMPLEMENT_OBJECT_GETTER(Selection, selection, IHTMLSelectionObject)
  IMPLEMENT_OBJECT_GETTER(StyleSheets, styleSheets, IHTMLStyleSheetsCollection)

};

/*==============================================================================
 * class HTMLDocument3Ptr
 *
 * specialization for IHTMLDocument3
 */
class HTMLDocument3Ptr : public DispatchPtr< IHTMLDocument3 >
{
public:
  HTMLDocument3Ptr() throw() : DispatchPtr< IHTMLDocument3 >() {
  }

  HTMLDocument3Ptr(_In_opt_ IHTMLDocument3* lp) throw() :
    DispatchPtr< IHTMLDocument3 >(lp) {
  }

  HTMLDocument3Ptr(_In_ const HTMLDocument3Ptr & lp) throw() :
    DispatchPtr< IHTMLDocument3 >(lp.p) {
  }

  HTMLDocument3Ptr(_In_opt_ IUnknown* lp) throw() :
    DispatchPtr< IHTMLDocument3 >(lp) {
  }

  IHTMLDocument3* operator=(_In_ const HTMLDocument3Ptr & lp) throw() {
    if(*this != lp) {
      return static_cast<IHTMLDocument3*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
    }
    return *this;
  }

  IHTMLDocument3* operator=(_In_opt_ IUnknown* lp) throw() {
    if(*this != lp) {
      return static_cast<IHTMLDocument3*>(AtlComQIPtrAssign((IUnknown**)&p, lp,
          __uuidof(IHTMLDocument3)));
    }
    return *this;
  }

  IMPLEMENT_OBJECT_GETTER(DocumentElement, documentElement, IHTMLElement)
  IMPLEMENT_OBJECT_GETTER(ParentDocument, parentDocument, IHTMLDocument2)
  // getElementById
  // getElementsByTagName

};

/*==============================================================================
 * class HTMLDocument4Ptr
 *
 * specialization for IHTMLDocument4
 */
class HTMLDocument4Ptr : public DispatchPtr< IHTMLDocument4 >
{
public:
  HTMLDocument4Ptr() throw() : DispatchPtr< IHTMLDocument4 >() {
  }

  HTMLDocument4Ptr(_In_opt_ IHTMLDocument4* lp) throw() :
    DispatchPtr< IHTMLDocument4 >(lp) {
  }

  HTMLDocument4Ptr(_In_ const HTMLDocument4Ptr & lp) throw() :
    DispatchPtr< IHTMLDocument4 >(lp.p) {
  }

  HTMLDocument4Ptr(_In_opt_ IUnknown* lp) throw() :
    DispatchPtr< IHTMLDocument4 >(lp) {
  }

  IHTMLDocument4* operator=(_In_ const HTMLDocument4Ptr & lp) throw() {
    if(*this != lp) {
      return static_cast<IHTMLDocument4*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
    }
    return *this;
  }

  IHTMLDocument4* operator=(_In_opt_ IUnknown* lp) throw() {
    if(*this != lp) {
      return static_cast<IHTMLDocument4*>(AtlComQIPtrAssign((IUnknown**)&p, lp,
          __uuidof(IHTMLDocument4)));
    }
    return *this;
  }

  IMPLEMENT_OBJECT_GETTER2(Namespaces, namespaces, IHTMLNamespaceCollection)

};

/*==============================================================================
 * class HTMLDocument5Ptr
 *
 * specialization for IHTMLDocument5
 */
class HTMLDocument5Ptr : public DispatchPtr< IHTMLDocument5 >
{
public:
  HTMLDocument5Ptr() throw() : DispatchPtr< IHTMLDocument5 >() {
  }

  HTMLDocument5Ptr(_In_opt_ IHTMLDocument5* lp) throw() :
    DispatchPtr< IHTMLDocument5 >(lp) {
  }

  HTMLDocument5Ptr(_In_ const HTMLDocument5Ptr & lp) throw() :
    DispatchPtr< IHTMLDocument5 >(lp.p) {
  }

  HTMLDocument5Ptr(_In_opt_ IUnknown* lp) throw() :
    DispatchPtr< IHTMLDocument5 >(lp) {
  }

  IHTMLDocument5* operator=(_In_ const HTMLDocument5Ptr & lp) throw() {
    if(*this != lp) {
      return static_cast<IHTMLDocument5*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
    }
    return *this;
  }

  IHTMLDocument5* operator=(_In_opt_ IUnknown* lp) throw() {
    if(*this != lp) {
      return static_cast<IHTMLDocument5*>(AtlComQIPtrAssign((IUnknown**)&p, lp,
          __uuidof(IHTMLDocument5)));
    }
    return *this;
  }

  IMPLEMENT_OBJECT_GETTER(Doctype, doctype, IHTMLDOMNode)

};

/*==============================================================================
 * class HTMLWindow2Ptr
 *
 * specialization for IHTMLWindow2
 */
class HTMLWindow2Ptr : public DispatchPtr< IHTMLWindow2 >
{
public:
  HTMLWindow2Ptr() throw() : DispatchPtr< IHTMLWindow2 >() {
  }

  HTMLWindow2Ptr(_In_opt_ IHTMLWindow2* lp) throw() :
    DispatchPtr< IHTMLWindow2 >(lp) {
  }

  HTMLWindow2Ptr(_In_ const HTMLWindow2Ptr & lp) throw() :
    DispatchPtr< IHTMLWindow2 >(lp.p) {
  }

  HTMLWindow2Ptr(_In_opt_ IUnknown* lp) throw() :
    DispatchPtr< IHTMLWindow2 >(lp) {
  }

  IHTMLWindow2* operator=(_In_ const HTMLWindow2Ptr & lp) throw() {
    if(*this != lp) {
      return static_cast<IHTMLWindow2*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
    }
    return *this;
  }

  IHTMLWindow2* operator=(_In_opt_ IUnknown* lp) throw() {
    if(*this != lp) {
      return static_cast<IHTMLWindow2*>(AtlComQIPtrAssign((IUnknown**)&p, lp,
          __uuidof(IHTMLWindow2)));
    }
    return *this;
  }

  IMPLEMENT_OBJECT_GETTER(Document, document, IHTMLDocument2)
  IMPLEMENT_OBJECT_GETTER(Event, event, IHTMLEventObj)
  IMPLEMENT_OBJECT_GETTER(External, external, IDispatch)
  IMPLEMENT_OBJECT_GETTER(Frames, frames, IHTMLFramesCollection2)
  IMPLEMENT_OBJECT_GETTER(History, history, IOmHistory)
  // next: Image

};

// cleanup
#undef IMPLEMENT_OBJECT_GETTER2
#undef IMPLEMENT_OBJECT_GETTER

} // namespace COMvenience
