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
#include <_implementQIPtr.h>

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
template<> class ConvDispPtr<IWebBrowser2> :
      public CComQIPtr<IWebBrowser2>,
      public DispatchConvenience<ConvDispPtr<IWebBrowser2> > {
public:
  IMPLEMENT_QI_PTR(ConvDispPtr, IWebBrowser2)

  _Check_return_ CComPtr<IHTMLDocument2> getDocument(_Out_opt_ HRESULT * aResultPtr = NULL) {
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

  _Check_return_ CComPtr<IHTMLWindow2> getWindow(_Out_opt_ HRESULT * aResultPtr = NULL) {
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

  _Check_return_ ConvIDispatchExPtr getScriptDispatchEx(_Out_opt_ HRESULT * aResultPtr = NULL) {
    return ConvIDispatchExPtr(getWindow(aResultPtr));
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
typedef ConvDispPtr<IWebBrowser2> WebBrowser2Ptr;

/*==============================================================================
 * class HTMLDocument2Ptr
 *
 * specialization for IHTMLDocument2
 */
template<> class ConvDispPtr<IHTMLDocument2> :
      public CComQIPtr<IHTMLDocument2>,
      public DispatchConvenience<ConvDispPtr<IHTMLDocument2> > {
public:
  IMPLEMENT_QI_PTR(ConvDispPtr, IHTMLDocument2)

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
typedef ConvDispPtr<IHTMLDocument2> HTMLDocument2Ptr;

/*==============================================================================
 * class HTMLDocument3Ptr
 *
 * specialization for IHTMLDocument3
 */
template<> class ConvDispPtr<IHTMLDocument3> :
      public CComQIPtr<IHTMLDocument3>,
      public DispatchConvenience<ConvDispPtr<IHTMLDocument3> > {
public:
  IMPLEMENT_QI_PTR(ConvDispPtr, IHTMLDocument3)

  IMPLEMENT_OBJECT_GETTER(DocumentElement, documentElement, IHTMLElement)
  IMPLEMENT_OBJECT_GETTER(ParentDocument, parentDocument, IHTMLDocument2)
  // getElementById
  // getElementsByTagName

};
typedef ConvDispPtr<IHTMLDocument3> HTMLDocument3Ptr;

/*==============================================================================
 * class HTMLDocument4Ptr
 *
 * specialization for IHTMLDocument4
 */
template<> class ConvDispPtr<IHTMLDocument4> :
      public CComQIPtr<IHTMLDocument4>,
      public DispatchConvenience<ConvDispPtr<IHTMLDocument4> > {
public:
  IMPLEMENT_QI_PTR(ConvDispPtr, IHTMLDocument4)

  IMPLEMENT_OBJECT_GETTER2(Namespaces, namespaces, IHTMLNamespaceCollection)

};
typedef ConvDispPtr<IHTMLDocument4> HTMLDocument4Ptr;

/*==============================================================================
 * class HTMLDocument5Ptr
 *
 * specialization for IHTMLDocument5
 */
template<> class ConvDispPtr<IHTMLDocument5> :
      public CComQIPtr<IHTMLDocument5>,
      public DispatchConvenience<ConvDispPtr<IHTMLDocument5> > {
public:
  IMPLEMENT_QI_PTR(ConvDispPtr, IHTMLDocument5)

  IMPLEMENT_OBJECT_GETTER(Doctype, doctype, IHTMLDOMNode)

};
typedef ConvDispPtr<IHTMLDocument5> HTMLDocument5Ptr;

// cleanup
#undef IMPLEMENT_OBJECT_GETTER2
#undef IMPLEMENT_OBJECT_GETTER
#undef IMPLEMENT_QI_PTR

} // namespace COMvenience
