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

// cleanup
#undef IMPLEMENT_QI_PTR

} // namespace COMvenience
