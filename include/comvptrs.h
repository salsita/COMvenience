/*******************************************************************************
 * COMvenience : COM/ATL a bit more convenient
 * Copyright 2013 Salsita Software http://www.salsitasoft.com/
 * Author: Arne Seib <kontakt@seiberspace.de>
 *
 * File: comvptrs.h
 *
 * Some helpful CComPtr extensions.
 * Classes in this file:
 *  template class DispatchConvenience
 *   Provides convenience functions for dealing with IDispatch interfaces.
 *   This template is used by ComIDispatchPtr and ComIDispatchExPtr.
 *  ComIDispatchPtr
 *   Derives from CComQIPtr<IDispatch> and implements DispatchConvenience.
 *  ComIDispatchExPtr
 *   Derives from CComQIPtr<IDispatchEx> and implements DispatchConvenience for
 *   the IDispatch part.
 ******************************************************************************/

#pragma once

#include <atlcomcli.h>
#include <comvDispatchArgs.h>

/*==============================================================================
 * namespace
 */
namespace COMvenience
{

/*==============================================================================
 * template class DispatchConvenience
 *
 * The methods returning something (getProperty, call..) return a CComVariant
 * instead of a HRESULT for convenience.
 * The HRESULT of the call is still available as an optional out argument.
 *
 */
template<class Timpl> class DispatchConvenience
{
public:
  //----------------------------------------------------------------------------
  // Get a DISPID from a name
  DISPID getDispID(_In_ LPCOLESTR aName, _Out_opt_ HRESULT * aResultPtr = NULL)
      throw() {
    IDispatch * dispPtr = static_cast<Timpl*>(this)->p;
    ATLASSERT(dispPtr);
    HRESULT hr = E_FAIL;
    DISPID dispID = DISPID_UNKNOWN;
    if (dispPtr) {
      hr = dispPtr->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR*>(&aName), 1,
          LOCALE_USER_DEFAULT, &dispID);
    }
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }
    return dispID;
  }

  //----------------------------------------------------------------------------
  // Get a property by DISPID
  CComVariant getProperty(_In_ const DISPID dispID,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    IDispatch * dispPtr = static_cast<Timpl*>(this)->p;
    return getProperty(dispPtr, dispID, aResultPtr);
  }

  //----------------------------------------------------------------------------
  // Get a property by name
  CComVariant getProperty(_In_ LPCOLESTR aName,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    HRESULT hr = E_FAIL;
    DISPID dispID = getDispID(aName, &hr);
    if (SUCCEEDED(hr)) {
      return getProperty(dispID, aResultPtr);
    }
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }
    return CComVariant(NULL);
  }

  //----------------------------------------------------------------------------
  // Put a property by DISPID
  _Check_return_ HRESULT putProperty(_In_ DISPID dispID, _In_ CComVariant aValue)
      throw() {
    IDispatch * dispPtr = static_cast<Timpl*>(this)->p;
    return putProperty(dispPtr, dispID, aValue);
  }

  //----------------------------------------------------------------------------
  // Put a property by name
  _Check_return_ HRESULT putProperty(_In_ LPCOLESTR aName, _In_ CComVariant aValue)
      throw() {
    HRESULT hr = E_FAIL;
    DISPID dispID = getDispID(aName, &hr);
    if (SUCCEEDED(hr)) {
      hr = putProperty(dispID, aValue);
    }
    return hr;
  }

  //----------------------------------------------------------------------------
  // Wrapper for Invoke as a function call
  _Check_return_ CComVariant call(_In_ const DISPID aDispID,
      _In_ DispatchArgs & aArguments, _Out_opt_ HRESULT * aResultPtr = NULL)
      throw()
  {
    IDispatch * dispPtr = static_cast<Timpl*>(this)->p;
    return call(dispPtr, aDispID, aArguments, aResultPtr);
  }

  //----------------------------------------------------------------------------
  // Convenience method: Call a method by name
  _Check_return_ CComVariant call(_In_ LPCOLESTR aName,
      _In_ DispatchArgs & aArguments, _Out_opt_ HRESULT * aResultPtr = NULL)
      throw()
  {
    // name must exist
    DISPID dispID = getDispID(aName, aResultPtr);
    CComVariant retVal;
    if (DISPID_UNKNOWN != dispID) {
      retVal = call(dispID, aArguments, aResultPtr);
    }
    return retVal;
  }

  //----------------------------------------------------------------------------
  // Convenience method: Call a method by DISPID with no parameters
  _Check_return_ CComVariant call0(_In_ const DISPID aDispID,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw()
  {
    DispatchArgs arguments;
    return call(aDispID, arguments, aResultPtr);
  }

  //----------------------------------------------------------------------------
  // Convenience method: Call a method by name with no parameters
  _Check_return_ CComVariant call0(_In_ LPCOLESTR aName,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    // name must exist
    DISPID dispID = getDispID(aName, aResultPtr);
    CComVariant retVal;
    if (DISPID_UNKNOWN != dispID) {
      retVal = call0(dispID, aResultPtr);
    }
    return retVal;
  }

  //----------------------------------------------------------------------------
  // Convenience method: Call a method by DISPID with a single parameter
  _Check_return_ CComVariant call1(_In_ const DISPID aDispID,
      _In_ CComVariant aParam1, _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    DispatchArgs arguments;
    arguments.add(aParam1);
    return call(aDispID, arguments, aResultPtr);
  }

  //----------------------------------------------------------------------------
  // Convenience method: Call a method by name with a single parameter
  _Check_return_ CComVariant call1(_In_ LPCOLESTR aName,
      _In_ CComVariant aParam1, _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    DISPID dispID = getDispID(aName, aResultPtr);  // name must exist
    CComVariant retVal;
    if (DISPID_UNKNOWN != dispID) {
      retVal = call1(dispID, aParam1, aResultPtr);
    }
    return retVal;
  }

  //----------------------------------------------------------------------------
  // STATIC: Get a property by DISPID
  _Check_return_ CComVariant static getProperty(_In_ IDispatch* p,
    _In_ const DISPID aDispID, _Out_opt_ HRESULT * aResultPtr = NULL) {
    ATLASSERT(p);
    CComVariant retVal;
    if(p == NULL) {
      if (aResultPtr) {
        (*aResultPtr) = E_INVALIDARG;
      }
      return retVal;
    }

    ATLTRACE(atlTraceCOM, 2, _T("DispatchConvenience::getProperty\n"));
    DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
    HRESULT hr = p->Invoke(aDispID, IID_NULL, LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYGET, &dispparamsNoArgs, &retVal, NULL, NULL);
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }
    return retVal;
  }

  //----------------------------------------------------------------------------
  // STATIC: Put a property by DISPID
  _Check_return_ static HRESULT putProperty(_Inout_ IDispatch* p,
      _In_ const DISPID aDispID, _In_ CComVariant aValue) throw() {
    ATLASSERT(p);
    if(!p) {
      return E_INVALIDARG;
    }

    ATLTRACE(atlTraceCOM, 2, _T("DispatchConvenience::putProperty\n"));
    DISPID dispidPut = DISPID_PROPERTYPUT;
    DISPPARAMS dispParams = {&aValue, &dispidPut, 1, 1};

    if (aValue.vt == VT_UNKNOWN || aValue.vt == VT_DISPATCH ||
      (aValue.vt & VT_ARRAY) || (aValue.vt & VT_BYREF)) {
      HRESULT hr = p->Invoke(aDispID, IID_NULL, LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYPUTREF, &dispParams, NULL, NULL, NULL);
      if (SUCCEEDED(hr)) {
        return hr;
      }
    }
    return p->Invoke(aDispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
        &dispParams, NULL, NULL, NULL);
  }

  //----------------------------------------------------------------------------
  // STATIC: Wrapper for Invoke as a function call
  _Check_return_ static CComVariant call(_In_ IDispatch* p,
      _In_ const DISPID aDispID, _In_ DispatchArgs & aArguments,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    ATLASSERT(p);
    CComVariant retVal;
    if(p == NULL) {
      if (aResultPtr) {
        (*aResultPtr) = E_INVALIDARG;
      }
      return retVal;
    }

    HRESULT hr = p->Invoke(aDispID, IID_NULL, LOCALE_USER_DEFAULT,
        DISPATCH_METHOD, aArguments, &retVal, NULL, NULL);
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }

    return retVal;
  }
};

/*==============================================================================
 * class ComIDispatchPtr
 *
 * Adds some convenience methods to CComQIPtr<IDispatch> and makes dealing with
 * activescript easier.
 */
class ComIDispatchPtr :
  public CComQIPtr<IDispatch>, public DispatchConvenience<ComIDispatchPtr>
{
public:
  //----------------------------------------------------------------------------
  // Methods overwritten from CComQIPtr<IDispatch>
  ComIDispatchPtr() throw() :
      CComQIPtr<IDispatch>() {}

  ComIDispatchPtr(IDispatch* aDispPtr) throw() :
      CComQIPtr<IDispatch>(aDispPtr) {}

  ComIDispatchPtr(const CComQIPtr<IDispatch>& aSafePtr) throw() :
      CComQIPtr<IDispatch>(aSafePtr) {}

  IDispatch* operator=(IDispatch* aDispPtr) throw() {
    return CComQIPtr<IDispatch>::operator =(aDispPtr);
  }

};


/*==============================================================================
 * class ComIDispatchExPtr
 *
 * The same like ComIDispatchPtr, but for the IDispatchEx interface
 */
class ComIDispatchExPtr :
  public CComQIPtr<IDispatchEx>, public DispatchConvenience<ComIDispatchExPtr>
{
public:
  //----------------------------------------------------------------------------
  // Methods overwritten from CComQIPtr<IDispatchEx>
  ComIDispatchExPtr() throw() :
      CComQIPtr<IDispatchEx>() {}

  ComIDispatchExPtr(IDispatch* aDispPtr) throw() :
      CComQIPtr<IDispatchEx>(aDispPtr) {}

  ComIDispatchExPtr(const CComQIPtr<IDispatchEx>& aSafePtr) throw() :
      CComQIPtr<IDispatchEx>(aSafePtr) {}

  IDispatch* operator=(IDispatch* aDispPtr) throw() {
    return CComQIPtr<IDispatchEx>::operator =(aDispPtr);
  }

  //----------------------------------------------------------------------------
  // Get an ex(pando) DISPID from a name
  _Check_return_ DISPID getExDispID(_In_ LPCOLESTR aName,
      _In_ const BOOL aEnsureExists = TRUE,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    DWORD gfrdex = (aEnsureExists) ? fdexNameEnsure : 0;
    DISPID dispID = DISPID_UNKNOWN;
    HRESULT hr = p->GetDispID(CComBSTR(aName), gfrdex, &dispID);
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }
    return dispID;
  }

  //----------------------------------------------------------------------------
  // Get an ex-property by DISPID
  _Check_return_ CComVariant getExProperty(_In_ const DISPID dispID,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    return getExProperty(p, dispID, aResultPtr);
  }

  //----------------------------------------------------------------------------
  // Get an ex(pando)-property by name
  _Check_return_ CComVariant getExProperty(_In_ LPCOLESTR aName,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    ATLASSERT(p);
    HRESULT hr = E_FAIL;
    DISPID dispID = getExDispID(aName, FALSE, &hr);  // name must exist
    if (SUCCEEDED(hr)) {
      return getExProperty(dispID, aResultPtr);
    }
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }
    return CComVariant(NULL);
  }

  //----------------------------------------------------------------------------
  // Put an ex-property by DISPID
  _Check_return_ HRESULT putExProperty(_In_ DISPID aDispID,
      _In_ CComVariant aValue) throw() {
    return putExProperty(p, aDispID, aValue);
  }

  //----------------------------------------------------------------------------
  // Put an ex-property by name, create if not exists
  _Check_return_ HRESULT putExProperty(_In_ LPCOLESTR aName,
      _In_ CComVariant aValue) throw() {
    HRESULT hr = E_FAIL;
    // name will be created if not exists
    DISPID dispID = getExDispID(aName, TRUE, &hr);
    if (SUCCEEDED(hr)) {
      hr = putExProperty(dispID, aValue);
    }
    return hr;
  }

  //----------------------------------------------------------------------------
  // Wrapper for InvokeEx as a function call
  _Check_return_ CComVariant callEx(_In_ const DISPID aDispID,
      _In_ DispatchArgs & aArguments, _Out_opt_ HRESULT * aResultPtr = NULL)
      throw()
  {
    return callEx(p, aDispID, aArguments, aResultPtr);
  }

  //----------------------------------------------------------------------------
  // Convenience method: Call a method by name
  _Check_return_ CComVariant callEx(_In_ LPCOLESTR aName,
      _In_ DispatchArgs & aArguments, _Out_opt_ HRESULT * aResultPtr = NULL)
      throw()
  {
    // name must exist
    DISPID dispID = getExDispID(aName, FALSE, aResultPtr);
    CComVariant retVal;
    if (DISPID_UNKNOWN != dispID) {
      retVal = callEx(dispID, aArguments, aResultPtr);
    }
    return retVal;
  }

  //----------------------------------------------------------------------------
  // Convenience method: Call a method by DISPID with no parameters
  _Check_return_ CComVariant callEx0(_In_ const DISPID aDispID,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw()
  {
    DispatchArgs arguments;
    return callEx(aDispID, arguments, aResultPtr);
  }

  //----------------------------------------------------------------------------
  // Convenience method: Call a method by name with no parameters
  _Check_return_ CComVariant callEx0(_In_ LPCOLESTR aName,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    // name must exist
    DISPID dispID = getExDispID(aName, FALSE, aResultPtr);
    CComVariant retVal;
    if (DISPID_UNKNOWN != dispID) {
      retVal = callEx0(dispID, aResultPtr);
    }
    return retVal;
  }

  //----------------------------------------------------------------------------
  // Convenience method: Call a method by DISPID with a single parameter
  _Check_return_ CComVariant callEx1(_In_ const DISPID aDispID,
      _In_ CComVariant aParam1, _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    DispatchArgs arguments;
    arguments.add(aParam1);
    return callEx(aDispID, arguments, aResultPtr);
  }

  //----------------------------------------------------------------------------
  // Convenience method: Call a method by name with a single parameter
  _Check_return_ CComVariant callEx1(_In_ LPCOLESTR aName,
      _In_ CComVariant aParam1, _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    DISPID dispID = getExDispID(aName, FALSE, aResultPtr);  // name must exist
    CComVariant retVal;
    if (DISPID_UNKNOWN != dispID) {
      retVal = callEx1(dispID, aParam1, aResultPtr);
    }
    return retVal;
  }

  //----------------------------------------------------------------------------
  // Convenience method: Get a property of a certain type, do a type conversion
  // if required.
  // T is the type of the return value.
  // VT is the resulting VARIANT type. Get(..) trys to do a conversion to this
  // type.
  // TCast is an optional type to cast the result to before assigning to the
  // return value.
  // This is required e.g. for BSTR
  template <class T, VARTYPE VT, class TCast> T getAs(_In_ LPOLESTR aName,
      _Out_opt_ HRESULT * aResultPtr = NULL) {
    HRESULT hr = E_FAIL;
    CComVariant vtResult(getExProperty(aName, &hr));
    if (SUCCEEDED(hr)) {
      hr = vtResult.ChangeType(VT);
      if (SUCCEEDED(hr)) {
        return T((T)(TCast)(vtResult.llVal));
      }
    }
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }
    return T();
  }

  //----------------------------------------------------------------------------
  // Convenience method: Wrapper used if T == TCast
  template <class T, VARTYPE VT> T getAs(_In_ LPOLESTR aName,
      _Out_opt_ HRESULT * aResultPtr = NULL) {
    return getAs<T, VT, T>(aName, aResultPtr);
  }

  //----------------------------------------------------------------------------
  // Convenience method: for strings.
  // T must be constructable from a BSTR
  template <class T> T getAs(_In_ LPOLESTR aName,
      _Out_opt_ HRESULT * aResultPtr = NULL) {
    return getAs<T, VT_BSTR, BSTR>(aName, aResultPtr);
  }

  //----------------------------------------------------------------------------
  // STATIC: Put an ex-property by DISPID
  _Check_return_ static HRESULT putExProperty(_Inout_ IDispatchEx* p,
      _In_ DISPID aDispID, _In_ VARIANT & aValue) throw() {
    ATLASSERT(p);
    if(p == NULL) {
      return E_INVALIDARG;
    }

    ATLTRACE(atlTraceCOM, 2, _T("ComIDispatchExPtr::putExProperty\n"));
    DISPID dispidPut = DISPID_PROPERTYPUT;
    DISPPARAMS dispParams = {&aValue, &dispidPut, 1, 1};

    if (aValue.vt == VT_UNKNOWN || aValue.vt == VT_DISPATCH ||
      (aValue.vt & VT_ARRAY) || (aValue.vt & VT_BYREF)) {
      HRESULT hr = p->InvokeEx(aDispID, LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYPUTREF, &dispParams, NULL, NULL, NULL);
      if (SUCCEEDED(hr)) {
        return hr;
      }
    }
    return p->InvokeEx(aDispID, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
        &dispParams, NULL, NULL, NULL);
  }

  //----------------------------------------------------------------------------
  // STATIC: Get an ex-property by DISPID
  _Check_return_ CComVariant static getExProperty(_In_ IDispatchEx* p,
    _In_ const DISPID aDispID, _Out_opt_ HRESULT * aResultPtr = NULL) {
    ATLASSERT(p);
    CComVariant retVal;
    if(p == NULL) {
      if (aResultPtr) {
        (*aResultPtr) = E_INVALIDARG;
      }
      return retVal;
    }

    ATLTRACE(atlTraceCOM, 2, _T("ComIDispatchExPtr::getExProperty\n"));
    DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
    HRESULT hr = p->InvokeEx(aDispID, LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYGET, &dispparamsNoArgs, &retVal, NULL, NULL);
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }
    return retVal;
  }

  //----------------------------------------------------------------------------
  // STATIC: Wrapper for InvokeEx as a function call
  _Check_return_ static CComVariant callEx(_In_ IDispatchEx* p,
      _In_ const DISPID aDispID, _In_ DispatchArgs & aArguments,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    ATLASSERT(p);
    CComVariant retVal;
    if(p == NULL) {
      if (aResultPtr) {
        (*aResultPtr) = E_INVALIDARG;
      }
      return retVal;
    }

    HRESULT hr = p->InvokeEx(aDispID, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
        aArguments, &retVal, NULL, NULL);
    if (aResultPtr) {
      (*aResultPtr) = hr;
    }

    return retVal;
  }

};

} // namespace COMvenience
