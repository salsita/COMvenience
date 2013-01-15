/*******************************************************************************
 * COMvenience : COM/ATL a bit more convenient
 * Copyright 2013 Salsita Software http://www.salsitasoft.com/
 * Author: Arne Seib <kontakt@seiberspace.de>
 *
 * File: comvptrs.h
 *
 * Some helpful CComPtr & co replacements.
 * Classes in this file:
 *  template <class T> class Ptr
 *   Derives from CComPtrBase<T> and replaces CComPtr<T>.
 *  template <class T, const IID* piid = &__uuidof(T)> class QIPtr
 *   Derives from Ptr<T> and replaces CComQIPtr<T, IID>.
 *  template <class T> class DispatchPtr
 *   Derives from QIPtr<T>.
 *   Provides convenience functions for dealing with IDispatch based
 *   interfaces.
 *  class DispatchExPtr
 *   Derives from DispatchPtr< IDispatchEx >.
 *   Provides convenience functions for dealing with IDispatchEx based
 *   interfaces.
 ******************************************************************************/

#pragma once

#include <atlcomcli.h>
#include <comvDispatchArgs.h>
#include <string>

/*==============================================================================
 * namespace
 */

namespace COMvenience
{

#ifndef _ATL_NO_EXCEPTIONS
// This macro sets the out argument correct AND still throws.
// The programmer might still choose to define AtlThrow.
// Also it propertly returns the expected returnvalue.
# define RET_HANDLE_ERROR_PTR(_hr, _phr, _retval) \
    if (_phr) { \
      (*_phr) = _hr; \
    } \
    if (FAILED(_hr)) { \
      AtlThrow(_hr); \
    } \
    return _retval;
#else
// This macro ONLY sets the out argument correct and returns.
# define RET_HANDLE_ERROR_PTR(_hr, _phr, _retval) \
    if (_phr) { \
      (*_phr) = _hr; \
    } \
    return _retval;
#endif


/*==============================================================================
 * template class Ptr
 *
 * Basic CComPtr replacement. We need a template type clean of any specialized
 * versions, because we will do our own specialization.
 * The code is a copy from the original CComPtr - only the formatting is changed
 * to fit the Salsita coding guidelines.
 *
 */
template <class T> class Ptr : public CComPtrBase<T>
{
public:
  Ptr() throw() {
  }

  Ptr(int nNull) throw() :
      CComPtrBase<T>(nNull) {
  }

  Ptr(T* lp) throw() :
      CComPtrBase<T>(lp) {
  }

  Ptr(_In_ const Ptr<T>& lp) throw() :
      CComPtrBase<T>(lp.p) {
  }

  // additional ctor for compatibility with CComPtr
  Ptr(_In_ const CComPtr<T>& lp) throw() :
      CComPtrBase<T>(lp.p) {
  }

  T* operator=(_In_opt_ T* lp) throw() {
    if(*this!=lp) {
      return static_cast<T*>(AtlComPtrAssign((IUnknown**)&p, lp));
    }
    return *this;
  }

  template <typename Q> T* operator=(_In_ const Ptr<Q>& lp) throw() {
    if (!IsEqualObject(lp)) {
      return static_cast<T*>(AtlComQIPtrAssign((IUnknown**)&p, lp, __uuidof(T)));
    }
    return *this;
  }

  T* operator=(_In_ const Ptr<T>& lp) throw() {
    if (*this != lp) {
      return static_cast<T*>(AtlComPtrAssign((IUnknown**)&p, lp));
    }
    return *this;
  }

  // additional operators for compatibility with CComPtr
  template <typename Q> T* operator=(_In_ const CComPtr<Q>& lp) throw() {
    if (!IsEqualObject(lp)) {
      return static_cast<T*>(AtlComQIPtrAssign((IUnknown**)&p, lp, __uuidof(T)));
    }
    return *this;
  }

  T* operator=(_In_ const CComPtr<T>& lp) throw() {
    if (*this != lp) {
      return static_cast<T*>(AtlComPtrAssign((IUnknown**)&p, lp));
    }
    return *this;
  }
};

/*==============================================================================
 * template class QIPtr
 *
 * Basic CComQIPtr replacement. See Ptr.
 */
template <class T, const IID* piid = &__uuidof(T)> class QIPtr : public Ptr<T>
{
public:
  QIPtr() throw() {
  }

  QIPtr(_In_opt_ T* lp) throw() :
      Ptr<T>(lp) {
  }

  QIPtr(_In_ const QIPtr<T,piid>& lp) throw() :
      Ptr<T>(lp.p) {
  }

  // additional ctor for compatibility with CComQIPtr
  QIPtr(_In_ const CComQIPtr<T,piid>& lp) throw() :
      Ptr<T>(lp.p) {
  }

  QIPtr(_In_opt_ IUnknown* lp) throw() {
    if (lp != NULL) {
      lp->QueryInterface(*piid, (void **)&p);
    }
  }

  T* operator=(_In_opt_ T* lp) throw() {
    if (*this != lp) {
      return static_cast<T*>(AtlComPtrAssign((IUnknown**)&p, lp));
    }
    return *this;
  }

  T* operator=(_In_ const QIPtr<T,piid>& lp) throw()
  {
    if (*this != lp) {
      return static_cast<T*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
    }
    return *this;
  }

  // additional operator for compatibility with CComQIPtr
  T* operator=(_In_ const CComQIPtr<T,piid>& lp) throw()
  {
    if (*this != lp) {
      return static_cast<T*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
    }
    return *this;
  }

  T* operator=(_In_opt_ IUnknown* lp) throw() {
    if (*this != lp) {
      return static_cast<T*>(AtlComQIPtrAssign((IUnknown**)&p, lp, *piid));
    }
    return *this;
  }
};

//Specialization to make it work
template<> class QIPtr<IUnknown, &IID_IUnknown> : public Ptr<IUnknown>
{
public:
  QIPtr() throw() {
  }

  QIPtr(_In_opt_ IUnknown* lp) throw() {
    //Actually do a QI to get identity
    if (lp != NULL) {
      lp->QueryInterface(__uuidof(IUnknown), (void **)&p);
    }
  }

  QIPtr(_In_ const QIPtr<IUnknown,&IID_IUnknown>& lp) throw() :
      Ptr<IUnknown>(lp.p) {
  }

  IUnknown* operator=(_In_opt_ IUnknown* lp) throw() {
    if (*this != lp) {
     //Actually do a QI to get identity
      return AtlComQIPtrAssign((IUnknown**)&p, lp, __uuidof(IUnknown));
    }
    return *this;
  }

  IUnknown* operator=(_In_ const QIPtr<IUnknown,&IID_IUnknown>& lp) throw()
  {
    if (*this != lp) {
      return AtlComPtrAssign((IUnknown**)&p, lp.p);
    }
    return *this;
  }

  // additional operator for compatibility with CComQIPtr
  IUnknown* operator=(_In_ const CComQIPtr<IUnknown,&IID_IUnknown>& lp) throw()
  {
    if (*this != lp) {
      return AtlComPtrAssign((IUnknown**)&p, lp.p);
    }
    return *this;
  }
};

/*==============================================================================
 * template class DispatchPtr
 *
 * The methods returning something (getProperty, call..) return a CComVariant
 * instead of a HRESULT for convenience.
 * The HRESULT of the call is still available as an optional out argument.
 *
 */
template <class T> class DispatchPtr : public QIPtr<T>
{
public:
  DispatchPtr() throw() : QIPtr<T>() {
  }

  DispatchPtr(_In_opt_ T* lp) throw() :
    QIPtr<T>(lp) {
  }

  DispatchPtr(_In_ const DispatchPtr & lp) throw() :
    QIPtr<T>(lp.p) {
  }

  DispatchPtr(_In_opt_ IUnknown* lp) throw() {
    if (lp != NULL) {
      lp->QueryInterface(__uuidof(T), (void **)&p);
    }
  }

  T* operator=(_In_ const DispatchPtr & lp) throw() {
    if(*this != lp) {
      return static_cast<T*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
    }
    return *this;
  }

  T* operator=(_In_opt_ IUnknown* lp) throw() {
    if(*this != lp) {
      return static_cast<T*>(AtlComQIPtrAssign((IUnknown**)&p, lp, __uuidof(T)));
    }
    return *this;
  }

  //----------------------------------------------------------------------------
  // Get a DISPID from a name
  DISPID getDispID(_In_ LPCOLESTR aName, _Out_opt_ HRESULT * aResultPtr = NULL)
      throw() {
    ATLASSERT(p);
    HRESULT hr = E_FAIL;
    DISPID dispID = DISPID_UNKNOWN;
    if (p) {
      hr = p->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR*>(&aName), 1,
          LOCALE_USER_DEFAULT, &dispID);
    }
    RET_HANDLE_ERROR_PTR(hr, aResultPtr, dispID);
    return dispID;
  }

  //----------------------------------------------------------------------------
  // Get a property by DISPID
  CComVariant getProperty(_In_ const DISPID dispID,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    return getProperty(p, dispID, aResultPtr);
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
    RET_HANDLE_ERROR_PTR(hr, aResultPtr, (CComVariant(NULL)));
    return CComVariant(NULL);
  }

  //----------------------------------------------------------------------------
  // Put a property by DISPID
  _Check_return_ HRESULT putProperty(_In_ DISPID dispID, _In_ CComVariant aValue)
      throw() {
    return putProperty(p, dispID, aValue);
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
    return call(p, aDispID, aArguments, aResultPtr);
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

  template<typename T> struct VariantCastTypeTraitAndAccessor { };

  template<> struct VariantCastTypeTraitAndAccessor<SHORT> {
    static const VARENUM vt_type = VT_I2;
    static SHORT get(VARIANT & vt) {return vt.iVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<LONG> {
    static const VARENUM vt_type = VT_I4;
    static LONG get(VARIANT & vt) {return vt.lVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<FLOAT> {
    static const VARENUM vt_type = VT_R4;
    static FLOAT get(VARIANT & vt) {return vt.fltVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<DOUBLE> {
    static const VARENUM vt_type = VT_R8;
    static DOUBLE get(VARIANT & vt) {return vt.dblVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<CY> {
    static const VARENUM vt_type = VT_CY;
    static CY get(VARIANT & vt) {return vt.cyVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<CComPtr<IDispatch> > {
    static const VARENUM vt_type = VT_DISPATCH;
    static CComPtr<IDispatch> get(VARIANT & vt) {return vt.pdispVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<CComPtr<IUnknown> > {
    static const VARENUM vt_type = VT_UNKNOWN;
    static CComPtr<IUnknown> get(VARIANT & vt) {return vt.punkVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<CComVariant> {
    static const VARENUM vt_type = VT_VARIANT;
    static CComVariant get(VARIANT & vt) {return *vt.pvarVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<CHAR> {
    static const VARENUM vt_type = VT_I1;
    static CHAR get(VARIANT & vt) {return vt.cVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<BYTE> {
    static const VARENUM vt_type = VT_UI1;
    static BYTE get(VARIANT & vt) {return vt.bVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<USHORT> {
    static const VARENUM vt_type = VT_UI2;
    static USHORT get(VARIANT & vt) {return vt.uiVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<ULONG> {
    static const VARENUM vt_type = VT_UI4;
    static ULONG get(VARIANT & vt) {return vt.ulVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<LONGLONG> {
    static const VARENUM vt_type = VT_I8;
    static LONGLONG get(VARIANT & vt) {return vt.llVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<ULONGLONG> {
    static const VARENUM vt_type = VT_UI8;
    static ULONGLONG get(VARIANT & vt) {return vt.ullVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<CComBSTR> {
    static const VARENUM vt_type = VT_BSTR;
    static CComBSTR get(VARIANT & vt) {return vt.bstrVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<CString> {
    static const VARENUM vt_type = VT_BSTR;
    static CString get(VARIANT & vt) {return vt.bstrVal;}
  };

  template<> struct VariantCastTypeTraitAndAccessor<std::wstring> {
    static const VARENUM vt_type = VT_BSTR;
    static std::wstring get(VARIANT & vt) {return vt.bstrVal;}
  };

  //----------------------------------------------------------------------------
  // Convenience method: Get a property of a certain type, do a type conversion
  // if required.
  // T is the type of the return value.
  template <class T> T getAs(_In_ LPOLESTR aName,
      _Out_opt_ HRESULT * aResultPtr = NULL) {
    HRESULT hr = E_FAIL;
    typedef VariantCastTypeTraitAndAccessor<T> accessor;
    CComVariant vtResult(getProperty(aName, &hr));
    if (SUCCEEDED(hr)) {
      // If you receive compiler error C2039 here you are trying to get a type
      // without an associated VariantCastTypeTraitAndAccessor.
      // Go ahead and create one.
      hr = vtResult.ChangeType(accessor::vt_type);
      if (SUCCEEDED(hr)) {
        return accessor::get(vtResult);
      }
    }
    RET_HANDLE_ERROR_PTR(hr, aResultPtr, (T()));
    return T();
  }

  //----------------------------------------------------------------------------
  // STATIC: Get a property by DISPID
  _Check_return_ CComVariant static getProperty(_In_ IDispatch* p,
    _In_ const DISPID aDispID, _Out_opt_ HRESULT * aResultPtr = NULL) {
    ATLASSERT(p);
    CComVariant retVal;
    if(p == NULL) {
      RET_HANDLE_ERROR_PTR(E_INVALIDARG, aResultPtr, retVal);
    }

    ATLTRACE(atlTraceCOM, 2, _T("DispatchPtr::getProperty\n"));
    DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
    HRESULT hr = p->Invoke(aDispID, IID_NULL, LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYGET, &dispparamsNoArgs, &retVal, NULL, NULL);
    RET_HANDLE_ERROR_PTR(hr, aResultPtr, retVal);
  }

  //----------------------------------------------------------------------------
  // STATIC: Put a property by DISPID
  _Check_return_ static HRESULT putProperty(_Inout_ IDispatch* p,
      _In_ const DISPID aDispID, _In_ CComVariant aValue) throw() {
    ATLASSERT(p);
    if(!p) {
      return E_INVALIDARG;
    }

    ATLTRACE(atlTraceCOM, 2, _T("DispatchPtr::putProperty\n"));
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
      RET_HANDLE_ERROR_PTR(E_INVALIDARG, aResultPtr, retVal);
    }

    HRESULT hr = p->Invoke(aDispID, IID_NULL, LOCALE_USER_DEFAULT,
        DISPATCH_METHOD, aArguments, &retVal, NULL, NULL);
    RET_HANDLE_ERROR_PTR(hr, aResultPtr, retVal);
  }
};

typedef DispatchPtr< IDispatch > ConvIDispatchPtr;

/*==============================================================================
 * class DispatchExPtr
 *
 * The same like DispatchPtr, but for the IDispatchEx interface
 */
class DispatchExPtr : public DispatchPtr< IDispatchEx >
{
public:
  DispatchExPtr() throw() : DispatchPtr< IDispatchEx >() {
  }

  DispatchExPtr(_In_opt_ IDispatchEx* lp) throw() :
    DispatchPtr< IDispatchEx >(lp) {
  }

  DispatchExPtr(_In_ const DispatchExPtr & lp) throw() :
    DispatchPtr< IDispatchEx >(lp.p) {
  }

  DispatchExPtr(_In_opt_ IUnknown* lp) throw() :
    DispatchPtr< IDispatchEx >(lp) {
  }

  IDispatchEx* operator=(_In_ const DispatchExPtr & lp) throw() {
    if(*this != lp) {
      return static_cast<IDispatchEx*>(AtlComPtrAssign((IUnknown**)&p, lp.p));
    }
    return *this;
  }

  IDispatchEx* operator=(_In_opt_ IUnknown* lp) throw() {
    if(*this != lp) {
      return static_cast<IDispatchEx*>(AtlComQIPtrAssign((IUnknown**)&p, lp,
          __uuidof(IDispatchEx)));
    }
    return *this;
  }

  //----------------------------------------------------------------------------
  // Get an ex(pando) DISPID from a name
  _Check_return_ DISPID getExDispID(_In_ LPCOLESTR aName,
      _In_ const BOOL aEnsureExists = TRUE,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    DWORD gfrdex = (aEnsureExists) ? fdexNameEnsure : 0;
    DISPID dispID = DISPID_UNKNOWN;
    HRESULT hr = p->GetDispID(CComBSTR(aName), gfrdex, &dispID);
    RET_HANDLE_ERROR_PTR(hr, aResultPtr, dispID);
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
    RET_HANDLE_ERROR_PTR(hr, aResultPtr, (CComVariant(NULL)));
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
  template <class T> T getExAs(_In_ LPOLESTR aName,
      _Out_opt_ HRESULT * aResultPtr = NULL) {
    HRESULT hr = E_FAIL;
    typedef VariantCastTypeTraitAndAccessor<T> trait;
    CComVariant vtResult(getExProperty(aName, &hr));
    if (SUCCEEDED(hr)) {
      hr = vtResult.ChangeType(trait::vt_type);
      if (SUCCEEDED(hr)) {
        return trait::get(vtResult);
      }
    }
    RET_HANDLE_ERROR_PTR(hr, aResultPtr, (T()));
  }

  //----------------------------------------------------------------------------
  // STATIC: Put an ex-property by DISPID
  _Check_return_ static HRESULT putExProperty(_Inout_ IDispatchEx* p,
      _In_ DISPID aDispID, _In_ VARIANT & aValue) throw() {
    ATLASSERT(p);
    if(p == NULL) {
      return E_INVALIDARG;
    }

    ATLTRACE(atlTraceCOM, 2, _T("DispatchExPtr::putExProperty\n"));
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
      RET_HANDLE_ERROR_PTR(E_INVALIDARG, aResultPtr, retVal);
    }

    ATLTRACE(atlTraceCOM, 2, _T("DispatchExPtr::getExProperty\n"));
    DISPPARAMS dispparamsNoArgs = {NULL, NULL, 0, 0};
    HRESULT hr = p->InvokeEx(aDispID, LOCALE_USER_DEFAULT,
        DISPATCH_PROPERTYGET, &dispparamsNoArgs, &retVal, NULL, NULL);
    RET_HANDLE_ERROR_PTR(hr, aResultPtr, retVal);
  }

  //----------------------------------------------------------------------------
  // STATIC: Wrapper for InvokeEx as a function call
  _Check_return_ static CComVariant callEx(_In_ IDispatchEx* p,
      _In_ const DISPID aDispID, _In_ DispatchArgs & aArguments,
      _Out_opt_ HRESULT * aResultPtr = NULL) throw() {
    ATLASSERT(p);
    CComVariant retVal;
    if(p == NULL) {
      RET_HANDLE_ERROR_PTR(E_INVALIDARG, aResultPtr, retVal);
    }

    HRESULT hr = p->InvokeEx(aDispID, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
        aArguments, &retVal, NULL, NULL);
    RET_HANDLE_ERROR_PTR(hr, aResultPtr, retVal);
  }

};

#undef RET_HANDLE_ERROR_PTR

} // namespace COMvenience
