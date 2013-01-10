/*******************************************************************************
 * COMvenience : COM/ATL a bit more convenient
 * Copyright 2013 Salsita Software http://www.salsitasoft.com/
 * Author: Arne Seib <kontakt@seiberspace.de>
 *
 * File: _implementQIPtr.h
 *
 ******************************************************************************/

// Helper macro. Constructors and copy operators for a CComQIPtr<> derived
// class. Used by specialized safepointer implementations like ConvDispPtr.
#define IMPLEMENT_QI_PTR(_clsname, _iface) \
  _clsname() throw() : CComQIPtr<_iface>() { \
  } \
  _clsname(_In_opt_ _iface* lp) throw() : \
    CComQIPtr<_iface>(lp) { \
  } \
  _clsname(_In_ const _clsname & lp) throw() : \
    CComQIPtr<_iface>(lp.p) { \
  } \
  _clsname(_In_opt_ IUnknown* lp) throw() { \
    if (lp != NULL) { \
      lp->QueryInterface(__uuidof(_iface), (void **)&p); \
    } \
  } \
  _iface* operator=(_In_ const _clsname & lp) throw() { \
    if(*this != lp) { \
      return static_cast<_iface*>(AtlComPtrAssign((IUnknown**)&p, lp.p)); \
    } \
    return *this; \
  } \
  _iface* operator=(_In_opt_ IUnknown* lp) throw() { \
    if(*this != lp) { \
      return static_cast<_iface*>(AtlComQIPtrAssign((IUnknown**)&p, lp, __uuidof(_iface))); \
    } \
    return *this; \
  }
