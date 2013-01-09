/*******************************************************************************
 * COMvenience : COM/ATL a bit more convenient
 * Copyright 2013 Salsita Software http://www.salsitasoft.com/
 * Author: Arne Seib <kontakt@seiberspace.de>
 *
 * File: comvmacros.h
 *
 * Some helpful macros for COM/ATL
 *
 * RETURN_IF_FAILED(hr), RETURN_IF_FAILED2(hr, retVal)
 *  does basically a
 *   if (FAILED(hr)) return hr;
 *  but with some extrastuff for assertions and tracing.
 *  Note: It is safely possible to do
 *   RETURN_IF_FAILED(someFunctionCall());
 *  someFunctionCall() will still be executed only once.
 *  RETURN_IF_FAILED2(hr, retVal) returns retVal instead of hr.
 *
 * RETURN_IF_PTR_NULL(val)
 *  Shortcut for the frequent out argument checking:
 *   if (!ptr) return E_POINTER;
 *
 ******************************************************************************/

#pragma once

#ifndef COMVASSERT
# ifdef COMV_DEBUG_BREAK
#  define COMVASSERT ATLASSERT
# else
#  define COMVASSERT
# endif
#endif //ndef COMVASSERT

#ifndef COMVTRACE
# ifdef ATLTRACE
#  define COMVTRACE ATLTRACE
# else
#  define COMVTRACE
# endif
#endif //ndef COMVTRACE

#define RETURN_IF_FAILED(_hr) \
  do \
  { \
    HRESULT _hr__ = _hr; \
    COMVASSERT(SUCCEEDED(_hr__)); \
    if (FAILED(_hr__)) \
    { \
      COMVTRACE(_T("ASSERTION FAILED: 0x%08x in "), _hr__); \
      COMVTRACE(__FILE__); \
      COMVTRACE(_T(" line %i\n"), __LINE__); \
      return _hr__; \
    } \
  } while(0);

#define RETURN_IF_FAILED2(_hr, _ret) \
  do \
  { \
    HRESULT _hr__ = _hr; \
    COMVASSERT(SUCCEEDED(_hr__)); \
    if (FAILED(_hr__)) \
    { \
      COMVTRACE(_T("ASSERTION FAILED: 0x%08x in "), _hr__); \
      COMVTRACE(__FILE__); \
      COMVTRACE(_T(" line %i\n"), __LINE__); \
      return _ret; \
    } \
  } while(0);

#define RETURN_IF_PTR_NULL(_val) \
  if (!_val) return E_POINTER;
