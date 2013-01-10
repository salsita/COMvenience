/*******************************************************************************
 * COMvenience : COM/ATL a bit more convenient
 * Copyright 2013 Salsita Software http://www.salsitasoft.com/
 * Author: Arne Seib <kontakt@seiberspace.de>
 *
 * File: comvDispatchArgs.h
 *
 * class DispatchArgs
 * A DISPPARAMS wrapper.
 *
 ******************************************************************************/

#pragma once

/*==============================================================================
 * namespace
 */
namespace COMvenience
{

/*==============================================================================
 * class DispatchArgs
 * Wraps DISPPARAMS.
 * DISPPARAMS argument lists are in reversed order. This class takes care of the
 * right order by managing a stack that grows downwards.
 * It has a built-in operator DISPPARAMS*, so it can be used directly in calls
 * to Invoke(Ex).
 * By default it allocates 4 values, and it grows by 4 values at a time. Most
 * Invoke calls will not have more than 4 arguments, so this should be
 * sufficiant.
 * It also allows setting a "this" pointer to allow proper member function
 * calling on JS objects.
 */
class DispatchArgs
{
public:
  //----------------------------------------------------------------------------
  DispatchArgs() throw() : mData(NULL), mCapacity(0), mCurrent(0) {
    memset(&mParams, 0, sizeof(mParams));
  }

  //----------------------------------------------------------------------------
  ~DispatchArgs() {
    clear();
  }

  //----------------------------------------------------------------------------
  _Check_return_ HRESULT add(_In_ const CComVariant vt) throw() {
    // Make sure we have enough place.
    HRESULT hr = resize(mCapacity - mCurrent + 1);
    if (FAILED(hr)) {
      return hr;
    }
    if (mParams.rgdispidNamedArgs) {
      // Have already a this pointer - shift it down by one.
      // Since it IS possible to modify the underlying DISPPARAMS via the
      // DISPPARAMS* operator let's make sure we really have named arguments
      // for "this".
      ATLASSERT(1 == mParams.cNamedArgs);
      ATLASSERT((*mParams.rgdispidNamedArgs) == sDISPID_THIS);
      ATLASSERT(mData);
      // The current index is the "this" pointer, so shift it down ...
      mData[mCurrent-1] = mData[mCurrent];
      // ... and set the new value at current...
      mData[mCurrent] = vt;
      // and THEN decrement current index.
      --mCurrent;
    }
    else {
      // There is no "this" pointer, so decrement FIRST...
      --mCurrent;
      // ... and set the new value at the new current.
      mData[mCurrent] = vt;
    }
    // update DISPPARAMS
    mParams.cArgs = mCapacity - mCurrent;
    mParams.rgvarg = mData + mCurrent;
    return S_OK;
  }

  //----------------------------------------------------------------------------
  _Check_return_ HRESULT setThis(_In_ IDispatch * aDispatch) throw() {
    ATLASSERT(aDispatch);
    if (!aDispatch) {
      return E_INVALIDARG;
    }
    if (mParams.rgdispidNamedArgs) {
      // Have already a "this" pointer - overwrite it.
      // Since it IS possible to modify the underlying DISPPARAMS via the
      // DISPPARAMS* operator let's make sure we really have named arguments for
      // "this".
      ATLASSERT(1 == mParams.cNamedArgs);
      ATLASSERT((*mParams.rgdispidNamedArgs) == sDISPID_THIS);
      ATLASSERT(mData);
      // The "this" pointer is in any case the bottommost (means indexed by
      // mCurrent) element.
      mData[mCurrent] = aDispatch;
    }
    else {
      // Don't have a "this" pointer yet, simply add...
      add(aDispatch);
      // ... and adjust DISPPARAMS
      mParams.rgdispidNamedArgs = const_cast<DISPID*>(&sDISPID_THIS);
      mParams.cNamedArgs = 1;
    }
    return S_OK;
  }

  //----------------------------------------------------------------------------
  void clear() throw() {
    if (mData) {
      for (UINT n = 0; n < mCapacity; n++) {
        mData[n].Clear();
      }
      free(mData);
      mData = NULL;
    }
    memset(&mParams, 0, sizeof(mParams));
    mCapacity = 0;
    mCurrent = 0;
  }

  //----------------------------------------------------------------------------
  operator DISPPARAMS * () throw() {
    return &mParams;
  }

private:
  //----------------------------------------------------------------------------
  static const UINT sGROW_BY = 4;
  static const DISPID sDISPID_THIS = DISPID_THIS;

  //----------------------------------------------------------------------------
  // Forbid copying
  DispatchArgs(_In_ const DispatchArgs &);
  DispatchArgs & operator = (_In_ const DispatchArgs &);

  //----------------------------------------------------------------------------
  _Check_return_ HRESULT resize(_In_ const UINT aNewCapacity) {
    if(aNewCapacity <= mCapacity) {
      return S_FALSE;
    }
    UINT newCapacity = ((aNewCapacity / sGROW_BY) + 1) * sGROW_BY;
    CComVariant * data = mData;
    mData = (CComVariant*) malloc(newCapacity * sizeof(CComVariant));
    if (!mData) {
      return E_OUTOFMEMORY;
    }
    // Instead of calling the constructor for each entry we can set the whole
    // allocated memory to 0. CComVariant is binary compatible to VARIANT, so
    // this is safe.
    memset(mData, 0, newCapacity * sizeof(CComVariant));
    if (data) {
      // Note: Copying the raw variant data transfers ownership of the
      // data to the newly created CComVariants, so there is no need
      // to call delete on each of the old entries.
      memcpy(mData + newCapacity - mCapacity, data,
          mCapacity * sizeof(CComVariant));
      // Free previously allocated array WITHOUT deleting each entry.
      free(data);
    }
    mCapacity = newCapacity;
    mCurrent = sGROW_BY;
    return S_OK;
  }

  //----------------------------------------------------------------------------
  CComVariant * mData;
  UINT          mCapacity;
  UINT          mCurrent;
  DISPPARAMS    mParams;
};

} // namespace COMvenience
