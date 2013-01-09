/*******************************************************************************
 * COMvenience : COM/ATL a bit more convenient
 * Copyright 2013 Salsita Software http://www.salsitasoft.com/
 * Author: Arne Seib <kontakt@seiberspace.de>
 *
 * File: comvtypes.h
 *
 * Some helpful and often used types.
 * 
 ******************************************************************************/

#pragma once

#include <map>
#include <vector>
#include <string>

/*==============================================================================
 * namespace
 */
namespace COMvenience
{
/*==============================================================================
 * frequently used types
 */

typedef std::map<std::wstring, CComPtr<IDispatch> > DispatchMap;
typedef std::vector<CComVariant> VariantVector;

} // namespace COMvenience
