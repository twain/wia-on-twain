/***************************************************************************
* Copyright © 2009 TWAIN Working Group:  
*   Adobe Systems Incorporated, AnyDoc Software Inc., Eastman Kodak Company, 
*   Fujitsu Computer Products of America, JFL Peripheral Solutions Inc., 
*   Ricoh Corporation, and Xerox Corporation.
* All rights reserved.
*
* Redistribution and use in source and binary forms, with or without
* modification, are permitted provided that the following conditions are met:
*     * Redistributions of source code must retain the above copyright
*       notice, this list of conditions and the following disclaimer.
*     * Redistributions in binary form must reproduce the above copyright
*       notice, this list of conditions and the following disclaimer in the
*       documentation and/or other materials provided with the distribution.
*     * Neither the name of the TWAIN Working Group nor the
*       names of its contributors may be used to endorse or promote products
*       derived from this software without specific prior written permission.
*
* THIS SOFTWARE IS PROVIDED BY TWAIN Working Group ``AS IS'' AND ANY
* EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
* WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
* DISCLAIMED. IN NO EVENT SHALL TWAIN Working Group BE LIABLE FOR ANY
* DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
* (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
* LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
* ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
* (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
* SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*
***************************************************************************/
/**
 * @file stdafx.h
 * include file for standard system include files,
 * or project specific include files that are used frequently, but
 * are changed infrequently
 * @author TWAIN Working Group
 * @date October 2009
 */

#pragma once

#pragma warning( disable : 6387 )
#pragma warning( disable : 6011 )
#pragma warning (disable : 6309 )

#ifndef WINVER                // Specifies that the minimum required platform is Windows Vista.
#define WINVER 0x0501         // Change this to the appropriate value to target other versions of Windows.
#endif

#ifndef _WIN32_WINNT          // Specifies that the minimum required platform is Windows Vista.
#define _WIN32_WINNT 0x0600   // Change this to the appropriate value to target other versions of Windows.
#endif

#ifndef _WIN32_WINDOWS        // Specifies that the minimum required platform is Windows 98.
#define _WIN32_WINDOWS 0x0500 // Change this to the appropriate value to target Windows Me or later.
#endif

#ifndef _WIN32_IE             // Specifies that the minimum required platform is Internet Explorer 7.0.
#define _WIN32_IE 0x0600      // Change this to the appropriate value to target other versions of IE.
#endif


#define WIN32_LEAN_AND_MEAN   // Exclude rarely-used stuff from Windows headers
#include <windows.h>

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS      // some CString constructors will be explicit
#define _SAFECRT_IMPL
#include <atlstr.h>
#include <initguid.h>
#include "basicstr.h"
#include "basicarray.h"       // CSimpleDynamicArray

#pragma warning( default : 6387 )
#pragma warning( default : 6011 )
#pragma warning (default : 6309 )

extern HINSTANCE g_hInst;
