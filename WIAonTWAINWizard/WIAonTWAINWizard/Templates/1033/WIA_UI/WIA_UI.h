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
 * @file WIA_UI.h
 * CWiaUIExtension2, CUIEXT2App, CSimpleWIACallback and CUIExt2ClassObject classes definitions
 * @author TWAIN Working Group
 * @date October 2009
 */
#pragma once

#ifndef __AFXWIN_H__
	#error "include 'stdafx.h' before including this file for PCH"
#endif
#include "resource.h"		// main symbols

class CWiaUIExtension2 : public IWiaUIExtension2, public IShellExtInit, public IShellPropSheetExt
{
public:
  //IUnknown
  STDMETHODIMP QueryInterface(const IID& iid_requested, void** ppInterfaceOut);
  STDMETHODIMP_(ULONG) AddRef(void);
  STDMETHODIMP_(ULONG) Release(void);
  //IWiaUIExtension2
  STDMETHODIMP DeviceDialog( PDEVICEDIALOGDATA2 pDeviceDialogData );
  STDMETHODIMP GetDeviceIcon( BSTR bstrDeviceId, HICON *phIcon, ULONG nSize ){return E_NOTIMPL;};

  // IShellExtInit
  STDMETHODIMP Initialize (LPCITEMIDLIST pidlFolder,LPDATAOBJECT lpdobj,HKEY hkeyProgID);

  // IShellPropSheetExt
  STDMETHODIMP AddPages (LPFNADDPROPSHEETPAGE lpfnAddPage,LPARAM lParam);
  STDMETHODIMP ReplacePage (UINT uPageID,LPFNADDPROPSHEETPAGE lpfnReplacePage,LPARAM lParam) {return E_NOTIMPL;};

  CWiaUIExtension2();
  ~CWiaUIExtension2();

  /**
  * Transfers the file from the given wia item to the file name specified in the pDeviceDialogData
  * @param[in] pDeviceDialogData a PDEVICEDIALOGDATA2.
  * @param[in] pWiaFlatbed a pointer to IWiaItem2. Child item - feeder or flatbed
  * @return S_OK on success
  */
  HRESULT TransferFromWiaItem( PDEVICEDIALOGDATA2 pDeviceDialogData, IWiaItem2 *pWiaFlatbed);

private:
  LONG                m_nRefCount;          /**< reference counter*/
  CPropPage           *m_PP;                /**< pointer to Property page class. Used with WIA 1.0 applications only*/
};


/**
* CUIEXT2App class. 
* The main DLL class.
* @author 
*/
class CUIEXT2App : public CWinApp
{
public:
	CUIEXT2App();

  LONG m_nLocks;                  /**< counter of COM class instances*/
  CStringArray m_FileNamesArray;  /**< list of transfered image files */
  CWiaUIExtension2* m_pThisClass; /**< pointer to last instantiated CWiaUIExtension2. */

  /**
  * Inc m_nLocks
  */
  void LockModule(void);   
  /**
  * Dec m_nLocks
  */
  void UnlockModule(void);	      

// Overrides
	virtual BOOL InitInstance();
	DECLARE_MESSAGE_MAP()
};

class CSimpleWIACallback : public IWiaTransferCallback
{
private:
  volatile LONG m_cRef;           /**< Reference counter*/
  CString  m_strPath;             /**< Path to dest image directory*/
  CString  m_strFileName;         /**< Template name for image file name*/
      
public:
  STDMETHODIMP QueryInterface(const IID &iid_requested, void** ppInterfaceOut);
  STDMETHODIMP_(ULONG)AddRef();
  STDMETHODIMP_(ULONG) Release();
  STDMETHODIMP GetNextStream(LONG lFlags, BSTR bstrItemName, BSTR bstrFullItemName, IStream **ppDestination);
  STDMETHODIMP TransferCallback(LONG lFlags, WiaTransferParams *pWiaTransferParams);
  
  /**
  * Constuctor
  * @param[in] strPath a CString. Path to dest image directory
  * @param[in] strFileName aCString. Template name for image file name
  * @return S_OK on success
  */
  CSimpleWIACallback(  CString strPath, CString strFileName  );
};

class CUIExt2ClassObject : public IClassFactory
{
public:
  STDMETHODIMP QueryInterface(const IID& iid_requested, void** ppInterfaceOut);
  STDMETHODIMP_(ULONG) AddRef(void);   
  STDMETHODIMP_(ULONG) Release(void);
  STDMETHODIMP CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv);
  STDMETHODIMP LockServer(BOOL bLock);
};



