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
 * @file WIA_UI.cpp
 * CWiaUIExtension2, CUIEXT2App, CSimpleWIACallback and CUIExt2ClassObject implementation
 * @author TWAIN Working Group
 * @date October 2009
 */

#include "stdafx.h"

#include <Shlwapi.h>
#include "PropPage.h"
#include "WIA_UI.h"
#include "UIDlg.h"
#include "TWAIN_API.h"
#include "utilities.h"

//{[!output GUID_CLASS_UI1]}
static const GUID CLSID_WiaUIExt2 = 
[!output GUID_CLASS_UI2];

HINSTANCE g_hInst = 0;
CUIEXT2App theApp;

/*****************************************************************************
 *
 *  CUIEXT2App
 *
 *******************************************************************************/
BEGIN_MESSAGE_MAP(CUIEXT2App, CWinApp)
END_MESSAGE_MAP()

CUIEXT2App::CUIEXT2App()
{
  m_nLocks = 0;
  m_FileNamesArray.RemoveAll();
  m_pThisClass=NULL;
}

BOOL CUIEXT2App::InitInstance()
{
	CWinApp::InitInstance();

  g_hInst = m_hInstance;
	return TRUE;
}

void CUIEXT2App::LockModule(void)   
{ 
  InterlockedIncrement(&m_nLocks); 
}

void CUIEXT2App::UnlockModule(void) 
{ 
  InterlockedDecrement(&m_nLocks); 
}

/*****************************************************************************
 *
 *  CWiaUIExtension2
 *
 *******************************************************************************/
CWiaUIExtension2::CWiaUIExtension2(): 
  m_nRefCount(0),
  m_PP(NULL) 
{
  if(theApp.m_pThisClass)
  {
    theApp.m_pThisClass->Release();
  }  
  theApp.m_pThisClass=this;
};

CWiaUIExtension2::~CWiaUIExtension2() 
{
  if(m_PP)
  {
    delete m_PP;
  }
};

STDMETHODIMP CWiaUIExtension2::Initialize (LPCITEMIDLIST pidlFolder,
                                    LPDATAOBJECT lpdobj, HKEY hkeyProgID)
{
  UNREFERENCED_PARAMETER(pidlFolder);
  UNREFERENCED_PARAMETER(hkeyProgID);

  LONG lType = 0;
  HRESULT hr = NOERROR;
  
  if (!lpdobj)
  {
      return E_INVALIDARG;
  }
  CComPtr<IWiaItem> pItem;


  // For singular selections, the WIA namespace should always provide a dataobject that also supports IWiaItem
  if (FAILED(lpdobj->QueryInterface (IID_IWiaItem, reinterpret_cast<LPVOID*>(&pItem))))
  {
    hr = E_FAIL;
  }
  if (SUCCEEDED(hr))
  {
    pItem->GetItemType (&lType);
    if ((lType & WiaItemTypeRoot))
    {
      hr = E_FAIL; // we only support changing the property on the child item
    }
    else
    {
      if(m_PP)
      {
        delete m_PP;
      }
      m_PP=new CPropPage(pItem, (IUnknown *)(void*)this);
    }
  }
  return hr;
}

STDMETHODIMP CWiaUIExtension2::AddPages (LPFNADDPROPSHEETPAGE lpfnAddPage,LPARAM lParam)
{
  HPROPSHEETPAGE hpsp;

  HRESULT hr = E_FAIL;
  
  if(m_PP)
  {
    PROPSHEETPAGE psp = m_PP->GetPSP();

    hpsp = CreatePropertySheetPage (&psp);
    if (hpsp)
    {
      hr = (*lpfnAddPage) (hpsp, lParam) ? S_OK:E_FAIL;
      if (SUCCEEDED(hr))
      {
        AddRef();// the propsheetpage will release us when it is destroyed
      }
    }
  }
  return hr;
}

STDMETHODIMP CWiaUIExtension2::QueryInterface(const IID& iid_requested, void** ppInterfaceOut)
{
  HRESULT hr = S_OK;

  hr = ppInterfaceOut ? S_OK : E_POINTER;

  if (SUCCEEDED(hr))
  {
    *ppInterfaceOut = NULL;
  }
  
  //  We support IID_IUnknown and IWiaUIExtension2
  if (SUCCEEDED(hr))
  {
    if (IID_IUnknown == iid_requested) 
    {
      *ppInterfaceOut = this;
    }
    else if (IID_IWiaUIExtension2 == iid_requested) 
    {
      *ppInterfaceOut = static_cast<IWiaUIExtension2*>(this);
    }
    else if (IID_IShellExtInit == iid_requested) 
    {
      *ppInterfaceOut = static_cast<IShellExtInit*>(this);
    }
    else if (IID_IShellPropSheetExt == iid_requested) 
    {
      *ppInterfaceOut = static_cast<IShellPropSheetExt*>(this);
    }
    else
    {
      hr = E_NOINTERFACE;
    }
  }

  if (SUCCEEDED(hr))
  {
    reinterpret_cast<IUnknown*>(*ppInterfaceOut)->AddRef();
  }

  return hr;
}

STDMETHODIMP_(ULONG) CWiaUIExtension2::AddRef(void)
{
  if (m_nRefCount == 0)
  {
    theApp.LockModule();
  }

  return InterlockedIncrement(&m_nRefCount);
}

STDMETHODIMP_(ULONG) CWiaUIExtension2::Release(void)
{
  ULONG nRetval = InterlockedDecrement(&m_nRefCount);

  if (0 == nRetval) 
  {
    delete this;
    theApp.UnlockModule();
    theApp.m_pThisClass = NULL;
  }

  return nRetval;
}

STDMETHODIMP CWiaUIExtension2::DeviceDialog(PDEVICEDIALOGDATA2 pDeviceDialogData)
{
  HRESULT             hr = S_OK;
  IEnumWiaItem2*      pEnumItem = NULL;
  IWiaItem2*          pWiaItem = NULL;
  GUID                guidCategory = WIA_CATEGORY_FLATBED;
  TCHAR               bufDialog[MAX_PATH] = {0};
  TCHAR               bufTitle[MAX_PATH]  = {0};

  CWnd* pWndParent = CWnd::FromHandle( pDeviceDialogData->hwndParent );
      
  CUIDlg uiDlg(pWndParent,pDeviceDialogData);

  INT_PTR iRes = uiDlg.DoModal();
  if( iRes == IDCANCEL )
  {
    return S_FALSE;
  }
  
  if( iRes == IDC_SHOW_DEF_UI )
  {
    return E_NOTIMPL;
  }
        
  if( iRes == IDOK)
  {
    CComBSTR bstrRootFullName( L"\0" ); 
     
    IWiaPropertyStorage * pPropertyRootStorage = NULL;
    hr = pDeviceDialogData->pIWiaItemRoot->QueryInterface( IID_IWiaPropertyStorage, (LPVOID *)&pPropertyRootStorage);
    if (SUCCEEDED(hr) )
    {      
      hr = WIA_ReadPropBSTR( pPropertyRootStorage, WIA_IPA_FULL_ITEM_NAME, bstrRootFullName );
      pPropertyRootStorage->Release(); 
    }
      
    if (SUCCEEDED(hr) )
    {        
      LONG lDocumentHandlingSelect = 0;
           
      CComBSTR bstrItemTransferName = bstrRootFullName;
      
      if( uiDlg.m_bFeeder)
      {
        bstrItemTransferName += L"\\Feeder";           
      }
      else
      {
        bstrItemTransferName += L"\\Flatbed";
      }
      hr = pDeviceDialogData->pIWiaItemRoot->FindItemByName(0, bstrItemTransferName, &pWiaItem );
    }

    if (SUCCEEDED(hr))
    {
       hr = TransferFromWiaItem(pDeviceDialogData, pWiaItem);
    }
  }
  return hr;
}

/*****************************************************************************
 *
 *  CWiaUIExtension2
 *
 *  Transfers the file from the given wia item to the file name specified in the 
 *   pDeviceDialogData
 *
 *****************************************************************************/
HRESULT CWiaUIExtension2::TransferFromWiaItem( PDEVICEDIALOGDATA2 pDeviceDialogData, IWiaItem2 *pWiaFlatbed)
{
  HRESULT hr = S_OK;
  IStream * pStream = NULL;
  IWiaTransfer * pWiaTransfer = NULL;
  IWiaPropertyStorage * pPropertyStorage = NULL;

  if (!pDeviceDialogData || !pWiaFlatbed)
  {
    hr = E_INVALIDARG;
  }

  WCHAR *pFileName = (WCHAR*) LocalAlloc(LPTR, sizeof(WCHAR)*MAX_PATH); 
  WCHAR *pUniqueFileName = (WCHAR*) LocalAlloc(LPTR, sizeof(WCHAR)*MAX_PATH); 

  if(!pFileName || !pUniqueFileName)
  {
    hr = E_OUTOFMEMORY;
  }

  if (SUCCEEDED(hr))
  {
    hr = pWiaFlatbed->QueryInterface(IID_IWiaPropertyStorage, (LPVOID *)&pPropertyStorage);

    if (SUCCEEDED(hr))
    {
      PROPSPEC pSpec[1] = {0};
      PROPVARIANT pVar[1] = {0};
      GUID guidFormat = WiaImgFmt_BMP;

      pSpec[0].ulKind = PRSPEC_PROPID;
      pSpec[0].propid = WIA_IPA_FORMAT;

      pVar[0].vt = VT_CLSID;
      pVar[0].puuid = &guidFormat;

      hr = pPropertyStorage->WriteMultiple(1, pSpec, pVar, WIA_IPS_FIRST);

      pPropertyStorage->Release();
    }
  }

  if (SUCCEEDED(hr))
  {
    hr = pWiaFlatbed->QueryInterface(IID_IWiaTransfer, (LPVOID *)&pWiaTransfer);
  }

  if (SUCCEEDED(hr))
  {
    CSimpleWIACallback * pCallback = NULL;

#pragma prefast(suppress:__WARNING_ALIASED_MEMORY_LEAK, "pTransferCallback is freed on release.")
    pCallback = new CSimpleWIACallback(pDeviceDialogData->bstrFolderName,pDeviceDialogData->bstrFilename);

    if (pCallback)
    {
      IWiaTransferCallback * pTransferCallback = NULL;

      hr = pCallback->QueryInterface(IID_IWiaTransferCallback, (LPVOID*)&pTransferCallback);
      
      if (SUCCEEDED(hr))
      {
        theApp.m_FileNamesArray.RemoveAll();
        hr = pWiaTransfer->Download(0, pTransferCallback);
        
        pTransferCallback ->Release();
      }

      pCallback->Release();
    }
  }

// transfer names of the files
  if (SUCCEEDED(hr))
  {
    pDeviceDialogData->lNumFiles = (long) theApp.m_FileNamesArray.GetSize();

    pWiaFlatbed->AddRef();
    pDeviceDialogData->pWiaItem = pWiaFlatbed;

    pDeviceDialogData->pbstrFilePaths = (BSTR *)CoTaskMemAlloc(sizeof(BSTR *)*theApp.m_FileNamesArray.GetSize());

    if (pDeviceDialogData->pbstrFilePaths)
    {
      for( int i = 0 ; i < theApp.m_FileNamesArray.GetSize() ; i++ )
      {            
        pDeviceDialogData->pbstrFilePaths[i] = theApp.m_FileNamesArray.GetAt(i).AllocSysString();            
      }
    }
    else
    {
      hr = E_FAIL;
    }
  }
  else
  {
    CString strErr;
    switch( hr )
    {    
      case WIA_ERROR_GENERAL_ERROR:
        strErr = L"WIA_ERROR_GENERAL_ERROR";
      break;
      case WIA_ERROR_PAPER_JAM:
        strErr = L"WIA_ERROR_PAPER_JAM";
      break;
      case WIA_ERROR_PAPER_EMPTY:
        strErr = L"WIA_ERROR_PAPER_EMPTY";
      break;
      case WIA_ERROR_PAPER_PROBLEM:
        strErr = L"WIA_ERROR_PAPER_PROBLEM";
      break;
      case WIA_ERROR_OFFLINE:
        strErr = L"WIA_ERROR_OFFLINE";
      break;
      case WIA_ERROR_BUSY:
        strErr = L"WIA_ERROR_BUSY";
      break;
      case WIA_ERROR_WARMING_UP:
        strErr = L"WIA_ERROR_WARMING_UP";
      break;
      case WIA_ERROR_USER_INTERVENTION:
        strErr = L"WIA_ERROR_USER_INTERVENTION";
      break;
      case WIA_ERROR_ITEM_DELETED:
        strErr = L"WIA_ERROR_ITEM_DELETED";
      break;
      case WIA_ERROR_DEVICE_COMMUNICATION:
        strErr = L"WIA_ERROR_DEVICE_COMMUNICATION";
      break;
      case WIA_ERROR_DEVICE_LOCKED:
        strErr = L"WIA_ERROR_DEVICE_LOCKED";
      break;
      case WIA_ERROR_COVER_OPEN:
        strErr = L"WIA_ERROR_COVER_OPEN";
      break;
      case WIA_ERROR_LAMP_OFF:
        strErr = L"WIA_ERROR_LAMP_OFF";
      break;             
      default:
        strErr.Format(L"Undefined error: %08X", hr );
      break;
    }    

    MessageBox(NULL, strErr, L"Scan Error", MB_ICONERROR|MB_OK );
     
    for( int i = 0 ; i < theApp.m_FileNamesArray.GetSize() ; i++ )
    {
      DeleteFile( theApp.m_FileNamesArray.GetAt( i ) );
    }
               
    theApp.m_FileNamesArray.RemoveAll();
  }

  if (pWiaTransfer)
  {
    pWiaTransfer->Release();
    pWiaTransfer = NULL;
  }

  if (pFileName)
  {
    LocalFree(pFileName);
    pFileName = NULL;
  }

  if (pUniqueFileName)
  {
    LocalFree(pUniqueFileName);
    pUniqueFileName = NULL;
  }

  return hr;
}


/*****************************************************************************
 *
 *  ClassFactory
 *
 *******************************************************************************/
STDMETHODIMP CUIExt2ClassObject::QueryInterface(const IID& iid_requested, void** ppInterfaceOut)
{
  HRESULT hr = S_OK;

  hr = ppInterfaceOut ? S_OK : E_POINTER;

  if (SUCCEEDED(hr))
  {
    *ppInterfaceOut = NULL;
  }

  //  We only support IID_IUnknown and IID_IClassFactory
  if (SUCCEEDED(hr))
  {
    if (IID_IUnknown == iid_requested) 
    {
      *ppInterfaceOut = static_cast<IUnknown*>(this);
    }
    else if (IID_IClassFactory == iid_requested) 
    {
      *ppInterfaceOut = static_cast<IClassFactory*>(this);
    }
    else
    {
      hr = E_NOINTERFACE;
    }
  }

  if (SUCCEEDED(hr))
  {
    reinterpret_cast<IUnknown*>(*ppInterfaceOut)->AddRef();
  }

  return hr;
}

STDMETHODIMP_(ULONG) CUIExt2ClassObject::AddRef(void)
{
  theApp.LockModule();
  return 2;
}

STDMETHODIMP_(ULONG) CUIExt2ClassObject::Release(void)
{
  theApp.UnlockModule();
  return 1;
}

STDMETHODIMP CUIExt2ClassObject::CreateInstance(IUnknown *pUnkOuter,REFIID riid, void **ppv)                        
{
  CWiaUIExtension2  *pExt = NULL;
  HRESULT      hr;

  hr = ppv ? S_OK : E_POINTER;

  if (SUCCEEDED(hr))
  {
    *ppv = 0;
  }

  if (SUCCEEDED(hr))
  {
    if (pUnkOuter)
    {
      hr = CLASS_E_NOAGGREGATION;
    }
  }

  if (SUCCEEDED(hr))
  {
#pragma prefast(suppress:__WARNING_ALIASED_MEMORY_LEAK, "pExt is freed on release.")
    pExt = new CWiaUIExtension2();

    hr = pExt ? S_OK : E_OUTOFMEMORY;
  }

  if (SUCCEEDED(hr))
  {
    pExt->AddRef();
    hr = pExt->QueryInterface(riid, ppv);
    pExt->Release();
  }

  return hr;
}

STDMETHODIMP CUIExt2ClassObject::LockServer(BOOL bLock)
{
  if (bLock)
  {
    theApp.LockModule();
  }
  else
  {
    theApp.UnlockModule();
  }

  return S_OK;
}

/*****************************************************************************
 *
 *  CSimpleWIACallback
 *
 *****************************************************************************/
STDMETHODIMP CSimpleWIACallback::QueryInterface(const IID &iid_requested, void** ppInterfaceOut)
{
  HRESULT hr = S_OK;

  hr = ppInterfaceOut ? S_OK : E_POINTER;

  if (SUCCEEDED(hr))
  {
    *ppInterfaceOut = NULL;
  }

  //  We support IID_IUnknown and IID_IWiaTransferCallback
  if (SUCCEEDED(hr))
  {
    if (IID_IUnknown == iid_requested) 
    {
      *ppInterfaceOut = static_cast<IWiaTransferCallback*>(this);
    }
    else if (IID_IWiaTransferCallback == iid_requested) 
    {
      *ppInterfaceOut = static_cast<IWiaTransferCallback*>(this);
    }
    else
    {
      hr = E_NOINTERFACE;
    }
  }

  if (SUCCEEDED(hr))
  {
    reinterpret_cast<IUnknown*>(*ppInterfaceOut)->AddRef();
  }

  return hr;
}

STDMETHODIMP_(ULONG) CSimpleWIACallback::AddRef()
{
  return InterlockedIncrement(&m_cRef);
}

STDMETHODIMP_(ULONG)CSimpleWIACallback::Release()
{
  ULONG ulRefCount = InterlockedDecrement(&m_cRef);

  if (0 == ulRefCount)
  {
    delete this;
  }

  return ulRefCount;
}
    
CSimpleWIACallback::CSimpleWIACallback(CString strPath, CString strFileName): m_cRef(1)
{
  if( strPath[strPath.GetLength() - 1 ] != L'\\' )
  {
    strPath += L'\\';
  }

  m_strPath = strPath;
  m_strFileName = strFileName;
}

STDMETHODIMP CSimpleWIACallback::GetNextStream(LONG lFlags, BSTR bstrItemName,
                                  BSTR    bstrFullItemName, IStream **ppDestination)
{
  UNREFERENCED_PARAMETER(lFlags);
  UNREFERENCED_PARAMETER(bstrItemName);
  UNREFERENCED_PARAMETER(bstrFullItemName);
  HRESULT hr = S_OK;
  
  CString strFileName = m_strPath;
  strFileName += m_strFileName;
  strFileName += L".BMP";
  
  WCHAR strPath[MAX_PATH];
          
  if( !PathYetAnotherMakeUniqueName( strPath, strFileName, NULL, NULL) )
  {
    hr = E_FAIL;
  }
  
  hr = SHCreateStreamOnFileEx( strPath, STGM_READWRITE, FILE_ATTRIBUTE_NORMAL, TRUE, 0, ppDestination);
  
  if( SUCCEEDED(hr) )
  {
    theApp.m_FileNamesArray.Add( strPath );                   
  }
  
  return hr;
}

STDMETHODIMP CSimpleWIACallback::TransferCallback(LONG lFlags, WiaTransferParams *pWiaTransferParams)
{
  UNREFERENCED_PARAMETER(lFlags);
  UNREFERENCED_PARAMETER(pWiaTransferParams);
//@TODO add progress indicator
  return S_OK;
}


/*****************************************************************************
 *
 *  Global functions
 *
 *****************************************************************************/
STDAPI DllCanUnloadNow(void)
{
  if(theApp.m_pThisClass)
  {
    theApp.m_pThisClass->Release();
  }
  theApp.m_pThisClass=NULL;
  return (theApp.m_nLocks == 0) ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(__in           REFCLSID   rclsid,
                         __in           REFIID     riid,
                         __deref_out    void       **ppv)
{
  static  CUIExt2ClassObject s_FilterClass;

  HRESULT     hr;

  hr = ppv ? S_OK : E_INVALIDARG;

  if (SUCCEEDED(hr))
  {
    if (rclsid == CLSID_WiaUIExt2)
    {
      hr = s_FilterClass.QueryInterface(riid, ppv);
    }
    else
    {
      *ppv = 0;
      hr = CLASS_E_CLASSNOTAVAILABLE;
    }
  }

  return hr;
}

STDAPI DllUnregisterServer()
{
  return S_OK;
}

STDAPI DllRegisterServer()
{
  return S_OK;
}
