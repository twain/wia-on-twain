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
 * @file wiadriver.cpp
 * @author TWAIN Working Group
 * @date October 2009
 */

#include <initguid.h>
#include "stdafx.h"
#include <strsafe.h>
#include <limits.h>
#include "wiadriver.h"
#include "resource.h"

#include <shellapi.h>
#include <shlobj.h>
#include <shlwapi.h>

HINSTANCE g_hInst = NULL;

#define MAX(a,b) ((a)>(b)?(a):(b))
#define MIN(a,b) ((a)<(b)?(a):(b))

/**
 * Unique ID for each WIA context
 */
#define CUSTOM_ROOT_PROP_ID WIA_PRIVATE_DEVPROP
#define CUSTOM_ROOT_PROP_ID_STR L"AppID"

/**
 * Path to profile. WIA GUI use it to set profile
 */
#define CUSTOM_ROOT_PROP_ID1 WIA_PRIVATE_DEVPROP+1
#define CUSTOM_ROOT_PROP_ID1_STR L"ProfileName"


/**
 * Minimum Horizontal Bed Size
 */
#define CUSTOM_ROOT_PROP_WIA_DPS_MIN_HORIZONTAL_BED_SIZE     WIA_PRIVATE_DEVPROP+2
#define CUSTOM_ROOT_PROP_WIA_DPS_MIN_HORIZONTAL_BED_SIZE_STR L"Minimum Horizontal Bed Size"

/**
 * Minimum Vertical Bed Size
 */
#define CUSTOM_ROOT_PROP_WIA_DPS_MIN_VERTICAL_BED_SIZE       WIA_PRIVATE_DEVPROP+3
#define CUSTOM_ROOT_PROP_WIA_DPS_MIN_VERTICAL_BED_SIZE_STR   L"Minimum Vertical Bed Size"


/**
 * Profiles directory
 */
#define PROFILELOCATION L"TWAIN Working Group\\WIAONTWAIN\\"
/**
 * Profiles file extension
 */
#define FILEEXTENTION L".TWP"

/**
 * Application ID
 */
#define kVER_MAJ 1
#define kVER_MIN 0
#define kVER_INFO "1.0 alpha"
#define kPROT_MAJ 2
#define kPROT_MIN 0
#define kMANUFACTURER "TWAIN Working Group"
#define kPRODUCT_FAMILY "WIA_ON_TWAIN"
#define kPRODUCT_NAME "WIA_ON_TWAIN"

/**
 * TWAIN DS ID
 */

#define kDS_VER_MAJ 1
#define kDS_VER_MIN 0
#define kDS_MANUFACTURER "TWAIN Working Group"
#define kDS_PRODUCT_NAME "TWAIN2 FreeImage Software Scanner"

/**
 * WIA_DPA_FIRMWARE_VERSION value
 */
#define kFIRMWARE_VERSION L"0.0.0"

/**
 * WIA_IPA_FILENAME_EXTENSION value
 */
#define kFILE_TYPE L"BMP"

/**
 * WIA_DPS_MAX_SCAN_TIME value
 */
#define kMAX_SCAN_TIME 30000 // Modify this if needed


/**
 * Prototype of Set capability function
 */
typedef HRESULT (CWIADriver::*SetCapFunc) (BYTE *, CTWAIN_API *);
/**
 * Prototype of Get capability function
 */
typedef HRESULT (CWIADriver::*GetCapFunc) (BYTE *, CTWAIN_API *);

/**
 * list of all opened TWAIN DS
 */
map<int,CTWAIN_API*>    g_TWAIN_APIs;


//Table for transfering WIA properties to TWAIN capabilities
const ULONG g_unWIAPropToTWAINcap[] = {
  WIA_DPS_DOCUMENT_HANDLING_SELECT,CAP_FEEDERENABLED,CAP_DUPLEXENABLED,0,
  WIA_IPS_DOCUMENT_HANDLING_SELECT,CAP_FEEDERENABLED,CAP_DUPLEXENABLED,0,
  WIA_IPS_CUR_INTENT,ICAP_PIXELTYPE,0,
  WIA_IPA_DATATYPE,ICAP_PIXELTYPE,0,
  WIA_IPA_DEPTH,ICAP_BITDEPTH,0,
  WIA_IPS_BRIGHTNESS,ICAP_BRIGHTNESS,0,
  WIA_IPS_CONTRAST,ICAP_CONTRAST,0,
  WIA_IPS_PAGES,CAP_XFERCOUNT,0,
  WIA_DPS_PAGES,CAP_XFERCOUNT,0,
  WIA_IPS_PHOTOMETRIC_INTERP,ICAP_PIXELFLAVOR,0,
  WIA_IPS_THRESHOLD,ICAP_THRESHOLD,0,
  WIA_IPS_XEXTENT,DAT_IMAGELAYOUT,0,
  WIA_IPS_XPOS,DAT_IMAGELAYOUT,0,
  WIA_IPS_XRES,ICAP_XRESOLUTION,0,
  WIA_IPS_YEXTENT,DAT_IMAGELAYOUT,0,
  WIA_IPS_YPOS,DAT_IMAGELAYOUT,0,
  WIA_IPS_YRES,ICAP_YRESOLUTION,0,
  0,0
};

//List of TWAIN capabilities to be set by the driver
const TW_UINT16 g_unSetCapOrder[] = {//keep it in sync with g_unSetCapOrderFunc
  CAP_FEEDERENABLED,
  CAP_AUTOFEED,
  CAP_DUPLEXENABLED,
  ICAP_PIXELTYPE,
  ICAP_BITDEPTH,
  ICAP_XRESOLUTION,
  ICAP_YRESOLUTION,
  ICAP_PIXELFLAVOR,
  ICAP_THRESHOLD,
  ICAP_BRIGHTNESS,
  ICAP_CONTRAST,
  DAT_IMAGELAYOUT ,
  CAP_XFERCOUNT,
  0
};

//List of TWAIN capabilities to be get by the driver
const TW_UINT16 g_unGetCapOrder[] = {//keep it in sync with g_unGetCapOrderFunc
  CAP_FEEDERENABLED,
  CAP_FEEDERALIGNMENT ,
  CAP_AUTOFEED,
  CAP_DUPLEXENABLED,
  ICAP_PIXELTYPE,
  ICAP_BITDEPTH,
  ICAP_XRESOLUTION,
  ICAP_YRESOLUTION,
  ICAP_PIXELFLAVOR,
  ICAP_THRESHOLD,
  ICAP_BRIGHTNESS,
  ICAP_CONTRAST,
  DAT_IMAGELAYOUT ,
  CAP_XFERCOUNT,
  0
};

//List of function for setting TWAIN capabilities
SetCapFunc g_unSetCapOrderFunc[]={//keep it in sync with g_unSetCapOrder
  &CWIADriver::SetFEEDERENABLED,
  &CWIADriver::SetAUTOFEED,
  &CWIADriver::SetDUPLEXENABLED,
  &CWIADriver::SetPIXELTYPE,
  &CWIADriver::SetBITDEPTH,
  &CWIADriver::SetXRESOLUTION,
  &CWIADriver::SetYRESOLUTION,
  &CWIADriver::SetPIXELFLAVOR,
  &CWIADriver::SetTHRESHOLD,
  &CWIADriver::SetBRIGHTNESS,
  &CWIADriver::SetCONTRAST,
  &CWIADriver::SetDAT_IMAGELAYOUT,
  &CWIADriver::SetXFERCOUNT,
  0
};

//List of function for getting TWAIN capabilities
GetCapFunc g_unGetCapOrderFunc[]={//keep it in sync with g_unGetCapOrder
  &CWIADriver::GetFEEDERENABLED,
  &CWIADriver::GetFEEDERALIGNMENT,
  &CWIADriver::GetAUTOFEED,
  &CWIADriver::GetDUPLEXENABLED,
  &CWIADriver::GetPIXELTYPE,
  &CWIADriver::GetBITDEPTH,
  &CWIADriver::GetXRESOLUTION,
  &CWIADriver::GetYRESOLUTION,
  &CWIADriver::GetPIXELFLAVOR,
  &CWIADriver::GetTHRESHOLD,
  &CWIADriver::GetBRIGHTNESS,
  &CWIADriver::GetCONTRAST,
  &CWIADriver::GetDAT_IMAGELAYOUT,
  &CWIADriver::GetXFERCOUNT,
  0
};

///////////////////////////////////////////////////////////////////////////////
// WIA driver GUID
///////////////////////////////////////////////////////////////////////////////

// {7247AA1E-C489-40f6-A19B-62FBCAA81334} SampleDS
DEFINE_GUID(CLSID_WIADriver, 
0x7247aa1e, 0xc489, 0x40f6, 0xa1, 0x9b, 0x62, 0xfb, 0xca, 0xa8, 0x13, 0x34);


/////////////////////////////////////////////////////////////////////////
// IClassFactory Interface Section (for all COM objects)     //
/////////////////////////////////////////////////////////////////////////

class CWIADriverClassFactory : public IClassFactory
{
public:
  CWIADriverClassFactory() : m_cRef(1) {}
  ~CWIADriverClassFactory(){}
  HRESULT __stdcall QueryInterface(REFIID riid, LPVOID *ppv)
  {
    if(!ppv)
    {
      WIAS_ERROR((g_hInst, "Invalid parameters were passed"));
      return E_INVALIDARG;
    }

    *ppv = NULL;
    HRESULT hr = E_NOINTERFACE;
    if(IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory))
    {
      *ppv = static_cast<IClassFactory*>(this);
      reinterpret_cast<IUnknown*>(*ppv)->AddRef();
      hr = S_OK;
    }
    return hr;
  }
  ULONG __stdcall AddRef()
  {
    return InterlockedIncrement(&m_cRef);
  }
  ULONG __stdcall Release()
  {
    ULONG ulRef = InterlockedDecrement(&m_cRef);
    if(ulRef == 0)
    {
      delete this;
      return 0;
    }
    return ulRef;
  }
  HRESULT __stdcall CreateInstance(IUnknown __RPC_FAR *pUnkOuter,REFIID riid,void __RPC_FAR *__RPC_FAR *ppvObject)
  {
    if (ppvObject == NULL)
    {
      return E_INVALIDARG;
    }
    *ppvObject = NULL;

    if((pUnkOuter)&&(!IsEqualIID(riid,IID_IUnknown)))
    {
      return CLASS_E_NOAGGREGATION;
    }

    HRESULT hr = E_NOINTERFACE;
#pragma prefast(suppress:__WARNING_ALIASED_MEMORY_LEAK, "pDev is freed on release.")
    CWIADriver *pDev = new CWIADriver(pUnkOuter);
    if(pDev)
    {
      hr = pDev->NonDelegatingQueryInterface(riid,ppvObject);
      pDev->NonDelegatingRelease();
    }
    else
    {
      hr = E_OUTOFMEMORY;
      WIAS_ERROR((g_hInst, "Failed to allocate WIA driver class object, hr = 0x%lx",hr));
    }

    return hr;
  }
  HRESULT __stdcall LockServer(BOOL fLock)
  {
    UNREFERENCED_PARAMETER(fLock);
  
    return S_OK;
  }
private:
  LONG m_cRef;
};

/////////////////////////////////////////////////////////////////////////
// DLL Entry Point Section (for all COM objects, in a DLL)       //
/////////////////////////////////////////////////////////////////////////

extern "C" __declspec(dllexport) BOOL APIENTRY DllMain(HINSTANCE hinst,DWORD dwReason, __reserved LPVOID lpReserved)
{
  UNREFERENCED_PARAMETER(lpReserved);

  g_hInst = hinst;
  switch(dwReason)
  {
  case DLL_PROCESS_ATTACH:
    DisableThreadLibraryCalls(hinst);
    break;
  }
  return TRUE;
}

extern "C" HRESULT __stdcall DllCanUnloadNow(void)
{
  return S_OK;
}
extern "C" HRESULT __stdcall DllGetClassObject(__in REFCLSID rclsid, __in REFIID riid, __deref_out LPVOID *ppv)
{
  if(!ppv)
  {
    WIAS_ERROR((g_hInst, "Invalid parameters were passed"));
    return E_INVALIDARG;
  }
  HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;
  *ppv = NULL;
  if(IsEqualCLSID(rclsid, CLSID_WIADriver))
  {
#pragma prefast(suppress:__WARNING_ALIASED_MEMORY_LEAK, "pcf is freed on release.")
    CWIADriverClassFactory *pcf = new CWIADriverClassFactory;
    if(pcf)
    {
      hr = pcf->QueryInterface(riid,ppv);
      pcf->Release();
    }
    else
    {
      hr = E_OUTOFMEMORY;
      WIAS_ERROR((g_hInst, "Failed to allocate WIA driver class factory object, hr = 0x%lx",hr));
    }
  }
  return hr;
}

extern "C" HRESULT __stdcall DllRegisterServer()
{
  return S_OK;
}

extern "C" HRESULT __stdcall DllUnregisterServer()
{
  return S_OK;
}


///////////////////////////////////////////////////////////////////////////
// Construction/Destruction Section
///////////////////////////////////////////////////////////////////////////

CWIADriver::CWIADriver(__in_opt LPUNKNOWN punkOuter) : m_cRef(1),
  m_punkOuter(NULL),
  m_pIDrvItemRoot(NULL),
  m_lClientsConnected(0),
  m_pFormats(NULL),
  m_ulNumFormats(0),
  m_bstrDeviceID(NULL),
  m_bstrRootFullItemName(NULL),
  m_pIStiDevice(NULL),
  m_nLockCounter(0)
{
  if(punkOuter)
  {
    m_punkOuter = punkOuter;
  }
  else
  {
    m_punkOuter = reinterpret_cast<IUnknown*>(static_cast<INonDelegatingUnknown*>(this));
  }

  OSVERSIONINFO osvi;
  ZeroMemory(&osvi, sizeof(OSVERSIONINFO));
  osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);

  GetVersionEx(&osvi);

  m_bIsWindowsVista = (osvi.dwMajorVersion >= 6);
  m_bHasFlatbed=true;
  m_bHasFeeder=true;
  //Get Profile path and store it in m_strProfilesPath
  TCHAR strProfilesPath[MAX_PATH];

  SHGetSpecialFolderPath(NULL, strProfilesPath, CSIDL_COMMON_APPDATA|CSIDL_FLAG_CREATE, TRUE);

  if(strProfilesPath[wcslen(strProfilesPath)-1] != '\\')
  {
    wcscat_s(strProfilesPath, MAX_PATH, L"\\");
  }

  wcscat_s(strProfilesPath, MAX_PATH, PROFILELOCATION);
  if(strProfilesPath[wcslen(strProfilesPath)-1] != '\\')
  {
    wcscat_s(strProfilesPath, MAX_PATH, L"\\");
  }

  m_strProfilesPath = strProfilesPath;
}

CWIADriver::~CWIADriver()
{
  HRESULT hr = S_OK;
  //Close DS, Close DSM and delete from list all opened DSM
  map<int,CTWAIN_API*>::iterator lIter=g_TWAIN_APIs.begin();
  for(;lIter!=g_TWAIN_APIs.end();lIter++)
  {
    lIter->second->CloseDS();
    lIter->second->CloseDSM();
    delete lIter->second;
  }
  g_TWAIN_APIs.clear();

  if(m_bstrDeviceID)
  {
    SysFreeString(m_bstrDeviceID);
    m_bstrDeviceID = NULL;
  }

  if(m_bstrRootFullItemName)
  {
    SysFreeString(m_bstrRootFullItemName);
    m_bstrRootFullItemName = NULL;
  }

  // Free cached driver capability array
  m_CapabilityManager.Destroy();

  // Free cached driver format array
  if(m_pFormats)
  {
    WIAS_TRACE((g_hInst,"Deleting WIA format array memory"));
    delete [] m_pFormats;
    m_pFormats = NULL;
    m_ulNumFormats = 0;
  }

  // Unlink and release the cached IWiaDrvItem root item interface.
  DestroyDriverItemTree();

}

///////////////////////////////////////////////////////////////////////////
// Standard COM Section
///////////////////////////////////////////////////////////////////////////

HRESULT CWIADriver::QueryInterface(__in REFIID riid, __out LPVOID * ppvObj)
{
  if (ppvObj == NULL)
  {
    return E_INVALIDARG;
  }
  *ppvObj = NULL;

  if(!m_punkOuter)
  {
    return E_NOINTERFACE;
  }
  return m_punkOuter->QueryInterface(riid,ppvObj);
}
ULONG CWIADriver::AddRef()
{
  if(!m_punkOuter)
  {
    return 0;
  }
  return m_punkOuter->AddRef();
}
ULONG CWIADriver::Release()
{
  if(!m_punkOuter)
  {
    return 0;
  }
  return m_punkOuter->Release();
}

///////////////////////////////////////////////////////////////////////////
// IStiUSD Interface Section (for all WIA drivers)
///////////////////////////////////////////////////////////////////////////

HRESULT CWIADriver::Initialize(__in   PSTIDEVICECONTROL pHelDcb,
                    DWORD       dwStiVersion,
                 __in   HKEY        hParametersKey)
{
  HRESULT hr = E_INVALIDARG;
  if((pHelDcb)&&(hParametersKey))
  {
    //Get Scanner configuration
    CTWAIN_API TwainApi;
    TW_IDENTITY AppID={0,{kVER_MAJ,kVER_MIN,TWLG_USA,TWCY_USA,kVER_INFO},kPROT_MAJ,kPROT_MIN,DG_CONTROL|DG_IMAGE|DF_APP2,kMANUFACTURER,kPRODUCT_FAMILY,kPRODUCT_NAME};
    DWORD dwRes = TwainApi.OpenDSM(AppID,0);
    if(dwRes!=0)
    {
      hr = WIA_ERROR_GENERAL_ERROR;
    }
    else
    {
      TW_IDENTITY DSID={0,{kDS_VER_MAJ,kDS_VER_MIN,0,0,""},0,0,0,kDS_MANUFACTURER,"",kDS_PRODUCT_NAME};
      dwRes = TwainApi.OpenDS(&DSID);
      if(dwRes!=0)
      {
        TwainApi.CloseDSM();
        hr = WIA_ERROR_OFFLINE;
      }
      else
      {
        hr = GetConfigParams(&TwainApi);
        TwainApi.CloseDS();
      }
      TwainApi.CloseDSM();
    }
    if(hr==S_OK)
    {
      hr = m_CapabilityManager.Initialize(g_hInst);
      if(FAILED(hr))
      {
        WIAS_ERROR((g_hInst, "Failed to initialize the WIA driver capability manager object, hr = 0x%lx",hr));
      }
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::GetCapabilities(__out PSTI_USD_CAPS pDevCaps)
{
  HRESULT hr = E_INVALIDARG;
  if(pDevCaps)
  {
    memset(pDevCaps, 0, sizeof(STI_USD_CAPS));
    if(m_bIsWindowsVista)
    {
      pDevCaps->dwVersion   = STI_VERSION_3;//WAI 2.0
    }
    else
    {
      pDevCaps->dwVersion   = STI_VERSION_REAL;//WAI 1.0
    }
    pDevCaps->dwGenericCaps = STI_GENCAP_WIA | STI_GENCAP_NOTIFICATIONS;

    WIAS_TRACE((g_hInst,"========================================================"));
    WIAS_TRACE((g_hInst,"STI Capabilities information reported to the WIA Service"));
    WIAS_TRACE((g_hInst,"Version:   0x%lx",pDevCaps->dwVersion));
    WIAS_TRACE((g_hInst,"GenericCaps: 0x%lx", pDevCaps->dwGenericCaps));
    WIAS_TRACE((g_hInst,"========================================================"));

    hr = S_OK;
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::GetStatus(__inout PSTI_DEVICE_STATUS pDevStatus)
{
  HRESULT hr = E_INVALIDARG;
  if(pDevStatus)
  {
    // assume successful status checks
    hr = S_OK;

    if(pDevStatus->StatusMask & STI_DEVSTATUS_ONLINE_STATE)
    {
      // check if the device is ON-LINE
      WIAS_TRACE((g_hInst,"Checking device online status..."));
      pDevStatus->dwOnlineState = 0L;

      if(SUCCEEDED(hr))
      {
        pDevStatus->dwOnlineState |= STI_ONLINESTATE_OPERATIONAL;
        WIAS_TRACE((g_hInst,"The device is online"));
      }
      else
      {
        WIAS_TRACE((g_hInst,"The device is offline"));
      }
    }

    if(pDevStatus->StatusMask & STI_DEVSTATUS_EVENTS_STATE)
    {
      // check for polled events
      pDevStatus->dwEventHandlingState &= ~STI_EVENTHANDLING_PENDING;

      hr = S_FALSE; // no are events detected

      if(hr == S_OK)
      {
        pDevStatus->dwEventHandlingState |= STI_EVENTHANDLING_PENDING;
        WIAS_TRACE((g_hInst,"The device reported a polled event"));
      }
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::DeviceReset()
{
  return S_OK;
}

HRESULT CWIADriver::Diagnostic(__out LPDIAG pBuffer)
{
  HRESULT hr = E_INVALIDARG;
  if(pBuffer)
  {
    memset(pBuffer,0,sizeof(DIAG));
    hr = S_OK;
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::Escape(                STI_RAW_CONTROL_CODE EscapeFunction,
               __in_bcount(cbInDataSize)   LPVOID         lpInData,
                             DWORD        cbInDataSize,
               __out_bcount(dwOutDataSize)   LPVOID         pOutData,
                             DWORD        dwOutDataSize,
               __out             LPDWORD        pdwActualData)
{
  UNREFERENCED_PARAMETER(EscapeFunction);
  UNREFERENCED_PARAMETER(lpInData);
  UNREFERENCED_PARAMETER(cbInDataSize);
  UNREFERENCED_PARAMETER(pOutData);
  UNREFERENCED_PARAMETER(dwOutDataSize);
  UNREFERENCED_PARAMETER(pdwActualData);

  WIAS_ERROR((g_hInst, "This method is not implemented or supported for this driver"));
  return E_NOTIMPL;
}

HRESULT CWIADriver::GetLastError(__out LPDWORD pdwLastDeviceError)
{
  HRESULT hr = E_INVALIDARG;
  if(pdwLastDeviceError)
  {
    *pdwLastDeviceError = 0;
    hr = S_OK;
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::LockDevice()
{
  return S_OK;
}

HRESULT CWIADriver::UnLockDevice()
{
  return S_OK;
}

HRESULT CWIADriver::RawReadData(__out_bcount(*lpdwNumberOfBytes)    LPVOID     lpBuffer,
                __out                 LPDWORD    lpdwNumberOfBytes,
                __out                 LPOVERLAPPED lpOverlapped)
{
  UNREFERENCED_PARAMETER(lpBuffer);
  UNREFERENCED_PARAMETER(lpdwNumberOfBytes);
  UNREFERENCED_PARAMETER(lpOverlapped);

  WIAS_ERROR((g_hInst, "This method is not implemented or supported for this driver"));
  return E_NOTIMPL;
}

HRESULT CWIADriver::RawWriteData(__in_bcount(dwNumberOfBytes)  LPVOID     lpBuffer,
                                 DWORD    dwNumberOfBytes,
                                 __out LPOVERLAPPED lpOverlapped)
{
  UNREFERENCED_PARAMETER(lpBuffer);
  UNREFERENCED_PARAMETER(dwNumberOfBytes);
  UNREFERENCED_PARAMETER(lpOverlapped);

  WIAS_ERROR((g_hInst, "This method is not implemented or supported for this driver"));
  return E_NOTIMPL;
}

HRESULT CWIADriver::RawReadCommand(__out_bcount(*lpdwNumberOfBytes) LPVOID     lpBuffer,
                   __out              LPDWORD    lpdwNumberOfBytes,
                   __out              LPOVERLAPPED lpOverlapped)
{
  UNREFERENCED_PARAMETER(lpBuffer);
  UNREFERENCED_PARAMETER(lpdwNumberOfBytes);
  UNREFERENCED_PARAMETER(lpOverlapped);

  WIAS_ERROR((g_hInst, "This method is not implemented or supported for this driver"));
  return E_NOTIMPL;
}

HRESULT CWIADriver::RawWriteCommand(__in_bcount(dwNumberOfBytes) LPVOID     lpBuffer,
                                 DWORD    dwNumberOfBytes,
                                 __out LPOVERLAPPED lpOverlapped)
{
  UNREFERENCED_PARAMETER(lpBuffer);
  UNREFERENCED_PARAMETER(dwNumberOfBytes);
  UNREFERENCED_PARAMETER(lpOverlapped);

  WIAS_ERROR((g_hInst, "This method is not implemented or supported for this driver"));
  return E_NOTIMPL;
}

HRESULT CWIADriver::SetNotificationHandle(__in HANDLE hEvent)
{
  UNREFERENCED_PARAMETER(hEvent);

  WIAS_ERROR((g_hInst, "This method is not implemented or supported for this driver"));
  return E_NOTIMPL;
}

HRESULT CWIADriver::GetNotificationData(__in LPSTINOTIFY lpNotify)
{
  UNREFERENCED_PARAMETER(lpNotify);

  WIAS_ERROR((g_hInst, "This method is not implemented or supported for this driver"));
  return E_NOTIMPL;
}

HRESULT CWIADriver::GetLastErrorInfo(STI_ERROR_INFO *pLastErrorInfo)
{
  HRESULT hr = E_INVALIDARG;
  if(pLastErrorInfo)
  {
    memset(pLastErrorInfo,0,sizeof(STI_ERROR_INFO));
    hr = S_OK;
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

/////////////////////////////////////////////////////////////////////////
// IWiaMiniDrv Interface Section (for all WIA drivers)         //
/////////////////////////////////////////////////////////////////////////

HRESULT CWIADriver::drvInitializeWia(__inout BYTE    *pWiasContext,
                       LONG    lFlags,
                   __in  BSTR    bstrDeviceID,
                   __in  BSTR    bstrRootFullItemName,
                   __in  IUnknown  *pStiDevice,
                   __in  IUnknown  *pIUnknownOuter,
                   __out   IWiaDrvItem **ppIDrvItemRoot,
                   __out   IUnknown  **ppIUnknownInner,
                   __out   LONG    *plDevErrVal)
{
  UNREFERENCED_PARAMETER(lFlags);
  UNREFERENCED_PARAMETER(pIUnknownOuter);
  UNREFERENCED_PARAMETER(ppIUnknownInner);
  HRESULT hr = S_OK;
 
  if((pWiasContext)&&(plDevErrVal)&&(ppIDrvItemRoot))
  {
    *plDevErrVal = 0;
    *ppIDrvItemRoot = NULL;

    if(!m_bstrDeviceID)
    {
      m_bstrDeviceID = SysAllocString(bstrDeviceID);
      if(!m_bstrDeviceID)
      {
        hr = E_OUTOFMEMORY;
        WIAS_ERROR((g_hInst, "Failed to allocate BSTR DeviceID string, hr = 0x%lx",hr));
      }
    }

    if(!m_pIStiDevice)
    {
      m_pIStiDevice = reinterpret_cast<IStiDevice*>(pStiDevice);
    }

    if(!m_bstrRootFullItemName)
    {
      m_bstrRootFullItemName = SysAllocString(bstrRootFullItemName);
      if(!m_bstrRootFullItemName)
      {
        hr = E_OUTOFMEMORY;
        WIAS_ERROR((g_hInst, "Failed to allocate BSTR Root full item name string, hr = 0x%lx",hr));
      }
    }

    if(SUCCEEDED(hr))
    {
      if(!m_pIDrvItemRoot)//build tree if it is not already build
      {
        hr = BuildDriverItemTree();
      }
      else
      {

        //
        // A WIA item tree already exists.  The root item of this item tree
        // should be returned to the WIA service.
        //

        hr = S_OK;
      }
    }

    //
    // Make PREfast happy by inspecting m_pIDrvItemRoot.
    // PREfast doesn't seem to figure out that m_pIDrvItemRoot is set by
    // BuildDriverTree()'s call to waisCreateDrvItem() only on success.
    //

    if(SUCCEEDED(hr) && !m_pIDrvItemRoot)
    {
      hr = E_UNEXPECTED;
      WIAS_ERROR((g_hInst, "Missing driver item tree root unexpected, hr = 0x%lx",hr));
    }

    //
    // Only increment the client connection count, when the driver
    // has successfully created all the necessary WIA items for
    // a client to use.
    //

    if(SUCCEEDED(hr))
    {
      *ppIDrvItemRoot = m_pIDrvItemRoot;
      InterlockedIncrement(&m_lClientsConnected);
      WIAS_TRACE((g_hInst,"%d client(s) are currently connected to this driver.",m_lClientsConnected));
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::drvAcquireItemDataWIA1(       LONG               lFlags,
                   __in     BYTE               *pWiasContext,
                   __in     PMINIDRV_TRANSFER_CONTEXT    pmdtc,
                   __out    LONG               *plDevErrVal)
{
  HRESULT  hr        = S_OK;
  long lPagesScanned  = 0;
  bool bUserCancel    = false;
  long lPagesToScan = 1;
  long lDataBufferSize = pmdtc->lBufferSize;
  HGLOBAL hBMP = NULL;
  DWORD dwRes;
  BYTE *pRootWiasContext=NULL;

  hr = wiasGetRootItem(pWiasContext,&pRootWiasContext);
  if(hr!=S_OK)
  {
    return hr;
  }
  CTWAIN_API *pTwainApi;
  hr = GetDS(pWiasContext,&pTwainApi);
  if(hr!=S_OK)
  {
    return hr;
  }
  
  bool bFeeder = false;
  hr = IsFeederItem(pWiasContext,&bFeeder);
  if(hr!=S_OK)
  {
    return hr;
  }

  //get pages to scan 
  if (!bFeeder)
  {
    lPagesToScan = 1;
  }
  else
  {
    hr = wiasReadPropLong(pRootWiasContext, WIA_DPS_PAGES, &lPagesToScan, NULL, TRUE);
    if (FAILED(hr))
    {
      WIAS_ERROR((g_hInst, "drvAcquireItemData: failure reading WIA_IPS_PAGES property for Feeder item, hr = 0x%lx", hr));
    }
    else if (ALL_PAGES == lPagesToScan)
    {
      lPagesToScan = -1;
    }
  }  
  
  try
  {
    if(pmdtc->bClassDrvAllocBuf)
    {
      if(!pmdtc->pTransferBuffer)
      {
        throw hr;
      }
    }

    LL llVal;
    WORD wType;

    //check for paper - some TWAIN DS may hang UIless mode if there is no paper in ADF
    if(bFeeder)
    {
      dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERLOADED,&llVal,&wType);
      if(dwRes==0&&wType==TWTY_BOOL&& (bool)llVal==false)
      {
        return WIA_ERROR_PAPER_EMPTY;
      }
    }
    //enable in UIless mode
    HANDLE hTransferEvent;
    dwRes = pTwainApi->EnableDS(&hTransferEvent);
    hr = TWAINtoWIAerror(dwRes);
    if(hr!=S_OK)
    {
      return hr;
    }     

    //wait for event from the scanner
    if(WaitForSingleObject(hTransferEvent,kMAX_SCAN_TIME) != WAIT_OBJECT_0 || pTwainApi->GetLastMsg()!=MSG_XFERREADY)
    {
      throw ERROR_CANCELLED;
    }

    //Scan All Pages
    while( lPagesToScan && SUCCEEDED( hr ) && ( hr != S_FALSE ) )
    {
      bool bEndOfImage = false;

      DWORD dwDIBsize;
      dwRes = pTwainApi->ScanNative(&hBMP,&dwDIBsize);
      hr = TWAINtoWIAerror(dwRes);
      if(hr!=S_OK) // exit on error
      {
        break;
      }  

      BYTE *pData = (BYTE*)pTwainApi->DSM_LockMemory(hBMP);
      if (pData==0)
      {
        WIAS_ERROR((g_hInst, "Failed to lock hDIB, hr = 0x%lx",hr));
        throw E_OUTOFMEMORY;
      }

      hr = wiasReadPropLong(pWiasContext, WIA_IPA_BUFFER_SIZE, &lDataBufferSize, NULL, TRUE);

      if (FAILED(hr)) 
      {
        throw hr;
      }

      //  Get BMP info header
      BITMAPINFOHEADER *pBMPinfo = (BITMAPINFOHEADER*)pData;

      if( pmdtc->tymed == TYMED_FILE )
      {
        //Add BMP file header
        lDataBufferSize +=dwDIBsize+ sizeof(BITMAPFILEHEADER);
      }

      if(pmdtc->lBufferSize<lDataBufferSize)//reallocate if smaller. WIA driver always allocates memory for WIA 1.0 transfer -WIA_IPA_ITEM_SIZE=0 
      {
        if(pmdtc->lBufferSize)
        {
          CoTaskMemFree( pmdtc->pTransferBuffer );
        }
        pmdtc->pTransferBuffer = (BYTE*) CoTaskMemAlloc( lDataBufferSize );

        if (!pmdtc->pTransferBuffer) 
        {
          throw E_OUTOFMEMORY;
        }
        pmdtc->lBufferSize = lDataBufferSize;  
      }
      //correct size according to real image 
      pmdtc->lWidthInPixels = pBMPinfo->biWidth; 
      pmdtc->lLines = pBMPinfo->biHeight;
      pmdtc->cbWidthInBytes = ((pBMPinfo->biWidth*pBMPinfo->biBitCount+31)/32)*4;
      
      //update pmdtc - do not call this before updating of image size
      hr = wiasGetImageInformation( pWiasContext, 0, pmdtc );
      if( FAILED( hr ) )
      {
        throw hr;
      }

      long  lPercentDone    = 0;
      long  lBytesRecieved    = pmdtc->lHeaderSize;
      DWORD  dwBytesTransferred  = 0;
      long  lBytesReadFromDIB  = 0;
       
      BITMAPINFOHEADER* phdr = NULL;
      if( ( pmdtc->tymed == TYMED_CALLBACK ) || ( pmdtc->tymed == TYMED_MULTIPAGE_CALLBACK ) )
      {
        phdr = (BITMAPINFOHEADER*) ( pmdtc->pTransferBuffer );
      }
      else
      {
        // add BMP file header
        BITMAPFILEHEADER* pfh = (BITMAPFILEHEADER*) ( pmdtc->pTransferBuffer );   
        phdr = (BITMAPINFOHEADER*) ( pmdtc->pTransferBuffer + sizeof( BITMAPFILEHEADER ) );    
      }

      phdr->biXPelsPerMeter  = 1 + (10000 * pmdtc->lXRes) / 254;
      phdr->biYPelsPerMeter  = 1 + (10000 * pmdtc->lYRes) / 254;
      

      if( lDataBufferSize >  pmdtc->lBufferSize )
      {
        throw E_OUTOFMEMORY;
      }

      if( pmdtc->tymed == TYMED_FILE )
      {
        //header is already there filled by wiasGetImageInformation
        dwBytesTransferred += sizeof( BITMAPFILEHEADER ) ;  
      }

      DWORD dwBytesRecieved = pmdtc->lBufferSize;
                
      //transfer image
      while( SUCCEEDED(hr) && ( hr != S_FALSE ) && !bEndOfImage)
      {        
        BYTE* pDataRead = pmdtc->pTransferBuffer;

        if( pmdtc->tymed == TYMED_FILE )
        {
          pDataRead = (BYTE*)(pmdtc->pTransferBuffer+dwBytesTransferred);
        }
        if(lBytesReadFromDIB+dwBytesRecieved >= dwDIBsize)//last read
        {
          bEndOfImage = true;
          dwBytesRecieved = dwDIBsize-lBytesReadFromDIB;
        }
        memcpy(pDataRead, pData+lBytesReadFromDIB,dwBytesRecieved);
        lBytesReadFromDIB +=dwBytesRecieved;
        
        lPercentDone = min(95, ((100*dwBytesTransferred)/dwDIBsize));
        if(pmdtc->tymed == TYMED_CALLBACK)
        {
          hr = pmdtc->pIWiaMiniDrvCallBack->MiniDrvCallback( IT_MSG_DATA,IT_STATUS_TRANSFER_TO_CLIENT, lPercentDone, pmdtc->cbOffset, dwBytesRecieved, pmdtc, 0 );
          pmdtc->cbOffset += dwBytesRecieved;      
        }
        else
        {
          hr = pmdtc->pIWiaMiniDrvCallBack->MiniDrvCallback( IT_MSG_STATUS, IT_STATUS_TRANSFER_TO_CLIENT, lPercentDone, 0, 0, pmdtc, 0 );                        
        }
        if( hr == S_FALSE ) 
        {              
          bUserCancel = true; 
        }  

        if( FAILED( hr ) )
        {
          throw hr;
        }

        dwBytesTransferred += dwBytesRecieved;
      }
      //free allocated resources
      if(hBMP)
      {
        pTwainApi->DSM_UnlockMemory(hBMP);
        pTwainApi->DSM_Free(hBMP);
        hBMP = NULL;
      }
      bool bMoreImages;
      dwRes = pTwainApi->EndTransfer(&bMoreImages);
      hr = TWAINtoWIAerror(dwRes);
      if(hr!=S_OK)
      {
        break;
      }  

      if( bUserCancel )//canceled by the user
      {
        break;
      }

      lPagesScanned++;  
      lPagesToScan--;
      if(!bMoreImages)
      {
        lPagesToScan=0;
      }
      //inform WIA service we are done with this transfer
      if ( ( pmdtc->tymed == TYMED_CALLBACK ) || ( pmdtc->tymed == TYMED_MULTIPAGE_CALLBACK ) )
      {
        hr = wiasSendEndOfPage( pWiasContext, lPagesScanned, pmdtc );
        pmdtc->cbOffset = 0;                                
      }
      else
      { 
        hr = wiasWritePageBufToFile( pmdtc );
      }
                
      if(SUCCEEDED(hr) && lPagesToScan==0)
      {
        hr = S_OK;
      }
    }//End Scan All Pages

  }
  catch( HRESULT e )
  {    
    hr = e;
    if( ( hr == WIA_ERROR_PAPER_EMPTY ) && lPagesScanned )// it is not error after first page 
    {  
      hr = WIA_STATUS_END_OF_MEDIA;
    }
  }
  catch(...)
  {    
  }

  if(pTwainApi)
  {
    pTwainApi->ResetTransfer();// Go to state 5
    pTwainApi->DisableDS();// Go to state 4
  }
  //free allocated resources
  if(hBMP)
  {
    pTwainApi->DSM_UnlockMemory(hBMP);
    pTwainApi->DSM_Free(hBMP);
    hBMP = NULL;
  }

  if(!pmdtc->bClassDrvAllocBuf && pmdtc->lBufferSize)
  {
    CoTaskMemFree( pmdtc->pTransferBuffer );
    pmdtc->pTransferBuffer = NULL;
    pmdtc->lBufferSize = 0;
  }
  
  hr = bUserCancel ? S_FALSE : hr;

  if(SUCCEEDED(hr) || (WIA_ERROR_PAPER_EMPTY == hr) || (WIA_STATUS_END_OF_MEDIA == hr))
  {
    *plDevErrVal = 0;
  }

  return hr;
}

HRESULT CWIADriver::DownloadToStream(       LONG               lFlags,
                   __in     BYTE               *pWiasContext,
                   __in     PMINIDRV_TRANSFER_CONTEXT    pmdtc,
                   __callback IWiaMiniDrvTransferCallback  *pTransferCallback,
                   __out    LONG               *plDevErrVal)
{
  UNREFERENCED_PARAMETER(lFlags);
  UNREFERENCED_PARAMETER(plDevErrVal);

  HRESULT hr          = S_OK;
  BSTR  bstrItemName    = NULL;
  BSTR  bstrFullItemName  = NULL;
  UINT  uiBitmapResourceID  = 0;


  if (S_OK == hr)
  {
    //  Get the item name
    hr = wiasReadPropStr(pWiasContext, WIA_IPA_ITEM_NAME, &bstrItemName, NULL, TRUE);
    if (SUCCEEDED(hr))
    {
      //  Get the full item name
      hr = wiasReadPropStr(pWiasContext, WIA_IPA_FULL_ITEM_NAME, &bstrFullItemName, NULL, TRUE);
      if (SUCCEEDED(hr))
      {
        //  Get the destination stream
        IStream *pDestination = NULL;
        hr = pTransferCallback->GetNextStream(0, bstrItemName, bstrFullItemName, &pDestination);
        if (hr == S_OK)
        {
          WiaTransferParams *pParams = (WiaTransferParams*)CoTaskMemAlloc(sizeof(WiaTransferParams));
          if (pParams)
          {
            memset(pParams, 0, sizeof(WiaTransferParams));
            if (SUCCEEDED(hr))
            { 
              if (S_OK == hr)
              {
                CTWAIN_API *pTwainApi;
                hr = GetDS(pWiasContext,&pTwainApi);
                HANDLE hBMP=NULL;
                DWORD dwSize;
                DWORD dwDataTransfered=0;
                //Get image
                DWORD dwRes = pTwainApi->ScanNative(&hBMP,&dwSize);
                hr = TWAINtoWIAerror(dwRes);
                if(hr!=S_OK)
                {
                  return hr;
                }  
                BYTE *pData = NULL;
                pData = (BYTE*)pTwainApi->DSM_LockMemory(hBMP);

                if (pData)
                {
                  //  Data transfer loop
                  //  Read from device
                  ULONG   ulBytesRead     = 0;
                  LONG  lPercentComplete  = 0;
                  ULONG   ulBytesWritten = 0;
                  
                  // file and transfer BMP file header
                  BITMAPFILEHEADER fileHeader;
                  memset(&fileHeader,0,sizeof(BITMAPFILEHEADER));
                  fileHeader.bfType = 0x4D42;
                  fileHeader.bfSize = sizeof(BITMAPFILEHEADER)+dwSize;
                  fileHeader.bfOffBits = fileHeader.bfSize - ((BITMAPINFOHEADER*)pData)->biSizeImage;
                  hr = pDestination->Write(&fileHeader, sizeof(BITMAPFILEHEADER), &ulBytesWritten);
                  pParams->ulTransferredBytes = ulBytesWritten;
                  
                  if (S_OK == hr)
                  {
                    //transfer the image
                    while(dwDataTransfered<dwSize)
                    {
                      ulBytesWritten = 0;
                      LARGE_INTEGER li = {0};
                      ulBytesRead=dwSize-dwDataTransfered;
                      if(ulBytesRead>dwSize/2)//because of WINQUAL. If it is image at once WINQUAL will not able to cancel the transfer
                      {
                        ulBytesRead=dwSize/2;
                      }
                      //
                      // Write to stream after seeking to end of stream as it could
                      // be randomized intially or during the callback
                      //
                      hr = pDestination->Seek(li, STREAM_SEEK_END, NULL);

                      if (S_OK == hr)
                      {
                        hr = pDestination->Write(pData+dwDataTransfered, ulBytesRead, &ulBytesWritten);
                      }
                      
                      if (S_OK == hr)
                      {
                        //
                        // Make progress callback
                        //
                        pParams->lMessage      = WIA_TRANSFER_MSG_STATUS;
                        pParams->lPercentComplete  = lPercentComplete;
                        pParams->ulTransferredBytes += ulBytesWritten;
                        dwDataTransfered +=ulBytesWritten;

                        hr = pTransferCallback->SendMessage(0, pParams);
                        if (FAILED(hr))
                        {
                          WIAS_ERROR((g_hInst, "Failed to send progress notification during download, hr = 0x%lx",hr));
                          break;
                        }
                        else if (S_FALSE == hr)
                        {
                          //
                          // Transfer cancelled
                          //
                          break;
                        }
                        else if (S_OK != hr)
                        {
                          WIAS_ERROR((g_hInst, "SendMessage returned unknown Success value, hr = 0x%lx",hr));
                          hr = E_UNEXPECTED;
                          break;
                        }                      
                      }
                    }
                  }
    
                  if (WIA_STATUS_END_OF_MEDIA == hr)
                  {
                    hr = S_OK;
                  }
                  
                  pTwainApi->DSM_UnlockMemory(hBMP);
                }
                else
                {
                  hr = E_OUTOFMEMORY;
                  WIAS_ERROR((g_hInst, "Failed to lock hDIB, hr = 0x%lx",hr));
                }
                if(hBMP)
                {
                  pTwainApi->DSM_Free(hBMP);
                }
              }
            }
            else
            {
              WIAS_ERROR((g_hInst, "Failed to allocate memory for transfer buffer, hr = 0x%lx",hr));
            }
            CoTaskMemFree(pParams);
            pParams = NULL;
          }
          else
          {
            hr = E_OUTOFMEMORY;
            WIAS_ERROR((g_hInst, "Failed to allocate memory for WiaTransferParams structure, hr = 0x%lx",hr));
          }
          pDestination->Release();
          pDestination = NULL;
        }
        else if(!((S_FALSE == hr) || (WIA_STATUS_SKIP_ITEM == hr)))
        {
          WIAS_ERROR((g_hInst, "GetNextStream returned unknown Success value, hr = 0x%lx",hr));
          hr = E_UNEXPECTED;
        }
        else
        {
          WIAS_ERROR((g_hInst, "Failed to get the destination stream for download, hr = 0x%lx",hr));
        }

        SysFreeString(bstrFullItemName);
        bstrFullItemName = NULL;
      }
      else
      {
        WIAS_ERROR((g_hInst, "Failed to read the WIA_IPA_FULL_ITEM_NAME property, hr = 0x%lx",hr));
      }
      SysFreeString(bstrItemName);
      bstrItemName = NULL;
    }
    else
    {
      WIAS_ERROR((g_hInst, "Failed to read the WIA_IPA_ITEM_NAME property, hr = 0x%lx",hr));
    }
  }
  return hr;
}

HRESULT CWIADriver::LoadProfileFromFile(BYTE *pWiasContext, HGLOBAL *phData, DWORD *pdwDataSize)
{
  if(phData==NULL || pdwDataSize==NULL)
  {
    return E_FAIL;
  }
  BYTE *pRootWiasContext=NULL;
  HRESULT hr = wiasGetRootItem(pWiasContext,&pRootWiasContext);

  LONG lAppID = 0;
  BSTR strPrfName;
  //Get profile path
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropStr(pRootWiasContext,CUSTOM_ROOT_PROP_ID1,&strPrfName,NULL,TRUE);
  }
  if(!SUCCEEDED(hr))
  {
    return E_FAIL;
  }

  CString strFileName = strPrfName;

  *phData=NULL;
  *pdwDataSize=0;
  //load profile in memory
  HGLOBAL hData;
  FILE *pFile = NULL;
  if(_wfopen_s(&pFile, strFileName, _T("rb"))!=0)
  {
    return E_FAIL;
  }

  fseek(pFile, 0, SEEK_END);
  DWORD dwDataSize = (DWORD)ftell(pFile);
  rewind(pFile);
  HRESULT hRes = S_OK;
  if(dwDataSize<=sizeof(TW_IDENTITY))
  {
    fclose(pFile);
    return E_FAIL;
  }
  dwDataSize -=sizeof(TW_IDENTITY);
  TW_IDENTITY tempID;
  if(fread(&tempID, sizeof(TW_IDENTITY), 1, pFile)!=1)
  {
    fclose(pFile);
    return E_FAIL;
  }
  TW_IDENTITY DSID={0,{kDS_VER_MAJ,kDS_VER_MIN,0,0,""},0,0,0,kDS_MANUFACTURER,"",kDS_PRODUCT_NAME};
//check if the profile is for this DS
  if(memcmp(&tempID, &DSID, sizeof(TW_IDENTITY))!=0)
  {
    fclose(pFile);
    return E_FAIL;
  }
  LPVOID pData = NULL;
  if(hData = GlobalAlloc(GMEM_MOVEABLE, dwDataSize))
  {
    pData = GlobalLock(hData);
    if(pData)
    {
      if(fread(pData, dwDataSize, 1, pFile)!=1)
      {
        hRes = E_FAIL;
      }
    }
    else
    {
      hRes = E_FAIL;
    }
  }
  else
  {
    hRes = E_FAIL;
  }
  fclose(pFile);

  if(hData)
  {
    GlobalUnlock(hData);
    if(hRes!=S_OK)
    {
      GlobalFree(hData);
      hData=NULL;
    } 
    else
    {
      *phData=hData; 
      *pdwDataSize=dwDataSize;
    }
  }
  
  return hRes;
}


HRESULT CWIADriver::drvAcquireItemData(__in  BYTE            *pWiasContext,
                         LONG            lFlags,
                     __in  PMINIDRV_TRANSFER_CONTEXT pmdtc,
                     __out   LONG            *plDevErrVal)
{
  HRESULT hr          = E_INVALIDARG;
  GUID  guidItemCategory  = GUID_NULL;

  if((pWiasContext)&&(pmdtc)&&(plDevErrVal))
  {
    //
    // Read the current transfer format that we are requested to use:
    //
    
    //
    // Read the WIA item category, to decide which data transfer handler should
    // be used.
    //
    CTWAIN_API *pTwainApi=NULL;
    hr = GetDS(pWiasContext,&pTwainApi);
    DWORD dwRes;
    TW_CUSTOMDSDATA Data;
    Data.hData=NULL;
    Data.InfoLength=0;
    if (SUCCEEDED(hr))
    {
      dwRes = pTwainApi->OpenDS();
      hr = TWAINtoWIAerror(dwRes);
    }
    if (SUCCEEDED(hr))
    {
      //trye to load profile - Data.hData will be 0 if it is not valid
      LoadProfileFromFile(pWiasContext,&Data.hData, &Data.InfoLength);
    }
    if (SUCCEEDED(hr))
    {
      if(Data.hData)//Valid profile?
      {
        dwRes = pTwainApi->SetDSCustomData(Data);
        hr = TWAINtoWIAerror(dwRes);
      }
      else
      {
        //set TWAIN capabilities depends on WIA properties
        hr = SetAllTWAIN_Caps(pWiasContext);
      }
    }
    // Update WIA properties 
    if (SUCCEEDED(hr))
    {
      hr = GetAllTWAIN_Caps(pWiasContext);
    }
    if(SUCCEEDED(hr))
    {
      hr = UpdateWIAPropDepend(pWiasContext);
    }

    if(!m_bIsWindowsVista)// different transfer for WIA 1.0
    {
      return drvAcquireItemDataWIA1(lFlags, pWiasContext, pmdtc, plDevErrVal);
    }

    if (SUCCEEDED(hr))
    {
      hr = wiasReadPropGuid(pWiasContext,WIA_IPA_ITEM_CATEGORY,&guidItemCategory,NULL,TRUE);
    }
    // get number of images to transfer 
    if (SUCCEEDED(hr))
    {
      LONG lXferCount = 0;
      
      if (!IsEqualGUID(guidItemCategory, WIA_CATEGORY_FEEDER))
      {
        lXferCount = 1;
      }
      else
      {
        hr = wiasReadPropLong(pWiasContext, WIA_IPS_PAGES, &lXferCount, NULL, TRUE);
        if (FAILED(hr))
        {
          WIAS_ERROR((g_hInst, "drvAcquireItemData: failure reading WIA_IPS_PAGES property for Feeder item, hr = 0x%lx", hr));
        }
        else if (ALL_PAGES == lXferCount)
        {
          lXferCount = -1;
        }
      }
      LL llVal;
      WORD wType;
      //check for paper - some TWAIN DS may hang UIless mode if there is no paper in ADF
      dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERLOADED,&llVal,&wType);
      if(dwRes==0&&wType==TWTY_BOOL&& (bool)llVal==false)
      {
        return WIA_ERROR_PAPER_EMPTY;
      }

      // enable DS in UIless mode
      HANDLE hTransferEvent;
      dwRes = pTwainApi->EnableDS(&hTransferEvent);
      hr = TWAINtoWIAerror(dwRes);
      if(hr!=S_OK)
      {
        return hr;
      }  

      if ((lFlags & WIA_MINIDRV_TRANSFER_DOWNLOAD) == 0)
      {
        return E_INVALIDARG;
      }
      // This is stream-based download
      IWiaMiniDrvTransferCallback *pTransferCallback = NULL;
      hr = GetTransferCallback(pmdtc, &pTransferCallback);
      if (!SUCCEEDED(hr))
      {
        WIAS_ERROR((g_hInst, "Could not get our IWiaMiniDrvTransferCallback for download"));
        return hr;
      }

      //wait for event from the scanner
      if(WaitForSingleObject(hTransferEvent,kMAX_SCAN_TIME) == WAIT_OBJECT_0 && pTwainApi->GetLastMsg()==MSG_XFERREADY)
      {
        bool bMoreImages=true;
        while ((SUCCEEDED(hr)) && bMoreImages && lXferCount)
        {
          hr = DownloadToStream(lFlags, pWiasContext, pmdtc, pTransferCallback, plDevErrVal);
          dwRes = pTwainApi->EndTransfer(&bMoreImages);//go to state 6
          HRESULT hr1 = TWAINtoWIAerror(dwRes);
          if(hr==S_OK)
          {
            hr = hr1;
          }
          lXferCount--;
        }
      }
      else
      {
        hr = ERROR_CANCELLED;
      }
      if(pTwainApi)
      {
        pTwainApi->ResetTransfer();//go to state 5
        pTwainApi->DisableDS();//go to state 4
      }
      pTransferCallback->Release();
      pTransferCallback = NULL;
    }
    else
    {
      WIAS_ERROR((g_hInst, "Failed to read the WIA_IPA_ITEM_CATEGORY property, hr = 0x%lx",hr));
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::InitializeRootItemProperties(
  __in  BYTE    *pWiasContext)
{
  HRESULT hr = E_INVALIDARG;
  LL llVal;
  LL_ARRAY llLst;
  WORD wCapType;
  DWORD dwRes;

  if(pWiasContext)
  {
    //Create instance of CTWAIN_API for this WIA context and store it in global table. Open DSM and DS
    int nCurAppID;
    CTWAIN_API* pTwainApi=NULL;
    hr = AddDS(pWiasContext,&nCurAppID,&pTwainApi);
    CWIAPropertyManager PropertyManager;
    if(SUCCEEDED(hr))
    {
      hr = PropertyManager.AddProperty(CUSTOM_ROOT_PROP_ID, CUSTOM_ROOT_PROP_ID_STR , RN, nCurAppID);
      if(FAILED(hr))
      {
        WIAS_ERROR((g_hInst, "Failed to add CUSTOM_ROOT_PROP_ID property to the property manager, hr = 0x%lx", hr));
      }
    }
    if(SUCCEEDED(hr))
    {
      //create temporary file for transfering profile from UI to the driver and store its name in WIA propety
      TCHAR szTempFileName[1024];  
      GetTempFileName((LPCTSTR)m_strProfilesPath,L"TWP",0,szTempFileName);
      BSTR bstrFirmware = SysAllocString(szTempFileName);
      if ( bstrFirmware ) 
      {
        hr = PropertyManager.AddProperty(CUSTOM_ROOT_PROP_ID1, CUSTOM_ROOT_PROP_ID1_STR, RN, bstrFirmware);
        if (FAILED(hr)) 
        {
          WIAS_ERROR((g_hInst, "Failed to add WIA_DPA_FIRMWARE_VERSION to prop manager, hr = 0x%lx", hr));
        }
        SysFreeString(bstrFirmware);
      }
      else 
      {
        hr = E_OUTOFMEMORY;
      }

    }    
    if(SUCCEEDED(hr))
    {
      GUID guidItemCategory = WIA_CATEGORY_ROOT;
      hr = PropertyManager.AddProperty(WIA_IPA_ITEM_CATEGORY, WIA_IPA_ITEM_CATEGORY_STR, RN, guidItemCategory);
      if(FAILED(hr))
      {
        WIAS_ERROR((g_hInst, "Failed to add WIA_IPA_ITEM_CATEGORY property to the property manager, hr = 0x%lx", hr));
      }
    }

    if(SUCCEEDED(hr))
    {
      LONG lDocumentHandlingCapabilities = 0;
      if(m_bHasFlatbed)
      {
        lDocumentHandlingCapabilities = lDocumentHandlingCapabilities | FLAT;
      }
      if(m_bHasFeeder)
      {
        lDocumentHandlingCapabilities = lDocumentHandlingCapabilities | FEED ;
        if(m_bDuplex)
        {
          lDocumentHandlingCapabilities = lDocumentHandlingCapabilities | DUP ;
        }
      }

      hr = PropertyManager.AddProperty(WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES, 
        WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES_STR , RN, lDocumentHandlingCapabilities);
      if(FAILED(hr))
      {
        WIAS_ERROR((g_hInst, "Failed to add WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES property to the property manager, hr = 0x%lx", hr));
      }
    }
    if(SUCCEEDED(hr))
    {
      long lDocumentHandlingStatus = 0;//default value
      if(m_bHasFlatbed)
      {
        lDocumentHandlingStatus=FLAT_READY;
      }
      hr = PropertyManager.AddProperty(WIA_DPS_DOCUMENT_HANDLING_STATUS, WIA_DPS_DOCUMENT_HANDLING_STATUS_STR , RN, lDocumentHandlingStatus);
      if(FAILED(hr))
      {
        WIAS_ERROR((g_hInst, "Failed to add WIA_DPS_DOCUMENT_HANDLING_STATUS property to the property manager, hr = 0x%lx", hr));
      }
    }

    if(SUCCEEDED(hr))
    {
      LONG lAccessRights = WIA_ITEM_READ | WIA_ITEM_WRITE;
      hr = PropertyManager.AddProperty(WIA_IPA_ACCESS_RIGHTS ,WIA_IPA_ACCESS_RIGHTS_STR ,RN, lAccessRights);
    }

    if(SUCCEEDED(hr))
    {
      LONG lMaxScanTime = kMAX_SCAN_TIME; 
      hr = PropertyManager.AddProperty(WIA_DPS_MAX_SCAN_TIME ,WIA_DPS_MAX_SCAN_TIME_STR ,RN, lMaxScanTime);
    }

    if (SUCCEEDED(hr))
    {
      BSTR bstrFirmware = SysAllocString(kFIRMWARE_VERSION);
      if ( bstrFirmware ) 
      {
        hr = PropertyManager.AddProperty(WIA_DPA_FIRMWARE_VERSION, WIA_DPA_FIRMWARE_VERSION_STR, RN, bstrFirmware);
        if (FAILED(hr)) 
        {
          WIAS_ERROR((g_hInst, "Failed to add WIA_DPA_FIRMWARE_VERSION to prop manager, hr = 0x%lx", hr));
        }
        SysFreeString(bstrFirmware);
      }
      else 
      {
        hr = E_OUTOFMEMORY;
      }
    }

    if(!m_bIsWindowsVista)
    {
      if(SUCCEEDED(hr))
      {
        LONG lOpticalXResolution = 0; 
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_XNATIVERESOLUTION ,&llVal,&wCapType);
        if(dwRes==0 && wCapType==TWTY_FIX32)
        {
          lOpticalXResolution = (short)llVal;
        }
        else if(dwRes==FAILURE(TWCC_BUMMER))
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        else // ICAP_XNATIVERESOLUTION not supported
        {
          dwRes = pTwainApi->GetCapConstrainedValues(ICAP_XRESOLUTION ,&llLst,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32 && llLst.size()>0)
          {
            llLst.sort();
            lOpticalXResolution = (short)llLst.back();
          }
          else
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }         
        }

        hr = PropertyManager.AddProperty(WIA_DPS_OPTICAL_XRES ,WIA_DPS_OPTICAL_XRES_STR ,RN,lOpticalXResolution);
      }

      if(SUCCEEDED(hr))
      {
        LONG lOpticalYResolution = 0; 
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_YNATIVERESOLUTION ,&llVal,&wCapType);
        if(dwRes==0 && wCapType==TWTY_FIX32)
        {
          lOpticalYResolution = (short)llVal;
        }
        else if(dwRes==FAILURE(TWCC_BUMMER))
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        else // ICAP_YNATIVERESOLUTION not supported
        {
          dwRes = pTwainApi->GetCapConstrainedValues(ICAP_YRESOLUTION ,&llLst,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32 && llLst.size()>0)
          {
            llLst.sort();
            lOpticalYResolution = (short)llLst.back();
          }
          else
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }         
        }
        hr = PropertyManager.AddProperty(WIA_DPS_OPTICAL_YRES ,WIA_DPS_OPTICAL_YRES_STR ,RN,lOpticalYResolution);
      }

      if(SUCCEEDED(hr))
      {
        //
        // Just basic duplex mode supported (no single back side scan):
        //
        LONG lDocumentHandlingSelectValidValues = 0;
        LONG lCurDocumentHandlingSelectValidValues = 0;
        if(m_bHasFlatbed)
        {
          lDocumentHandlingSelectValidValues |= FLATBED|FRONT_ONLY;     
          lCurDocumentHandlingSelectValidValues = FLATBED|FRONT_ONLY;
        }
        if(m_bHasFeeder)
        {
          lDocumentHandlingSelectValidValues |= FEEDER|FRONT_ONLY;   
          dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERENABLED,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_BOOL && llVal)
          {
            lCurDocumentHandlingSelectValidValues = FEEDER|FRONT_ONLY;
          }
          if(dwRes==FAILURE(TWCC_BUMMER))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          if(m_bDuplex)
          {
            lDocumentHandlingSelectValidValues |= DUPLEX;     
            if(lCurDocumentHandlingSelectValidValues&FEEDER)
            {
              dwRes = pTwainApi->GetCapCurrentValue(CAP_DUPLEXENABLED,&llVal,&wCapType);
              if(dwRes==0 && wCapType==TWTY_BOOL && llVal)
              {
                lCurDocumentHandlingSelectValidValues = FEEDER|DUPLEX;
              }
              if(dwRes==FAILURE(TWCC_BUMMER))
              {
                return WIA_ERROR_EXCEPTION_IN_DRIVER;
              }
            }
          }
        }
        hr = PropertyManager.AddProperty(WIA_DPS_DOCUMENT_HANDLING_SELECT ,WIA_DPS_DOCUMENT_HANDLING_SELECT_STR ,RWF,lCurDocumentHandlingSelectValidValues,lDocumentHandlingSelectValidValues);
      }

      if (SUCCEEDED(hr))
      {
        CBasicDynamicArray<LONG> lPageSizeArray;
        lPageSizeArray.Append(WIA_PAGE_CUSTOM);
        LONG lCurVal = WIA_PAGE_CUSTOM;
        hr = PropertyManager.AddProperty(WIA_DPS_PAGE_SIZE, WIA_DPS_PAGE_SIZE_STR, RWL, lCurVal, WIA_PAGE_CUSTOM, &lPageSizeArray);
      }

      TW_IMAGELAYOUT ImgLayout;
      if (SUCCEEDED(hr)) 
      {
        dwRes = pTwainApi->GetImageLayout(&ImgLayout);
        if(dwRes!=0)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        LONG lHorizontalSize = (long)(((LL)ImgLayout.Frame.Right-(LL)ImgLayout.Frame.Left)*1000);
        hr = PropertyManager.AddProperty(WIA_DPS_PAGE_WIDTH, WIA_DPS_PAGE_WIDTH_STR, RN, lHorizontalSize);
      }

      if (SUCCEEDED(hr)) 
      {
        LONG lVerticalSize = (long)(((LL)ImgLayout.Frame.Bottom-(LL)ImgLayout.Frame.Top)*1000);
        hr = PropertyManager.AddProperty(WIA_DPS_PAGE_HEIGHT, WIA_DPS_PAGE_HEIGHT_STR, RN, lVerticalSize);
      }

      bool bFeederEnableState = false;
      if(SUCCEEDED(hr))
      {
        dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERENABLED ,&llVal,&wCapType);
        if(dwRes==0 && wCapType==TWTY_BOOL)
        {
          bFeederEnableState = (bool)llVal;
        }
      }
      if (SUCCEEDED(hr)) {
        CBasicDynamicArray<LONG> lPreview;
        lPreview.Append(WIA_FINAL_SCAN);
        hr = PropertyManager.AddProperty(WIA_DPS_PREVIEW, WIA_DPS_PREVIEW_STR, RWL, lPreview[0], lPreview[0], &lPreview);
      }
      if(SUCCEEDED(hr))
      {
        LONG lShowPreviewControl = WIA_DONT_SHOW_PREVIEW_CONTROL;
        hr = PropertyManager.AddProperty(WIA_DPS_SHOW_PREVIEW_CONTROL ,WIA_DPS_SHOW_PREVIEW_CONTROL_STR ,RN,lShowPreviewControl);
      }
      if(SUCCEEDED(hr))
      {
        LONG lMaxPages      = 1;
        LONG lPages         = 1;
        dwRes = pTwainApi->GetCapConstrainedValues(CAP_XFERCOUNT,&llLst,&wCapType);
        if(dwRes || wCapType!=TWTY_INT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }

        llLst.sort();
        if(llLst.front()==-1)
        {
          lPages=0;
        }
        else
        {
          lPages = (long)llLst.front(); 
        }

        if(llLst.back()==-1)
        {
          lMaxPages=0;
        }
        else
        {
          lMaxPages = (long)llLst.back(); 
        }
        if(!m_bHasFeeder)
        {
          lPages=lMaxPages=1;
        }

        hr = PropertyManager.AddProperty(WIA_DPS_PAGES ,WIA_DPS_PAGES_STR ,RWR,lPages,lPages,lPages,lMaxPages,1);
      }

      if(m_bHasFeeder)
      {
        if(SUCCEEDED(hr))
        {
          dwRes = pTwainApi->SetCapability(CAP_FEEDERENABLED ,true,TWTY_BOOL);
          if(dwRes==FAILURE(TWCC_BUMMER))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERENABLED,&llVal,&wCapType);
          if(dwRes!=0 || wCapType!=TWTY_BOOL || !(bool)llVal)
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
        }

        if(SUCCEEDED(hr))
        {
          LONG lADFBedWidth;
          dwRes = pTwainApi->GetCapCurrentValue(ICAP_PHYSICALWIDTH ,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32)
          {
            lADFBedWidth = (long)(llVal*1000);
          }
          else
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          hr = PropertyManager.AddProperty(WIA_DPS_HORIZONTAL_SHEET_FEED_SIZE ,WIA_DPS_HORIZONTAL_SHEET_FEED_SIZE_STR ,RN,lADFBedWidth);
        }
        if(SUCCEEDED(hr))
        {
          LONG lADFBedHeight;
          dwRes = pTwainApi->GetCapCurrentValue(ICAP_PHYSICALHEIGHT ,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32)
          {
            lADFBedHeight = (long)(llVal*1000);
          }
          else
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          hr = PropertyManager.AddProperty(WIA_DPS_VERTICAL_SHEET_FEED_SIZE ,WIA_DPS_VERTICAL_SHEET_FEED_SIZE_STR ,RN,lADFBedHeight);
        }
        if(SUCCEEDED(hr))
        {
          LONG lminADFBedWidth = 0;
          dwRes = pTwainApi->GetCapCurrentValue(ICAP_MINIMUMWIDTH ,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32)
          {
            lminADFBedWidth = (long)(llVal*1000);
          }
          else if(dwRes==FAILURE(TWCC_BUMMER))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          hr = PropertyManager.AddProperty(WIA_DPS_MIN_HORIZONTAL_SHEET_FEED_SIZE ,WIA_DPS_MIN_HORIZONTAL_SHEET_FEED_SIZE_STR ,RN,lminADFBedWidth);
        }
        if(SUCCEEDED(hr))
        {
          LONG lminADFBedHeight = 0;
          dwRes = pTwainApi->GetCapCurrentValue(ICAP_MINIMUMHEIGHT ,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32)
          {
            lminADFBedHeight = (long)(llVal*1000);
          }
          else if(dwRes==FAILURE(TWCC_BUMMER))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          hr = PropertyManager.AddProperty(WIA_DPS_MIN_VERTICAL_SHEET_FEED_SIZE ,WIA_DPS_MIN_VERTICAL_SHEET_FEED_SIZE_STR ,RN,lminADFBedHeight);
        }

        if(SUCCEEDED(hr))
        {
          LONG lSheetFeederRegistration = LEFT_JUSTIFIED;
          dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERALIGNMENT,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_UINT16)
          {
            if(llVal==TWFA_CENTER)
            {
              lSheetFeederRegistration = CENTERED;
            }
            else if(llVal==TWFA_RIGHT)
            {
              lSheetFeederRegistration = RIGHT_JUSTIFIED;
            }
          }
          else if(dwRes==FAILURE(TWCC_BUMMER))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }

          hr = PropertyManager.AddProperty(WIA_DPS_SHEET_FEEDER_REGISTRATION ,WIA_DPS_SHEET_FEEDER_REGISTRATION_STR ,RN,lSheetFeederRegistration);
        }
      }
      if(m_bHasFlatbed)
      {
        if(SUCCEEDED(hr))
        {
          dwRes = pTwainApi->SetCapability(CAP_FEEDERENABLED ,false,TWTY_BOOL);
          if(dwRes && dwRes!=FAILURE(TWCC_CAPUNSUPPORTED))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          if(!dwRes)//if supported
          {
            dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERENABLED,&llVal,&wCapType);
            if(dwRes!=0 || wCapType!=TWTY_BOOL || (bool)llVal)
            {
              return WIA_ERROR_EXCEPTION_IN_DRIVER;
            }
          }
        }

        if(SUCCEEDED(hr))
        {
          LONG lADFBedWidth;
          dwRes = pTwainApi->GetCapCurrentValue(ICAP_PHYSICALWIDTH ,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32)
          {
            lADFBedWidth = (long)(llVal*1000);
          }
          else
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          hr = PropertyManager.AddProperty(WIA_DPS_HORIZONTAL_BED_SIZE ,WIA_DPS_HORIZONTAL_BED_SIZE_STR ,RN,lADFBedWidth);
        }

        if(SUCCEEDED(hr))
        {
          LONG lADFBedHeight;
          dwRes = pTwainApi->GetCapCurrentValue(ICAP_PHYSICALHEIGHT ,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32)
          {
            lADFBedHeight = (long)(llVal*1000);
          }
          else
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          hr = PropertyManager.AddProperty(WIA_DPS_VERTICAL_BED_SIZE ,WIA_DPS_VERTICAL_BED_SIZE_STR ,RN,lADFBedHeight);
        }

        if(SUCCEEDED(hr))
        {
          LONG lminADFBedWidth = 0;
          dwRes = pTwainApi->GetCapCurrentValue(ICAP_MINIMUMWIDTH ,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32)
          {
            lminADFBedWidth = (long)(llVal*1000);
          }
          else if(dwRes==FAILURE(TWCC_BUMMER))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          hr = PropertyManager.AddProperty(CUSTOM_ROOT_PROP_WIA_DPS_MIN_HORIZONTAL_BED_SIZE ,CUSTOM_ROOT_PROP_WIA_DPS_MIN_HORIZONTAL_BED_SIZE_STR ,RN,lminADFBedWidth);
        }
        if(SUCCEEDED(hr))
        {
          LONG lminADFBedHeight = 0;
          dwRes = pTwainApi->GetCapCurrentValue(ICAP_MINIMUMHEIGHT ,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32)
          {
            lminADFBedHeight = (long)(llVal*1000);
          }
          else if(dwRes==FAILURE(TWCC_BUMMER))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          hr = PropertyManager.AddProperty(CUSTOM_ROOT_PROP_WIA_DPS_MIN_VERTICAL_BED_SIZE ,CUSTOM_ROOT_PROP_WIA_DPS_MIN_VERTICAL_BED_SIZE_STR ,RN,lminADFBedHeight);
        }

        if(SUCCEEDED(hr))
        {
          LONG lSheetFeederRegistration = LEFT_JUSTIFIED;
          hr = PropertyManager.AddProperty(WIA_DPS_HORIZONTAL_BED_REGISTRATION ,WIA_DPS_HORIZONTAL_BED_REGISTRATION_STR ,RN,lSheetFeederRegistration);
        }
        if(SUCCEEDED(hr))
        {
          LONG lSheetFeederRegistration = TOP_JUSTIFIED;
          hr = PropertyManager.AddProperty(WIA_DPS_VERTICAL_BED_REGISTRATION ,WIA_DPS_VERTICAL_BED_REGISTRATION_STR ,RN,lSheetFeederRegistration);
        }
      }
      if(SUCCEEDED(hr))
      {
        dwRes = pTwainApi->SetCapability(CAP_FEEDERENABLED ,bFeederEnableState,TWTY_BOOL);
        if(dwRes && dwRes!=FAILURE(TWCC_CAPUNSUPPORTED))
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        if(!dwRes)//if supported
        {
          dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERENABLED,&llVal,&wCapType);
          if(dwRes!=0 || wCapType!=TWTY_BOOL || ((bool)llVal)!=bFeederEnableState)
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
        }
      }
    }

    if(SUCCEEDED(hr))
    {
      hr = PropertyManager.SetItemProperties(pWiasContext);
      if(FAILED(hr))
      {
        WIAS_ERROR((g_hInst, "CWIAPropertyManager::SetItemProperties failed to set WIA root item properties, hr = 0x%lx",hr));
      }
    }
    else
    {
      WIAS_ERROR((g_hInst, "Failed to add WIA_IPA_ITEM_CATEGORY property to the property manager, hr = 0x%lx",hr));
    }
  }
  else
  {
    WIAS_ERROR((g_hInst, "Invalid parameters were passed"));
  }
  return hr;
}

HRESULT CWIADriver::InitializeWIAItemProperties(
  __in  BYTE    *pWiasContext,
      GUID guidItemCategory)
{
  HRESULT hr = E_INVALIDARG;
  LL llVal;
  LL llDefVal;
  LL_ARRAY llLst;
  WORD wCapType;
  DWORD dwRes;
  CTWAIN_API *pTwainApi=NULL;

  if(pWiasContext)
  {
    //get CTWAIN_API associated with this context
    hr = GetDS(pWiasContext,&pTwainApi);
    //Open TWAIN DS so we can get capabilities 
    if (SUCCEEDED(hr))
    {
      dwRes = pTwainApi->OpenDS();
      hr = TWAINtoWIAerror(dwRes);
    }

    if (SUCCEEDED(hr))
    {
      CWIAPropertyManager PropertyManager;

      LONG lXPosition     = 0;
      LONG lYPosition     = 0;
      LONG lXExtent       = 1;
      LONG lYExtent       = 1;
      LONG lPixWidthMax      = 1;
      LONG lPixHeightMax     = 1;
      LONG lPixWidthMin      = 1;
      LONG lPixHeightMin     = 1;
      LONG lItemType      = 0;     
      LONG lPixelType      = WIA_DATA_THRESHOLD;     

      hr = wiasGetItemType(pWiasContext,&lItemType);
      if(SUCCEEDED(hr))
      {
        if(lItemType & WiaItemTypeGenerated)
        {
          WIAS_TRACE((g_hInst,"WIA item was created by application."));
        }
      }
      else
      {
        WIAS_ERROR((g_hInst, "Failed to get the WIA item type, hr = 0x%lx",hr));
      }
      bool bFeeder= (WIA_CATEGORY_FEEDER==guidItemCategory)?true:false;

      if(SUCCEEDED(hr))
      {
        dwRes = pTwainApi->SetCapability(CAP_FEEDERENABLED ,bFeeder,TWTY_BOOL);
        if(dwRes && dwRes!=FAILURE(TWCC_CAPUNSUPPORTED))
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        if(!dwRes)//if supported
        {
          dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERENABLED,&llVal,&wCapType);
          if(dwRes!=0 || wCapType!=TWTY_BOOL || (bool)llVal!=bFeeder)
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
        }
      }

      //
      // Add all common item properties first
      //
      TW_IMAGELAYOUT ImgLayout;
      LL llBitDepth = 1;
      LL llXRes = 0;
      LL llYRes = 0;
      LL llHorizontalSize=0;
      LL llVerticalSize=0;
      LL llMinHorizontalSize = 0;
      LL llMinVerticalSize = 0;
      dwRes = pTwainApi->GetCapCurrentValue(ICAP_YRESOLUTION ,&llYRes,&wCapType);
      if(dwRes!=0 || wCapType!=TWTY_FIX32)
      {
        return WIA_ERROR_EXCEPTION_IN_DRIVER;
      }

      dwRes = pTwainApi->GetCapCurrentValue(ICAP_XRESOLUTION ,&llXRes,&wCapType);
      if(dwRes!=0 || wCapType!=TWTY_FIX32)
      {
        return WIA_ERROR_EXCEPTION_IN_DRIVER;
      }

      if (SUCCEEDED(hr)) 
      {
        dwRes = pTwainApi->GetImageLayout(&ImgLayout);
        if(dwRes!=0)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        lXPosition = (long)(((LL)ImgLayout.Frame.Left)*llXRes);
        lYPosition = (long)(((LL)ImgLayout.Frame.Top)*llYRes);
        lXExtent = (long)((((LL)ImgLayout.Frame.Right)-((LL)ImgLayout.Frame.Left))*llXRes);
        lYExtent = (long)((((LL)ImgLayout.Frame.Bottom)-((LL)ImgLayout.Frame.Top))*llYRes);

        dwRes = pTwainApi->GetCapCurrentValue(ICAP_PHYSICALWIDTH ,&llHorizontalSize,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_FIX32)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_PHYSICALHEIGHT ,&llVerticalSize,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_FIX32)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }


        dwRes = pTwainApi->GetCapCurrentValue(ICAP_MINIMUMWIDTH ,&llMinHorizontalSize,&wCapType);
        if(dwRes==FAILURE(TWCC_BUMMER))
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }          

        dwRes = pTwainApi->GetCapCurrentValue(ICAP_MINIMUMHEIGHT ,&llMinVerticalSize,&wCapType);
        if(dwRes==FAILURE(TWCC_BUMMER))
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }   

        lPixWidthMax      = (long)(llHorizontalSize*llXRes);
        lPixHeightMax     = (long)(llVerticalSize*llYRes);
        lPixWidthMin      = (long)(llMinHorizontalSize*llXRes);
        lPixHeightMin     = (long)(llMinVerticalSize*llYRes);

        if(lXPosition<0)
        {
          lXPosition=0;
        }
        if(lXPosition>lPixWidthMax)
        {
          lXPosition=lPixWidthMax;
        }
        if(lXExtent<lPixWidthMin)
        {
          lXExtent=lPixWidthMin;
        }
        if(lXExtent>lPixWidthMax)
        {
          lXExtent=lPixWidthMax;
        }
        if(lYPosition<0)
        {
          lYPosition=0;
        }
        if(lYPosition>lPixHeightMax)
        {
          lYPosition=lPixHeightMax;
        }
        if(lYExtent<lPixHeightMin)
        {
          lYExtent=lPixHeightMin;
        }
        if(lYExtent>lPixHeightMax)
        {
          lYExtent=lPixHeightMax;
        }
      }
      if (SUCCEEDED(hr)) 
      {
        LONG lAccessRights = WIA_ITEM_READ|WIA_ITEM_WRITE;
        hr = PropertyManager.AddProperty(WIA_IPA_ACCESS_RIGHTS ,WIA_IPA_ACCESS_RIGHTS_STR ,RF, lAccessRights, lAccessRights);
      }

      if(SUCCEEDED(hr))
      {
        hr = PropertyManager.AddProperty(WIA_IPA_ITEM_CATEGORY, WIA_IPA_ITEM_CATEGORY_STR, RN, guidItemCategory);
        if(FAILED(hr))
        {
          WIAS_ERROR((g_hInst, "Failed to add WIA_IPA_ITEM_CATEGORY property to the property manager, hr = 0x%lx", hr));
        }
      }

      if (SUCCEEDED(hr)) 
      {
        CBasicDynamicArray<LONG> lOrientationArray;
        lOrientationArray.Append(PORTRAIT);
        long lCurVal = PORTRAIT;
        hr = PropertyManager.AddProperty(WIA_IPS_ORIENTATION, WIA_IPS_ORIENTATION_STR, RWL, lCurVal, PORTRAIT, &lOrientationArray);
      }

      if(SUCCEEDED(hr))
      {
        LONG lNumberOfLines = (long)(((LL)(ImgLayout.Frame.Bottom)-(LL)(ImgLayout.Frame.Top))*llYRes);
        hr = PropertyManager.AddProperty(WIA_IPA_NUMBER_OF_LINES ,WIA_IPA_NUMBER_OF_LINES_STR ,RN,lNumberOfLines);
      }
      LONG lPixelsPerLine = 0;
      if(SUCCEEDED(hr))
      {
        lPixelsPerLine = (long)(((LL)(ImgLayout.Frame.Right)-(LL)(ImgLayout.Frame.Left))*llXRes);
        hr = PropertyManager.AddProperty(WIA_IPA_PIXELS_PER_LINE ,WIA_IPA_PIXELS_PER_LINE_STR ,RN,lPixelsPerLine);
      }

      if(SUCCEEDED(hr))
      {
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_BITDEPTH ,&llBitDepth,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_UINT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        LONG lBytesPerLine = ((lPixelsPerLine*(long)llBitDepth+31)/32)*4;

        hr = PropertyManager.AddProperty(WIA_IPA_BYTES_PER_LINE ,WIA_IPA_BYTES_PER_LINE_STR ,RN,lBytesPerLine);
      }

      if(SUCCEEDED(hr))
      {
        dwRes = pTwainApi->GetCapConstrainedValues(ICAP_XRESOLUTION ,&llLst,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_FIX32 || llLst.size()==0)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_XRESOLUTION ,&llVal,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_FIX32)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        dwRes = pTwainApi->GetCapDefaultValue(ICAP_XRESOLUTION ,&llDefVal,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_FIX32)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }

        CBasicDynamicArray<LONG> lXResolutionArray;
        LL_ARRAY::iterator llIter=llLst.begin();
        for(;llIter!=llLst.end();llIter++)
        {
          lXResolutionArray.Append((short)(*llIter));
        }
        hr = PropertyManager.AddProperty(WIA_IPS_XRES ,WIA_IPS_XRES_STR ,RWLC,(short)(llVal),(short)(llDefVal),&lXResolutionArray);
      }

      if(SUCCEEDED(hr))
      {
        dwRes = pTwainApi->GetCapConstrainedValues(ICAP_YRESOLUTION ,&llLst,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_FIX32 || llLst.size()==0)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_YRESOLUTION ,&llVal,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_FIX32)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        dwRes = pTwainApi->GetCapDefaultValue(ICAP_YRESOLUTION ,&llDefVal,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_FIX32)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }

        CBasicDynamicArray<LONG> lYResolutionArray;
        LL_ARRAY::iterator llIter=llLst.begin();
        for(;llIter!=llLst.end();llIter++)
        {
          lYResolutionArray.Append((short)(*llIter));
        }
        hr = PropertyManager.AddProperty(WIA_IPS_YRES ,WIA_IPS_YRES_STR ,RWLC,(short)(llVal),(short)(llDefVal),&lYResolutionArray);
      }

      if(SUCCEEDED(hr))
      {
        hr = PropertyManager.AddProperty(WIA_IPS_XPOS, WIA_IPS_XPOS_STR, RWRC, lXPosition, 0, 0, lPixWidthMax - lPixWidthMin, 1);
      }

      if(SUCCEEDED(hr))
      {
        hr = PropertyManager.AddProperty(WIA_IPS_YPOS, WIA_IPS_YPOS_STR, RWRC, lYPosition, 0, 0, lPixHeightMax -lPixHeightMin, 1);
      }

      if(SUCCEEDED(hr))
      {
        hr = PropertyManager.AddProperty(WIA_IPS_XEXTENT ,WIA_IPS_XEXTENT_STR ,RWRC, lXExtent, lXExtent, lPixWidthMin, lPixWidthMax - lXPosition, 1);
      }

      if(SUCCEEDED(hr))
      {
        hr = PropertyManager.AddProperty(WIA_IPS_YEXTENT ,WIA_IPS_YEXTENT_STR ,RWRC, lYExtent, lYExtent, lPixHeightMin, lPixHeightMax - lYPosition, 1);
      }

      if(SUCCEEDED(hr))
      {
        llDefVal = 0L;
        llVal = 0L;
        LL llMax=1000L;
        LL llMin=-1000L;
        LL llStep=1L;
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_BRIGHTNESS ,&llVal,&wCapType);

        if(dwRes==FAILURE(TWCC_BUMMER) ||  (dwRes==0 && wCapType!=TWTY_FIX32))
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        if(dwRes==0)
        {
          dwRes = pTwainApi->GetCapDefaultValue(ICAP_BRIGHTNESS ,&llDefVal,&wCapType);
          if(dwRes==FAILURE(TWCC_BUMMER) ||  (dwRes==0 && wCapType!=TWTY_FIX32))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
        }
        if(dwRes==0)
        {
          dwRes = pTwainApi->GetCapMinMaxValues(ICAP_BRIGHTNESS,&llMin,&llMax,&llStep,&wCapType);
          if(dwRes==FAILURE(TWCC_BUMMER) ||  (dwRes==0 && wCapType!=TWTY_FIX32))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          if((long)llStep<=0)
          {
            llStep = 1L;
          }
        }
        hr = PropertyManager.AddProperty(WIA_IPS_BRIGHTNESS,WIA_IPS_BRIGHTNESS_STR,RWRC,(short)(llVal),(short)(llDefVal),-1000,1000,(BYTE)llStep);
      }

      if(SUCCEEDED(hr))
      {
        llDefVal = 0L;
        LL llMax=1000L;
        LL llMin=-1000L;
        LL llStep=1L;
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_CONTRAST ,&llVal,&wCapType);

        if(dwRes==FAILURE(TWCC_BUMMER) ||  (dwRes==0 && wCapType!=TWTY_FIX32))
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        if(dwRes==0)
        {
          dwRes = pTwainApi->GetCapDefaultValue(ICAP_CONTRAST ,&llDefVal,&wCapType);
          if(dwRes==FAILURE(TWCC_BUMMER) ||  (dwRes==0 && wCapType!=TWTY_FIX32))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
        }
        if(dwRes==0)
        {
          dwRes = pTwainApi->GetCapMinMaxValues(ICAP_CONTRAST,&llMin,&llMax,&llStep,&wCapType);
          if(dwRes==FAILURE(TWCC_BUMMER) ||  (dwRes==0 && wCapType!=TWTY_FIX32))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          if((long)llStep<=0)
          {
            llStep = 1L;
          }
        }
        hr = PropertyManager.AddProperty(WIA_IPS_CONTRAST,WIA_IPS_CONTRAST_STR,RWRC,(short)(llVal),(short)(llDefVal),-1000,1000,(BYTE)llStep);
      }

      if(SUCCEEDED(hr))
      {
        llVal = (BYTE)128;
        LL llMax=(BYTE)255;
        LL llMin=(BYTE)0;
        LL llStep=(BYTE)1;
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_THRESHOLD  ,&llVal,&wCapType);
        if(dwRes==FAILURE(TWCC_BUMMER) ||  (dwRes==0 && wCapType!=TWTY_FIX32))
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        if(dwRes==0)
        {
          dwRes = pTwainApi->GetCapDefaultValue(ICAP_THRESHOLD  ,&llDefVal,&wCapType);
          if(dwRes==FAILURE(TWCC_BUMMER) ||  (dwRes==0 && wCapType!=TWTY_FIX32))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
        }
        if(dwRes==0)
        {
          dwRes = pTwainApi->GetCapMinMaxValues(ICAP_THRESHOLD,&llMin,&llMax,&llStep,&wCapType);
          if(dwRes==FAILURE(TWCC_BUMMER) ||  (dwRes==0 && wCapType!=TWTY_FIX32))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          if((long)llStep<=0)
          {
            llStep = 1L;
          }
        }
        hr = PropertyManager.AddProperty(WIA_IPS_THRESHOLD,WIA_IPS_THRESHOLD_STR,RWRC,(BYTE)(llVal),(BYTE)(llDefVal),0,255,(BYTE)llStep);
      }

      if(SUCCEEDED(hr))
      {
        LONG lItemSize = 0;//keep 0 to allow WIA driver to allocate memory. We needed because TWAIN DS may return biger image than expected
        hr = PropertyManager.AddProperty(WIA_IPA_ITEM_SIZE ,WIA_IPA_ITEM_SIZE_STR ,RN,lItemSize);
      }

      if(SUCCEEDED(hr))
      {
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_PIXELTYPE  ,&llVal,&wCapType);

        if(dwRes!=0 || wCapType!=TWTY_UINT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }

        dwRes = pTwainApi->GetCapDefaultValue(ICAP_PIXELTYPE  ,&llDefVal,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_UINT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }

        dwRes = pTwainApi->GetCapConstrainedValues(ICAP_PIXELTYPE  ,&llLst,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_UINT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }

        CBasicDynamicArray<LONG> lDataTypeArray;
        LL_ARRAY::iterator llIter=llLst.begin();
        for(;llIter!=llLst.end();llIter++)
        {
          switch((WORD)*llIter)
          {
            case TWPT_BW:
              lDataTypeArray.Append(WIA_DATA_THRESHOLD);
              break;
            case TWPT_GRAY:
              lDataTypeArray.Append(WIA_DATA_GRAYSCALE);
              break;
            case TWPT_RGB:
              lDataTypeArray.Append(WIA_DATA_COLOR);
              break;
            default:
              break;
          }
        }
        switch((WORD)llVal)
        {
          case TWPT_BW:
            lPixelType = WIA_DATA_THRESHOLD;
            break;
          case TWPT_GRAY:
            lPixelType = WIA_DATA_GRAYSCALE;
            break;
          case TWPT_RGB:
            lPixelType = WIA_DATA_COLOR;
            break;
          default:
            break;
        }
        switch((WORD)llDefVal)
        {
          case TWPT_BW:
            llDefVal = (WORD)WIA_DATA_THRESHOLD;
            break;
          case TWPT_GRAY:
            llDefVal = (WORD)WIA_DATA_GRAYSCALE;
            break;
          case TWPT_RGB:
            llDefVal = (WORD)WIA_DATA_COLOR;
            break;
          default:
            break;
        }
        
        hr = PropertyManager.AddProperty(WIA_IPA_DATATYPE ,WIA_IPA_DATATYPE_STR ,RWL,lPixelType,(long)llDefVal,&lDataTypeArray);
      }

      if(SUCCEEDED(hr))
      {
        LONG lCurrentIntent             = WIA_INTENT_NONE;
        LONG lCurrentIntentValidValues  =   WIA_INTENT_NONE | WIA_INTENT_MINIMIZE_SIZE | WIA_INTENT_MAXIMIZE_QUALITY;
        CBasicDynamicArray<LONG> lDataTypeArray;
        LL_ARRAY::iterator llIter=llLst.begin();
        for(;llIter!=llLst.end();llIter++)
        {
          switch((WORD)*llIter)
          {
            case TWPT_BW:
              lCurrentIntentValidValues |=WIA_INTENT_IMAGE_TYPE_TEXT;
              break;
            case TWPT_GRAY:
              lCurrentIntentValidValues |=WIA_INTENT_IMAGE_TYPE_GRAYSCALE;
              break;
            case TWPT_RGB:
              lCurrentIntentValidValues |=WIA_INTENT_IMAGE_TYPE_COLOR;
              break;
            default:
              break;
          }
        }

        hr = PropertyManager.AddProperty(WIA_IPS_CUR_INTENT ,WIA_IPS_CUR_INTENT_STR ,RWF,lCurrentIntent,lCurrentIntentValidValues);
      }

      if(SUCCEEDED(hr))
      {
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_BITDEPTH  ,&llVal,&wCapType);

        if(dwRes!=0 || wCapType!=TWTY_UINT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }

        dwRes = pTwainApi->GetCapDefaultValue(ICAP_BITDEPTH  ,&llDefVal,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_UINT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }

        dwRes = pTwainApi->GetCapConstrainedValues(ICAP_BITDEPTH  ,&llLst,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_UINT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        LL_ARRAY::iterator llIter=llLst.begin();
        CBasicDynamicArray<LONG> lBitDepthArray;
        for(;llIter!=llLst.end();llIter++)
        {
          lBitDepthArray.Append((LONG)*llIter);
        }

        hr = PropertyManager.AddProperty(WIA_IPA_DEPTH ,WIA_IPA_DEPTH_STR ,RWLC,(LONG)llVal,(LONG)llDefVal,&lBitDepthArray);
      }

      if(SUCCEEDED(hr))
      {
        GUID guidPreferredFormat = WiaImgFmt_BMP;
        if(!m_bIsWindowsVista)
        {
          guidPreferredFormat = WiaImgFmt_MEMORYBMP;
        }
        hr = PropertyManager.AddProperty(WIA_IPA_PREFERRED_FORMAT ,WIA_IPA_PREFERRED_FORMAT_STR ,RN,guidPreferredFormat);
      }

      if(SUCCEEDED(hr))
      {
        CBasicDynamicArray<GUID> guidFormatArray;
        guidFormatArray.Append(WiaImgFmt_BMP);
        if(!m_bIsWindowsVista)
        {
          guidFormatArray.Append(WiaImgFmt_MEMORYBMP);
        }
        hr = PropertyManager.AddProperty(WIA_IPA_FORMAT ,WIA_IPA_FORMAT_STR ,RWL,guidFormatArray[0],guidFormatArray[0],&guidFormatArray);
      }

      if(SUCCEEDED(hr))
      {
        CBasicDynamicArray<LONG> lCompressionArray;
        lCompressionArray.Append(WIA_COMPRESSION_NONE);
        hr = PropertyManager.AddProperty(WIA_IPA_COMPRESSION ,WIA_IPA_COMPRESSION_STR ,RWL, lCompressionArray[0], lCompressionArray[0], &lCompressionArray);
      }

      if(SUCCEEDED(hr))
      {
        CBasicDynamicArray<LONG> lTymedArray;
        lTymedArray.Append(TYMED_FILE);
        if(!m_bIsWindowsVista)
        {
          lTymedArray.Append(TYMED_CALLBACK);
        }
        hr = PropertyManager.AddProperty(WIA_IPA_TYMED ,WIA_IPA_TYMED_STR ,RWL,lTymedArray[0],lTymedArray[0],&lTymedArray);
      }

      if(SUCCEEDED(hr))
      {
        LONG lChannelsPerPixel = 1;
        if(lPixelType==WIA_DATA_COLOR)
        {
          lChannelsPerPixel =3;
        }
        hr = PropertyManager.AddProperty(WIA_IPA_CHANNELS_PER_PIXEL ,WIA_IPA_CHANNELS_PER_PIXEL_STR ,RN,lChannelsPerPixel);
      }

      if(SUCCEEDED(hr))
      {
        LONG lBitsPerChannel = (LONG)llBitDepth;
        if(lPixelType==WIA_DATA_COLOR)
        {
          lBitsPerChannel /=3;
        }
        hr = PropertyManager.AddProperty(WIA_IPA_BITS_PER_CHANNEL ,WIA_IPA_BITS_PER_CHANNEL_STR ,RN,lBitsPerChannel);
      }

      if(SUCCEEDED(hr))
      {
        dwRes = pTwainApi->GetCapCurrentValue(ICAP_PIXELFLAVOR  ,&llVal,&wCapType);

        if(dwRes!=0 || wCapType!=TWTY_UINT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }

        if(llVal == TWPF_CHOCOLATE )
        {
          llVal = (WORD)WIA_PHOTO_WHITE_1;
        }
        else
        {
          llVal = (WORD)WIA_PHOTO_WHITE_0;
        }
        dwRes = pTwainApi->GetCapDefaultValue(ICAP_PIXELFLAVOR   ,&llDefVal,&wCapType);
        if(dwRes!=0 || wCapType!=TWTY_UINT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        if(llDefVal == TWPF_CHOCOLATE )
        {
          llDefVal = (WORD)WIA_PHOTO_WHITE_1;
        }
        else
        {
          llDefVal = (WORD)WIA_PHOTO_WHITE_0;
        }
        dwRes = pTwainApi->GetCapConstrainedValues(ICAP_PIXELFLAVOR  ,&llLst,&wCapType);

        if(dwRes!=0 || wCapType!=TWTY_UINT16)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        LL_ARRAY::iterator llIter=llLst.begin();
        CBasicDynamicArray<LONG> lPhotometricInterpArray;
        for(;llIter!=llLst.end();llIter++)
        {
          if(*llIter == TWPF_CHOCOLATE )
          {
            lPhotometricInterpArray.Append(WIA_PHOTO_WHITE_1);
          }
          else
          {
            lPhotometricInterpArray.Append(WIA_PHOTO_WHITE_0);
          }
        }

        hr = PropertyManager.AddProperty(WIA_IPS_PHOTOMETRIC_INTERP ,WIA_IPS_PHOTOMETRIC_INTERP_STR ,RWL,(LONG)llVal,(LONG)llDefVal,&lPhotometricInterpArray);

      }

      if(SUCCEEDED(hr))
      {
        LONG lPlanar = WIA_PACKED_PIXEL;//Scanning DIB only
        hr = PropertyManager.AddProperty(WIA_IPA_PLANAR ,WIA_IPA_PLANAR_STR ,RN,lPlanar);
      }

      if(SUCCEEDED(hr))
      {
        TW_SETUPMEMXFER memxfer;
        dwRes=pTwainApi->GetMemTransferCfg(&memxfer);
        if(dwRes!=0)
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }

        hr = PropertyManager.AddProperty(WIA_IPA_MIN_BUFFER_SIZE ,WIA_IPA_MIN_BUFFER_SIZE_STR ,RN,memxfer.Preferred);
      }
      
      if(m_bIsWindowsVista)
      {
        if(SUCCEEDED(hr))
        {
          LONG lVal = 0;
          hr = PropertyManager.AddProperty(WIA_IPS_SUPPORTS_CHILD_ITEM_CREATION, WIA_IPS_SUPPORTS_CHILD_ITEM_CREATION_STR, RN, lVal);
        }
        if(SUCCEEDED(hr))
        {
          LONG lVal = WIA_DONT_SHOW_PREVIEW_CONTROL;
          hr = PropertyManager.AddProperty(WIA_IPS_SHOW_PREVIEW_CONTROL, WIA_IPS_SHOW_PREVIEW_CONTROL_STR, RN, lVal);
        }
        if(SUCCEEDED(hr))
        {
          CBasicDynamicArray<LONG> lArray;
          lArray.Append(WIA_FINAL_SCAN);
          LONG lCurVal = WIA_FINAL_SCAN;
          hr = PropertyManager.AddProperty(WIA_IPS_PREVIEW, WIA_IPS_PREVIEW_STR, RWL, lCurVal, WIA_FINAL_SCAN, &lArray);
        }
        if(SUCCEEDED(hr))
        {
          BSTR bstrFileType = SysAllocString(kFILE_TYPE);
          if ( bstrFileType ) 
          {
            hr = PropertyManager.AddProperty(WIA_IPA_FILENAME_EXTENSION, WIA_IPA_FILENAME_EXTENSION_STR, RN, bstrFileType);
            if (FAILED(hr)) 
            {
              WIAS_ERROR((g_hInst, "Failed to add WIA_DPA_FIRMWARE_VERSION to prop manager, hr = 0x%lx", hr));
            }
            SysFreeString(bstrFileType);
          }
          else 
          {
            hr = E_OUTOFMEMORY;
          }
        }

        if(SUCCEEDED(hr))
        {
          CBasicDynamicArray<LONG> lPageSizeArray;
          lPageSizeArray.Append(WIA_PAGE_CUSTOM);
          LONG lCurVal = WIA_PAGE_CUSTOM;
          hr = PropertyManager.AddProperty(WIA_IPS_PAGE_SIZE, WIA_IPS_PAGE_SIZE_STR, RWL, lCurVal, WIA_PAGE_CUSTOM, &lPageSizeArray);
        }

        if(bFeeder)//because of WINQUAL
        {
          if (SUCCEEDED(hr)) 
          {
            LONG lHorizontalSize = (long)(((LL)ImgLayout.Frame.Right-(LL)ImgLayout.Frame.Left)*1000);
            hr = PropertyManager.AddProperty(WIA_IPS_PAGE_WIDTH, WIA_IPS_PAGE_WIDTH_STR, RN, lHorizontalSize);
          }

          if (SUCCEEDED(hr)) 
          {
            LONG lVerticalSize = (long)(((LL)ImgLayout.Frame.Bottom-(LL)ImgLayout.Frame.Top)*1000);
            hr = PropertyManager.AddProperty(WIA_IPS_PAGE_HEIGHT, WIA_IPS_PAGE_HEIGHT_STR, RN, lVerticalSize);
          }
        }
        if (SUCCEEDED(hr)) 
        {
          LONG lRotation=0;
          CBasicDynamicArray<LONG> lRotationArray;
          lRotationArray.Append(0);
          hr = PropertyManager.AddProperty(WIA_IPS_ROTATION, WIA_IPS_ROTATION_STR, RN, lRotation);
          hr = PropertyManager.AddProperty(WIA_IPS_ROTATION, WIA_IPS_ROTATION_STR, RWL, lRotation, lRotation, &lRotationArray);
        }
        if(SUCCEEDED(hr))
        {
          //
          // Just basic duplex mode supported (no single back side scan):
          //
          LONG lDocumentHandlingSelectValidValues = FRONT_ONLY;
          LONG lCurDocumentHandlingSelectValidValues = FRONT_ONLY;
          if(bFeeder)
          {
            if(m_bDuplex)
            {
              lDocumentHandlingSelectValidValues |= DUPLEX;     
              dwRes = pTwainApi->GetCapCurrentValue(CAP_DUPLEXENABLED,&llVal,&wCapType);
              if(dwRes==0 && wCapType==TWTY_BOOL && llVal)
              {
                lCurDocumentHandlingSelectValidValues = DUPLEX;
              }
              if(dwRes==FAILURE(TWCC_BUMMER))
              {
                return WIA_ERROR_EXCEPTION_IN_DRIVER;
              }
            }
          }
          hr = PropertyManager.AddProperty(WIA_IPS_DOCUMENT_HANDLING_SELECT ,WIA_IPS_DOCUMENT_HANDLING_SELECT_STR ,RWF,lCurDocumentHandlingSelectValidValues,lDocumentHandlingSelectValidValues);
        }

        if(SUCCEEDED(hr))
        {
          LONG lHorizontalSize = (long)(llHorizontalSize*1000);
          hr = PropertyManager.AddProperty(WIA_IPS_MAX_HORIZONTAL_SIZE ,WIA_IPS_MAX_HORIZONTAL_SIZE_STR ,RN,lHorizontalSize);
        }

        if(SUCCEEDED(hr))
        {
          LONG lVerticalSize = (long)(llVerticalSize*1000);
          hr = PropertyManager.AddProperty(WIA_IPS_MAX_VERTICAL_SIZE ,WIA_IPS_MAX_VERTICAL_SIZE_STR ,RN,lVerticalSize);
        }

        if(SUCCEEDED(hr))
        {
          LONG lMinHorizontalSize = (long)(llMinHorizontalSize*1000);
          lMinHorizontalSize = MAX(lMinHorizontalSize,1);
          hr = PropertyManager.AddProperty(WIA_IPS_MIN_HORIZONTAL_SIZE ,WIA_IPS_MIN_HORIZONTAL_SIZE_STR ,RN,lMinHorizontalSize);
        }

        if(SUCCEEDED(hr))
        {
          LONG lMinVerticalSize = (long)(llMinVerticalSize*1000);
          lMinVerticalSize = MAX(lMinVerticalSize,1);
          hr = PropertyManager.AddProperty(WIA_IPS_MIN_VERTICAL_SIZE ,WIA_IPS_MIN_VERTICAL_SIZE_STR ,RN,lMinVerticalSize);
        }

        if(SUCCEEDED(hr))
        {
          LONG lOpticalXResolution = 0; 
          dwRes = pTwainApi->GetCapCurrentValue(ICAP_XNATIVERESOLUTION ,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32)
          {
            lOpticalXResolution = (short)llVal;
          }
          else if(dwRes==FAILURE(TWCC_BUMMER))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          else // ICAP_XNATIVERESOLUTION not supported
          {
            dwRes = pTwainApi->GetCapConstrainedValues(ICAP_XRESOLUTION ,&llLst,&wCapType);
            if(dwRes==0 && wCapType==TWTY_FIX32 && llLst.size()>0)
            {
              llLst.sort();
              lOpticalXResolution = (short)llLst.back();
            }
            else
            {
              return WIA_ERROR_EXCEPTION_IN_DRIVER;
            }         
          }
          hr = PropertyManager.AddProperty(WIA_IPS_OPTICAL_XRES ,WIA_IPS_OPTICAL_XRES_STR ,RN,lOpticalXResolution);
        }

        if(SUCCEEDED(hr))
        {
          LONG lOpticalYResolution = 0; 
          dwRes = pTwainApi->GetCapCurrentValue(ICAP_YNATIVERESOLUTION ,&llVal,&wCapType);
          if(dwRes==0 && wCapType==TWTY_FIX32)
          {
            lOpticalYResolution = (short)llVal;
          }
          else if(dwRes==FAILURE(TWCC_BUMMER))
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
          else // ICAP_YNATIVERESOLUTION not supported
          {
            dwRes = pTwainApi->GetCapConstrainedValues(ICAP_YRESOLUTION ,&llLst,&wCapType);
            if(dwRes==0 && wCapType==TWTY_FIX32 && llLst.size()>0)
            {
              llLst.sort();
              lOpticalYResolution = (short)llLst.back();
            }
            else
            {
              return WIA_ERROR_EXCEPTION_IN_DRIVER;
            }         
          }
          hr = PropertyManager.AddProperty(WIA_IPS_OPTICAL_YRES ,WIA_IPS_OPTICAL_YRES_STR ,RN,lOpticalYResolution);
        }

        if(SUCCEEDED(hr))
        {
          LONG lMaxPages      = 1;
          LONG lPages         = 1;
          dwRes = pTwainApi->GetCapConstrainedValues(CAP_XFERCOUNT,&llLst,&wCapType);
          if(dwRes || wCapType!=TWTY_INT16)
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }

          llLst.sort();
          if(llLst.front()==-1)
          {
            lPages=0;
          }
          else
          {
            lPages = (long)llLst.front(); 
          }

          if(llLst.back()==-1)
          {
            lMaxPages=0;
          }
          else
          {
            lMaxPages = (long)llLst.back(); 
          }
          hr = PropertyManager.AddProperty(WIA_IPS_PAGES ,WIA_IPS_PAGES_STR ,RWR,lPages,lPages,lPages,lMaxPages,1);
        }

        if(bFeeder)
        {
          if(SUCCEEDED(hr))
          {
            LONG lSheetFeederRegistration = LEFT_JUSTIFIED;
            dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERALIGNMENT,&llVal,&wCapType);
            if(dwRes==0 && wCapType==TWTY_UINT16)
            {
              if(llVal==TWFA_CENTER)
              {
                lSheetFeederRegistration = CENTERED;
              }
              else if(llVal==TWFA_RIGHT)
              {
                lSheetFeederRegistration = RIGHT_JUSTIFIED;
              }
            }
            else if(dwRes==FAILURE(TWCC_BUMMER))
            {
              return WIA_ERROR_EXCEPTION_IN_DRIVER;
            }
            hr = PropertyManager.AddProperty(WIA_IPS_SHEET_FEEDER_REGISTRATION ,WIA_IPS_SHEET_FEEDER_REGISTRATION_STR ,RN,lSheetFeederRegistration);
          }
        }
      }

      if(SUCCEEDED(hr))
      {
        hr = PropertyManager.SetItemProperties(pWiasContext);
        if(FAILED(hr))
        {
          WIAS_ERROR((g_hInst, "CWIAPropertyManager::SetItemProperties failed to set WIA flatbed item properties, hr = 0x%lx",hr));
        }
      }
    }
    else
    {
      WIAS_ERROR((g_hInst, "Failed to obtain driver item context data"));
    }
  }
  else
  {
    WIAS_ERROR((g_hInst, "Invalid parameters were passed"));
  }

  return hr;
}


HRESULT CWIADriver::drvInitItemProperties(__inout   BYTE *pWiasContext,
                          LONG lFlags,
                      __out   LONG *plDevErrVal)
{
  HRESULT hr      = E_INVALIDARG;
  LONG  lItemFlags  = 0;
  if((pWiasContext)&&(plDevErrVal))
  {
    //
    // Initialize individual storage item properties using the CWIAStorage object
    //
    hr = wiasReadPropLong(pWiasContext,WIA_IPA_ITEM_FLAGS,&lItemFlags,NULL,TRUE);
    if(SUCCEEDED(hr))
    {
      if((lItemFlags & WiaItemTypeRoot))
      {
        //
        // Add any root item properties needed.  
        //
        hr = InitializeRootItemProperties(pWiasContext);
        if(FAILED(hr))
        {
          WIAS_ERROR((g_hInst, "Failed to initialize generic WIA root item properties, hr = 0x%lx",hr));
        }
      }
      else
      {
        //
        // Add any non-root item properties needed.
        //
        GUID guidItemCategory = GUID_NULL;

        //
        // Item is not a generated item, assume that this was created by this WIA driver
        // and the WIA_ITEM_CATEGORY setting can be read from the WIA_DRIVER_ITEM_CONTEXT
        // structure stored with the WIA driver item.
        //
        BSTR bstrItemName = NULL;
        IWiaDrvItem *pIWiaDrvItem   = NULL;
        hr = wiasGetDrvItem(pWiasContext, &pIWiaDrvItem);
        if(SUCCEEDED(hr))
        {
          hr = pIWiaDrvItem->GetItemName(&bstrItemName);
        }
        if(SUCCEEDED(hr))
        {
          //
          // Initialize the WIA item property set according to the category specified
          //

          if(wcscmp(bstrItemName,WIA_DRIVER_FLATBED_NAME)==0)
          {
            hr = InitializeWIAItemProperties(pWiasContext,WIA_CATEGORY_FLATBED);
            if(FAILED(hr))
            {
              WIAS_ERROR((g_hInst, "Failed to initialize the flatbed item's property set. hr = 0x%lx",hr));
            }
          }
          else if(wcscmp(bstrItemName , WIA_DRIVER_FEEDER_NAME)==0)
          {
            hr = InitializeWIAItemProperties(pWiasContext,WIA_CATEGORY_FEEDER);
            if(FAILED(hr))
            {
              WIAS_ERROR((g_hInst, "Failed to initialize the feeder item's property set. hr = 0x%lx",hr));
            }
          }
          else
          {
            hr = S_OK;
          }
        }
        if(bstrItemName)
        {
          SysFreeString(bstrItemName);
        }
      }

    }
    else
    {
      WIAS_ERROR((g_hInst, "Failed to read WIA_IPA_ITEM_FLAGS property, hr = 0x%lx",hr));
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }

  if ((FAILED(hr)) && (plDevErrVal))
  {
    *plDevErrVal = (E_INVALIDARG == hr) ? WIA_ERROR_INVALID_COMMAND : WIA_ERROR_GENERAL_ERROR;
  }

  return hr;
}

HRESULT CWIADriver::drvValidateItemProperties(__inout   BYTE       *pWiasContext,
                            LONG       lFlags,
                            ULONG      nPropSpec,
                        __in    const PROPSPEC *pPropSpec,
                        __out   LONG       *plDevErrVal)
{
  UNREFERENCED_PARAMETER(lFlags);

  BYTE       *pWiasContextOrg = pWiasContext;
  HRESULT hr = E_INVALIDARG;
  if((pWiasContext)&&(pPropSpec)&&(plDevErrVal)&&(nPropSpec))
  {
    LONG lItemType = 0;
    hr = wiasGetItemType(pWiasContext,&lItemType);

    if(SUCCEEDED(hr))
    {
      if(lItemType & WiaItemTypeRoot)
      {  
        BSTR bstrRootItemName;
        if(SUCCEEDED(hr))
        {
          hr = wiasReadPropStr(pWiasContext, WIA_IPA_FULL_ITEM_NAME,  &bstrRootItemName,0,TRUE);
        }
        if(SUCCEEDED(hr))
        {
          if(m_bIsWindowsVista)
          {
             return E_INVALIDARG;
          }
          else
          {
            BSTR strChild= WIA_DRIVER_FEEDER_NAME;
            if(m_bHasFlatbed)
            {
              strChild=WIA_DRIVER_FLATBED_NAME;
            }
            size_t nStrLen = wcslen(bstrRootItemName)+ wcslen(L"\\")+wcslen(strChild)+1;
            BSTR bstrChildItemName = new WCHAR[nStrLen];
            wcscpy_s(bstrChildItemName,nStrLen,bstrRootItemName);
            wcscat_s(bstrChildItemName,nStrLen,L"\\");
            wcscat_s(bstrChildItemName,nStrLen,strChild);
            BYTE *pWiasContextChild=NULL;

            hr = wiasGetContextFromName(pWiasContext,0,bstrChildItemName,&pWiasContextChild);
            delete[] bstrChildItemName;
            if(SUCCEEDED(hr))
            {
              pWiasContext = pWiasContextChild;
            }
          }
        }
      }
      else
      {
        //
        // Validate child item property settings, if needed.
        //
        GUID          guidFormat        = GUID_NULL;  
        WIA_PROPERTY_CONTEXT  PropertyContext     = {0};
       //
        // Use the WIA item category to help classify and gather WIA item information
        // needed to validate the property set.
        //


        if(SUCCEEDED(hr))
        {
          //
          // Validate the selection area against the entire scanning surface of the device.
          // The scanning surface may be different sizes depending on the type of WIA item.
          // (ie. Flatbed glass platen sizes may be different to film scanning surfaces, and
          // feeder sizes.)
          //
          HRESULT FreePropContextHR;
          hr = FreePropContextHR= wiasCreatePropContext(nPropSpec,(PROPSPEC*)pPropSpec,0,NULL,&PropertyContext);

          if(SUCCEEDED(hr))
          {
            hr = UpdateIntent(pWiasContext, &PropertyContext);
          }

          if(SUCCEEDED(hr))
          {
            hr = wiasUpdateValidFormat(pWiasContext,&PropertyContext,(IWiaMiniDrv*)this);
          }

          if(SUCCEEDED(FreePropContextHR))
          {
            FreePropContextHR = wiasFreePropContext(&PropertyContext);
            if(FAILED(FreePropContextHR))
            {
              WIAS_ERROR((g_hInst, "wiasFreePropContext failed, hr = 0x%lx",FreePropContextHR));
            }
          }

          if(FAILED(FreePropContextHR))
          {
            WIAS_ERROR((g_hInst, "wiasFreePropContext failed, hr = 0x%lx",FreePropContextHR));
          }
        }        
      }


      //
      // Only call wiasValidateItemProperties if the validation above
      // succeeded.
      //

      if(SUCCEEDED(hr))
      {
        // get list of changed properties
        LL_ARRAY llLst;
        for(ULONG i =0; i<nPropSpec;i++)
        {
          ULONG ulID = (ULONG)pPropSpec[i].propid;
          if(pPropSpec[i].ulKind == PRSPEC_LPWSTR)
          {
            int j=0;
            while(g_wiaPropIdToName[j].propid)
            {
              if(wcscmp(g_wiaPropIdToName[j].pszName,pPropSpec[i].lpwstr)==0)
              {
                ulID = g_wiaPropIdToName[j].propid;
                break;
              }
              j++;
            }
          }
          llLst.push_back((ULONG)pPropSpec[i].propid);
        }
        LL_ARRAY llLstTWAINcaps;
        LL_ARRAY llLstOrdered;
        LL_ARRAY::iterator llIter=llLst.begin();
        for(;llIter!=llLst.end();llIter++)
        {
          ULONG lWIAProp = (ULONG)(*llIter);
          for(ULONG j=0; g_unWIAPropToTWAINcap[j];j++)
          {
            if(g_unWIAPropToTWAINcap[j]==lWIAProp)
            {
              j++;//move to first TWAIN cap
              while(g_unWIAPropToTWAINcap[j])
              {
                llLstTWAINcaps.push_back((short)g_unWIAPropToTWAINcap[j]);
                j++;
              }
            }
            else
            {
              while(g_unWIAPropToTWAINcap[j])
              {
                j++;
              }
            }
          }
        }

      //sort and remove duplicated
        llLstTWAINcaps.sort();
        llLstTWAINcaps.unique();
      //Sort by capabilities order - see White Paper:Capability Ordering on http://www.twain.org
        for(int i=0; g_unSetCapOrder[i]; i++)
        {
          if(llLstTWAINcaps.IfExist(g_unSetCapOrder[i]))
          {
            llLstOrdered.push_back(g_unSetCapOrder[i]);
          }  
        }

        if(llLstOrdered.size())// is there is something to verify?
        {
          //Get CTWAIN_API associated with this WIA context
          CTWAIN_API *pTwainApi=NULL;
          hr = GetDS(pWiasContext,&pTwainApi);
          if(SUCCEEDED(hr))
          {
            DWORD dwRes = pTwainApi->OpenDS();
            hr = TWAINtoWIAerror(dwRes);
          }          
          if(SUCCEEDED(hr))
          {
            hr = ValidateThroughTWAINDS(pWiasContext,llLstOrdered);
          }
          // Update WIA properties 
          if (SUCCEEDED(hr))
          {
            hr = GetAllTWAIN_Caps(pWiasContext);
          }
        }
      }
      if(SUCCEEDED(hr))
      {
        // update dependent properties
        hr = UpdateWIAPropDepend(pWiasContext);
      }

      if(SUCCEEDED(hr))
      {
        hr = wiasValidateItemProperties(pWiasContextOrg,nPropSpec,pPropSpec);
        if(FAILED(hr))
        {
          WIAS_ERROR((g_hInst, "Failed to validate remaining properties using wiasValidateItemProperties, hr = 0x%lx",hr));
        }
      }       
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::drvWriteItemProperties(__inout  BYTE            *pWiasContext,
                          LONG            lFlags,
                       __in   PMINIDRV_TRANSFER_CONTEXT pmdtc,
                       __out  LONG            *plDevErrVal)
{
  UNREFERENCED_PARAMETER(lFlags);

  HRESULT hr = E_INVALIDARG;
  if ((pWiasContext) && (pmdtc) && (plDevErrVal))
  {
    GUID  guidItemCategory  = GUID_NULL;
    hr = wiasReadPropGuid(pWiasContext,WIA_IPA_ITEM_CATEGORY,&guidItemCategory,NULL,TRUE);
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::drvReadItemProperties(__in    BYTE       *pWiasContext,
                          LONG       lFlags,
                          ULONG      nPropSpec,
                      __in    const PROPSPEC *pPropSpec,
                      __out   LONG       *plDevErrVal)
{
  UNREFERENCED_PARAMETER(lFlags);
  UNREFERENCED_PARAMETER(nPropSpec);

  HRESULT hr = E_INVALIDARG;
  if((pWiasContext)&&(pPropSpec)&&(plDevErrVal))
  {
    hr= S_OK;
    if(pPropSpec[0].propid==WIA_DPS_DOCUMENT_HANDLING_STATUS) // update scanner ready bit
    {
      long lVal = 0;
      if(m_bHasFlatbed)
      {
        lVal=FLAT_READY;
      }
      if(m_bHasFeeder)
      {
        CTWAIN_API *pTwainApi=NULL;
        hr = GetDS(pWiasContext,&pTwainApi);
        if(hr!=S_OK)
        {
          return hr;
        } 
        DWORD dwRes = pTwainApi->OpenDS();
        hr = TWAINtoWIAerror(dwRes);
        if(hr!=S_OK)
        {
          return hr;
        }    

        LL llVal;
        WORD wType;
        bool bPaperPresent=true;
        dwRes = pTwainApi->SetCapability(CAP_FEEDERENABLED ,true,TWTY_BOOL);
        if(dwRes==FAILURE(TWCC_BADVALUE))
        {
          return E_INVALIDARG;
        }
        if(dwRes && dwRes!=FAILURE(TWCC_CAPUNSUPPORTED)&& dwRes!=FAILURE(TWCC_CAPBADOPERATION))
        {
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        }
        if(!dwRes)//if supported
        {
          WORD wCapType;
          dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERENABLED,&llVal,&wCapType);
          if(dwRes!=0 || wCapType!=TWTY_BOOL || (bool)llVal!=true)
          {
            return WIA_ERROR_EXCEPTION_IN_DRIVER;
          }
        }
        dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERLOADED,&llVal,&wType);
        if(dwRes==0&&wType==TWTY_BOOL&& (bool)llVal==false)
        {
          bPaperPresent=false;
        }

        if(bPaperPresent)
        {
          lVal=lVal|FEED_READY;
          if(m_bDuplex)
          {
            lVal= lVal|DUP_READY;
          }
        }      
      }


      hr = wiasWritePropLong( pWiasContext, WIA_DPS_DOCUMENT_HANDLING_STATUS, lVal );
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::drvLockWiaDevice(__in     BYTE *pWiasContext,
                        LONG lFlags,
                   __out    LONG *plDevErrVal)
{
  UNREFERENCED_PARAMETER(lFlags);
  HRESULT hr = E_INVALIDARG;
  if((pWiasContext)&&(plDevErrVal))
  {
    if(m_pIStiDevice)
    {
      hr = m_pIStiDevice->LockDevice(DEFAULT_LOCK_TIMEOUT);
    }
    else
    {
      hr = S_OK;
    }
    if(hr==S_OK)
    {
      m_bUnint = false;
      m_nLockCounter++;
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::drvUnLockWiaDevice(__in   BYTE *pWiasContext,
                        LONG lFlags,
                     __out  LONG *plDevErrVal)
{
  UNREFERENCED_PARAMETER(lFlags);

  HRESULT hr = E_INVALIDARG;
  if((pWiasContext)&&(plDevErrVal))
  {
    m_nLockCounter--;
    if(!m_bUnint && m_nLockCounter<=0)
    {
      // close DS if it is not after drvUnInitializeWia
      CTWAIN_API *pTwainApi;
      hr = GetDS(pWiasContext,&pTwainApi);
      if(SUCCEEDED(hr))
      {
        pTwainApi->CloseDS();
      }
    }
    if(m_pIStiDevice)
    {
      hr = m_pIStiDevice->UnLockDevice();
    }
    else
    {
      hr = S_OK;
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::drvAnalyzeItem(__in   BYTE *pWiasContext,
                      LONG lFlags,
                   __out  LONG *plDevErrVal)
{
  UNREFERENCED_PARAMETER(pWiasContext);
  UNREFERENCED_PARAMETER(lFlags);
  UNREFERENCED_PARAMETER(plDevErrVal);

  WIAS_ERROR((g_hInst, "This method is not implemented or supported for this driver"));
  return E_NOTIMPL;
}
HRESULT CWIADriver::drvGetDeviceErrorStr(     LONG   lFlags,
                          LONG   lDevErrVal,
                     __out  LPOLESTR *ppszDevErrStr,
                     __out  LONG   *plDevErr)
{
  UNREFERENCED_PARAMETER(lFlags);
  UNREFERENCED_PARAMETER(lDevErrVal);
  UNREFERENCED_PARAMETER(ppszDevErrStr);
  UNREFERENCED_PARAMETER(plDevErr);

  WIAS_ERROR((g_hInst, "This method is not implemented or supported for this driver"));
  return E_NOTIMPL;
}
HRESULT CWIADriver::DestroyDriverItemTree()
{
  HRESULT hr = S_OK;

  if(m_pIDrvItemRoot)
  {
    WIAS_TRACE((g_hInst,"Unlinking WIA item tree"));
    hr = m_pIDrvItemRoot->UnlinkItemTree(WiaItemTypeDisconnected);
    if(FAILED(hr))
    {
      WIAS_ERROR((g_hInst, "Failed to unlink WIA item tree before being released, hr = 0x%lx",hr));
    }

    WIAS_TRACE((g_hInst,"Releasing IDrvItemRoot interface"));
    m_pIDrvItemRoot->Release();
    m_pIDrvItemRoot = NULL;
  }

  return hr;
}

HRESULT CWIADriver::BuildDriverItemTree()
{
  HRESULT hr = S_OK;
  if(!m_pIDrvItemRoot)
  {
    LONG lItemFlags = WiaItemTypeFolder | WiaItemTypeDevice | WiaItemTypeRoot;
    BSTR bstrRootItemName = SysAllocString(WIA_DRIVER_ROOT_NAME);
    if(bstrRootItemName)
    {
      //
      // Create a default WIA root item
      //
#pragma warning (disable : 6309)//Param 6 can be NULL
#pragma warning (disable : 6387)//Param 6 can be NULL
      hr = wiasCreateDrvItem(lItemFlags,
                   bstrRootItemName,
                   m_bstrRootFullItemName,
                   (IWiaMiniDrv*)this,
                   0,
                   NULL,
                   &m_pIDrvItemRoot);
#pragma warning (default : 6309)
#pragma warning (default : 6387)
      //
      // Create child items that represent the data or programmable data sources.
      //

      if(m_bIsWindowsVista)
      {
        lItemFlags  = WiaItemTypeImage | WiaItemTypeTransfer | WiaItemTypeFile | WiaItemTypeProgrammableDataSource;
      }
      else
      {
        lItemFlags  = WiaItemTypeImage | WiaItemTypeFile;
      }
      
      if(SUCCEEDED(hr) && m_bHasFlatbed)
      {
        hr = CreateWIAChildItem(WIA_DRIVER_FLATBED_NAME,(IWiaMiniDrv*)this,m_pIDrvItemRoot,lItemFlags, WIA_CATEGORY_FLATBED,NULL);

        if(FAILED(hr))
        {
        WIAS_ERROR((g_hInst, "Failed to create WIA flatbed item, hr = 0x%lx",hr));
        }
      }

      if(SUCCEEDED(hr) && ((m_bIsWindowsVista && m_bHasFeeder)||(!m_bIsWindowsVista && m_bHasFeeder && !m_bHasFlatbed)))
      {
        hr = CreateWIAChildItem(WIA_DRIVER_FEEDER_NAME,(IWiaMiniDrv*)this,m_pIDrvItemRoot,lItemFlags, WIA_CATEGORY_FEEDER,NULL);
        if(FAILED(hr))
        {
          WIAS_ERROR((g_hInst, "Failed to create WIA feeder item, hr = 0x%lx",hr));
        }
      }

      SysFreeString(bstrRootItemName);
      bstrRootItemName = NULL;
    }
    else
    {
      hr = E_OUTOFMEMORY;
      WIAS_ERROR((g_hInst, "Failed to allocate memory for the root item name, hr = 0x%lx",hr));
    }
  }

  return hr;
}


HRESULT CWIADriver::DoSynchronizeCommand(
  __inout BYTE *pWiasContext)
{
  HRESULT hr = S_OK;

  hr = DestroyDriverItemTree();
  if (SUCCEEDED(hr))
  {
    hr = BuildDriverItemTree();

    //
    //  Queue tree updated event, regardless ofwhether it
    //  succeeded, since we can't guarantee that the tree
    //  was left in the same condition.
    //
    QueueWIAEvent(pWiasContext, WIA_EVENT_TREE_UPDATED);
  }
  else
  {
    WIAS_ERROR((g_hInst, " failed, hr = 0x%lx", hr));
  }

  return hr;
}

HRESULT CWIADriver::drvDeviceCommand(__inout  BYTE    *pWiasContext,
                        LONG    lFlags,
                   __in   const GUID  *pguidCommand,
                   __out  IWiaDrvItem **ppWiaDrvItem,
                   __out  LONG    *plDevErrVal)
{
  UNREFERENCED_PARAMETER(lFlags);
  UNREFERENCED_PARAMETER(ppWiaDrvItem);
  UNREFERENCED_PARAMETER(plDevErrVal);

  HRESULT hr = E_NOTIMPL;

  if (pguidCommand)
  {
    if (*pguidCommand == WIA_CMD_SYNCHRONIZE)
    {
      hr = DoSynchronizeCommand(pWiasContext);
    }
  }
  else
  {
    hr = E_NOTIMPL;
    WIAS_ERROR((g_hInst, "This method is not implemented or supported for this driver"));
  }

  return hr;
}

HRESULT CWIADriver::drvGetCapabilities(__in   BYTE      *pWiasContext,
                        LONG      ulFlags,
                     __out  LONG      *pcelt,
                     __out  WIA_DEV_CAP_DRV **ppCapabilities,
                     __out  LONG      *plDevErrVal)
{
  UNREFERENCED_PARAMETER(pWiasContext);

  HRESULT hr = E_INVALIDARG;
  if((pcelt)&&(ppCapabilities)&&(plDevErrVal))
  {
    hr = S_OK;
    if(m_CapabilityManager.GetNumCapabilities() == 0)
    {
      hr = m_CapabilityManager.AddCapability(WIA_EVENT_DEVICE_CONNECTED,
                           IDS_EVENT_DEVICE_CONNECTED_NAME,
                           IDS_EVENT_DEVICE_CONNECTED_DESCRIPTION,
                           WIA_NOTIFICATION_EVENT,
                           WIA_ICON_DEVICE_CONNECTED);
      if(SUCCEEDED(hr))
      {
        hr = m_CapabilityManager.AddCapability(WIA_EVENT_DEVICE_DISCONNECTED,
                             IDS_EVENT_DEVICE_DISCONNECTED_NAME,
                             IDS_EVENT_DEVICE_DISCONNECTED_DESCRIPTION,
                             WIA_NOTIFICATION_EVENT,
                             WIA_ICON_DEVICE_DISCONNECTED);
        if(SUCCEEDED(hr))
        {
          hr = m_CapabilityManager.AddCapability(WIA_CMD_SYNCHRONIZE,
                               IDS_CMD_SYNCHRONIZE_NAME,
                               IDS_CMD_SYNCHRONIZE_DESCRIPTION,
                               0,
                               WIA_ICON_SYNCHRONIZE);
          if(FAILED(hr))
          {
            WIAS_ERROR((g_hInst, "Failed to add WIA_CMD_SYNCHRONIZE to capability manager, hr = 0x%lx",hr));
          }
        }
        else
        {
          WIAS_ERROR((g_hInst, "Failed to add WIA_EVENT_DEVICE_DISCONNECTED to capability manager, hr = 0x%lx",hr));
        }
      }
      else
      {
        WIAS_ERROR((g_hInst, "Failed to add WIA_EVENT_TREE_UPDATED to capability manager, hr = 0x%lx",hr));
      }
    }

    if(SUCCEEDED(hr))
    {
      if(((ulFlags & WIA_DEVICE_COMMANDS) == WIA_DEVICE_COMMANDS)&&(ulFlags & WIA_DEVICE_EVENTS) == WIA_DEVICE_EVENTS)
      {
        *ppCapabilities = m_CapabilityManager.GetCapabilities();
        *pcelt      = m_CapabilityManager.GetNumCapabilities();
        WIAS_TRACE((g_hInst,"Application is asking for Commands and Events, and we have %d total capabilities",*pcelt));
      }
      else if((ulFlags & WIA_DEVICE_COMMANDS) == WIA_DEVICE_COMMANDS)
      {
        *ppCapabilities = m_CapabilityManager.GetCommands();
        *pcelt      = m_CapabilityManager.GetNumCommands();
        WIAS_TRACE((g_hInst,"Application is asking for Commands, and we have %d",*pcelt));
      }
      else if((ulFlags & WIA_DEVICE_EVENTS) == WIA_DEVICE_EVENTS)
      {
        *ppCapabilities = m_CapabilityManager.GetEvents();
        *pcelt      = m_CapabilityManager.GetNumEvents();
        WIAS_TRACE((g_hInst,"Application is asking for Events, and we have %d",*pcelt));
      }

      WIAS_TRACE((g_hInst,"========================================================"));
      WIAS_TRACE((g_hInst,"WIA driver capability information"));
      WIAS_TRACE((g_hInst,"========================================================"));

      WIA_DEV_CAP_DRV *pCapabilities    = m_CapabilityManager.GetCapabilities();
      LONG      lNumCapabilities  = m_CapabilityManager.GetNumCapabilities();

      for(LONG i = 0; i < lNumCapabilities; i++)
      {
        if(pCapabilities[i].ulFlags & WIA_NOTIFICATION_EVENT)
        {
          WIAS_TRACE((g_hInst,"Event Name:    %ws",pCapabilities[i].wszName));
          WIAS_TRACE((g_hInst,"Event Description: %ws",pCapabilities[i].wszDescription));
        }
        else
        {
          WIAS_TRACE((g_hInst,"Command Name:    %ws",pCapabilities[i].wszName));
          WIAS_TRACE((g_hInst,"Command Description: %ws",pCapabilities[i].wszDescription));
        }
      }
      WIAS_TRACE((g_hInst,"========================================================"));
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::drvDeleteItem(__inout BYTE *pWiasContext,
                      LONG lFlags,
                  __out   LONG *plDevErrVal)
{
  HRESULT hr = E_NOTIMPL;
  return hr;
}
HRESULT CWIADriver::drvFreeDrvItemContext(    LONG lFlags,
                      __in  BYTE *pSpecContext,
                      __out   LONG *plDevErrVal)
{
  HRESULT hr = E_NOTIMPL;
  return hr;

}
HRESULT CWIADriver::drvGetWiaFormatInfo(__in  BYTE        *pWiasContext,
                        LONG        lFlags,
                    __out   LONG        *pcelt,
                    __out   WIA_FORMAT_INFO   **ppwfi,
                    __out   LONG        *plDevErrVal)
{
  UNREFERENCED_PARAMETER(lFlags);

  HRESULT hr = E_INVALIDARG;
  if((plDevErrVal)&&(pcelt)&&(ppwfi))
  {
    if(m_pFormats)
    {
      delete [] m_pFormats;
      m_pFormats = NULL;
    }

    if(!m_pFormats)
    {
      m_ulNumFormats = DEFAULT_NUM_DRIVER_FORMATS;

      //
      // add the default formats to the corresponding arrays
      //
      if(pWiasContext)
      {
        //
        // Create a format list that is specific to the WIA item.
        //

        LONG lItemType = 0;

        hr = wiasGetItemType(pWiasContext,&lItemType);
        if(SUCCEEDED(hr))
        {
          if((lItemType & WiaItemTypeImage))
          {
            if(!m_bIsWindowsVista)
            {
                m_ulNumFormats = DEFAULT_NUM_DRIVER_FORMATS_WIA1;
            }
          }
        }
        else
        {
          WIAS_ERROR((g_hInst, "Failed to get WIA item type, hr = 0x%lx",hr));
        }
      }
      else
      {
        //
        // Create a default format list
        //

        // FIX: for this device, we are assuming that the majority of data
        //    transferred will be image data, so when a query for formats fails,
        //    it is safe to default to bmp and raw as the formats.

        hr = S_OK;
      }

      m_pFormats    = new WIA_FORMAT_INFO[m_ulNumFormats];
      if(m_pFormats)
      {
        //
        // add file (TYMED_FILE) formats to format array
        //

          m_pFormats[0].guidFormatID = WiaImgFmt_BMP;
          m_pFormats[0].lTymed     = TYMED_FILE;
          if(m_ulNumFormats==2)
          {
            m_pFormats[0].guidFormatID = WiaImgFmt_MEMORYBMP;
            m_pFormats[0].lTymed     = TYMED_CALLBACK;
            m_pFormats[1].guidFormatID = WiaImgFmt_BMP;
            m_pFormats[1].lTymed     = TYMED_FILE;
          }
          else
          {
            m_pFormats[0].guidFormatID = WiaImgFmt_BMP;
            m_pFormats[0].lTymed     = TYMED_FILE;
          }

        hr = S_OK;
      }
      else
      {
        hr        = E_OUTOFMEMORY;
        m_ulNumFormats  = NULL;
        *ppwfi      = NULL;
        WIAS_ERROR((g_hInst, "Failed to allocate memory for WIA_FORMAT_INFO structure array, hr = 0x%lx",hr));
      }
    }

    if(m_pFormats)
    {
      *pcelt  = m_ulNumFormats;
      *ppwfi  = &m_pFormats[0];
      hr    = S_OK;
    }
    else
    {
      *pcelt = 0;
      *ppwfi = NULL;
    }

    *plDevErrVal = 0;
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::drvNotifyPnpEvent(__in    const GUID *pEventGUID,
                    __in    BSTR     bstrDeviceID,
                        ULONG    ulReserved)
{
  UNREFERENCED_PARAMETER(bstrDeviceID);
  UNREFERENCED_PARAMETER(ulReserved);

  HRESULT hr = E_INVALIDARG;
  if(pEventGUID)
  {
    // TBD: Add any special event handling here.
    //    Power management, canceling pending I/O etc.
    hr = S_OK;
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::drvUnInitializeWia(__inout BYTE *pWiasContext)
{

  HRESULT hr = E_INVALIDARG;
  if(pWiasContext)
  {
    if(InterlockedDecrement(&m_lClientsConnected) < 0)
    {
      WIAS_TRACE((g_hInst, "The client connection counter decremented below zero. Assuming no clients are currently connected and automatically setting to 0"));
      m_lClientsConnected = 0;
    }

    WIAS_TRACE((g_hInst,"%d client(s) are currently connected to this driver.",m_lClientsConnected));
    //close DS, close DSM and delete CTWAIN_API associated with this WIA context
    hr = DeleteDS(pWiasContext);
    
    if(m_lClientsConnected == 0)
    {
      //
      // When the last client disconnects, destroy the WIA item tree.
      // This should reduce the idle memory foot print of this driver
      //

      DestroyDriverItemTree();
    }
    m_bUnint=true;
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

/////////////////////////////////////////////////////////////////////////
// INonDelegating Interface Section (for all WIA drivers)        //
/////////////////////////////////////////////////////////////////////////

HRESULT CWIADriver::NonDelegatingQueryInterface(REFIID  riid,LPVOID  *ppvObj)
{
  if(!ppvObj)
  {
    WIAS_ERROR((g_hInst, "Invalid parameters were passed"));
    return E_INVALIDARG;
  }

  *ppvObj = NULL;

  if(IsEqualIID( riid, IID_IUnknown ))
  {
    *ppvObj = static_cast<INonDelegatingUnknown*>(this);
  }
  else if(IsEqualIID( riid, IID_IStiUSD ))
  {
    *ppvObj = static_cast<IStiUSD*>(this);
  }
  else if(IsEqualIID( riid, IID_IWiaMiniDrv ))
  {
    *ppvObj = static_cast<IWiaMiniDrv*>(this);
  }
  else
  {
    return E_NOINTERFACE;
  }

  reinterpret_cast<IUnknown*>(*ppvObj)->AddRef();
  return S_OK;
}

ULONG CWIADriver::NonDelegatingAddRef()
{
  return InterlockedIncrement(&m_cRef);
}

ULONG CWIADriver::NonDelegatingRelease()
{
  ULONG ulRef = InterlockedDecrement(&m_cRef);
  if(ulRef == 0)
  {
    delete this;
    return 0;
  }
  return ulRef;
}

void CWIADriver::QueueWIAEvent(
  __in  BYTE    *pWiasContext,
      const GUID  &guidWIAEvent)
{
  HRESULT hr          = S_OK;
  BSTR  bstrDeviceID    = NULL;
  BSTR  bstrFullItemName  = NULL;
  BYTE  *pRootItemContext   = NULL;

  hr = wiasReadPropStr(pWiasContext, WIA_IPA_FULL_ITEM_NAME, &bstrFullItemName, NULL,TRUE);
  if(SUCCEEDED(hr))
  {
    hr = wiasGetRootItem(pWiasContext,&pRootItemContext);
    if(SUCCEEDED(hr))
    {
      hr = wiasReadPropStr(pRootItemContext, WIA_DIP_DEV_ID,&bstrDeviceID, NULL,TRUE);
      if(SUCCEEDED(hr))
      {
        hr = wiasQueueEvent(bstrDeviceID,&guidWIAEvent,bstrFullItemName);
        if(FAILED(hr))
        {
          WIAS_ERROR((g_hInst, "Failed to queue WIA event, hr = 0x%lx",hr));
        }
      }
      else
      {
        WIAS_ERROR((g_hInst, "Failed to read the WIA_DIP_DEV_ID property, hr = 0x%lx",hr));
      }
    }
    else
    {
      WIAS_ERROR((g_hInst, "Failed to get the Root item from child item, using wiasGetRootItem, hr = 0x%lx",hr));
    }
  }
  else
  {
    WIAS_ERROR((g_hInst, "Failed to read WIA_IPA_FULL_ITEM_NAME property, hr = %lx",hr));
  }

  if(bstrFullItemName)
  {
    SysFreeString(bstrFullItemName);
    bstrFullItemName = NULL;
  }

  if(bstrDeviceID)
  {
    SysFreeString(bstrDeviceID);
    bstrDeviceID = NULL;
  }
}

HRESULT CWIADriver::GetTransferCallback(
  __in    PMINIDRV_TRANSFER_CONTEXT     pmdtc,
  __callback  IWiaMiniDrvTransferCallback   **ppIWiaMiniDrvTransferCallback)
{
  HRESULT hr = E_INVALIDARG;
  if (pmdtc && ppIWiaMiniDrvTransferCallback)
  {
    if (pmdtc->pIWiaMiniDrvCallBack)
    {
      hr = pmdtc->pIWiaMiniDrvCallBack->QueryInterface(IID_IWiaMiniDrvTransferCallback,
                               (void**) ppIWiaMiniDrvTransferCallback);
    }
    else
    {
      hr = E_UNEXPECTED;
      WIAS_ERROR((g_hInst, "A NULL pIWiaMiniDrvCallBack was passed in the MINIDRV_TRANSFER_CONTEXT structure, hr = 0x%lx",hr));
    }
  }
  else
  {
    WIAS_ERROR((g_hInst, "Invalid parameters were passed"));
  }
  return hr;
}

HRESULT CWIADriver::CreateWIAChildItem(
  __in      LPOLESTR  wszItemName,
  __in      IWiaMiniDrv *pIWiaMiniDrv,
  __in      IWiaDrvItem *pParent,
          LONG    lItemFlags,
          GUID    guidItemCategory,
  __out_opt     IWiaDrvItem **ppChild,
  __in_opt    const WCHAR *wszStoragePath)
{
  HRESULT hr = E_INVALIDARG;
  if((wszItemName)&&(pIWiaMiniDrv)&&(pParent))
  {
    BSTR    bstrItemName    = SysAllocString(wszItemName);
    BSTR    bstrFullItemName  = NULL;
    IWiaDrvItem *pIWiaDrvItem     = NULL;

    if (bstrItemName)
    {
      hr = MakeFullItemName(pParent,bstrItemName,&bstrFullItemName);
      if(SUCCEEDED(hr))
      {
#pragma warning (disable : 6309)//Param 6 can be NULL
#pragma warning (disable : 6387)//Param 6 can be NULL
        hr = wiasCreateDrvItem(lItemFlags,
                  bstrItemName,
                  bstrFullItemName,
                  pIWiaMiniDrv,
                  0,
                  NULL,
                  &pIWiaDrvItem);
#pragma warning (default : 6309)
#pragma warning (default : 6387)

        if(SUCCEEDED(hr))
        {
          if(SUCCEEDED(hr))
          {
            hr = pIWiaDrvItem->AddItemToFolder(pParent);
            if(FAILED(hr))
            {
              WIAS_ERROR((g_hInst, "Failed to add the new WIA item (%ws) to the specified parent item, hr = 0x%lx",bstrFullItemName,hr));
              pIWiaDrvItem->Release();
              pIWiaDrvItem = NULL;
            }

            //
            // If a child iterface pointer parameter was specified, then the caller
            // expects to have the newly created child interface pointer returned to
            // them. (do not release the newly created item, in this case)
            //

            if(ppChild)
            {
              *ppChild    = pIWiaDrvItem;
              pIWiaDrvItem  = NULL;
            }
            else if (pIWiaDrvItem)
            {
              //
              // The newly created child has been added to the tree, and is no longer
              // needed.  Release it.
              //

              pIWiaDrvItem->Release();
              pIWiaDrvItem = NULL;
            }
          }
        }
        else
        {
          WIAS_ERROR((g_hInst, "Failed to create the new WIA driver item, hr = 0x%lx",hr));
        }

        SysFreeString(bstrItemName);
        bstrItemName = NULL;
        SysFreeString(bstrFullItemName);
        bstrFullItemName = NULL;
      }
      else
      {
        WIAS_ERROR((g_hInst, "Failed to create the new WIA item's full item name, hr = 0x%lx",hr));
      }
    }
    else
    {
      //
      // Failed to allocate memory for bstrItemName.
      //
      hr = E_OUTOFMEMORY;
      WIAS_ERROR((g_hInst, "Failed to allocate memory for BSTR storage item name"));
    }

  }
  else
  {
    WIAS_ERROR((g_hInst, "Invalid parameters were passed"));
  }
  return hr;
}

HRESULT CWIADriver::MakeFullItemName(
  __in  IWiaDrvItem *pParent,
  __in  BSTR    bstrItemName,
  __out   BSTR    *pbstrFullItemName)
{
  HRESULT hr = S_OK;
  if((pParent)&&(bstrItemName)&&(pbstrFullItemName))
  {
    BSTR bstrParentFullItemName = NULL;
    hr = pParent->GetFullItemName(&bstrParentFullItemName);
    if(SUCCEEDED(hr))
    {
      CBasicStringWide cswFullItemName;
      cswFullItemName.Format(TEXT("%ws\\%ws"),bstrParentFullItemName,bstrItemName);
      *pbstrFullItemName = SysAllocString(cswFullItemName.String());
      if(*pbstrFullItemName)
      {
        hr = S_OK;
      }
      else
      {
        hr = E_OUTOFMEMORY;
        WIAS_ERROR((g_hInst, "Failed to allocate memory for BSTR full item name, hr = 0x%lx",hr));
      }
      SysFreeString(bstrParentFullItemName);
      bstrParentFullItemName = NULL;
    }
    else
    {
      WIAS_ERROR((g_hInst, "Failed to get full item name from parent IWiaDrvItem, hr = 0x%lx",hr));
    }
  }
  else
  {
    hr = E_INVALIDARG;
    WIAS_ERROR((g_hInst, "Invalid parameters were passed, hr = 0x%lx",hr));
  }
  return hr;
}

HRESULT CWIADriver::GetConfigParams(CTWAIN_API *pTwainApi)
{
  LL_ARRAY lst;
  LL llVal;
  WORD wCapType;
  m_bHasFlatbed = true;
  m_bHasFeeder = false;
  m_bDuplex = false;

  DWORD dwRes = pTwainApi->GetCapCurrentValue(CAP_UICONTROLLABLE,&llVal,&wCapType);
  if(dwRes!=0 || wCapType != TWTY_BOOL || (TW_BOOL)llVal==0 )
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  dwRes = pTwainApi->SetCapability(CAP_INDICATORS, false,TWTY_BOOL);
  if(dwRes && dwRes!=FAILURE(TWCC_CAPUNSUPPORTED))
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  if(dwRes==0)//if supported it must be false
  {
    dwRes = pTwainApi->GetCapCurrentValue(CAP_INDICATORS,&llVal,&wCapType);
    if(dwRes!=0 || wCapType != TWTY_BOOL || (TW_BOOL)llVal )
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  dwRes = pTwainApi->GetBoolCapConstrainedValues(CAP_FEEDERENABLED,&lst);
  if(dwRes==0)
  {
    LL llFirst=lst.front();
    LL llSecond=lst.back();

    if(llFirst || llSecond)//there is feeder
    {
      m_bHasFeeder = true;
    }
    if(llFirst && llSecond)//there is no flatbed
    {
      m_bHasFlatbed = false;
    }
  }
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  dwRes = 0;
  if(m_bHasFeeder)
  {
    dwRes = pTwainApi->GetCapCurrentValue(CAP_DUPLEX,&llVal,&wCapType);
    if(dwRes==0 && wCapType == TWTY_UINT16 && (TW_UINT16)llVal!=TWDX_NONE )
    {
      m_bDuplex = true;
    }
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return 0;
}

HRESULT CWIADriver::GetDS(BYTE *pWiasContext, CTWAIN_API** pTwainApi)
{
  BYTE *pRootWiasContext=NULL;
  HRESULT hr = wiasGetRootItem(pWiasContext,&pRootWiasContext);

  LONG lAppID = 0;
  hr = wiasReadPropLong(pRootWiasContext,CUSTOM_ROOT_PROP_ID,&lAppID,NULL,TRUE);
  if(SUCCEEDED(hr))
  {
    *pTwainApi = g_TWAIN_APIs.find(lAppID)->second;
  }
  return hr;
}

HRESULT CWIADriver::DeleteDS(BYTE *pRootWiasContext)
{
  LONG lAppID = 0;
  HRESULT hr = wiasReadPropLong(pRootWiasContext,CUSTOM_ROOT_PROP_ID,&lAppID,NULL,TRUE);
  if(SUCCEEDED(hr))
  {
    //find CTWAIN_API associated with this WIA context
    map<int,CTWAIN_API*>::iterator lIter=g_TWAIN_APIs.find(lAppID);
    if(lIter!=g_TWAIN_APIs.end())//because of bug in WIA service - it can call drvUnInitializeWIA before drvInitializeWIA
    {
      lIter->second->CloseDS();
      lIter->second->CloseDSM();
      delete lIter->second;
      g_TWAIN_APIs.erase(lAppID);
    }
  }
  //delete temporary profile file 
  BSTR strFileName;
  hr = wiasReadPropStr(pRootWiasContext,CUSTOM_ROOT_PROP_ID1,&strFileName,NULL,TRUE);

  if(SUCCEEDED(hr))
  {
    hr =  DeleteFile(strFileName)?S_OK:E_FAIL;
  }
  return hr;
}

HRESULT CWIADriver::AddDS(BYTE *pRootWiasContext, int *pnAppID, CTWAIN_API** pTwainApi)
{
  //find first free ID
  for(*pnAppID=0; *pnAppID<INT_MAX;(*pnAppID)++)
  {
    bool bFound=false;
    map<int,CTWAIN_API*>::iterator lIter=g_TWAIN_APIs.begin();
    for(;lIter!=g_TWAIN_APIs.end();lIter++)
    {
      if(lIter->first==*pnAppID)
      {
        bFound = true;
        break;
      }
    }
    if(!bFound)
    {
      break;
    }
  }
  //Creates CTWAIN_API and open DS
  TW_IDENTITY AppID={0,{kVER_MAJ,kVER_MIN,TWLG_USA,TWCY_USA,kVER_INFO},kPROT_MAJ,kPROT_MIN,DG_CONTROL|DG_IMAGE|DF_APP2,kMANUFACTURER,kPRODUCT_FAMILY,kPRODUCT_NAME};
  sprintf_s(AppID.ProductName,"%s%d",kPRODUCT_NAME,*pnAppID);
  *pTwainApi = new CTWAIN_API();
  DWORD dwRes = (*pTwainApi)->OpenDSM(AppID,*pnAppID);
  if(dwRes!=0)
  {
    delete *pTwainApi;
    *pTwainApi = NULL;
    return WIA_ERROR_BUSY;
  }
  TW_IDENTITY DSID={0,{kDS_VER_MAJ,kDS_VER_MIN,0,0,""},0,0,0,kDS_MANUFACTURER,"",kDS_PRODUCT_NAME};
  dwRes = (*pTwainApi)->OpenDS(&DSID);
  if(dwRes!=0)
  {
    (*pTwainApi)->CloseDSM();
    delete *pTwainApi;
    *pTwainApi = NULL;
    return WIA_ERROR_OFFLINE;
  }
  //Store CTWAIN_API in global table
  g_TWAIN_APIs[*pnAppID]= *pTwainApi; 
  return S_OK;
}

HRESULT CWIADriver::GetFEEDERENABLED(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  bool bFeeder;
  bool bDuplex=false;
  HRESULT hr = IsFeederItem(pWiasContext,&bFeeder);
  if(SUCCEEDED(hr))
  {
    LL llVal;
    WORD wCapType;
    DWORD dwRes = pTwainApi->GetCapCurrentValue(CAP_DUPLEXENABLED,&llVal,&wCapType);
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
    if(dwRes==0)
    {
      if(wCapType!=TWTY_BOOL)
      {
        return WIA_ERROR_EXCEPTION_IN_DRIVER;
      }
      bDuplex = ((bool)llVal);
    }    
    dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERENABLED,&llVal,&wCapType);
    if(dwRes!=0 || wCapType!=TWTY_BOOL)
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }


    if(m_bIsWindowsVista)
    {
      if(bFeeder!=(bool)llVal)//out of sync
      {
        hr = SetFEEDERENABLED(pWiasContext, pTwainApi);
      }
      if(SUCCEEDED(hr))
      {
        LONG lFeeder = (bool)llVal?(bDuplex?(DUPLEX):(FRONT_ONLY)):(FRONT_ONLY);
        hr = wiasWritePropLong(pWiasContext,WIA_IPS_DOCUMENT_HANDLING_SELECT,lFeeder);
      }
    }
    else
    {
      BYTE *pWiasRootContext=NULL;
      hr = wiasGetRootItem(pWiasContext,&pWiasRootContext);
      if(SUCCEEDED(hr))
      {
        LONG lFeeder = (bool)llVal?(bDuplex?(DUPLEX|FEEDER):(FEEDER|FRONT_ONLY)):(FLATBED|FRONT_ONLY);
        hr = wiasWritePropLong(pWiasRootContext,WIA_DPS_DOCUMENT_HANDLING_SELECT,lFeeder);
      }
    }
  }
  return S_OK;
}

HRESULT CWIADriver::SetFEEDERENABLED(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  bool bFeeder;
  HRESULT hr = IsFeederItem(pWiasContext,&bFeeder);
  if(SUCCEEDED(hr))
  {
    DWORD dwRes = pTwainApi->SetCapability(CAP_FEEDERENABLED ,bFeeder,TWTY_BOOL);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes && dwRes!=FAILURE(TWCC_CAPUNSUPPORTED)&& dwRes!=FAILURE(TWCC_CAPBADOPERATION))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
    if(!dwRes)//if supported
    {
      LL llVal;
      WORD wCapType;
      dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERENABLED,&llVal,&wCapType);
      if(dwRes!=0 || wCapType!=TWTY_BOOL || (bool)llVal!=bFeeder)
      {
        return WIA_ERROR_EXCEPTION_IN_DRIVER;
      }
    }
  }
  return S_OK;
}

HRESULT CWIADriver::GetFEEDERALIGNMENT(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  bool bFeeder;
  HRESULT hr = IsFeederItem(pWiasContext,&bFeeder);
  if(hr!=S_OK)
  {
    return hr;
  }
  if(!bFeeder)//this works for feeder only
  {
    return S_OK;
  }

  LONG lSheetFeederRegistration = LEFT_JUSTIFIED;
  LL llVal;
  WORD wCapType;

  DWORD dwRes = pTwainApi->GetCapCurrentValue(CAP_FEEDERALIGNMENT,&llVal,&wCapType);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  if(dwRes==0 && wCapType==TWTY_UINT16)
  {
    if(llVal==TWFA_CENTER)
    {
      lSheetFeederRegistration = CENTERED;
    }
    else if(llVal==TWFA_RIGHT)
    {
      lSheetFeederRegistration = RIGHT_JUSTIFIED;
    }
  }

  if(m_bIsWindowsVista)
  {
    hr = wiasWritePropLong(pWiasContext, WIA_IPS_SHEET_FEEDER_REGISTRATION, lSheetFeederRegistration);
  }
  else
  {
    BYTE *pRootWiasContext=NULL;
    hr = wiasGetRootItem(pWiasContext,&pRootWiasContext);
    if(SUCCEEDED(hr))
    {
      hr = wiasWritePropLong(pRootWiasContext, WIA_DPS_SHEET_FEEDER_REGISTRATION, lSheetFeederRegistration);
    }
  } 
  return hr;
}

HRESULT CWIADriver::SetAUTOFEED(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  DWORD dwRes = pTwainApi->SetCapability(CAP_AUTOFEED,true,TWTY_BOOL);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  return S_OK;
}

HRESULT CWIADriver::GetDUPLEXENABLED(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  bool bFeeder;
  HRESULT hr = IsFeederItem(pWiasContext,&bFeeder);
  if(hr!=S_OK)
  {
    return hr;
  }

  long lDocumentHandlingSelect= FLATBED|FRONT_ONLY;
  if(bFeeder)
  {
    lDocumentHandlingSelect= FEEDER|FRONT_ONLY;
  }
  if(m_bIsWindowsVista)
  {
    lDocumentHandlingSelect= FRONT_ONLY;
  }
  if(bFeeder)
  {
    LL llVal;
    WORD wCapType;
    DWORD dwRes = pTwainApi->GetCapCurrentValue(CAP_DUPLEXENABLED,&llVal,&wCapType);
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
    if(dwRes==0 && wCapType==TWTY_BOOL && (bool)llVal)
    {
      lDocumentHandlingSelect = m_bIsWindowsVista?DUPLEX:(DUPLEX|FEEDER);
    }
  }

  if(m_bIsWindowsVista)
  {
    hr = wiasWritePropLong(pWiasContext, WIA_IPS_DOCUMENT_HANDLING_SELECT, lDocumentHandlingSelect);
  }
  else
  {
    BYTE *pRootWiasContext=NULL;
    hr = wiasGetRootItem(pWiasContext,&pRootWiasContext);
    if(SUCCEEDED(hr))
    {
      hr = wiasWritePropLong(pRootWiasContext, WIA_DPS_DOCUMENT_HANDLING_SELECT, lDocumentHandlingSelect);
    }
  }
  
  return hr;
}

HRESULT CWIADriver::SetDUPLEXENABLED(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  long lDocumentHandlingSelect= 0;
  HRESULT hr;
  bool bDuplex=false;
  bool bFeeder=false; 
  hr=IsFeederItem(pWiasContext,&bFeeder);

  if(SUCCEEDED(hr) && bFeeder)
  {
    if(m_bIsWindowsVista)
    {
      hr = wiasReadPropLong(pWiasContext, WIA_IPS_DOCUMENT_HANDLING_SELECT, &lDocumentHandlingSelect, NULL, TRUE);
    }
    else
    {
      BYTE *pRootWiasContext=NULL;
      hr = wiasGetRootItem(pWiasContext,&pRootWiasContext);
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootWiasContext, WIA_DPS_DOCUMENT_HANDLING_SELECT, &lDocumentHandlingSelect, NULL, TRUE);
      }
    }
    bDuplex = (lDocumentHandlingSelect & DUPLEX)!=0;
  }

  if(SUCCEEDED(hr))
  {
    DWORD dwRes = pTwainApi->SetCapability(CAP_DUPLEXENABLED,bDuplex,TWTY_BOOL);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes && dwRes!=FAILURE(TWCC_CAPUNSUPPORTED)&& dwRes!=FAILURE(TWCC_CAPBADOPERATION))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
    if(!dwRes)//if supported
    {
      LL llVal;
      WORD wCapType;
      dwRes = pTwainApi->GetCapCurrentValue(CAP_DUPLEXENABLED,&llVal,&wCapType);
      if(dwRes==FAILURE(TWCC_BUMMER))
      {
        return WIA_ERROR_EXCEPTION_IN_DRIVER;
      }
      if(dwRes==0 && (wCapType!=TWTY_BOOL || ((bool)llVal)!=bDuplex))
      {
        return WIA_ERROR_EXCEPTION_IN_DRIVER;
      }
    }
  }
  return hr;
}

HRESULT CWIADriver::GetPIXELTYPE(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  LL llVal;
  LL llDefVal;
  LL_ARRAY llLst;
  WORD wCapType;
  DWORD dwRes;

  dwRes = pTwainApi->GetCapCurrentValue(ICAP_PIXELTYPE  ,&llVal,&wCapType);

  if(dwRes!=0 || wCapType!=TWTY_UINT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  dwRes = pTwainApi->GetCapDefaultValue(ICAP_PIXELTYPE  ,&llDefVal,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_UINT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  dwRes = pTwainApi->GetCapConstrainedValues(ICAP_PIXELTYPE  ,&llLst,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_UINT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  CBasicDynamicArray<LONG> lDataTypeArray;
  LL_ARRAY::iterator llIter=llLst.begin();
  for(;llIter!=llLst.end();llIter++)
  {
    switch((WORD)*llIter)
    {
      case TWPT_BW:
        lDataTypeArray.Append(WIA_DATA_THRESHOLD);
        break;
      case TWPT_GRAY:
        lDataTypeArray.Append(WIA_DATA_GRAYSCALE);
        break;
      case TWPT_RGB:
        lDataTypeArray.Append(WIA_DATA_COLOR);
        break;
      default:
        break;
    }
  }
  switch((WORD)llVal)
  {
    case TWPT_BW:
      llVal = (WORD)WIA_DATA_THRESHOLD;
      break;
    case TWPT_GRAY:
      llVal = (WORD)WIA_DATA_GRAYSCALE;
      break;
    case TWPT_RGB:
      llVal = (WORD)WIA_DATA_COLOR;
      break;
    default:
      break;
  }
  switch((WORD)llDefVal)
  {
    case TWPT_BW:
      llDefVal = (WORD)WIA_DATA_THRESHOLD;
      break;
    case TWPT_GRAY:
      llDefVal = (WORD)WIA_DATA_GRAYSCALE;
      break;
    case TWPT_RGB:
      llDefVal = (WORD)WIA_DATA_COLOR;
      break;
    default:
      break;
  }
  HRESULT hr= wiasSetValidListLong(pWiasContext, WIA_IPA_DATATYPE,lDataTypeArray.Size(), (long)llDefVal,lDataTypeArray.GetBuffer(0));
  if(SUCCEEDED(hr))
  {
    hr= wiasWritePropLong(pWiasContext, WIA_IPA_DATATYPE,(long)llVal);
  }
  return hr;
}

HRESULT CWIADriver::SetPIXELTYPE(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  long lDataType= 0;
  HRESULT hr = wiasReadPropLong(pWiasContext, WIA_IPA_DATATYPE, &lDataType, NULL, TRUE);
  if(SUCCEEDED(hr))
  {
    switch(lDataType)
    {
      case WIA_DATA_THRESHOLD:
        lDataType = TWPT_BW;
        break;
      case WIA_DATA_GRAYSCALE:
        lDataType = TWPT_GRAY;
        break;
      case WIA_DATA_COLOR:
        lDataType = TWPT_RGB;
        break;
      default:
        break;
    }
    DWORD dwRes = pTwainApi->SetCapability(ICAP_PIXELTYPE,lDataType,TWTY_UINT16);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
    LL llVal;
    WORD wCapType;
    dwRes = pTwainApi->GetCapCurrentValue(ICAP_PIXELTYPE,&llVal,&wCapType);
    if(dwRes!=0 || wCapType!=TWTY_UINT16 || (long)llVal!=lDataType)
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return hr;
}

HRESULT CWIADriver::SetBITDEPTH(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  long lBitDepth= 0;
  HRESULT hr = wiasReadPropLong(pWiasContext, WIA_IPA_DEPTH, &lBitDepth, NULL, TRUE);
  if(SUCCEEDED(hr))
  {
    DWORD dwRes = pTwainApi->SetCapability(ICAP_BITDEPTH,lBitDepth,TWTY_UINT16);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
    LL llVal;
    WORD wCapType;
    dwRes = pTwainApi->GetCapCurrentValue(ICAP_BITDEPTH,&llVal,&wCapType);
    if(dwRes!=0 || wCapType!=TWTY_UINT16 || (short)llVal!=lBitDepth)
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return hr;
}

HRESULT CWIADriver::GetBITDEPTH(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  LL llVal;
  LL llPTVal;
  LL llDefVal;
  LL_ARRAY llLst;
  WORD wCapType;
  DWORD dwRes;

  dwRes = pTwainApi->GetCapCurrentValue(ICAP_BITDEPTH  ,&llVal,&wCapType);

  if(dwRes!=0 || wCapType!=TWTY_UINT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  dwRes = pTwainApi->GetCapCurrentValue(ICAP_PIXELTYPE  ,&llPTVal,&wCapType);

  if(dwRes!=0 || wCapType!=TWTY_UINT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  dwRes = pTwainApi->GetCapDefaultValue(ICAP_BITDEPTH  ,&llDefVal,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_UINT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  dwRes = pTwainApi->GetCapConstrainedValues(ICAP_BITDEPTH  ,&llLst,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_UINT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  LL_ARRAY::iterator llIter=llLst.begin();
  CBasicDynamicArray<LONG> lBitDepthArray;
  for(;llIter!=llLst.end();llIter++)
  {
    lBitDepthArray.Append((LONG)*llIter);
  }
  HRESULT hr= wiasSetValidListLong(pWiasContext, WIA_IPA_DEPTH,lBitDepthArray.Size(), (long)llDefVal,lBitDepthArray.GetBuffer(0));
  if(SUCCEEDED(hr))
  {
    hr= wiasWritePropLong(pWiasContext, WIA_IPA_DEPTH,(long)llVal);
  }

  return hr;
}

HRESULT CWIADriver::GetXRESOLUTION(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  LL llVal;
  LL llDefVal;
  LL_ARRAY llLst;
  WORD wCapType;
  DWORD dwRes;

  dwRes = pTwainApi->GetCapConstrainedValues(ICAP_XRESOLUTION ,&llLst,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_FIX32 || llLst.size()==0)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  dwRes = pTwainApi->GetCapCurrentValue(ICAP_XRESOLUTION ,&llVal,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_FIX32)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  dwRes = pTwainApi->GetCapDefaultValue(ICAP_XRESOLUTION ,&llDefVal,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_FIX32)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  CBasicDynamicArray<LONG> lXResolutionArray;
  LL_ARRAY::iterator llIter=llLst.begin();
  for(;llIter!=llLst.end();llIter++)
  {
    lXResolutionArray.Append((short)(*llIter));
  }
  HRESULT hr= wiasSetValidListLong(pWiasContext, WIA_IPS_XRES,lXResolutionArray.Size(), (long)llDefVal,lXResolutionArray.GetBuffer(0));
  if(SUCCEEDED(hr))
  {
    hr= wiasWritePropLong(pWiasContext, WIA_IPS_XRES,(long)llVal);
  }
  return hr;
}

HRESULT CWIADriver::SetXRESOLUTION(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  long lXRes= 0;
  HRESULT hr = wiasReadPropLong(pWiasContext, WIA_IPS_XRES, &lXRes, NULL, TRUE);
  if(SUCCEEDED(hr))
  {
    LL llVal = lXRes;
    DWORD dwRes = pTwainApi->SetCapability(ICAP_XRESOLUTION,llVal,TWTY_FIX32);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return hr;
}

HRESULT CWIADriver::GetYRESOLUTION(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  LL llVal;
  LL llDefVal;
  LL_ARRAY llLst;
  WORD wCapType;
  DWORD dwRes;

  dwRes = pTwainApi->GetCapConstrainedValues(ICAP_YRESOLUTION ,&llLst,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_FIX32 || llLst.size()==0)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  dwRes = pTwainApi->GetCapCurrentValue(ICAP_YRESOLUTION ,&llVal,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_FIX32)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  dwRes = pTwainApi->GetCapDefaultValue(ICAP_YRESOLUTION ,&llDefVal,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_FIX32)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  CBasicDynamicArray<LONG> lYResolutionArray;
  LL_ARRAY::iterator llIter=llLst.begin();
  for(;llIter!=llLst.end();llIter++)
  {
    lYResolutionArray.Append((short)(*llIter));
  }
  HRESULT hr= wiasSetValidListLong(pWiasContext, WIA_IPS_YRES,lYResolutionArray.Size(), (long)llDefVal,lYResolutionArray.GetBuffer(0));
  if(SUCCEEDED(hr))
  {
    hr= wiasWritePropLong(pWiasContext, WIA_IPS_YRES,(long)llVal);
  }
  return hr;
}

HRESULT CWIADriver::SetYRESOLUTION(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  long lYRes= 0;
  HRESULT hr = wiasReadPropLong(pWiasContext, WIA_IPS_YRES, &lYRes, NULL, TRUE);
  if(SUCCEEDED(hr))
  {
    LL llVal = lYRes;
    DWORD dwRes = pTwainApi->SetCapability(ICAP_YRESOLUTION,llVal,TWTY_FIX32);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return hr;
}

HRESULT CWIADriver::GetPIXELFLAVOR(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  LL llVal;
  LL llDefVal;
  LL_ARRAY llLst;
  WORD wCapType;
  DWORD dwRes;

  dwRes = pTwainApi->GetCapCurrentValue(ICAP_PIXELFLAVOR  ,&llVal,&wCapType);

  if(dwRes!=0 || wCapType!=TWTY_UINT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  if(llVal == TWPF_CHOCOLATE )
  {
    llVal = (WORD)WIA_PHOTO_WHITE_1;
  }
  else
  {
    llVal = (WORD)WIA_PHOTO_WHITE_0;
  }

  dwRes = pTwainApi->GetCapDefaultValue(ICAP_PIXELFLAVOR   ,&llDefVal,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_UINT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  if(llDefVal == TWPF_CHOCOLATE )
  {
    llDefVal = (WORD)WIA_PHOTO_WHITE_1;
  }
  else
  {
    llDefVal = (WORD)WIA_PHOTO_WHITE_0;
  }
  dwRes = pTwainApi->GetCapConstrainedValues(ICAP_PIXELFLAVOR  ,&llLst,&wCapType);

  if(dwRes!=0 || wCapType!=TWTY_UINT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  LL_ARRAY::iterator llIter=llLst.begin();
  CBasicDynamicArray<LONG> lPhotometricInterpArray;
  for(;llIter!=llLst.end();llIter++)
  {
    if(*llIter == TWPF_CHOCOLATE )
    {
      lPhotometricInterpArray.Append(WIA_PHOTO_WHITE_1);
    }
    else
    {
      lPhotometricInterpArray.Append(WIA_PHOTO_WHITE_0);
    }
  }

  HRESULT hr= wiasSetValidListLong(pWiasContext, WIA_IPS_PHOTOMETRIC_INTERP,lPhotometricInterpArray.Size(), (long)llDefVal,lPhotometricInterpArray.GetBuffer(0));
  if(SUCCEEDED(hr))
  {
    hr= wiasWritePropLong(pWiasContext, WIA_IPS_PHOTOMETRIC_INTERP,(long)llVal);
  }
  return hr;
}

HRESULT CWIADriver::SetPIXELFLAVOR(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  long lPhoto= 0;
  HRESULT hr = wiasReadPropLong(pWiasContext, WIA_IPS_PHOTOMETRIC_INTERP, &lPhoto, NULL, TRUE);
  if(SUCCEEDED(hr))
  {
    LL llVal = (short)TWPF_CHOCOLATE;
    if(lPhoto == WIA_PHOTO_WHITE_0)
    {
      llVal = (short)TWPF_VANILLA;
    }

    DWORD dwRes = pTwainApi->SetCapability(ICAP_PIXELFLAVOR,llVal,TWTY_UINT16);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
    LL llValR;
    WORD wCapType;
    dwRes = pTwainApi->GetCapCurrentValue(ICAP_PIXELFLAVOR,&llValR,&wCapType);
    if(dwRes!=0 || wCapType!=TWTY_UINT16 || llValR!=llVal)
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return hr;
}

HRESULT CWIADriver::GetTHRESHOLD(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  LL llVal;
  WORD wCapType;
  DWORD dwRes;

  dwRes = pTwainApi->GetCapCurrentValue(ICAP_THRESHOLD  ,&llVal,&wCapType);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  if(dwRes==0 && wCapType==TWTY_FIX32)
  {
    return wiasWritePropLong(pWiasContext, WIA_IPS_THRESHOLD,(long)llVal);
  }
  return S_OK;
}

HRESULT CWIADriver::SetTHRESHOLD(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  long lThreshold= 0;
  HRESULT hr = wiasReadPropLong(pWiasContext, WIA_IPS_THRESHOLD, &lThreshold, NULL, TRUE);
  if(SUCCEEDED(hr))
  {
    LL llVal = lThreshold;
    DWORD dwRes = pTwainApi->SetCapability(ICAP_THRESHOLD,llVal,TWTY_FIX32);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return hr;
}

HRESULT CWIADriver::GetBRIGHTNESS(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  LL llVal;
  WORD wCapType;
  DWORD dwRes;

  dwRes = pTwainApi->GetCapCurrentValue(ICAP_BRIGHTNESS  ,&llVal,&wCapType);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  if(dwRes==0 && wCapType==TWTY_FIX32)
  {
    return wiasWritePropLong(pWiasContext, WIA_IPS_BRIGHTNESS,(long)llVal);
  }
  return S_OK;
}

HRESULT CWIADriver::SetBRIGHTNESS(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  long lBrightness= 0;
  HRESULT hr = wiasReadPropLong(pWiasContext, WIA_IPS_BRIGHTNESS, &lBrightness, NULL, TRUE);
  if(SUCCEEDED(hr))
  {
    LL llVal = lBrightness;
    DWORD dwRes = pTwainApi->SetCapability(ICAP_BRIGHTNESS,llVal,TWTY_FIX32);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return hr;
}

HRESULT CWIADriver::GetCONTRAST(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  LL llVal;
  WORD wCapType;
  DWORD dwRes;

  dwRes = pTwainApi->GetCapCurrentValue(ICAP_CONTRAST  ,&llVal,&wCapType);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  if(dwRes==0 && wCapType==TWTY_FIX32)
  {
    return wiasWritePropLong(pWiasContext, WIA_IPS_CONTRAST,(long)llVal);
  }
  return S_OK;
}


HRESULT CWIADriver::SetCONTRAST(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  long lContrast= 0;
  HRESULT hr = wiasReadPropLong(pWiasContext, WIA_IPS_CONTRAST, &lContrast, NULL, TRUE);
  if(SUCCEEDED(hr))
  {
    LL llVal = lContrast;
    DWORD dwRes = pTwainApi->SetCapability(ICAP_CONTRAST,llVal,TWTY_FIX32);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return hr;
}

HRESULT CWIADriver::GetDAT_IMAGELAYOUT (BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  HRESULT hr=S_OK;
  TW_IMAGELAYOUT ImgLayout;
  DWORD dwRes;
  WORD wCapType;
  LL llXRes = 0;
  LL llYRes = 0;
  LL llHorizontalSize=0;
  LL llVerticalSize=0;
  LL llMinHorizontalSize = 0;
  LL llMinVerticalSize = 0;
  LONG lXPosition     = 0;
  LONG lYPosition     = 0;
  LONG lXExtent       = 1;
  LONG lYExtent       = 1;
  LONG lPixWidthMax      = 1;
  LONG lPixHeightMax     = 1;
  LONG lPixWidthMin      = 1;
  LONG lPixHeightMin     = 1;

  dwRes = pTwainApi->GetCapCurrentValue(ICAP_YRESOLUTION ,&llYRes,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_FIX32)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  dwRes = pTwainApi->GetCapCurrentValue(ICAP_XRESOLUTION ,&llXRes,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_FIX32)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  dwRes = pTwainApi->GetImageLayout(&ImgLayout);
  if(dwRes!=0)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  lXPosition = (long)(((LL)ImgLayout.Frame.Left)*llXRes);
  lYPosition = (long)(((LL)ImgLayout.Frame.Top)*llYRes);
  lXExtent = (long)((((LL)ImgLayout.Frame.Right)-((LL)ImgLayout.Frame.Left))*llXRes);
  lYExtent = (long)((((LL)ImgLayout.Frame.Bottom)-((LL)ImgLayout.Frame.Top))*llYRes);

  dwRes = pTwainApi->GetCapCurrentValue(ICAP_PHYSICALWIDTH ,&llHorizontalSize,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_FIX32)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  dwRes = pTwainApi->GetCapCurrentValue(ICAP_PHYSICALHEIGHT ,&llVerticalSize,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_FIX32)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }


  dwRes = pTwainApi->GetCapCurrentValue(ICAP_MINIMUMWIDTH ,&llMinHorizontalSize,&wCapType);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }          

  dwRes = pTwainApi->GetCapCurrentValue(ICAP_MINIMUMHEIGHT ,&llMinVerticalSize,&wCapType);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }   

  lPixWidthMax      = (long)(llHorizontalSize*llXRes);
  lPixHeightMax     = (long)(llVerticalSize*llYRes);
  lPixWidthMin      = (long)(llMinHorizontalSize*llXRes);
  lPixHeightMin     = (long)(llMinVerticalSize*llYRes);

  if(lXPosition<0)
  {
    lXPosition=0;
  }
  if(lXPosition>lPixWidthMax)
  {
    lXPosition=lPixWidthMax;
  }
  if(lXExtent<lPixWidthMin)
  {
    lXExtent=lPixWidthMin;
  }
  if(lXExtent>lPixWidthMax)
  {
    lXExtent=lPixWidthMax;
  }
  if(lYPosition<0)
  {
    lYPosition=0;
  }
  if(lYPosition>lPixHeightMax)
  {
    lYPosition=lPixHeightMax;
  }
  if(lYExtent<lPixHeightMin)
  {
    lYExtent=lPixHeightMin;
  }
  if(lYExtent>lPixHeightMax)
  {
    lYExtent=lPixHeightMax;
  }

  if(SUCCEEDED(hr))
  {
    LONG lNumberOfLines = (long)(((LL)(ImgLayout.Frame.Bottom)-(LL)(ImgLayout.Frame.Top))*llYRes);
    hr = wiasWritePropLong(pWiasContext,WIA_IPA_NUMBER_OF_LINES,lNumberOfLines);
  }
  LONG lPixelsPerLine = 0;
  if(SUCCEEDED(hr))
  {
    lPixelsPerLine = (long)(((LL)(ImgLayout.Frame.Right)-(LL)(ImgLayout.Frame.Left))*llXRes);
    hr = wiasWritePropLong(pWiasContext,WIA_IPA_PIXELS_PER_LINE ,lPixelsPerLine);
  }

  if(SUCCEEDED(hr))
  {
    LL llBitDepth;
    dwRes = pTwainApi->GetCapCurrentValue(ICAP_BITDEPTH ,&llBitDepth,&wCapType);
    if(dwRes!=0 || wCapType!=TWTY_UINT16)
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
    LONG lBytesPerLine = ((lPixelsPerLine*(long)llBitDepth+31)/32)*4;

    hr = wiasWritePropLong(pWiasContext,WIA_IPA_BYTES_PER_LINE ,lBytesPerLine);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasSetValidRangeLong(pWiasContext,WIA_IPS_XPOS, 0,0, lPixWidthMax - lPixWidthMin, 1);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasWritePropLong(pWiasContext,WIA_IPS_XPOS ,lXPosition);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasSetValidRangeLong(pWiasContext,WIA_IPS_YPOS, 0,0, lPixHeightMax -lPixHeightMin, 1);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasWritePropLong(pWiasContext,WIA_IPS_YPOS ,lYPosition);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasSetValidRangeLong(pWiasContext,WIA_IPS_XEXTENT,lPixWidthMin, lXExtent,lPixWidthMax - lXPosition, 1);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasWritePropLong(pWiasContext,WIA_IPS_XEXTENT ,lXExtent);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasSetValidRangeLong(pWiasContext,WIA_IPS_YEXTENT, lPixHeightMin, lYExtent, lPixHeightMax - lYPosition, 1);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasWritePropLong(pWiasContext,WIA_IPS_YEXTENT ,lYExtent);
  }

  if(m_bIsWindowsVista)
  {
    bool bFeeder = false;
    if(SUCCEEDED(hr))
    {
      hr = IsFeederItem(pWiasContext,&bFeeder);
    }  
    if(bFeeder)
    {
      if (SUCCEEDED(hr)) 
      {
        LONG lHorizontalSize = (long)(((LL)ImgLayout.Frame.Right-(LL)ImgLayout.Frame.Left)*1000);
        hr = wiasWritePropLong(pWiasContext,WIA_IPS_PAGE_WIDTH, lHorizontalSize);
      }

      if (SUCCEEDED(hr)) 
      {
        LONG lVerticalSize = (long)(((LL)ImgLayout.Frame.Bottom-(LL)ImgLayout.Frame.Top)*1000);
        hr = wiasWritePropLong(pWiasContext,WIA_IPS_PAGE_HEIGHT, lVerticalSize);
      }
    }
  }
  else
  {
    if (SUCCEEDED(hr)) 
    {
      LONG lHorizontalSize = (long)(((LL)ImgLayout.Frame.Right-(LL)ImgLayout.Frame.Left)*1000);
      hr = wiasWritePropLong(pWiasContext,WIA_DPS_PAGE_WIDTH, lHorizontalSize);
    }

    if (SUCCEEDED(hr)) 
    {
      LONG lVerticalSize = (long)(((LL)ImgLayout.Frame.Bottom-(LL)ImgLayout.Frame.Top)*1000);
      hr = wiasWritePropLong(pWiasContext,WIA_DPS_PAGE_HEIGHT, lVerticalSize);
    }
  }

  return hr;
}

HRESULT CWIADriver::SetDAT_IMAGELAYOUT (BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  long lXRes= 0;
  long lYRes= 0;
  long lXPos= 0;
  long lYPos= 0;
  long lXExt= 0;
  long lYExt= 0;
  long lPixWidthMin ;
  long lPixWidthMax ;
  long lPixHeightMin;
  long lPixHeightMax;

  TW_IMAGELAYOUT ImgLayout;
  memset(&ImgLayout,0,sizeof(TW_IMAGELAYOUT));
  ImgLayout.FrameNumber=1;
  ImgLayout.DocumentNumber=1;
  ImgLayout.PageNumber=1;
  HRESULT hr = wiasReadPropLong(pWiasContext, WIA_IPS_XRES, &lXRes, NULL, TRUE);
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_YRES, &lYRes, NULL, TRUE);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_XPOS, &lXPos, NULL, TRUE);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_YPOS, &lYPos, NULL, TRUE);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_XEXTENT, &lXExt, NULL, TRUE);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_YEXTENT, &lYExt, NULL, TRUE);
  }

///////////////////////
  if(m_bIsWindowsVista)
  {
    if(SUCCEEDED(hr))
    {
      hr = wiasReadPropLong(pWiasContext, WIA_IPS_MAX_HORIZONTAL_SIZE, &lPixWidthMax, NULL, TRUE);
    }
    if(SUCCEEDED(hr))
    {
      hr = wiasReadPropLong(pWiasContext, WIA_IPS_MAX_VERTICAL_SIZE, &lPixHeightMax, NULL, TRUE);
    }
    if(SUCCEEDED(hr))
    {
      hr = wiasReadPropLong(pWiasContext, WIA_IPS_MIN_HORIZONTAL_SIZE, &lPixWidthMin, NULL, TRUE);
    }
    if(SUCCEEDED(hr))
    {
      hr = wiasReadPropLong(pWiasContext, WIA_IPS_MIN_VERTICAL_SIZE, &lPixHeightMin, NULL, TRUE);
    }
  }
  else
  {
    bool bFeeder = false;
    hr = IsFeederItem(pWiasContext,&bFeeder);
    BYTE  *pRootItemContext   = NULL;

    if(SUCCEEDED(hr))
    {
      hr = wiasGetRootItem(pWiasContext,&pRootItemContext);
    }
    if(bFeeder)
    {
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_HORIZONTAL_SHEET_FEED_SIZE, &lPixWidthMax, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_VERTICAL_SHEET_FEED_SIZE, &lPixHeightMax, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_MIN_HORIZONTAL_SHEET_FEED_SIZE, &lPixWidthMin, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_MIN_VERTICAL_SHEET_FEED_SIZE, &lPixHeightMin, NULL, TRUE);
      }
    }
    else
    {
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_HORIZONTAL_BED_SIZE, &lPixWidthMax, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_VERTICAL_BED_SIZE, &lPixHeightMax, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, CUSTOM_ROOT_PROP_WIA_DPS_MIN_HORIZONTAL_BED_SIZE, &lPixWidthMin, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, CUSTOM_ROOT_PROP_WIA_DPS_MIN_VERTICAL_BED_SIZE, &lPixHeightMin, NULL, TRUE);
      }
    }
  }
  lPixWidthMin = (lPixWidthMin*lXRes)/1000;
  lPixWidthMax = (lPixWidthMax*lXRes)/1000;
  lPixHeightMin = (lPixHeightMin*lYRes)/1000;
  lPixHeightMax = (lPixHeightMax*lYRes)/1000;
  if(lPixWidthMax - lXPos<lXExt)
  {
    lXExt = lPixWidthMax - lXPos;
  }
  if(lPixHeightMax - lYPos<lYExt)
  {
    lYExt = lPixHeightMax - lYPos;
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasWritePropLong(pWiasContext,WIA_IPS_XEXTENT,lXExt);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasWritePropLong(pWiasContext,WIA_IPS_YEXTENT,lYExt);
  }

//////////////////////

  if(SUCCEEDED(hr))
  {
    LL llTemp;
    llTemp = ((LL)lXPos)/lXRes;
    ImgLayout.Frame.Left = (TW_FIX32)llTemp;
    llTemp = ((LL)lYPos)/lYRes;
    ImgLayout.Frame.Top = (TW_FIX32)llTemp;
    llTemp = ((LL)(lXExt+lXPos))/lXRes;
    ImgLayout.Frame.Right = (TW_FIX32)llTemp;
    llTemp = ((LL)(lYExt+lYPos))/lYRes;
    ImgLayout.Frame.Bottom = (TW_FIX32)llTemp;
    DWORD dwRes = pTwainApi->SetImageLayout(ImgLayout);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes && ((WORD)dwRes!=TWRC_CHECKSTATUS))//ignore if TWRC_CHECKSTATUS
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return hr;
}
HRESULT CWIADriver::GetXFERCOUNT(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  HRESULT hr;
  DWORD dwRes;
  bool bFeeder = false;
  hr = IsFeederItem(pWiasContext,&bFeeder);
  if(hr!=S_OK)
  {
    return hr;
  }
  long lPagesToScan = 1;
  LONG lMinPages      = 1;
  LONG lMaxPages      = 1;

  LL llVal;
  LL_ARRAY llLst;

  WORD wCapType;
  dwRes = pTwainApi->GetCapCurrentValue(CAP_XFERCOUNT,&llVal,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_INT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }
  if((long)llVal==-1)
  {
    lPagesToScan=0;
  }
  else
  {
    lPagesToScan=(long)llVal;
  }
  dwRes = pTwainApi->GetCapConstrainedValues(CAP_XFERCOUNT,&llLst,&wCapType);
  if(dwRes!=0 || wCapType!=TWTY_INT16)
  {
    return WIA_ERROR_EXCEPTION_IN_DRIVER;
  }

  llLst.sort();
  if(llLst.front()==-1)
  {
    lMinPages=0;
  }
  else
  {
    lMinPages = (long)llLst.front(); 
  }

  if(llLst.back()==-1)
  {
    lMaxPages=0;
  }
  else
  {
    lMaxPages = (long)llLst.back(); 
  }

  if(m_bIsWindowsVista)
  {
    if(SUCCEEDED(hr))
    {
      hr = wiasSetValidRangeLong(pWiasContext,WIA_IPS_PAGES, lMinPages, lPagesToScan, lMaxPages, 1);
    }
    if(SUCCEEDED(hr))
    {
      hr = wiasWritePropLong(pWiasContext,WIA_IPS_PAGES,lPagesToScan);
    }
  }
  else
  {
    BYTE *pRootWiasContext=NULL;
    hr = wiasGetRootItem(pWiasContext,&pRootWiasContext);

    if(SUCCEEDED(hr))
    {
      hr = wiasSetValidRangeLong(pRootWiasContext,WIA_DPS_PAGES, lMinPages, lPagesToScan, lMaxPages, 1);
    }
    if(SUCCEEDED(hr))
    {
      hr = wiasWritePropLong(pRootWiasContext,WIA_DPS_PAGES,lPagesToScan);
    }
  }
  return hr;
}

HRESULT CWIADriver::SetXFERCOUNT(BYTE *pWiasContext, CTWAIN_API *pTwainApi)
{
  HRESULT hr;
  long lPagesToScan = 0;
  if(m_bIsWindowsVista)
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_PAGES, &lPagesToScan, NULL, TRUE);
  }
  else
  {
    BYTE *pRootWiasContext=NULL;
    hr = wiasGetRootItem(pWiasContext,&pRootWiasContext);
    hr = wiasReadPropLong(pRootWiasContext,WIA_DPS_PAGES,&lPagesToScan,NULL,TRUE);
  }
  if(SUCCEEDED(hr))
  {
    if(lPagesToScan==0)
    {
      lPagesToScan=-1;
    }
    DWORD dwRes = pTwainApi->SetCapability(CAP_XFERCOUNT,lPagesToScan,TWTY_INT16);
    if(dwRes==FAILURE(TWCC_BADVALUE))
    {
      return E_INVALIDARG;
    }
    if(dwRes==FAILURE(TWCC_BUMMER))
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
    LL llVal;
    WORD wCapType;
    dwRes = pTwainApi->GetCapCurrentValue(CAP_XFERCOUNT,&llVal,&wCapType);
    if(dwRes!=0 || wCapType!=TWTY_INT16 || (long)llVal!=lPagesToScan)
    {
      return WIA_ERROR_EXCEPTION_IN_DRIVER;
    }
  }
  return hr;
}

HRESULT CWIADriver::SetAllTWAIN_Caps(BYTE *pWiasContext)
{
  HRESULT hr;
  CTWAIN_API *pTwainApi;
  hr = GetDS(pWiasContext,&pTwainApi);

  for(int i=0; g_unSetCapOrderFunc[i] && SUCCEEDED(hr); i++)
  {
    hr = (this->*g_unSetCapOrderFunc[i])(pWiasContext,pTwainApi);
  }  
  return hr;
}

HRESULT CWIADriver::GetAllTWAIN_Caps(BYTE *pWiasContext)
{
  HRESULT hr;
  CTWAIN_API *pTwainApi;
  hr = GetDS(pWiasContext,&pTwainApi);

  for(int i=0; g_unGetCapOrderFunc[i] && SUCCEEDED(hr); i++)
  {
    hr = (this->*g_unGetCapOrderFunc[i])(pWiasContext,pTwainApi);
  }  
  return hr;
}

HRESULT CWIADriver::ValidateThroughTWAINDS(BYTE *pWiasContext, LL_ARRAY llLst)
{
  HRESULT hr;

  if(llLst.size()==0)
  {
    return S_OK;//nothing to set
  }

  CTWAIN_API *pTwainApi;
  hr = GetDS(pWiasContext,&pTwainApi);
  bool bFound=false;

  // Set all TWAIN capabilities untill it reach first changed. This is because we open TWAIN DS every time and we have to be sure that it runs in the same conditions every time
  // After this point we get back all other capailitiess, because they can depend on previously set 
  LL_ARRAY::iterator llIter=llLst.begin();
  for(int i=0; g_unSetCapOrder[i]&& SUCCEEDED(hr); i++)
  {
    if(llIter!=llLst.end())
    {
      if(g_unSetCapOrder[i]==(short)(*llIter) )//set changed one
      {
        hr = (this->*g_unSetCapOrderFunc[i])(pWiasContext,pTwainApi);
        llIter++;
        bFound = true;
        continue;
      } 
    }
    if(!bFound)
    {
      hr = (this->*g_unSetCapOrderFunc[i])(pWiasContext,pTwainApi);//set all until it finds first
    }
    else
    { // Read back capability and update WIA properties
      for(int k=0; g_unGetCapOrder[k]; k++)//find read function
      {
        if(g_unGetCapOrder[k]==g_unSetCapOrder[i])
        {
          hr = (this->*g_unGetCapOrderFunc[k])(pWiasContext,pTwainApi); // read the capability
          break;
        }
      }
    }
  }
  return hr;
}

HRESULT CWIADriver::IsFeederItem(BYTE *pWiasContext, bool *pbFeeder)
{
  HRESULT hr;
  *pbFeeder = false;

  if(m_bIsWindowsVista)
  {
    // WIA 2.0 get Item GUID - one WIA item per source
    GUID          guidItemCategory    = GUID_NULL;
    hr = wiasReadPropGuid(pWiasContext,WIA_IPA_ITEM_CATEGORY,&guidItemCategory,NULL,TRUE);
    *pbFeeder = (guidItemCategory == WIA_CATEGORY_FEEDER)?true:false;
  }
  else
  {
    // WIA 1.0 get source by reading WIA_DPS_DOCUMENT_HANDLING_SELECT - only one WIA item
    BYTE *pWiasRootContext=NULL;
    hr = wiasGetRootItem(pWiasContext,&pWiasRootContext);
    if(SUCCEEDED(hr))
    {
      LONG lFeeder = 0;
      hr = wiasReadPropLong(pWiasRootContext,WIA_DPS_DOCUMENT_HANDLING_SELECT,&lFeeder,NULL,TRUE);
      if(SUCCEEDED(hr))
      {
        if(lFeeder&(FEEDER|DUPLEX))
        {
          *pbFeeder = true;
        }
      }
    }
  }
  return hr;
}

HRESULT CWIADriver::UpdateIntent(BYTE* pWiasContext, WIA_PROPERTY_CONTEXT *pContext)
{
  HRESULT hr;
  WIAS_CHANGED_VALUE_INFO Intent;		
  memset(&Intent,0,sizeof(WIAS_CHANGED_VALUE_INFO));

  hr = wiasGetChangedValueLong(pWiasContext, pContext, FALSE, WIA_IPS_CUR_INTENT, &Intent);

  if(SUCCEEDED(hr) && Intent.bChanged)
  {
    switch(Intent.Current.lVal & WIA_INTENT_IMAGE_TYPE_MASK) 
    {
    case WIA_INTENT_NONE:
      break;
    case WIA_INTENT_IMAGE_TYPE_COLOR:	
      hr = wiasWritePropLong(pWiasContext, WIA_IPA_DATATYPE, WIA_DATA_COLOR);
      break;
    case WIA_INTENT_IMAGE_TYPE_GRAYSCALE:	
      hr = wiasWritePropLong(pWiasContext, WIA_IPA_DATATYPE, WIA_DATA_GRAYSCALE);
      break;
    case WIA_INTENT_IMAGE_TYPE_TEXT:		
      hr = wiasWritePropLong(pWiasContext, WIA_IPA_DATATYPE, WIA_DATA_THRESHOLD);
      break;
    default:
      hr = E_INVALIDARG;
      break;
    }
  }

  return hr;
}

HRESULT CWIADriver::UpdateWIAPropDepend(BYTE *pWiasContext)
{
  long lDepth;
  long lChannels;
  long lDataType;
  long lXRes;
  long lYRes;
  long lXPos;
  long lYPos;
  long lXExt;
  long lYExt;
  long lPixWidthMin;
  long lPixWidthMax;
  long lPixHeightMin;
  long lPixHeightMax;

  //Collect data first
  HRESULT hr = wiasReadPropLong(pWiasContext, WIA_IPA_DATATYPE, &lDataType, NULL, TRUE);
  if(SUCCEEDED(hr))
  {
    switch(lDataType)
    {
      case WIA_DATA_THRESHOLD:
        lChannels = 1;  
        break;
      case WIA_DATA_GRAYSCALE:
        lChannels = 1;  
        break;
      case WIA_DATA_COLOR:
        lChannels = 3;  
        break;
      default:
        hr = E_INVALIDARG;
        break;
    }
  }

  if(SUCCEEDED( hr ))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPA_DEPTH, &lDepth, NULL, TRUE);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_XRES, &lXRes, NULL, TRUE);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_YRES, &lYRes, NULL, TRUE);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_XPOS, &lXPos, NULL, TRUE);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_YPOS, &lYPos, NULL, TRUE);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_XEXTENT, &lXExt, NULL, TRUE);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasReadPropLong(pWiasContext, WIA_IPS_YEXTENT, &lYExt, NULL, TRUE);
  }

  bool bFeeder = false;
  if(SUCCEEDED(hr))
  {
    hr = IsFeederItem(pWiasContext,&bFeeder);
  }  

  if(m_bIsWindowsVista)
  {
    if(SUCCEEDED(hr))
    {
      hr = wiasReadPropLong(pWiasContext, WIA_IPS_MAX_HORIZONTAL_SIZE, &lPixWidthMax, NULL, TRUE);
    }
    if(SUCCEEDED(hr))
    {
      hr = wiasReadPropLong(pWiasContext, WIA_IPS_MAX_VERTICAL_SIZE, &lPixHeightMax, NULL, TRUE);
    }
    if(SUCCEEDED(hr))
    {
      hr = wiasReadPropLong(pWiasContext, WIA_IPS_MIN_HORIZONTAL_SIZE, &lPixWidthMin, NULL, TRUE);
    }
    if(SUCCEEDED(hr))
    {
      hr = wiasReadPropLong(pWiasContext, WIA_IPS_MIN_VERTICAL_SIZE, &lPixHeightMin, NULL, TRUE);
    }
  }
  else
  {

    BYTE  *pRootItemContext   = NULL;

    if(SUCCEEDED(hr))
    {
      hr = wiasGetRootItem(pWiasContext,&pRootItemContext);
    }
    if(bFeeder)
    {
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_HORIZONTAL_SHEET_FEED_SIZE, &lPixWidthMax, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_VERTICAL_SHEET_FEED_SIZE, &lPixHeightMax, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_MIN_HORIZONTAL_SHEET_FEED_SIZE, &lPixWidthMin, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_MIN_VERTICAL_SHEET_FEED_SIZE, &lPixHeightMin, NULL, TRUE);
      }
    }
    else
    {
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_HORIZONTAL_BED_SIZE, &lPixWidthMax, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, WIA_DPS_VERTICAL_BED_SIZE, &lPixHeightMax, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, CUSTOM_ROOT_PROP_WIA_DPS_MIN_HORIZONTAL_BED_SIZE, &lPixWidthMin, NULL, TRUE);
      }
      if(SUCCEEDED(hr))
      {
        hr = wiasReadPropLong(pRootItemContext, CUSTOM_ROOT_PROP_WIA_DPS_MIN_VERTICAL_BED_SIZE, &lPixHeightMin, NULL, TRUE);
      }
    }
  }
  if(SUCCEEDED( hr ))
  {
    lPixWidthMin = (lPixWidthMin*lXRes)/1000;
    lPixWidthMax = (lPixWidthMax*lXRes)/1000;
    lPixHeightMin = (lPixHeightMin*lYRes)/1000;
    lPixHeightMax = (lPixHeightMax*lYRes)/1000;
    if(lPixWidthMax - lXPos<lXExt)
    {
      lXExt = lPixWidthMax - lXPos;
    }
    if(lPixHeightMax - lYPos<lYExt)
    {
      lYExt = lPixHeightMax - lYPos;
    }  
  }

//Update dependent properties
  if(SUCCEEDED( hr ))
  {
    hr = wiasWritePropLong(pWiasContext, WIA_IPA_CHANNELS_PER_PIXEL, lChannels);
  }

  if(SUCCEEDED( hr ))
  {
    hr = wiasWritePropLong(pWiasContext, WIA_IPA_BITS_PER_CHANNEL, lDepth/lChannels);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasSetValidRangeLong(pWiasContext,WIA_IPS_XEXTENT,lPixWidthMin, lXExt,lPixWidthMax - lXPos, 1);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasSetValidRangeLong(pWiasContext,WIA_IPS_YEXTENT, lPixHeightMin, lYExt, lPixHeightMax - lYPos, 1);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasWritePropLong(pWiasContext,WIA_IPS_XEXTENT,lXExt);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasWritePropLong(pWiasContext,WIA_IPS_YEXTENT,lYExt);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasSetValidRangeLong(pWiasContext,WIA_IPS_YEXTENT, lPixHeightMin, lYExt, lPixHeightMax - lYPos, 1);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasSetValidRangeLong(pWiasContext,WIA_IPS_XPOS, 0,0, lPixWidthMax - lPixWidthMin, 1);
  }

  if(SUCCEEDED(hr))
  {
    hr = wiasSetValidRangeLong(pWiasContext,WIA_IPS_YPOS, 0,0, lPixHeightMax -lPixHeightMin, 1);
  }
  if(m_bIsWindowsVista)
  {
    if(bFeeder)
    {
      if (SUCCEEDED(hr)) 
      {
        LONG lHorizontalSize = (lXExt*1000)/lXRes;
        hr = wiasWritePropLong(pWiasContext,WIA_IPS_PAGE_WIDTH, lHorizontalSize);
      }

      if (SUCCEEDED(hr)) 
      {
        LONG lVerticalSize = (lYExt*1000)/lYRes;
        hr = wiasWritePropLong(pWiasContext,WIA_IPS_PAGE_HEIGHT, lVerticalSize);
      }
    }
  }
  else
  {
    if (SUCCEEDED(hr)) 
    {
      LONG lHorizontalSize = (lXExt*1000)/lXRes;
      hr = wiasWritePropLong(pWiasContext,WIA_DPS_PAGE_WIDTH, lHorizontalSize);
    }

    if (SUCCEEDED(hr)) 
    {
      LONG lVerticalSize = (lYExt*1000)/lYRes;
      hr = wiasWritePropLong(pWiasContext,WIA_DPS_PAGE_HEIGHT, lVerticalSize);
    }
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasWritePropLong(pWiasContext,WIA_IPA_NUMBER_OF_LINES,lYExt);
  }
  if(SUCCEEDED(hr))
  {
    hr = wiasWritePropLong(pWiasContext,WIA_IPA_PIXELS_PER_LINE ,lXExt);
  }

  if(SUCCEEDED(hr))
  {
    LONG lBytesPerLine = ((lXExt*(long)lDepth+31)/32)*4;
    hr = wiasWritePropLong(pWiasContext,WIA_IPA_BYTES_PER_LINE ,lBytesPerLine);
  }

  return hr;
}

HRESULT CWIADriver::TWAINtoWIAerror(DWORD dwError)
{
  switch(LOWORD(dwError))
  {
    case TWRC_SUCCESS:
    case TWRC_CHECKSTATUS:
    case TWRC_XFERDONE:
      return S_OK;
    case TWRC_CANCEL:
      return ERROR_CANCELLED;
    default:
      return E_UNEXPECTED;
    case TWRC_FAILURE:
      switch(HIWORD(dwError))
      {
        case TWCC_SUCCESS:
          return WIA_ERROR_GENERAL_ERROR;
        case TWCC_BUMMER:
          return WIA_ERROR_EXCEPTION_IN_DRIVER;
        case TWCC_LOWMEMORY:
          return E_OUTOFMEMORY;
        case TWCC_NODS:
          return WIA_ERROR_OFFLINE;
        case TWCC_MAXCONNECTIONS:
          return WIA_ERROR_BUSY;
        case TWCC_OPERATIONERROR:
          return WIA_ERROR_GENERAL_ERROR;
        case TWCC_BADCAP:
          return WIA_ERROR_INVALID_DRIVER_RESPONSE;
        case TWCC_BADPROTOCOL:
          return WIA_ERROR_INVALID_DRIVER_RESPONSE;
        case TWCC_BADVALUE:
          return E_INVALIDARG;
        case TWCC_SEQERROR:
          return WIA_ERROR_INVALID_DRIVER_RESPONSE;
        case TWCC_BADDEST:
          return E_INVALIDARG;
        case TWCC_CAPUNSUPPORTED:
          return WIA_ERROR_INVALID_DRIVER_RESPONSE;
        case TWCC_CAPBADOPERATION:
          return WIA_ERROR_INVALID_DRIVER_RESPONSE;
        case TWCC_CAPSEQERROR:
          return WIA_ERROR_INVALID_DRIVER_RESPONSE;
        case TWCC_DENIED:
          return WIA_ERROR_GENERAL_ERROR;
        case TWCC_FILEEXISTS:
          return WIA_ERROR_GENERAL_ERROR;
        case TWCC_FILENOTFOUND:
          return WIA_ERROR_GENERAL_ERROR;
        case TWCC_NOTEMPTY:
          return WIA_ERROR_GENERAL_ERROR;
        case TWCC_PAPERJAM:
          return WIA_ERROR_PAPER_JAM;
        case TWCC_PAPERDOUBLEFEED:
          return WIA_ERROR_PAPER_PROBLEM;
        case TWCC_FILEWRITEERROR:
          return WIA_ERROR_GENERAL_ERROR;
        case TWCC_CHECKDEVICEONLINE:
          return WIA_ERROR_DEVICE_COMMUNICATION;
        case TWCC_INTERLOCK:
          return WIA_ERROR_DEVICE_LOCKED;
        case TWCC_DAMAGEDCORNER:
          return WIA_ERROR_PAPER_PROBLEM;
        case TWCC_FOCUSERROR:
          return WIA_ERROR_PAPER_PROBLEM;
        case TWCC_DOCTOOLIGHT:
          return WIA_ERROR_PAPER_PROBLEM;
        case TWCC_DOCTOODARK:
          return WIA_ERROR_PAPER_PROBLEM;
        default:
          return WIA_ERROR_GENERAL_ERROR;

      }
  }
  return E_UNEXPECTED;
}
