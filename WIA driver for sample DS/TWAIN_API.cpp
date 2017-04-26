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
 * @file TWAIN_API.cpp
 * CTWAIN_API implementation file
 * @author TWAIN Working Group
 * @date October 2009
 */

#include "StdAfx.h"
#include <stdio.h>
#include <tchar.h>
#include <windows.h>
#include "TWAIN_API.h"

extern HINSTANCE g_hInst;
#define DSM2NAME L"system32\\TWAINDSM.DLL"  //TWAIN DSM location
#define CLASS_NAME L"CLASS_WIA_ON_TWAIN"    //Class name

extern map<int,CTWAIN_API*> g_TWAIN_APIs;   //list of all opened TWAIN DS

const TW_UINT16 g_unUnitDepCap[] = 
{// list of all TW_FIX32 depend on ICAP_UNITS
  ICAP_PHYSICALWIDTH,
  ICAP_PHYSICALHEIGHT,
  ICAP_MINIMUMHEIGHT,
  ICAP_MINIMUMWIDTH,
  0
};

const TW_UINT16 g_un1_UnitDepCap[] = 
{// list of all TW_FIX32 depend on 1/ICAP_UNITS
  ICAP_XRESOLUTION,
  ICAP_YRESOLUTION,
  ICAP_XNATIVERESOLUTION,
  ICAP_YNATIVERESOLUTION,
  0
};

//WindowProc for the window created by CTWAIN_API class
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

CTWAIN_API::CTWAIN_API(HWND hWindow)
{
  m_hWindow = NULL;
  m_bOwnWindow = false;

  if(hWindow)
  {
    m_hWindow = hWindow;
    m_bOwnWindow=true;
  }
  memset(&m_AppID,0,sizeof(TW_IDENTITY));
  memset(&m_DS,0,sizeof(TW_IDENTITY));
  m_twState = DSM_LOADED;
  m_hDSMEvent = CreateEvent(NULL,FALSE,FALSE,NULL);
}

CTWAIN_API::~CTWAIN_API()
{
  if(m_hDSMEvent)
  {
    CloseHandle(m_hDSMEvent);
  }
}

DWORD CTWAIN_API::OpenDSM(TW_IDENTITY AppID, int nID)
{
  DWORD dwRes=0;

  if(m_hDSMEvent==NULL)//validate m_hDSMEvent allocated in constructor
  {
    return FAILURE(TWCC_LOWMEMORY);
  }

  if (m_twState != DSM_LOADED )
  {
    return FAILURE(TWCC_SEQERROR);
  }

  m_nID = nID;
  m_AppID = AppID;
  m_hDSM = NULL;
  m_fnDSM_Entry = NULL;
  memset(&m_DSM_Entry, 0, sizeof(TW_ENTRYPOINT));
  m_AppID.ProductName[33]=0;
  CString strProductName;
  strProductName = m_AppID.ProductName;//convert to WCHAR
  m_strClassName.Format(L"%s_%s_%d",CLASS_NAME,(LPCTSTR)strProductName,m_nID);

  WCHAR         WinDir[MAX_PATH];
  WCHAR         DSMName[MAX_PATH];

  memset(WinDir, 0, MAX_PATH*sizeof(WCHAR));
  memset(DSMName, 0, MAX_PATH*sizeof(WCHAR));

  if(GetWindowsDirectory (WinDir, MAX_PATH)==0)
  {
    WinDir[0]=0;
  }

  if (WinDir[wcslen(WinDir)-1] != '\\')
  {
    wcscat_s(WinDir,MAX_PATH, L"\\");
  }    

  wcscpy_s (DSMName,MAX_PATH, WinDir);
  wcscat_s (DSMName,MAX_PATH, DSM2NAME);


  if ((m_hDSM = LoadLibrary(DSMName)) != NULL)//error loading DSM
  {
    if((m_fnDSM_Entry = (DSMENTRYPROC)GetProcAddress(m_hDSM, MAKEINTRESOURCEA (1))) == NULL)
    {
      FreeLibrary (m_hDSM);
      m_hDSM=NULL;
    }
  }

  if(m_fnDSM_Entry == NULL)//error accessing DSM_Entry point
  {
    return FAILURE(TWCC_BUMMER);
  }

  try 
  {
    if(m_hWindow == NULL)
    {
      WNDCLASS  wc;
      wc.hCursor        = NULL;
      wc.hIcon          = NULL;
      wc.lpszMenuName   = NULL;
      wc.lpszClassName  = (LPCTSTR)m_strClassName;
      wc.hbrBackground  = (HBRUSH)(COLOR_BACKGROUND);
      wc.hInstance      = g_hInst;
      wc.style          = 0;
      wc.lpfnWndProc    = WindowProc;
      wc.cbWndExtra     = sizeof(DWORD);
      wc.cbClsExtra     = 0;
      if (!RegisterClass(&wc))
      {
        if(GetLastError()!= ERROR_CLASS_ALREADY_EXISTS)//Already registered?
        {
          dwRes = FAILURE(TWCC_BUMMER);
        }
      }
      
      if( dwRes == TWRC_SUCCESS)
      {
        m_hWindow = CreateWindow(wc.lpszClassName, NULL, WS_OVERLAPPED, 0, 0, 0, 0, HWND_MESSAGE  
                          ,(HMENU) NULL, g_hInst, NULL);
      }
      if (m_hWindow==NULL)
      {
        dwRes = FAILURE(TWCC_BUMMER);
      }
    }
    if( dwRes == TWRC_SUCCESS)
    {
      dwRes = (*m_fnDSM_Entry)(&m_AppID, NULL, DG_CONTROL, DAT_PARENT, MSG_OPENDSM, &m_hWindow);
    }
    if( dwRes == TWRC_SUCCESS)
    {
      if((m_AppID.SupportedGroups & DF_DSM2) == DF_DSM2)
      {
        m_DSM_Entry.Size = sizeof(TW_ENTRYPOINT);
        if( (*m_fnDSM_Entry)(&m_AppID, NULL, DG_CONTROL, DAT_ENTRYPOINT, MSG_GET, &m_DSM_Entry) != TWRC_SUCCESS)
        {
          memset(&m_DSM_Entry, 0, sizeof(TW_ENTRYPOINT));
        }
      }          
    }
  }
  catch (...)
  {
    dwRes = FAILURE(TWCC_BUMMER);
  }

  if( dwRes == TWRC_SUCCESS)//if OK change state to DSM_OPENED
  {
    m_twState = DSM_OPENED;
    m_DS.Id = NULL;
  }
  else
  {
    FreeLibrary (m_hDSM); // if not free all resources
    m_hDSM=NULL;
    if(!m_bOwnWindow && m_hWindow)
    {
      DestroyWindow(m_hWindow);
      m_hWindow = NULL;
    }
  }
  return dwRes;
}

DWORD CTWAIN_API::CloseDSM()
{
  DWORD dwRes=0;
  if (m_twState >= DSM_OPENED)
  {
    try
    {
      dwRes = (*m_fnDSM_Entry)(&m_AppID, NULL, DG_CONTROL, DAT_PARENT, MSG_CLOSEDSM, m_hWindow);
    }
    catch (...)
    {
      dwRes = FAILURE(TWCC_BUMMER);
    }
  }

  m_AppID.Id = NULL;

  if(m_hDSM)
  {
    FreeLibrary (m_hDSM);
    m_hDSM=NULL;
  }

  UnregisterClass((LPCTSTR)m_strClassName,g_hInst);
  if(!m_bOwnWindow && m_hWindow)
  {
    DestroyWindow(m_hWindow);
    m_hWindow = NULL;
  }
  if (m_twState >= DSM_OPENED)
  {
    m_twState = DSM_LOADED;
    return dwRes;
  }
  return FAILURE(TWCC_SEQERROR);
}

DWORD CTWAIN_API::OpenDS(TW_IDENTITY *pDS_ID)
{
  DWORD dwRes;
  TW_IDENTITY DSid;

  if(pDS_ID)
  {
    if (m_twState != DSM_OPENED)
    {
      return FAILURE(TWCC_SEQERROR);
    }
    memset(&m_DS,0,sizeof(TW_IDENTITY));
    try
    {
      dwRes = (*m_fnDSM_Entry)(&m_AppID, NULL, DG_CONTROL, DAT_IDENTITY, MSG_GETFIRST, &DSid);
    }
    catch (...)
    {
      return FAILURE(TWCC_BUMMER);
    }

    string strMfg = pDS_ID->Manufacturer;
    string strModel = pDS_ID->ProductName;

    while (dwRes == TWRC_SUCCESS)
    {
      char strDSidManufacturer[sizeof(TW_STR32)+1];
      char strDSidProductName[sizeof(TW_STR32)+1];
      memset(strDSidManufacturer,0,sizeof(strDSidManufacturer));
      memset(strDSidProductName,0,sizeof(strDSidProductName));
      memcpy(strDSidManufacturer, DSid.Manufacturer,sizeof(TW_STR32));
      memcpy(strDSidProductName, DSid.ProductName,sizeof(TW_STR32));
      if(strMfg.compare(strDSidManufacturer)==0 && strModel.compare(strDSidProductName)==0 && DSid.Version.MajorNum>= pDS_ID->Version.MajorNum && (DSid.Version.MinorNum>=pDS_ID->Version.MinorNum || DSid.Version.MajorNum>pDS_ID->Version.MajorNum))
      {
        m_DS = DSid;
        break;
      }
      try
      {
        dwRes = (*m_fnDSM_Entry)(&m_AppID, NULL, DG_CONTROL, DAT_IDENTITY, MSG_GETNEXT, &DSid);
      }
      catch (...)
      {
        return FAILURE(TWCC_BUMMER);
      }
    }
  }
  else
  {
    if (m_twState < DSM_OPENED)
    {
      return FAILURE(TWCC_SEQERROR);
    }
    if (m_twState > DSM_OPENED)
    {
      return S_OK;
    }
  }
  if (m_DS.Id==0)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, NULL, DG_CONTROL, DAT_IDENTITY, MSG_OPENDS, &m_DS);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes == TWRC_FAILURE || dwRes == TWRC_CHECKSTATUS)
  {
    if(dwRes == TWRC_CHECKSTATUS)
    {
      CloseDS();
    }
    return GetDSMConditionCode((WORD)dwRes);
  }

  TW_CALLBACK callback = { 0 };
  callback.CallBackProc = (void*)(TWAIN_callback);
  callback.RefCon = m_nID;
  dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_CALLBACK, MSG_REGISTER_CALLBACK,(TW_MEMREF)&callback);
  if(dwRes)
  {
    m_twState = DS_OPENED;
    CloseDS();
    return dwRes;
  }
  m_twState = DS_OPENED;

  LL llVal;
  WORD wCapType;
 
  dwRes = SetCapability(CAP_INDICATORS, false,TWTY_BOOL);
  if(dwRes && dwRes!=FAILURE(TWCC_CAPUNSUPPORTED))
  {
    CloseDS();
    return dwRes;
  }
  if(dwRes==0)
  {
    dwRes = GetCapCurrentValue(CAP_INDICATORS,&llVal,&wCapType);
    if(dwRes!=0 || wCapType != TWTY_BOOL || ((TW_BOOL)llVal)!=0 )
    {
      CloseDS();
      return dwRes?dwRes:FAILURE(TWCC_BUMMER);
    }
  }

  dwRes = SetCapability(ICAP_XFERMECH, TWSX_NATIVE,TWTY_UINT16);
  if(dwRes)
  {
    CloseDS();
    return dwRes;
  }
  dwRes = GetCapCurrentValue(ICAP_XFERMECH,&llVal,&wCapType);
  if(dwRes!=0 || wCapType != TWTY_UINT16 || ((TW_UINT16)llVal)!=TWSX_NATIVE )
  {
    CloseDS();
    return dwRes?dwRes:FAILURE(TWCC_BUMMER);
  }
  
  dwRes = SetCapability(ICAP_UNITS,TWUN_INCHES,TWTY_UINT16);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    CloseDS();
    return dwRes;
  }
  if(dwRes)
  {
    LL_ARRAY llLst;
    dwRes = GetCapConstrainedValues(ICAP_UNITS,&llLst,&wCapType);
    if(dwRes)
    {
      CloseDS();
      return dwRes;
    }
    LL_ARRAY::iterator llIter=llLst.begin();
    for(;llIter!=llLst.end();llIter++)
    {
      if(((short)*llIter)!=TWUN_PIXELS)
      {
        dwRes = SetCapability(ICAP_UNITS,*llIter,TWTY_UINT16);
        if(dwRes)
        {
          CloseDS();
          return dwRes;
        }
        break;
      }
    }

  }
  dwRes = GetCapCurrentValue(ICAP_UNITS ,&llVal,&wCapType);
  if(dwRes==0 && wCapType==TWTY_UINT16 && (short)llVal!=TWUN_PIXELS)
  {
    m_wDSunits = (WORD)llVal;
  }
  else
  {
    CloseDS();
    return dwRes?dwRes:FAILURE(TWCC_BUMMER);
  }

  return 0;
}

DWORD CTWAIN_API::CloseDS()
{
  DWORD dwRes;
  if (m_twState < DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  m_twState = DSM_OPENED;
  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, NULL, DG_CONTROL, DAT_IDENTITY, MSG_CLOSEDS, &m_DS);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes == TWRC_FAILURE || dwRes == TWRC_CHECKSTATUS)
  {
    return MAKELONG(dwRes,GetDSMConditionCode((WORD)dwRes));
  }
  return dwRes;
}

DWORD CTWAIN_API::EnableDS(HANDLE * phEvent)
{
  DWORD dwRes;
  if (m_twState != DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  TW_USERINTERFACE guif;
  memset(&guif,0,sizeof(TW_USERINTERFACE));
  guif.hParent = m_hWindow;
  m_twState = DS_ENABLED;
  try
  {
    m_msgLast = 0;
    ResetEvent(m_hDSMEvent);
    *phEvent = m_hDSMEvent;
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_USERINTERFACE, MSG_ENABLEDS, &guif);
  }
  catch (...)
  {
    m_twState = DS_OPENED;
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes == TWRC_FAILURE || dwRes == TWRC_CHECKSTATUS)
  {
    m_twState = DS_OPENED;
    return GetDSConditionCode((WORD)dwRes);
  }
  return dwRes;
}

DWORD CTWAIN_API::EnableDSOnly(HANDLE *phEvent)
{
  DWORD dwRes;
  if (m_twState != DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  TW_USERINTERFACE guif;
  memset(&guif,0,sizeof(TW_USERINTERFACE));
  guif.hParent = m_hWindow;
  guif.ShowUI=1;
  m_twState = DS_ENABLED;
  try
  {
    m_msgLast = 0;
    ResetEvent(m_hDSMEvent);
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_USERINTERFACE, MSG_ENABLEDSUIONLY, &guif);
    *phEvent = m_hDSMEvent;

    if (dwRes == TWRC_FAILURE || dwRes == TWRC_CHECKSTATUS)
    {
      m_twState = DS_OPENED;
      return GetDSConditionCode((WORD)dwRes);
    }  
    if (dwRes)
    {
      m_twState = DS_OPENED;
      return dwRes;
    }  
  }
  catch (...)
  {
    m_twState = DS_OPENED;
    return FAILURE(TWCC_BUMMER);
  }


  return dwRes;
}

DWORD CTWAIN_API::DisableDS()
{
  DWORD dwRes;
  if (m_twState >= DS_ENABLED)
  {
    TW_USERINTERFACE guif;
    memset(&guif,0,sizeof(TW_USERINTERFACE));
    guif.hParent = m_hWindow;
    m_twState = DS_OPENED;
    try
    {
      dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_USERINTERFACE, MSG_DISABLEDS, &guif);
    }
    catch (...)
    {
      return FAILURE(TWCC_BUMMER);
    }
    return GetDSConditionCode((WORD)dwRes);
  }
  return FAILURE(TWCC_SEQERROR);
}

DWORD CTWAIN_API::SetDSCustomData(TW_CUSTOMDSDATA Data)
{
  DWORD dwRes;

  if (m_twState != DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_CUSTOMDSDATA, MSG_SET, &Data);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }
  return GetDSConditionCode((WORD)dwRes);
}

DWORD CTWAIN_API::GetDSCustomData(TW_CUSTOMDSDATA *pData)
{
  DWORD dwRes;

  if (m_twState != DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_CUSTOMDSDATA, MSG_GET, pData);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }
  return GetDSConditionCode((WORD)dwRes);
}

DWORD CTWAIN_API::GetImageInfo(TW_IMAGEINFO *pImageInfo)
{
  DWORD dwRes;
  memset(pImageInfo,0,sizeof(TW_IMAGEINFO));

  if (m_twState < TRANSFER_READY)
  {
    return FAILURE(TWCC_SEQERROR);
  }
  
  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_IMAGE, DAT_IMAGEINFO, MSG_GET, pImageInfo);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }
  return GetDSConditionCode((WORD)dwRes);
}

DWORD CTWAIN_API::GetImageLayout(TW_IMAGELAYOUT *pImageLayout)
{
  DWORD dwRes;
  memset(pImageLayout,0,sizeof(TW_IMAGELAYOUT));

  if (m_twState != DS_OPENED && m_twState != DS_ENABLED && m_twState != TRANSFER_READY)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_IMAGE, DAT_IMAGELAYOUT, MSG_GET, pImageLayout);
    if(dwRes==0)
    {
      pImageLayout->Frame.Bottom =  (TW_FIX32)ConvertToInch(pImageLayout->Frame.Bottom);
      pImageLayout->Frame.Left =  (TW_FIX32)ConvertToInch(pImageLayout->Frame.Left);
      pImageLayout->Frame.Right =  (TW_FIX32)ConvertToInch(pImageLayout->Frame.Right);
      pImageLayout->Frame.Top =  (TW_FIX32)ConvertToInch(pImageLayout->Frame.Top);
    }
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }
  return GetDSConditionCode((WORD)dwRes);
}

DWORD CTWAIN_API::SetImageLayout(TW_IMAGELAYOUT ImageLayout)
{
  DWORD dwRes;

  if (m_twState != DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  ImageLayout.Frame.Bottom =  (TW_FIX32)ConvertFromInch(ImageLayout.Frame.Bottom);
  ImageLayout.Frame.Left =  (TW_FIX32)ConvertFromInch(ImageLayout.Frame.Left);
  ImageLayout.Frame.Right =  (TW_FIX32)ConvertFromInch(ImageLayout.Frame.Right);
  ImageLayout.Frame.Top =  (TW_FIX32)ConvertFromInch(ImageLayout.Frame.Top);

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_IMAGE, DAT_IMAGELAYOUT, MSG_SET, &ImageLayout);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }
  return GetDSConditionCode((WORD)dwRes);
}

DWORD CTWAIN_API::ScanNative(HANDLE *phBMP, DWORD *pdwSize)
{
  DWORD dwRes;
  *phBMP = NULL;
  *pdwSize = 0;

  if (m_twState != TRANSFER_READY)
  {
    return FAILURE(TWCC_SEQERROR);
  }
  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_IMAGE, DAT_IMAGENATIVEXFER, MSG_GET, phBMP);
    BITMAPINFO *pBMPInfo = (BITMAPINFO*)DSM_LockMemory(*phBMP);
    if (pBMPInfo==0)
    {
      DSM_Free(*phBMP);
      return FAILURE(TWCC_LOWMEMORY);
    }
    *pdwSize = pBMPInfo->bmiHeader.biSizeImage;
    if(*pdwSize == 0)
    {
      pBMPInfo->bmiHeader.biSizeImage = ((pBMPInfo->bmiHeader.biWidth*pBMPInfo->bmiHeader.biBitCount+31)/8)*pBMPInfo->bmiHeader.biHeight;
    }
    *pdwSize += pBMPInfo->bmiHeader.biSize;
    *pdwSize += pBMPInfo->bmiHeader.biClrUsed * sizeof(RGBQUAD);
    DSM_UnlockMemory(*phBMP);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }        
  if(dwRes == TWRC_XFERDONE)
  {
    m_twState = TRANSFERRING;
  }

  return GetDSConditionCode((WORD)dwRes);
}

DWORD CTWAIN_API::GetMemTransferCfg(TW_SETUPMEMXFER *pSetupMemTransfer)
{
  DWORD dwRes;
  memset(pSetupMemTransfer,0,sizeof(TW_SETUPMEMXFER));

  if (m_twState != DS_OPENED && m_twState != DS_ENABLED && m_twState != TRANSFER_READY)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_SETUPMEMXFER, MSG_GET, pSetupMemTransfer);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }
  return GetDSConditionCode((WORD)dwRes);
}

DWORD CTWAIN_API::EndTransfer(bool *pbMoreImages)
{
  DWORD dwRes;

  *pbMoreImages = false;

  if (m_twState < DS_ENABLED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  TW_PENDINGXFERS pxfr;
  memset(&pxfr,0,sizeof(TW_PENDINGXFERS));
  m_twState = TRANSFER_READY;
  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_PENDINGXFERS, MSG_ENDXFER, &pxfr);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes == TWRC_SUCCESS)
  {
    if (pxfr.Count != 0)
    {
      *pbMoreImages = true;
      m_twState = TRANSFER_READY;
    }
    else
    {
      m_twState = DS_ENABLED;
    }
  }
  return GetDSConditionCode((WORD)dwRes);
}

DWORD CTWAIN_API::ResetTransfer()
{
  DWORD dwRes;
  if (m_twState < DS_ENABLED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  TW_PENDINGXFERS pxfr;
  memset(&pxfr,0,sizeof(TW_PENDINGXFERS));
  m_twState = DS_ENABLED;

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_PENDINGXFERS, MSG_RESET, &pxfr);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }
  return GetDSConditionCode((WORD)dwRes);
}

DWORD CTWAIN_API::ProcessEvent(MSG *pMsg, WORD *pTwainMsg, bool *pbDSevent)
{
  DWORD dwRes;
  *pTwainMsg = 0;
  *pbDSevent = false;

  if (m_twState < DS_ENABLED)
  {
    return FAILURE(TWCC_SEQERROR);
  }
     
  TW_EVENT TwainEvent;
  TwainEvent.pEvent  =  pMsg;
  TwainEvent.TWMessage = 0;
  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_EVENT, MSG_PROCESSEVENT, &TwainEvent);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }
  if (dwRes == TWRC_NOTDSEVENT)
  {
    return TWRC_SUCCESS;
  }
  else if (dwRes == TWRC_FAILURE || dwRes == TWRC_CHECKSTATUS)
  {
    return GetDSConditionCode((WORD)dwRes);
  }

  *pTwainMsg = TwainEvent.TWMessage;

  if (TwainEvent.TWMessage == MSG_XFERREADY)
  {
    m_twState = TRANSFER_READY;
  }
  *pbDSevent = true;
  return TWRC_SUCCESS;
}

DWORD CTWAIN_API::GetDSConditionCode(TW_UINT16 err)
{
  DWORD dwRes;
  TW_STATUS stat;

  if (err != TWRC_FAILURE && err != TWRC_CHECKSTATUS)
  {
    return err;
  }

  if (m_twState == PRE_SESSION || m_twState == DSM_LOADED || m_twState == DSM_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }    

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_STATUS, MSG_GET, &stat);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes != TWCC_SUCCESS)
  {
    return dwRes;
  }      

  if (stat.ConditionCode == TWCC_SUCCESS)
  {
    return TWCC_SUCCESS;
  }

  return MAKELONG(err,stat.ConditionCode);
}

DWORD CTWAIN_API::GetDSMConditionCode(TW_UINT16 err)
{
  DWORD dwRes;
  TW_STATUS stat;

  if (m_twState == PRE_SESSION || m_twState == DSM_LOADED)
  {
    return FAILURE(TWCC_SEQERROR);
  }    

  if (err != TWRC_FAILURE && err != TWRC_CHECKSTATUS)
  {
    return err;
  }

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, NULL, DG_CONTROL, DAT_STATUS, MSG_GET, &stat);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes != TWCC_SUCCESS)
  {
    return dwRes;
  }   

  if (stat.ConditionCode == TWCC_SUCCESS)
  {
    return TWCC_SUCCESS;
  }

  return MAKELONG(err,stat.ConditionCode);
}

DWORD CTWAIN_API::QueryCapSupport(WORD wCapID, DWORD *pdwSupport)
{
  DWORD dwRes;
  TW_CAPABILITY tw_cap;

  *pdwSupport = 0;

  if (m_twState < DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  tw_cap.Cap = wCapID;
  tw_cap.ConType = TWON_ONEVALUE;
  tw_cap.hContainer = NULL;

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_CAPABILITY, MSG_QUERYSUPPORT, &tw_cap);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes != TWCC_SUCCESS)
  {
    if (tw_cap.hContainer != NULL)
    {
      DSM_Free(tw_cap.hContainer);
    }
    return dwRes;
  }

  if (tw_cap.hContainer == NULL)
  {
    return FAILURE(TWCC_BADVALUE);
  }

  TW_ONEVALUE * pOneValue;
  pOneValue = (TW_ONEVALUE*)DSM_LockMemory(tw_cap.hContainer);

  if(pOneValue)
  {
    *pdwSupport = (DWORD)pOneValue->Item;
    DSM_UnlockMemory(tw_cap.hContainer);
  }
  else
  {
    dwRes = FAILURE(TWCC_LOWMEMORY);
  }

  DSM_Free(tw_cap.hContainer);
  return dwRes;
}

DWORD CTWAIN_API::GetCapCurrentValue(WORD wCapID, LL *pllValue, WORD *pwType)
{
  DWORD dwRes;

  if (m_twState < DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  TW_CAPABILITY tw_cap;

  tw_cap.Cap = wCapID;
  tw_cap.ConType = TWON_DONTCARE16;
  tw_cap.hContainer = NULL;

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_CAPABILITY, MSG_GETCURRENT, &tw_cap);
    dwRes = GetDSConditionCode((WORD)dwRes);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes != TWCC_SUCCESS)
  {
    if (tw_cap.hContainer != NULL)
    {
      DSM_Free(tw_cap.hContainer);
    }
    return dwRes;
  }
  if (tw_cap.hContainer == NULL)
  {
    return FAILURE(TWCC_BUMMER);
  }

  TW_ONEVALUE *pOneValue = (TW_ONEVALUE *)DSM_LockMemory(tw_cap.hContainer);
  if (pOneValue == NULL)
  {
    DSM_Free(tw_cap.hContainer);
    return FAILURE(TWCC_LOWMEMORY);
  }
  DWORD dwValue=0;
  TW_UINT16  ItemType = pOneValue->ItemType;
  if(pwType)
  {
    *pwType = ItemType;
  }
  switch (tw_cap.ConType)
  {
    case TWON_ONEVALUE:
      dwValue = pOneValue->Item;
      break;

    case TWON_ARRAY:
      {
        TW_ARRAY *pArrayValue = (TW_ARRAY *)pOneValue;
        if(pArrayValue->NumItems==0)
        {
          dwRes = FAILURE(TWCC_BADVALUE);
        }
        else
        {
          switch (pArrayValue->ItemType)
          {
            case TWTY_INT8:
            case TWTY_UINT8:
              dwValue = pArrayValue->ItemList[0];
              break;
            case TWTY_INT16:
            case TWTY_UINT16:
            case TWTY_BOOL:
              dwValue = *((WORD*)pArrayValue->ItemList);
              break;
            case TWTY_INT32:
            case TWTY_UINT32:
            case TWTY_FIX32:
              dwValue = *((DWORD*)pArrayValue->ItemList);
              break;
            default:
              dwRes = FAILURE(TWCC_BADVALUE);
              break;
          }
        }
        break;
      }
    case TWON_RANGE:
      {
        TW_RANGE *pRangeValue = (TW_RANGE *)pOneValue;
        dwValue = pRangeValue->CurrentValue;
        break;
      }
    case TWON_ENUMERATION:
      {
        TW_ENUMERATION *pEnumValue = (TW_ENUMERATION *)pOneValue;
        if(pEnumValue->NumItems==0 || pEnumValue->CurrentIndex>=pEnumValue->NumItems)
        {
          dwRes = FAILURE(TWCC_BADVALUE);
        }
        else
        {
          switch (pEnumValue->ItemType)
          {
            case TWTY_INT8:
            case TWTY_UINT8:
              dwValue = pEnumValue->ItemList[pEnumValue->CurrentIndex];
              break;
            case TWTY_INT16:
            case TWTY_UINT16:
            case TWTY_BOOL:
              dwValue = *((WORD*)(pEnumValue->ItemList+pEnumValue->CurrentIndex*2));
              break;
            case TWTY_INT32:
            case TWTY_UINT32:
            case TWTY_FIX32:
              dwValue = *((DWORD*)(pEnumValue->ItemList+pEnumValue->CurrentIndex*4));
              break;
            default:
              dwRes = FAILURE(TWCC_BADVALUE);
              break;
          }
        }
      }
    default:
      dwRes = FAILURE(TWCC_BADVALUE);
      break;
  }
  DSM_UnlockMemory(tw_cap.hContainer);
  DSM_Free(tw_cap.hContainer);

  switch (ItemType)
  {
    case TWTY_INT8:
      *pllValue = (char)dwValue;
      break;
    case TWTY_INT16:
      *pllValue = (short)dwValue;
      break;
    case TWTY_INT32:
      *pllValue = (long)dwValue;
      break;
    case TWTY_UINT8:
      *pllValue = (BYTE)dwValue;
      break;
    case TWTY_UINT16:
      *pllValue = (WORD)dwValue;
      break;
    case TWTY_UINT32:
      *pllValue = (DWORD)dwValue;
      break;
    case TWTY_BOOL:
      *pllValue = (dwValue != 0);
      break;
    case TWTY_FIX32:
      *pllValue = *((TW_FIX32*)(&dwValue));
      break;
    default:
      dwRes = FAILURE(TWCC_BADVALUE);
      break;
  }
  if(dwRes==0 && ItemType==TWTY_FIX32)
  {
    for(int i=0; g_unUnitDepCap[i]; i++)
    {
      if(g_unUnitDepCap[i]==wCapID)
      {
        *pllValue = ConvertToInch(*pllValue);
        return dwRes;
      }
    }
    for(int i=0; g_un1_UnitDepCap[i]; i++)
    {
      if(g_un1_UnitDepCap[i]==wCapID)
      {
        *pllValue = ConvertFromInch(*pllValue);
        return dwRes;
      }
    }
  }
  return dwRes;
}


DWORD CTWAIN_API::GetCapDefaultValue(WORD wCapID, LL *pllValue, WORD *pwType)
{
  DWORD dwRes;

  if (m_twState < DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  TW_CAPABILITY tw_cap;

  tw_cap.Cap = wCapID;
  tw_cap.ConType = TWON_DONTCARE16;
  tw_cap.hContainer = NULL;

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_CAPABILITY, MSG_GETDEFAULT, &tw_cap);
    dwRes = GetDSConditionCode((WORD)dwRes);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes != TWCC_SUCCESS)
  {
    if (tw_cap.hContainer != NULL)
    {
      DSM_Free(tw_cap.hContainer);
    }
    return dwRes;
  }
  if (tw_cap.hContainer == NULL)
  {
    return FAILURE(TWCC_BUMMER);
  }

  TW_ONEVALUE *pOneValue = (TW_ONEVALUE *)DSM_LockMemory(tw_cap.hContainer);
  if (pOneValue == NULL)
  {
    DSM_Free(tw_cap.hContainer);
    return FAILURE(TWCC_LOWMEMORY);
  }
  DWORD dwValue=0;
  TW_UINT16  ItemType = pOneValue->ItemType;
  if(pwType)
  {
    *pwType = ItemType;
  }
  switch (tw_cap.ConType)
  {
    case TWON_ONEVALUE:
      dwValue = pOneValue->Item;
      break;

    case TWON_ARRAY:
      {
        TW_ARRAY *pArrayValue = (TW_ARRAY *)pOneValue;
        if(pArrayValue->NumItems==0)
        {
          dwRes = FAILURE(TWCC_BADVALUE);
        }
        else
        {
          switch (pArrayValue->ItemType)
          {
            case TWTY_INT8:
            case TWTY_UINT8:
              dwValue = pArrayValue->ItemList[0];
              break;
            case TWTY_INT16:
            case TWTY_UINT16:
            case TWTY_BOOL:
              dwValue = *((WORD*)pArrayValue->ItemList);
              break;
            case TWTY_INT32:
            case TWTY_UINT32:
            case TWTY_FIX32:
              dwValue = *((DWORD*)pArrayValue->ItemList);
              break;
            default:
              dwRes = FAILURE(TWCC_BADVALUE);
              break;
          }
        }
        break;
      }
    case TWON_RANGE:
      {
        TW_RANGE *pRangeValue = (TW_RANGE *)pOneValue;
        dwValue = pRangeValue->DefaultValue;
        break;
      }
    case TWON_ENUMERATION:
      {
        TW_ENUMERATION *pEnumValue = (TW_ENUMERATION *)pOneValue;
        if(pEnumValue->NumItems==0 || pEnumValue->DefaultIndex>=pEnumValue->NumItems)
        {
          dwRes = FAILURE(TWCC_BADVALUE);
        }
        else
        {
          switch (pEnumValue->ItemType)
          {
            case TWTY_INT8:
            case TWTY_UINT8:
              dwValue = pEnumValue->ItemList[pEnumValue->DefaultIndex];
              break;
            case TWTY_INT16:
            case TWTY_UINT16:
            case TWTY_BOOL:
              dwValue = *((WORD*)(pEnumValue->ItemList+pEnumValue->DefaultIndex*2));
              break;
            case TWTY_INT32:
            case TWTY_UINT32:
            case TWTY_FIX32:
              dwValue = *((DWORD*)(pEnumValue->ItemList+pEnumValue->DefaultIndex*4));
              break;
            default:
              dwRes = FAILURE(TWCC_BADVALUE);
              break;
          }
        }
      }
    default:
      dwRes = FAILURE(TWCC_BADVALUE);
      break;
  }

  DSM_UnlockMemory(tw_cap.hContainer);
  DSM_Free(tw_cap.hContainer);

  switch (ItemType)
  {
    case TWTY_INT8:
      *pllValue = (char)dwValue;
      break;
    case TWTY_INT16:
      *pllValue = (short)dwValue;
      break;
    case TWTY_INT32:
      *pllValue = (long)dwValue;
      break;
    case TWTY_UINT8:
      *pllValue = (BYTE)dwValue;
      break;
    case TWTY_UINT16:
      *pllValue = (WORD)dwValue;
      break;
    case TWTY_UINT32:
      *pllValue = (DWORD)dwValue;
      break;
    case TWTY_BOOL:
      *pllValue = (dwValue != 0);
      break;
    case TWTY_FIX32:
      *pllValue = *((TW_FIX32*)(&dwValue));
      break;
    default:
      dwRes = FAILURE(TWCC_BADVALUE);
      break;
  }
  if(dwRes==0 && ItemType==TWTY_FIX32)
  {
    for(int i=0; g_unUnitDepCap[i]; i++)
    {
      if(g_unUnitDepCap[i]==wCapID)
      {
        *pllValue = ConvertToInch(*pllValue);
        return TWCC_SUCCESS;
      }
    }
    for(int i=0; g_un1_UnitDepCap[i]; i++)
    {
      if(g_un1_UnitDepCap[i]==wCapID)
      {
        *pllValue = ConvertFromInch(*pllValue);
        return TWCC_SUCCESS;
      }
    }
  }
  return TWCC_SUCCESS;
}

DWORD CTWAIN_API::GetCapMinMaxValues(WORD wCapID, LL *pllMin, LL *pllMax, LL *pllStep, WORD *pwType)
{
  DWORD dwRes;

  if (m_twState < DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  TW_CAPABILITY tw_cap;

  tw_cap.Cap = wCapID;
  tw_cap.ConType = TWON_DONTCARE16;
  tw_cap.hContainer = NULL;

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_CAPABILITY, MSG_GET, &tw_cap);
    dwRes = GetDSConditionCode((WORD)dwRes);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes != TWCC_SUCCESS)
  {
    if (tw_cap.hContainer != NULL)
    {
      DSM_Free(tw_cap.hContainer);
    }
    return dwRes;
  }
  if (tw_cap.hContainer == NULL)
  {
    return FAILURE(TWCC_BUMMER);
  }

  TW_ONEVALUE *pOneValue = (TW_ONEVALUE *)DSM_LockMemory(tw_cap.hContainer);
  if(pOneValue==NULL)
  {
    DSM_Free(tw_cap.hContainer);
    return FAILURE(TWCC_LOWMEMORY);
  }
  TW_UINT16  ItemType = pOneValue->ItemType;
  if(pwType)
  {
    *pwType = ItemType;
  }
  LL llValue;
  *pllMin = (long long)0x7FFFFFFFFFFFFFFF;
  *pllMax =(long long) 0x8000000000000000;
  *pllStep=0L;
  switch (tw_cap.ConType)
  {
    case TWON_ONEVALUE:
      switch (ItemType)
      {
        case TWTY_INT8:
          llValue = (char)pOneValue->Item;
          break;
        case TWTY_UINT8:
          llValue = (BYTE)pOneValue->Item;
          break;
        case TWTY_INT16:
          llValue = (short)pOneValue->Item;
          break;
        case TWTY_UINT16:
          llValue = (WORD)pOneValue->Item;
          break;
        case TWTY_BOOL:
          llValue = ((pOneValue->Item&0xFFFF) != 0);
          break;
        case TWTY_INT32:
          llValue = (long)pOneValue->Item;
          break;
        case TWTY_UINT32:
          llValue = (DWORD)pOneValue->Item;
          break;
        case TWTY_FIX32:
          llValue = *((TW_FIX32*)(&pOneValue->Item));
          break;
        default:
          dwRes = FAILURE(TWCC_BADVALUE);
          break;
      }
      if(dwRes == TWCC_SUCCESS)
      {
        *pllMin=*pllMax = llValue;
      }
      break;
    case TWON_ARRAY:
      {
        TW_ARRAY *pArrayValue = (TW_ARRAY *)pOneValue;
        if(pArrayValue->NumItems==0)
        {
          dwRes = FAILURE(TWCC_BADVALUE);
        }
        else if(pArrayValue->NumItems==0xFFFFFFFF)
        {
          dwRes = FAILURE(TWCC_LOWMEMORY);
        }
        else
        {        

          for(DWORD i=0; i<pArrayValue->NumItems; i++)
          {
            switch (pArrayValue->ItemType)
            {
              case TWTY_INT8:
                llValue = (char)pArrayValue->ItemList[i];
                break;
              case TWTY_UINT8:
                llValue = pArrayValue->ItemList[i];
                break;
              case TWTY_INT16:
                llValue = *((short*)(pArrayValue->ItemList + i*2));
                break;
              case TWTY_UINT16:
                llValue = *((WORD*)(pArrayValue->ItemList + i*2));
                break;
              case TWTY_BOOL:
                llValue = (*((WORD*)(pArrayValue->ItemList + i*2)))!= 0;
                break;
              case TWTY_INT32:
                llValue = *((long*)(pArrayValue->ItemList + i*4));
                break;
              case TWTY_UINT32:
                llValue = *((DWORD*)(pArrayValue->ItemList + i*4));
                break;
              case TWTY_FIX32:
                llValue = *((TW_FIX32*)(pArrayValue->ItemList + i*4));
                break;
              default:
                dwRes = FAILURE(TWCC_BADVALUE);
                break;
            }
            if(dwRes == TWCC_SUCCESS)
            {
              if(llValue > *pllMax)
              {
                *pllMax = llValue;
              }
              if(llValue < *pllMin)
              {
                *pllMin = llValue;
              }
            }
          }
        }
        break;
      }
    case TWON_RANGE:
      {
        TW_RANGE *pRangeValue = (TW_RANGE *)pOneValue;
        LL llMin;
        LL llMax;
        LL llStep;;
        switch (pRangeValue->ItemType)
        {
          case TWTY_INT8:
            llMin  = (char)pRangeValue->MinValue;
            llMax  = (char)pRangeValue->MaxValue;
            llStep = (char)pRangeValue->StepSize;
            break;
          case TWTY_UINT8:
            llMin  = (BYTE)pRangeValue->MinValue;
            llMax  = (BYTE)pRangeValue->MaxValue;
            llStep = (BYTE)pRangeValue->StepSize;
            break;
          case TWTY_INT16:
            llMin  = (short)pRangeValue->MinValue;
            llMax  = (short)pRangeValue->MaxValue;
            llStep = (short)pRangeValue->StepSize;
            break;
          case TWTY_UINT16:
            llMin  = (WORD)pRangeValue->MinValue;
            llMax  = (WORD)pRangeValue->MaxValue;
            llStep = (WORD)pRangeValue->StepSize;
            break;
          case TWTY_BOOL:
            llMin  = ((pRangeValue->MinValue  &0xFFFF) != 0);
            llMax  = ((pRangeValue->MaxValue  &0xFFFF) != 0);
            llStep = 0L;
            break;
          case TWTY_INT32:
            llMin  = (long)pRangeValue->MinValue;
            llMax  = (long)pRangeValue->MaxValue;
            llStep = (long)pRangeValue->StepSize;
            break;
          case TWTY_UINT32:
            llMin  = (DWORD)pRangeValue->MinValue;
            llMax  = (DWORD)pRangeValue->MaxValue;
            llStep = (DWORD)pRangeValue->StepSize;
            break;
          case TWTY_FIX32:
            llMin  = *((TW_FIX32*)(&pRangeValue->MinValue));
            llMax  = *((TW_FIX32*)(&pRangeValue->MaxValue));
            llStep = *((TW_FIX32*)(&pRangeValue->StepSize));
            break;
          default:
            dwRes = FAILURE(TWCC_BADVALUE);
            break;
        }
        if(dwRes == TWCC_SUCCESS)
        {
          if(llMin > llMax)
          {
            dwRes = FAILURE(TWCC_BADVALUE);
          }
          else
          {
            *pllMin = llMin;
            *pllMax = llMax;
            *pllStep = llStep;
          }
        }
        break;
      }
    case TWON_ENUMERATION:
      {
        TW_ENUMERATION *pEnumValue = (TW_ENUMERATION *)pOneValue;
        if(pEnumValue->NumItems==0)
        {
          dwRes = FAILURE(TWCC_BADVALUE);
        }
        else if(pEnumValue->NumItems==0xFFFFFFFF)
        {
          dwRes = FAILURE(TWCC_LOWMEMORY);
        }
        else
        {
          for(DWORD i=0; i<pEnumValue->NumItems; i++)
          {
            switch (pEnumValue->ItemType)
            {
              case TWTY_INT8:
                llValue = (char)pEnumValue->ItemList[i];
                break;
              case TWTY_UINT8:
                llValue = pEnumValue->ItemList[i];
                break;
              case TWTY_INT16:
                llValue = *((short*)(pEnumValue->ItemList + i*2));
                break;
              case TWTY_UINT16:
                llValue = *((WORD*)(pEnumValue->ItemList + i*2));
                break;
              case TWTY_BOOL:
                llValue = (*((WORD*)(pEnumValue->ItemList + i*2)))!= 0;
                break;
              case TWTY_INT32:
                llValue = *((long*)(pEnumValue->ItemList + i*4));
                break;
              case TWTY_UINT32:
                llValue = *((DWORD*)(pEnumValue->ItemList + i*4));
                break;
              case TWTY_FIX32:
                llValue = *((TW_FIX32*)(pEnumValue->ItemList + i*4));
                break;
              default:
                dwRes = FAILURE(TWCC_BADVALUE);
                break;
            }
            if(dwRes == TWCC_SUCCESS)
            {
              if(llValue > *pllMax)
              {
                *pllMax = llValue;
              }
              if(llValue < *pllMin)
              {
                *pllMin = llValue;
              }
            }
          }
        }
        break;
      }      
    default:
      dwRes = FAILURE(TWCC_BADVALUE);
      break;
  }
  DSM_UnlockMemory(tw_cap.hContainer);
  DSM_Free(tw_cap.hContainer);

  if(dwRes==0 && ItemType==TWTY_FIX32)
  {
    for(int i=0; g_unUnitDepCap[i]; i++)
    {
      if(g_unUnitDepCap[i]==wCapID)
      {
        *pllMax = ConvertToInch(*pllMax);
        *pllMin = ConvertToInch(*pllMin);
        return dwRes;
      }
    }
    for(int i=0; g_un1_UnitDepCap[i]; i++)
    {
      if(g_un1_UnitDepCap[i]==wCapID)
      {
        *pllMax = ConvertFromInch(*pllMax);
        *pllMin = ConvertFromInch(*pllMin);
        return dwRes;
      }
    }

  }
  return dwRes;
}

DWORD CTWAIN_API::GetXFERCOUNTConstrainedValues(LL_ARRAY *pllList)
{
  pllList->clear();
  LL llValOrg;
  WORD wType;
  DWORD dwRes = GetCapCurrentValue(CAP_XFERCOUNT,&llValOrg,&wType);
  if(dwRes==0)
  {
    if(wType!=TWTY_INT16)
    {
      return FAILURE(TWCC_BUMMER);
    }
  }

  dwRes = SetCapability(CAP_XFERCOUNT,-1,TWTY_INT16);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return dwRes;
  }
  LL llVal;
  dwRes = GetCapCurrentValue(CAP_XFERCOUNT,&llVal,&wType);
  if(dwRes==0 && wType==TWTY_INT16 && llVal == -1)
  {
    pllList->push_back(LL(-1L));
  }
  dwRes = SetCapability(CAP_XFERCOUNT,1,TWTY_INT16);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return dwRes;
  }

  dwRes = GetCapCurrentValue(CAP_XFERCOUNT,&llVal,&wType);
  if(dwRes==0 && wType==TWTY_INT16 && llVal == 1)
  {
    pllList->push_back(LL(1L));
  }

  dwRes = SetCapability(CAP_XFERCOUNT,2,TWTY_INT16);
  if(dwRes==FAILURE(TWCC_BUMMER))
  {
    return dwRes;
  }

  dwRes = GetCapCurrentValue(CAP_XFERCOUNT,&llVal,&wType);
  if(dwRes==0 && wType==TWTY_INT16 && llVal == 2)
  {
    pllList->push_back(LL(2L));
  }

  dwRes = SetCapability(CAP_XFERCOUNT,llValOrg,TWTY_INT16);

  return dwRes;
}

DWORD CTWAIN_API::GetCapConstrainedValues(WORD wCapID, LL_ARRAY *pllList, WORD *pwType)
{
  DWORD dwRes;

  if (m_twState < DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  pllList->clear();
  if(wCapID==CAP_XFERCOUNT)
  {
    if(pwType)
    {
      *pwType = TWTY_INT16;
    }
    return GetXFERCOUNTConstrainedValues(pllList);
  }

  TW_CAPABILITY tw_cap;

  tw_cap.Cap = wCapID;
  tw_cap.ConType = TWON_DONTCARE16;
  tw_cap.hContainer = NULL;

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_CAPABILITY, MSG_GET, &tw_cap);
    dwRes = GetDSConditionCode((WORD)dwRes);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes != TWCC_SUCCESS)
  {
    if (tw_cap.hContainer != NULL)
    {
      DSM_Free(tw_cap.hContainer);
    }
    return dwRes;
  }
  if (tw_cap.hContainer == NULL)
  {
    return FAILURE(TWCC_BUMMER);
  }

  TW_ONEVALUE *pOneValue = (TW_ONEVALUE *)DSM_LockMemory(tw_cap.hContainer);
  if(pOneValue==NULL)
  {
    DSM_Free(tw_cap.hContainer);
    return FAILURE(TWCC_LOWMEMORY);
  }
  TW_UINT16  ItemType = pOneValue->ItemType;
  if(pwType)
  {
    *pwType = ItemType;
  }

  LL llValue;
  switch (tw_cap.ConType)
  {
    case TWON_ONEVALUE:
      switch (ItemType)
      {
        case TWTY_INT8:
          llValue = (char)(pOneValue->Item&0xFF);
          break;
        case TWTY_UINT8:
          llValue = (BYTE)(pOneValue->Item&0xFF);
          break;
        case TWTY_INT16:
          llValue = (short)(pOneValue->Item&0xFFFF);
          break;
        case TWTY_UINT16:
          llValue = (WORD)(pOneValue->Item&0xFFFF);
          break;
        case TWTY_BOOL:
          llValue = ((pOneValue->Item&0xFFFF) != 0);
          break;
        case TWTY_INT32:
          llValue = (long)(pOneValue->Item);
          break;
        case TWTY_UINT32:
          llValue = (DWORD)(pOneValue->Item);
          break;
        case TWTY_FIX32:
          llValue = *((TW_FIX32*)(&pOneValue->Item));
          break;
        default:
          dwRes = FAILURE(TWCC_BADVALUE);
          break;
      }
      if(dwRes == TWCC_SUCCESS)
      {
        pllList->push_back(llValue);
      }
      break;
    case TWON_ARRAY:
      {
        TW_ARRAY *pArrayValue = (TW_ARRAY *)pOneValue;
        if(pArrayValue->NumItems==0)
        {
          dwRes = FAILURE(TWCC_BADVALUE);
        }
        else if(pArrayValue->NumItems==0xFFFFFFFF)
        {
          dwRes = FAILURE(TWCC_LOWMEMORY);
        }
        else
        {
          for(DWORD i=0; i<pArrayValue->NumItems; i++)
          {
            switch (pArrayValue->ItemType)
            {
              case TWTY_INT8:
                llValue = (char)pArrayValue->ItemList[i];
                break;
              case TWTY_UINT8:
                llValue = pArrayValue->ItemList[i];
                break;
              case TWTY_INT16:
                llValue = *((short*)(pArrayValue->ItemList + i*2));
                break;
              case TWTY_UINT16:
                llValue = *((WORD*)(pArrayValue->ItemList + i*2));
                break;
              case TWTY_BOOL:
                llValue = (*((WORD*)(pArrayValue->ItemList + i*2)))!= 0;
                break;
              case TWTY_INT32:
                llValue = *((long*)(pArrayValue->ItemList + i*4));
                break;
              case TWTY_UINT32:
                llValue = *((DWORD*)(pArrayValue->ItemList + i*4));
                break;
              case TWTY_FIX32:
                llValue = *((TW_FIX32*)(pArrayValue->ItemList + i*4));
                break;
              default:
                dwRes = FAILURE(TWCC_BADVALUE);
                break;
            }
            if(dwRes == TWCC_SUCCESS)
            {
              try
              {
                pllList->push_back(llValue);
              }
              catch(...)
              {
                pllList->clear();
                dwRes = FAILURE(TWCC_LOWMEMORY);
              }
            }
          }
        }
        break;
      }
    case TWON_RANGE:
      {
        TW_RANGE *pRangeValue = (TW_RANGE *)pOneValue;
        LL llCur;
        LL llStep;
        LL llMin;
        LL llMax;

        switch (pRangeValue->ItemType)
        {
          case TWTY_INT8:
            llCur  = (char)pRangeValue->CurrentValue;
            llStep = (char)pRangeValue->StepSize;
            llMin  = (char)pRangeValue->MinValue;
            llMax  = (char)pRangeValue->MaxValue;
            break;
          case TWTY_UINT8:
            llCur  = (BYTE)pRangeValue->CurrentValue;
            llStep = (BYTE)pRangeValue->StepSize;
            llMin  = (BYTE)pRangeValue->MinValue;
            llMax  = (BYTE)pRangeValue->MaxValue;
            break;
          case TWTY_INT16:
            llCur  = (short)pRangeValue->CurrentValue;
            llStep = (short)pRangeValue->StepSize;
            llMin  = (short)pRangeValue->MinValue;
            llMax  = (short)pRangeValue->MaxValue;
            break;
          case TWTY_UINT16:
            llCur  = (WORD)pRangeValue->CurrentValue;
            llStep = (WORD)pRangeValue->StepSize;
            llMin  = (WORD)pRangeValue->MinValue;
            llMax  = (WORD)pRangeValue->MaxValue;
            break;
          case TWTY_BOOL:
            llCur  = ((pRangeValue->CurrentValue&0xFFFF) != 0);
            llStep = ((pRangeValue->StepSize&0xFFFF) != 0);
            llMin  = ((pRangeValue->MinValue&0xFFFF) != 0);
            llMax  = ((pRangeValue->MaxValue&0xFFFF) != 0);
            break;
          case TWTY_INT32:
            llCur  = (long)pRangeValue->CurrentValue;
            llStep = (long)pRangeValue->StepSize;
            llMin  = (long)pRangeValue->MinValue;
            llMax  = (long)pRangeValue->MaxValue;
            break;
          case TWTY_UINT32:
            llCur  = (WORD)pRangeValue->CurrentValue;
            llStep = (WORD)pRangeValue->StepSize;
            llMin  = (WORD)pRangeValue->MinValue;
            llMax  = (WORD)pRangeValue->MaxValue;
            break;
          case TWTY_FIX32:
            llCur  = *((TW_FIX32*)(&pRangeValue->CurrentValue));
            llStep = *((TW_FIX32*)(&pRangeValue->StepSize));
            llMin  = *((TW_FIX32*)(&pRangeValue->MinValue));
            llMax  = *((TW_FIX32*)(&pRangeValue->MaxValue));
            break;
          default:
            dwRes = FAILURE(TWCC_BADVALUE);
            break;
        }
        if(dwRes == TWCC_SUCCESS)
        {
          if(llCur<llMin || llCur>llMax || llStep<=0)
          {
            dwRes = FAILURE(TWCC_BADVALUE);
          }
          else
          {
            llMin  = llCur-((llCur-llMin)/llStep)*llStep;
            llMax  = llCur+((llMax-llCur)/llStep)*llStep;
            for(llCur=llMin;llCur<=llMax;llCur+= llStep)
            {
              try
              {
                pllList->push_back(llCur);
              }
              catch(...)
              {
                pllList->clear();
                dwRes = FAILURE(TWCC_LOWMEMORY);
              }
            }
          }
        }
        break;
      }
    case TWON_ENUMERATION:
      {
        TW_ENUMERATION *pEnumValue = (TW_ENUMERATION *)pOneValue;
        if(pEnumValue->NumItems==0)
        {
          dwRes = FAILURE(TWCC_BADVALUE);
        }
        else if(pEnumValue->NumItems==0xFFFFFFFF)
        {
          dwRes = FAILURE(TWCC_LOWMEMORY);
        }
        else
        {
          for(DWORD i=0; i<pEnumValue->NumItems; i++)
          {
            switch (pEnumValue->ItemType)
            {
              case TWTY_INT8:
                llValue = (char)pEnumValue->ItemList[i];
                break;
              case TWTY_UINT8:
                llValue = pEnumValue->ItemList[i];
                break;
              case TWTY_INT16:
                llValue = *((short*)(pEnumValue->ItemList + i*2));
                break;
              case TWTY_UINT16:
                llValue = *((WORD*)(pEnumValue->ItemList + i*2));
                break;
              case TWTY_BOOL:
                llValue = (*((WORD*)(pEnumValue->ItemList + i*2)))!= 0;
                break;
              case TWTY_INT32:
                llValue = *((long*)(pEnumValue->ItemList + i*4));
                break;
              case TWTY_UINT32:
                llValue = *((DWORD*)(pEnumValue->ItemList + i*4));
                break;
              case TWTY_FIX32:
                llValue = *((TW_FIX32*)(pEnumValue->ItemList + i*4));
                break;
              default:
                dwRes = FAILURE(TWCC_BADVALUE);
                break;
            }
            if(dwRes == TWCC_SUCCESS)
            {
              try
              {
                pllList->push_back(llValue);
              }
              catch(...)
              {
                pllList->clear();
                dwRes = FAILURE(TWCC_LOWMEMORY);
              }
            }
          }
        }
        break;
      }      
    default:
      dwRes = FAILURE(TWCC_BADVALUE);
      break;
  }
  DSM_UnlockMemory(tw_cap.hContainer);
  DSM_Free(tw_cap.hContainer);
  if(dwRes==0 && ItemType==TWTY_FIX32)
  {
    for(int i=0; g_unUnitDepCap[i]; i++)
    {
      if(g_unUnitDepCap[i]==wCapID)
      {
        LL_ARRAY::iterator llIter=pllList->begin();
        for(;llIter!=pllList->end();llIter++)
        {
          *llIter = ConvertToInch(*llIter);
        }
        return dwRes;
      }
    }
    for(int i=0; g_un1_UnitDepCap[i]; i++)
    {
      if(g_un1_UnitDepCap[i]==wCapID)
      {
        LL_ARRAY::iterator llIter=pllList->begin();
        for(;llIter!=pllList->end();llIter++)
        {
          *llIter = ConvertFromInch(*llIter);
        }
        return dwRes;
      }
    }
  }
  return dwRes;
}

DWORD CTWAIN_API::SetCapability(WORD wCapID, LL llValue, WORD wType)
{
  if (m_twState != DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }
  
  DWORD dwRes = TWRC_SUCCESS;
  if(wType==TWTY_FIX32)
  {
    for(int i=0; g_unUnitDepCap[i]; i++)
    {
      if(g_unUnitDepCap[i]==wCapID)
      {
        llValue = ConvertFromInch(llValue);
        break;
      }
    }
    for(int i=0; g_un1_UnitDepCap[i]; i++)
    {
      if(g_un1_UnitDepCap[i]==wCapID)
      {
        llValue = ConvertToInch(llValue);
        break;
      }
    }
  }
  TW_CAPABILITY tw_cap;
  tw_cap.Cap = wCapID;
  tw_cap.ConType = TWON_ONEVALUE;
  tw_cap.hContainer = DSM_Alloc(sizeof(TW_ONEVALUE));
  if(tw_cap.hContainer==NULL)
  {
    return FAILURE(TWCC_LOWMEMORY);
  }

  TW_ONEVALUE *cap = (TW_ONEVALUE *) DSM_LockMemory(tw_cap.hContainer);

  if(cap==NULL)
  {
    DSM_Free(tw_cap.hContainer);
    return FAILURE(TWCC_LOWMEMORY);
  }  
  
  switch (wType)
  {
    case TWTY_INT8:
      cap->Item = (char)llValue;
      break;
    case TWTY_UINT8:
      cap->Item = (BYTE)llValue;
      break;
    case TWTY_INT16:
      cap->Item = (short)llValue;
      break;
    case TWTY_UINT16:
      cap->Item = (WORD)llValue;
      break;
    case TWTY_BOOL:
      cap->Item = ((bool)llValue)?1:0;
      break;
    case TWTY_INT32:
      cap->Item = (long)llValue;
      break;
    case TWTY_UINT32:
      cap->Item = (DWORD)llValue;
      break;
    case TWTY_FIX32:
      cap->Item = *((TW_UINT32*)(&((TW_FIX32)llValue)));
      break;
    default:
      dwRes = FAILURE(TWCC_BADVALUE);
      break;
  }

  cap->ItemType = wType;

  DSM_UnlockMemory(tw_cap.hContainer);

  if(dwRes == TWRC_SUCCESS)
  {
    try
    {
      dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_CAPABILITY, MSG_SET, &tw_cap);
      dwRes = GetDSConditionCode((WORD)dwRes);
    }
    catch (...)
    {
      dwRes = FAILURE(TWCC_BUMMER);
    }
  }
  DSM_Free(tw_cap.hContainer);
  
  return dwRes;
}

TW_HANDLE CTWAIN_API::DSM_Alloc(TW_UINT32 dwSize)
{
  if(m_DSM_Entry.DSM_MemAllocate)
  {
    return m_DSM_Entry.DSM_MemAllocate(dwSize);
  }

  return GlobalAlloc(GHND, dwSize);
}

void CTWAIN_API::DSM_Free(TW_HANDLE hMemory)
{
  if(m_DSM_Entry.DSM_MemFree)
  {
    m_DSM_Entry.DSM_MemFree(hMemory);
    return;
  }

  GlobalFree(hMemory);
}

TW_MEMREF CTWAIN_API::DSM_LockMemory(TW_HANDLE hMemory)
{
  if(m_DSM_Entry.DSM_MemLock)
  {
    return m_DSM_Entry.DSM_MemLock(hMemory);
  }

  return (TW_MEMREF)GlobalLock(hMemory);
}

void CTWAIN_API::DSM_UnlockMemory(TW_MEMREF hMemory)
{
  if(m_DSM_Entry.DSM_MemUnlock)
  {
    m_DSM_Entry.DSM_MemUnlock(hMemory);
    return;
  }

  GlobalUnlock(hMemory);
}

DWORD CTWAIN_API::GetBoolCapConstrainedValues(WORD wCapID, LL_ARRAY *pllList)
{
  DWORD dwRes;

  if (m_twState < DS_OPENED)
  {
    return FAILURE(TWCC_SEQERROR);
  }

  pllList->clear();
  TW_CAPABILITY tw_cap;

  tw_cap.Cap = wCapID;
  tw_cap.ConType = TWON_DONTCARE16;
  tw_cap.hContainer = NULL;

  try
  {
    dwRes = (*m_fnDSM_Entry)(&m_AppID, &m_DS, DG_CONTROL, DAT_CAPABILITY, MSG_GET, &tw_cap);
    dwRes = GetDSConditionCode((WORD)dwRes);
  }
  catch (...)
  {
    return FAILURE(TWCC_BUMMER);
  }

  if (dwRes != TWCC_SUCCESS)
  {
    if (tw_cap.hContainer != NULL)
    {
      DSM_Free(tw_cap.hContainer);
    }
    return dwRes;
  }
  if (tw_cap.hContainer == NULL)
  {
    return FAILURE(TWCC_BUMMER);
  }

  TW_ONEVALUE *pOneValue = (TW_ONEVALUE *)DSM_LockMemory(tw_cap.hContainer);
  if(pOneValue==NULL)
  {
    DSM_Free(tw_cap.hContainer);
    return FAILURE(TWCC_LOWMEMORY);
  }
  LL llValue;

  switch (tw_cap.ConType)
  {
    case TWON_ONEVALUE:// TDS <2.0 or TDS>=2.0 but bad implemented
      if(pOneValue->ItemType != TWTY_BOOL)
      {
        dwRes = FAILURE(TWCC_BADVALUE);
      }
      else
      {
        llValue = ((pOneValue->Item&0xFFFF) != 0);
      }

      if(dwRes == TWCC_SUCCESS)
      {
        pllList->push_back(llValue);
        LL llInvVal = llValue?0:1;
        DWORD dwTempRes = SetCapability(wCapID,llInvVal,TWTY_BOOL);//set inverse value
        if(dwTempRes==TWCC_BUMMER)
        {
          dwRes = FAILURE(TWCC_BUMMER);
        }
        else
        {
          WORD wType;
          dwTempRes=GetCapCurrentValue(wCapID,&llInvVal,&wType);//try to get reverse value
          if(dwTempRes==TWCC_BUMMER)
          {
            dwRes = FAILURE(TWCC_BUMMER);
          }
          else if(dwTempRes==TWCC_SUCCESS)
          {
            if(llInvVal!=llValue)
            {
              pllList->push_back(llInvVal);
            }
          }
          dwTempRes = SetCapability(wCapID,llValue,TWTY_BOOL);//restore original value
          if(dwTempRes==TWCC_BUMMER)
          {
            dwRes = FAILURE(TWCC_BUMMER);
          }
        }
      }
      break;
     
    case TWON_ENUMERATION:
      {
        TW_ENUMERATION *pEnumValue = (TW_ENUMERATION *)pOneValue;
        if(pEnumValue->NumItems==0)
        {
          dwRes = FAILURE(TWCC_BADVALUE);
        }
        else if(pEnumValue->NumItems==0xFFFFFFFF)
        {
          dwRes = FAILURE(TWCC_LOWMEMORY);
        }
        else
        {
          if(pEnumValue->ItemType != TWTY_BOOL)
          {
            dwRes = FAILURE(TWCC_BADVALUE);
          }
          else
          {
            for(DWORD i=0; i<pEnumValue->NumItems; i++)
            {
              llValue = (*((WORD*)(pEnumValue->ItemList + i*2)))!=0 ;
              try
              {
                pllList->push_back(llValue);
              }
              catch(...)
              {
                pllList->clear();
                dwRes = FAILURE(TWCC_LOWMEMORY);
              }
            }
          }
        }
        break;
      }      
    default:
      dwRes = FAILURE(TWCC_BADVALUE);
      break;
  }
  DSM_UnlockMemory(tw_cap.hContainer);
  DSM_Free(tw_cap.hContainer);

  return dwRes;
}

LL CTWAIN_API::ConvertToInch(LL llVal)
{
  switch(m_wDSunits)
  {
  case TWUN_INCHES:
    return llVal;
  case TWUN_CENTIMETERS:
    return (llVal*100)/254;
  case TWUN_MILLIMETERS:
    return (llVal*10)/254;
  case TWUN_PICAS:
    return llVal/6;
  case TWUN_POINTS:
    return llVal/72;
  case TWUN_TWIPS:
    return llVal/1440;
  case TWUN_PIXELS:
  default:
    return 0;
  }
  return 0;
}

LL CTWAIN_API::ConvertFromInch(LL llVal)
{
  switch(m_wDSunits)
  {
  case TWUN_INCHES:
    return llVal;
  case TWUN_CENTIMETERS:
    return (llVal*254)/100;
  case TWUN_MILLIMETERS:
    return (llVal*254)/10;
  case TWUN_PICAS:
    return llVal*6;
  case TWUN_POINTS:
    return llVal*72;
  case TWUN_TWIPS:
    return llVal*1440;
  case TWUN_PIXELS:
  default:
    return 0;
  }
  return 0;
}

bool LL_ARRAY::IfExist(LL llVal)
{
  LL_ARRAY::iterator llIter=begin();
  for(;llIter!=end();llIter++)
  {
    if(*llIter==llVal)
    {
      return true;
    }
  }
  return false;
}

TW_UINT16 FAR PASCAL CTWAIN_API::TWAIN_callback(pTW_IDENTITY pOrigin,pTW_IDENTITY pDest,TW_UINT32 DG,
                                     TW_UINT16 DAT,TW_UINT16 MSG,TW_MEMREF pData)
{
  int nID =  (int)pData;
  CTWAIN_API *pTwainApi = g_TWAIN_APIs.find(nID)->second;
  
  pTwainApi->m_msgLast = MSG;
  if (MSG == MSG_XFERREADY)
  {
    pTwainApi->m_twState = TRANSFER_READY;
  }

  if(pTwainApi->m_hDSMEvent)
  {
    SetEvent(pTwainApi->m_hDSMEvent);
  }
  return TWRC_SUCCESS;
}

