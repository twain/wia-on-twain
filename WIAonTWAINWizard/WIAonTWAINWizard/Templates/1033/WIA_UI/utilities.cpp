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
 * @file utilities.cpp
 * Utility functions common for CUIDlg and CPropPage
 * @author TWAIN Working Group
 * @date October 2009
 */

#include "stdafx.h"
#include "TWAIN_API.h"
#include "utilities.h"
#include "ProfName.h"

#define kVER_MAJ [!output PRODUCT_VERSION1]
#define kVER_MIN [!output PRODUCT_VERSION2]
#define kVER_INFO "[!output PRODUCT_VERSION1].[!output PRODUCT_VERSION2] [!output PRODUCT_DESCRIPTION]"
#define kPROT_MAJ 2
#define kPROT_MIN 0
#define kMANUFACTURER "[!output MANUFACTURE_NAME]"
#define kPRODUCT_FAMILY "[!output PRODUCT_DESCRIPTION]"
#define kPRODUCT_NAME "[!output PRODUCT_DESCRIPTION]"

#define kDS_VER_MAJ [!output TWAIN_DS_VER_MAJOR] 
#define kDS_VER_MIN [!output TWAIN_DS_VER_MINOR] 
#define kDS_MANUFACTURER "[!output TWAIN_DS_MANUFACTURE]"
#define kDS_PRODUCT_NAME "[!output TWAIN_DS_PRODUCT]"

map<int,CTWAIN_API*>    g_TWAIN_APIs;

CString g_strProfilesDirectory;
#define PROFILELOCATION L"[!output MANUFACTURE_NAME]\\[!output TWAIN_DS_PRODUCT]\\"
#define FILEEXTENTION L".TWP"

#define CUSTOM_ROOT_PROP_ID WIA_PRIVATE_DEVPROP
#define CUSTOM_ROOT_PROP_ID_STR L"AppID"
#define CUSTOM_ROOT_PROP_ID1 WIA_PRIVATE_DEVPROP+1
#define CUSTOM_ROOT_PROP_ID1_STR L"ProfileName"

const TW_IDENTITY kAppID={0,{kVER_MAJ,kVER_MIN,TWLG_USA,TWCY_USA,kVER_INFO},kPROT_MAJ,kPROT_MIN,DG_CONTROL|DG_IMAGE|DF_APP2,kMANUFACTURER,kPRODUCT_FAMILY,kPRODUCT_NAME};
const TW_IDENTITY kDSID={0,{kDS_VER_MAJ,kDS_VER_MIN,0,0,""},0,0,0,kDS_MANUFACTURER,"",kDS_PRODUCT_NAME};

//WIA functions
HRESULT WIA_ReadPropBSTR( IWiaPropertyStorage* pIPS, PROPID propID, CComBSTR& bstrVal )
{
  PROPVARIANT Props;
  PROPSPEC PropSpecs;

  PropSpecs.ulKind = PRSPEC_PROPID;
  PropSpecs.propid = propID;
  Props.vt = VT_BSTR;
     
  HRESULT hr = pIPS->ReadMultiple( 1, &PropSpecs, &Props );            
              
  if( SUCCEEDED(hr) ) 
  {           
    bstrVal = Props.bstrVal;                  
  }
  
  return hr;    
}

//TWAIN functions
bool TW_Get_Set_DS_Data(HWND hWnd, HGLOBAL *phData, DWORD *pdwDataSize)
{
  CTWAIN_API TwainApi(hWnd);
  //store in global table. UI creates only one instance of CTWAIN_API
  g_TWAIN_APIs[0]= &TwainApi; 
  DWORD dwRes = TwainApi.OpenDSM(kAppID,0);
  if(dwRes==0)
  {
    TW_IDENTITY DSID=kDSID;
    dwRes = TwainApi.OpenDS(&DSID);
    if(dwRes)
    {
      MessageBox(NULL,L"The device is offline!", L"Error", MB_OK | MB_ICONEXCLAMATION | MB_APPLMODAL | MB_SETFOREGROUND);
      TwainApi.CloseDSM();
      return false;
    }
  }

  TW_CUSTOMDSDATA Data;
  Data.hData = NULL;

  if(*pdwDataSize && *phData)//set Custom DS data
  {
    if(dwRes==0)
    {
      BYTE *pSrc =(BYTE *)GlobalLock(*phData);
      if(pSrc==0)
      {
        dwRes = FAILURE(TWCC_LOWMEMORY);
      }
      if(dwRes==0)
      {
        Data.hData = TwainApi.DSM_Alloc(*pdwDataSize);
        Data.InfoLength = *pdwDataSize;
        if(Data.hData==0)
        {
          dwRes = FAILURE(TWCC_LOWMEMORY);
        }
      }
      BYTE *pDest=0;
      if(dwRes==0)
      {
        pDest =(BYTE *) TwainApi.DSM_LockMemory(Data.hData);
        if(pDest==0)
        {
          dwRes = FAILURE(TWCC_LOWMEMORY);
        }
      }
      if(dwRes==0)
      {
        memcpy(pDest,pSrc,*pdwDataSize);
        TwainApi.DSM_UnlockMemory(Data.hData);
      }

      GlobalUnlock(*phData);
      GlobalFree(*phData);
      *phData=NULL;
      *pdwDataSize=0;

      if(dwRes==0)
      {
        dwRes = TwainApi.SetDSCustomData(Data);
      }
      //free allocated resources
      if(Data.hData)
      {
        TwainApi.DSM_Free(Data.hData);
        Data.hData=NULL;
      }
    }   
  }

  HANDLE hTransferEvent;
  if(dwRes==0)
  {
    //show TWAIN UI
    dwRes = TwainApi.EnableDSOnly(&hTransferEvent);
  }

  MSG msg;
  bool bOK=false;

  while(dwRes==0)
  {
    // pump messages and send it to TWAIN DS
    if (::PeekMessage(&msg, 0, 0, 0, PM_REMOVE))
    {
      if (msg.message == WM_QUIT) 
      {
        break;
      }
      WORD wMsgTW;
      bool bTwainMsg;

      if ((dwRes = TwainApi.ProcessEvent(&msg, &wMsgTW, &bTwainMsg)) != TWRC_SUCCESS)
      {
        break;
      }

      if (!bTwainMsg)
      {//Translate and Dispach non TWAIN messages
        ::TranslateMessage(&msg);
        ::DispatchMessage(&msg);
      }
      else
      {
        //checks for signal
        DWORD dwWait = WaitForSingleObject(hTransferEvent,0);
        if(dwWait == WAIT_OBJECT_0)
        {
          switch (TwainApi.GetLastMsg())
          {
            case MSG_CLOSEDSREQ: //Cancel button
              dwRes = TWRC_CANCEL;
              break;
            case MSG_CLOSEDSOK: // OK/Save button
              dwRes = TWRC_CANCEL;
              bOK = true;
              break;
            case MSG_DEVICEEVENT://not expected
              dwRes = TWRC_FAILURE;
              break;
            case MSG_XFERREADY://not expected
              dwRes = TWRC_FAILURE;
              break;
            default:
              break;
          }
        }
        else if(dwWait==WAIT_TIMEOUT)
        {
          continue;
        }
        else
        {
          dwRes = TWRC_FAILURE;
        }
      }
    }
    else
    {
      // 10mS sleep if no messages
      Sleep(10);
    }
  }
  // go to state 4
  TwainApi.DisableDS(); 

  if(dwRes == TWRC_CANCEL)// OK/Save and Cancel
  {
    if(bOK)
    {
      dwRes = 0;
      // Get Custom DS data
      dwRes = TwainApi.GetDSCustomData(&Data);
      if(dwRes==0)
      {
        if(Data.hData==0 || Data.InfoLength==0)
        {
          dwRes = TWRC_FAILURE;
        }
        else
        {
          BYTE *pSrc =(BYTE *) TwainApi.DSM_LockMemory(Data.hData);
          if(pSrc==0)
          {
            dwRes = FAILURE(TWCC_LOWMEMORY);
          }
          if(dwRes==0)
          {
            *phData = GlobalAlloc(GMEM_MOVEABLE,Data.InfoLength);
            *pdwDataSize = Data.InfoLength;
            if(*phData==0)
            {
              dwRes = FAILURE(TWCC_LOWMEMORY);;
            }
          }
          BYTE *pDest=0;

          if(dwRes==0)
          {
            pDest =(BYTE *)GlobalLock(*phData);
            if(pDest==0)
            {
              dwRes = FAILURE(TWCC_LOWMEMORY);
            }
          }
          if(dwRes==0)
          {
            memcpy(pDest,pSrc,Data.InfoLength);
            GlobalUnlock(*phData);
          }
          TwainApi.DSM_UnlockMemory(Data.hData);
          TwainApi.DSM_Free(Data.hData);
          Data.hData= NULL;
        }
        //free allocated resources
        if(Data.hData)
        {
          TwainApi.DSM_Free(Data.hData);
        }
      }
    }
  }
  // release DS amd DSM
  TwainApi.CloseDS();
  TwainApi.CloseDSM();
  if(dwRes)
  {
    MessageBox(NULL,L"Error saving profile", L"Error", MB_OK | MB_ICONEXCLAMATION | MB_APPLMODAL | MB_SETFOREGROUND);
  }
  return dwRes==0;
}

bool TW_Get_FeederEnabled(CComboBox *pcbxProfiles, BOOL *pbFeederEnabled)
{
  CString strProfileName;
  int nIndex = pcbxProfiles->GetCurSel();
  pcbxProfiles->GetWindowText(strProfileName);
  CString strFileName  = g_strProfilesDirectory;
  if(nIndex<0 || strProfileName.IsEmpty())
  {
    return true; // no profiles selected
  }
  
  strFileName += strProfileName;
  strFileName += FILEEXTENTION;

  HGLOBAL hData;
  DWORD dwDataSize;
  // Get Custom DS data from selected profile
  if(!TW_LoadProfileFromFile(strFileName, &hData, &dwDataSize))
  {
    return false;
  }  
   // Open DSM and DS
  CTWAIN_API TwainApi(pcbxProfiles->GetParent()->m_hWnd);
  //store in global table. UI creates only one instance of CTWAIN_API
  g_TWAIN_APIs[0]= &TwainApi; 
  DWORD dwRes = TwainApi.OpenDSM(kAppID,0);
  if(dwRes==0)
  {
    TW_IDENTITY DSID=kDSID;
    dwRes = TwainApi.OpenDS(&DSID);
  }
  TW_CUSTOMDSDATA Data;
  Data.hData = NULL;

  if(dwDataSize && hData)//set Custom DS data
  {
    if(dwRes==0)
    {
      BYTE *pSrc =(BYTE *)GlobalLock(hData);
      if(pSrc==0)
      {
        dwRes = FAILURE(TWCC_LOWMEMORY);
      }
      if(dwRes==0)
      {
        Data.hData = TwainApi.DSM_Alloc(dwDataSize);
        Data.InfoLength = dwDataSize;
        if(Data.hData==0)
        {
          dwRes = FAILURE(TWCC_LOWMEMORY);
        }
      }
      BYTE *pDest=0;
      if(dwRes==0)
      {
        pDest =(BYTE *) TwainApi.DSM_LockMemory(Data.hData);
        if(pDest==0)
        {
          dwRes = FAILURE(TWCC_LOWMEMORY);
        }
      }
      if(dwRes==0)
      {
        memcpy(pDest,pSrc,dwDataSize);
        TwainApi.DSM_UnlockMemory(Data.hData);
      }
      // Free allocated resources
      GlobalUnlock(hData);
      GlobalFree(hData);
      hData=NULL;
      dwDataSize=0;

      if(dwRes==0)
      {
        dwRes = TwainApi.SetDSCustomData(Data);
      }
      if(Data.hData)
      {
        // Free allocated resources
        TwainApi.DSM_Free(Data.hData);
        Data.hData=NULL;
      }
    }   
  }

  LL llVal;
  WORD wType;
  *pbFeederEnabled = false;
  //reads CAP_FEEDERENABLED
  dwRes = TwainApi.GetCapCurrentValue(CAP_FEEDERENABLED,&llVal,&wType);
  if(dwRes==0&&wType==TWTY_BOOL&& (bool)llVal==true)
  {
    *pbFeederEnabled=true;
  }
  // Release DS and DSM
  TwainApi.CloseDS();
  TwainApi.CloseDSM();
  return dwRes==0;
}


//Profile functions
void TW_InitilizeProfiles(CComboBox *pcbxProfiles)
{
  TCHAR strProfilesPath[MAX_PATH];
  TCHAR strOldPath[MAX_PATH];

  pcbxProfiles->Clear();
  pcbxProfiles->AddString(L"");
  GetCurrentDirectory(MAX_PATH, strOldPath);
  // get temp Application data directory for all users
  if(!SHGetSpecialFolderPath(NULL, strProfilesPath, CSIDL_COMMON_APPDATA|CSIDL_FLAG_CREATE, TRUE)) 
  {
    return;
  }
  

  if(strProfilesPath[wcslen(strProfilesPath)-1] != '\\')
  {
    wcscat_s(strProfilesPath, MAX_PATH, L"\\");
  }

  wcscat_s(strProfilesPath, MAX_PATH, PROFILELOCATION);
  if(strProfilesPath[wcslen(strProfilesPath)-1] != '\\')
  {
    wcscat_s(strProfilesPath, MAX_PATH, L"\\");
  }
  // Set current directory to profiles directory 
  if(!SetCurrentDirectory(strProfilesPath))
  {
    TCHAR     szPath[MAX_PATH];
    TCHAR     szTempPath[MAX_PATH];
    TCHAR     *pPath;
    TCHAR     *pSeparator;

    wcscpy_s(szPath, MAX_PATH, strProfilesPath);
    pPath    = szPath;
    // creates profiles path if it is needed
    while( pSeparator = wcschr(pPath, '\\') )
    {
      *pSeparator = '\0';
      wcscpy_s(szTempPath, MAX_PATH, pPath);
      wcscat_s(szTempPath, MAX_PATH, L"\\");

      if(!SetCurrentDirectory(szTempPath))
      {
        if( !CreateDirectory(szTempPath, NULL) )
        {
          return;
        }
        if(!SetCurrentDirectory(szTempPath))
        {
          return;
        }
      }
      pPath = pSeparator+1;
    }
  }
   // keep profiles path in global variable
  g_strProfilesDirectory = strProfilesPath;
  
  WIN32_FIND_DATA FindFileData;
  HANDLE          hFind   = NULL;
  TCHAR           szFileName[MAX_PATH];
  TCHAR           *pDot;
  
  // check all files with extension FILEEXTENTION in profiles directory
  hFind = FindFirstFile(L"*"FILEEXTENTION, &FindFileData);
  while(hFind != INVALID_HANDLE_VALUE)
  {
    wcscpy_s(szFileName, MAX_PATH, FindFileData.cFileName);

    try
    {
      HGLOBAL hData;
      DWORD dwSize;
      // check if profile is valid
      if(TW_LoadProfileFromFile(FindFileData.cFileName, &hData, &dwSize))
      {
        pDot = wcsrchr(szFileName, '.');
        if(pDot)
        {
          *pDot = '\0';
        }
        //add profile name into combobox
        pcbxProfiles->InsertString(-1,szFileName);
      }
    }
    catch(...)
    {
    }
    // finish?
    if(!FindNextFile(hFind, &FindFileData))
    {
      FindClose(hFind);
      break;
    }
  }

  // set to empty name -> use Default UI.
  pcbxProfiles->SetCurSel(0); 

  // Restore current path
  SetCurrentDirectory(strOldPath);
}

bool TW_SaveProfileToFile(CString strFileName, HGLOBAL hData, DWORD dwDataSize)
{
  if(hData==0)
  {
    return false;
  }
  if(dwDataSize==0)
  {
    GlobalFree(hData);
    return false;
  }

  void* pData = GlobalLock(hData);
  if(!pData)
  {
    GlobalFree(hData);
    return false;
  }

  FILE *pFile  = NULL;
  bool  bError = false;
  //opens file
  if(_wfopen_s(&pFile, strFileName, _T("wb") )==0 && pFile)
  {
    bError = true;
    //store DS identity
    if(fwrite(&kDSID, sizeof(TW_IDENTITY), 1, pFile)!=1)
    {
      bError = false;
    }
    else
    {
      //store CustomDSdata
      if(fwrite(pData, dwDataSize, 1, pFile)!=1)
      {
        bError = false;
      }
    }
    fclose(pFile);
    //remove file on error
    if(!bError)
    {
      int i =_wremove(strFileName);
    }
  }
  
  GlobalUnlock(hData);
  GlobalFree(hData);

  return bError;
}

bool TW_LoadProfileFromFile(CString strFileName, HGLOBAL *phData, DWORD *pdwDataSize)
{
  if(phData==NULL || pdwDataSize==NULL)
  {
    return false;
  }

  *phData=NULL;
  *pdwDataSize=0;

  HGLOBAL hData;
  FILE *pFile = NULL;
  //open file
  if(_wfopen_s(&pFile, strFileName, _T("rb"))!=0)
  {
    return false;
  }
  // get file size
  fseek(pFile, 0, SEEK_END);
  DWORD dwDataSize = (DWORD)ftell(pFile);
  rewind(pFile);
  bool bRes = true;
  // it has to contains at least DS identity
  if(dwDataSize<=sizeof(TW_IDENTITY))
  {
    fclose(pFile);
    return false;
  }
  // real CustomDSdata size
  dwDataSize -=sizeof(TW_IDENTITY);
  TW_IDENTITY tempID;
  // Get DS identity stored in the file
  if(fread(&tempID, sizeof(TW_IDENTITY), 1, pFile)!=1)
  {
    fclose(pFile);
    return false;
  }
  // stored DS identity has to match TWAIN DS identity
  if(memcmp(&tempID, &kDSID, sizeof(TW_IDENTITY))!=0)
  {
    fclose(pFile);
    return false;
  }
  // allocate storage and read CustomDSdata
  LPVOID pData = NULL;
  if(hData = GlobalAlloc(GMEM_MOVEABLE, dwDataSize))
  {
    pData = GlobalLock(hData);
    if(pData)
    {
      if(fread(pData, dwDataSize, 1, pFile)!=1)
      {
        bRes = false;
      }
    }
    else
    {
      bRes = false;
    }
  }
  else
  {
    bRes = false;
  }
  fclose(pFile);
  
  if(hData)
  {
    GlobalUnlock(hData);
    if(!bRes)
    {
      // free resource on error
      GlobalFree(hData);
      hData=NULL;
    } 
    else
    {
      *phData=hData; 
      *pdwDataSize=dwDataSize;
    }
  }
  
	return bRes;
}

bool TW_NewProfile(CComboBox *pcbxProfiles)
{
  CProfName dlgPrfName;
  //display dialog for creating new name
  if(dlgPrfName.DoModal()!=IDOK)
  {
    return false; //on cancel
  }
  //check if name exists
  if(pcbxProfiles->FindString(-1, dlgPrfName.m_edtProfName) >=0)
  {
    MessageBox(NULL,L"Profile already exists", L"Error", MB_OK | MB_ICONEXCLAMATION | MB_APPLMODAL | MB_SETFOREGROUND);
    return false; //name exists
  }

  CString strProfileName=dlgPrfName.m_edtProfName;
  CString strFileName  = g_strProfilesDirectory;
  strFileName += strProfileName;
  strFileName += FILEEXTENTION;

  HGLOBAL hData=NULL;
  DWORD dwDataSize=NULL;

  // get Custom DS data from TWAIN DS.
  bool bRes = TW_Get_Set_DS_Data(pcbxProfiles->GetParent()->m_hWnd, &hData,&dwDataSize);
  if(bRes)
  {
    //save Custom DS data to a file
    bRes = TW_SaveProfileToFile(strFileName,hData,dwDataSize);
    if(!bRes)
    {
      MessageBox(NULL,L"Error saving profile", L"Error", MB_OK | MB_ICONEXCLAMATION | MB_APPLMODAL | MB_SETFOREGROUND);
    }
    else
    {
      //slect profile name in combobox
      pcbxProfiles->SetCurSel(pcbxProfiles->InsertString(-1,strProfileName));
    }
  }
  return bRes;
}

bool TW_EditProfile(CComboBox *pcbxProfiles)
{
  CString strProfileName;
  int nIndex = pcbxProfiles->GetCurSel();
  pcbxProfiles->GetWindowText(strProfileName);
  CString strFileName  = g_strProfilesDirectory;
  if(nIndex<0 || strProfileName.IsEmpty())
  {
    return true;// no profile selected
  }
  
  strFileName += strProfileName;
  strFileName += FILEEXTENTION;

  HGLOBAL hData;
  DWORD dwDataSize;
  //load profile
  if(!TW_LoadProfileFromFile(strFileName, &hData, &dwDataSize))
  {
    return false;
  }  
  // set Custom DS data , display TWAIN UI and gets back Custom DS data
  bool bRes = TW_Get_Set_DS_Data(pcbxProfiles->GetParent()->m_hWnd, &hData, &dwDataSize);
  if(bRes)
  {
    // store changed profile
    bRes = TW_SaveProfileToFile(strFileName,hData, dwDataSize);
    if(!bRes)
    {
      MessageBox(NULL,L"Error saving profile", L"Error", MB_OK | MB_ICONEXCLAMATION | MB_APPLMODAL | MB_SETFOREGROUND);
    }
  }
  // free allocated resources
  if(hData)
  {
    GlobalFree(hData);
  }
  return bRes;
}

bool TW_DeleteProfile(CComboBox *pcbxProfiles)
{
  int nIndex = pcbxProfiles->GetCurSel();
  CString strName;
  pcbxProfiles->GetWindowText(strName);
  if(nIndex>=0 && !strName.IsEmpty())// if profile is selected
  {
    CString   strFileName;
    CString   strMessage;
    
    strMessage.Format(L"Do you want to delete profile - %s", (LPCTSTR)strName);
    if(IDYES == MessageBox(NULL,strMessage, L"Error", MB_YESNO | MB_ICONEXCLAMATION | MB_APPLMODAL | MB_SETFOREGROUND) )
    {
      strFileName  = g_strProfilesDirectory;
      strFileName += strName;
      strFileName += FILEEXTENTION;

      //delete file from the disk
      if(_wremove(strFileName)==0)
      {
        // if OK delete profile name from combobox
        pcbxProfiles->DeleteString(nIndex);
        pcbxProfiles->SetCurSel(nIndex-1);
      }
      else
      {
        MessageBox(NULL,L"Error deleting profile", L"Error", MB_OK | MB_ICONEXCLAMATION | MB_APPLMODAL | MB_SETFOREGROUND);
        return false;
      }
    }
    return true;
  }
  else
  {
    return false;
  }
}

bool WIA_SelectProfile(CComboBox *pcbxProfiles, IWiaPropertyStorage* pIPS)
{
  int nIndex = pcbxProfiles->GetCurSel();
  CString strProfileName;
  pcbxProfiles->GetWindowText(strProfileName);
  if(nIndex>=0)// if profile is selected
  {
    long lID=-1;
    CComBSTR bstrVal;
    // get temporary profile file
    HRESULT hr = WIA_ReadPropBSTR(pIPS,CUSTOM_ROOT_PROP_ID1,bstrVal);
    FILE *pfDest = NULL;
    // open file
    if( SUCCEEDED(hr) ) 
    {           
      if(_wfopen_s(&pfDest, bstrVal.m_str, _T("wb"))!=0)
      {
        hr= E_FAIL;
      }  
    }
    FILE *pfSrc = NULL;
    if(strProfileName.GetLength()>0 && SUCCEEDED(hr) )// if real profile - open it  
    {      
      CString strFileName  = g_strProfilesDirectory;
      strFileName += strProfileName;
      strFileName += FILEEXTENTION;

      if(_wfopen_s(&pfSrc, (LPCTSTR)strFileName, _T("rb"))!=0)
      {
        hr= E_FAIL;
      }  
    }
    if(pfSrc&& SUCCEEDED(hr) ) //if real profile
    {   
      BYTE line[1024];
      size_t bytes;
      //copy the profile to temporary file
      while((bytes = fread(line,1,sizeof(line),pfSrc)) > 0)
      {
        fwrite(line,1, bytes,pfDest);  
      }
    }
    // free allocated resources
    if(pfDest)
    {
      fclose(pfDest);
    }
    if(pfSrc)
    {
      fclose(pfSrc);
    }
    if(!SUCCEEDED(hr))
    {
      MessageBox(NULL,L"Error setting profile!", L"Error", MB_OK | MB_ICONEXCLAMATION | MB_APPLMODAL | MB_SETFOREGROUND);
    }
    return SUCCEEDED(hr);
  }
  else
  {
    return false;
  }
}

