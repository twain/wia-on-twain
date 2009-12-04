// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include <atlstr.h>
#include <msi.h>
#include <Msiquery.h>
#include <shlobj.h>

#include <atlstr.h>
#include <devguid.h>
#include <shlwapi.h>

#define MAX_DEST_PATH 1024
#define kREGKEY L"Microsoft\\VisualStudio\\9.0\\NewProjectTemplates\\TemplateDirs\\{FFE04EC1-301F-11D3-BF4B-00C04F79EFBD}"


BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
					 )
{
	switch (ul_reason_for_call)
	{
	case DLL_PROCESS_ATTACH:
	case DLL_THREAD_ATTACH:
	case DLL_THREAD_DETACH:
	case DLL_PROCESS_DETACH:
		break;
	}
	return TRUE;
}
typedef UINT (WINAPI* GETSYSTEMWOW64DIRECTORY)(LPTSTR, UINT);

bool IsWow64()
{
  GETSYSTEMWOW64DIRECTORY getSystemWow64Directory = NULL;
  HMODULE hKernel32 = NULL;
  TCHAR strWow64Directory[MAX_PATH];

  hKernel32 = ::GetModuleHandle(TEXT("kernel32.dll"));
  if (NULL==hKernel32)
  {
    //
    // This shouldn't happen, but if we can't get 
    // kernel32's module handle then assume we are 
    //on x86.  We won't ever install 32-bit drivers
    // on 64-bit machines, we just want to catch it 
    // up front to give users a better error message.
    //
    return FALSE;
  }

  getSystemWow64Directory = (GETSYSTEMWOW64DIRECTORY)GetProcAddress(hKernel32, "GetSystemWow64DirectoryW");
  if (getSystemWow64Directory == NULL) 
  {
    //
    // This most likely means we are running 
    // on Windows 2000, which didn't have this API 
    // and didn't have a 64-bit counterpart.
    //
    return FALSE;
  }

  if ((getSystemWow64Directory(strWow64Directory, sizeof(strWow64Directory)/sizeof(TCHAR)) == 0) &&
    (GetLastError() == ERROR_CALL_NOT_IMPLEMENTED)) 
  {
    return FALSE;
  }

  //
  // GetSystemWow64Directory succeeded 
  // so we are on a 64-bit OS.
  //
  return TRUE;
}
UINT _stdcall PrepareInstallation(MSIHANDLE hInstall)
{
  UINT errMsi = ERROR_SUCCESS;
  //assume general failure
  int nRet = 0;
  
  SetCursor( LoadCursor(NULL, IDC_WAIT) );
  //allocate memory to query the property
  TCHAR *pstrPath = new TCHAR[MAX_DEST_PATH];
  if(pstrPath)
  {
    DWORD dwLength = MAX_DEST_PATH;
  
    //retrieve the CustomActionData which is supposed to contain our full path
    errMsi = MsiGetProperty(hInstall, _T("CustomActionData"), pstrPath, &dwLength);
    if(ERROR_SUCCESS==errMsi)
    {
      //append the inf file path to the parameter line with quotes
      CString strParams;
      strParams = pstrPath;
      int i = strParams.Find(L";",0);
      if(i<0)
      {
        errMsi = ERROR_BAD_ENVIRONMENT;
      }
      else
      {
        CString strFilePath = strParams.Left(i);
        CString strTemplatePath = strParams.Mid(i+1);

        FILE *f;
        if(_wfopen_s(&f,(LPCTSTR)strFilePath,L"r+b"))
        {
          errMsi = ERROR_BAD_ENVIRONMENT;
        }
        else
        {
          fseek(f,0,SEEK_END);
          long lSize = ftell(f)+1+strTemplatePath.GetLength()*2;
          fseek(f,0,SEEK_SET);
          BYTE * pbyData = (BYTE*)malloc(lSize);
          if(pbyData)
          {
            memset(pbyData,0,lSize);
            fread(pbyData,lSize,1,f);
            pbyData[lSize]=0;
            strParams = (char*)pbyData;

            i = strParams.Find(L"@",0);
            strParams =strParams.Left(i)+ strTemplatePath+strParams.Mid(i+1);
            i++;
            i = strParams.Find(L"@",0);
            strParams =strParams.Left(i)+ strTemplatePath+strParams.Mid(i+1);
            WideCharToMultiByte(CP_ACP,0,(LPCTSTR)strParams,strParams.GetLength(),(LPSTR)pbyData,lSize,0,0);
            fseek(f,0,SEEK_SET);
            fwrite(pbyData,strParams.GetLength(),1,f);
          }
          else
          {
            errMsi = ERROR_OUTOFMEMORY;
          }
          fclose(f);
        }
        strTemplatePath = strFilePath.Left(strFilePath.ReverseFind('\\'));
        HKEY      hKey;
        DWORD     disposition;	//create the key; show if we have errors or not
        CString strRegKey = kREGKEY;
        if(IsWow64())
        {
          strRegKey = L"Wow6432Node\\"+strRegKey;
        }
        strRegKey =  L"SOFTWARE\\"+strRegKey;

        if(ERROR_SUCCESS == RegCreateKeyEx(HKEY_LOCAL_MACHINE,(LPCTSTR)strRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, &disposition))
        {
          HKEY      hKey1;
          if(ERROR_SUCCESS == RegCreateKeyEx(hKey,L"/1", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey1, &disposition))
          {
            if(ERROR_SUCCESS != RegSetValueEx(hKey1,0,0,REG_SZ,(BYTE*)L"TWAIN Working Group projects",((DWORD)wcslen(L"TWAIN Working Group projects")+1)*sizeof(TCHAR)))
            {
              errMsi = ERROR_BAD_ENVIRONMENT;
            }
            if(ERROR_SUCCESS != RegSetValueEx(hKey1,L"TemplatesDir",0,REG_SZ,(BYTE*)(LPCTSTR)strTemplatePath,(strTemplatePath.GetLength()+1)*sizeof(TCHAR)))
            {
              errMsi = ERROR_BAD_ENVIRONMENT;
            }
            DWORD dwVal=255;
            if(ERROR_SUCCESS != RegSetValueEx(hKey1,L"SortPriority",0,REG_DWORD,(BYTE*)&dwVal,sizeof(DWORD)))
            {
              errMsi = ERROR_BAD_ENVIRONMENT;
            }
            RegCloseKey(hKey1);
          }
          else
          {
            errMsi = ERROR_BAD_ENVIRONMENT;
          }
          if(ERROR_SUCCESS == RegCreateKeyEx(hKey,L"/2", 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey1, &disposition))
          {
            DWORD dwVal=255;
            if(ERROR_SUCCESS != RegSetValueEx(hKey1,L"SortPriority",0,REG_DWORD,(BYTE*)&dwVal,sizeof(DWORD)))
            {
              errMsi = ERROR_BAD_ENVIRONMENT;
            }
            RegCloseKey(hKey1);
          }
          else
          {
            errMsi = ERROR_BAD_ENVIRONMENT;
          }
          RegCloseKey(hKey);
        }
        else
        {
          errMsi = ERROR_BAD_ENVIRONMENT;
        }
      }
    }
    //cleanup property memory
    delete [] pstrPath;
    pstrPath = NULL;
  }
  if((0!=nRet)&&(0==errMsi))
  {
    errMsi = ERROR_BAD_ENVIRONMENT;
  }
  
  SetCursor( LoadCursor(NULL, IDC_ARROW) );
  return errMsi;
}

UINT _stdcall CleanupInstallation(MSIHANDLE hInstall)
{
  CString strRegKey = kREGKEY;
  if(IsWow64())
  {
    strRegKey = L"Wow6432Node\\"+strRegKey;
  }
  strRegKey =  L"SOFTWARE\\"+strRegKey;

  UINT errMsi = ERROR_SUCCESS;
  SHDeleteKey(HKEY_LOCAL_MACHINE,(LPCTSTR)strRegKey);
  return errMsi;
}
