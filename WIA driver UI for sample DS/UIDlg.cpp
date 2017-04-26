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
 * @file UIDlg.cpp
 * CUIDlg implementation file
 * @author TWAIN Working Group
 * @date October 2009
 */

#include "stdafx.h"
#include "UIDlg.h"
#include "TWAIN_API.h"
#include "utilities.h"

// CUIDlg dialog

IMPLEMENT_DYNAMIC(CUIDlg, CDialog)

CUIDlg::CUIDlg( CWnd* pParent, PDEVICEDIALOGDATA2 pDeviceDialogData )
  : CDialog(CUIDlg::IDD, pParent)
  , m_bFeeder(FALSE)
{  
  m_pDeviceDlgData = pDeviceDialogData;
}

CUIDlg::~CUIDlg()
{
}

void CUIDlg::DoDataExchange(CDataExchange* pDX)
{
  CDialog::DoDataExchange(pDX);
  DDX_Control(pDX, IDC_CBX_PROF, m_cbxProfiles);
}


BEGIN_MESSAGE_MAP(CUIDlg, CDialog)
  ON_BN_CLICKED(IDOK, &CUIDlg::OnBnClickedOk)
  ON_BN_CLICKED(IDC_SHOW_DEF_UI, &CUIDlg::OnBnClickedOk2)
  ON_BN_CLICKED(IDC_BTN_NEW, &CUIDlg::OnBnClickedBtnNew)
  ON_BN_CLICKED(IDC_BTN_EDIT, &CUIDlg::OnBnClickedBtnEdit)
  ON_BN_CLICKED(IDC_BTN_DEL, &CUIDlg::OnBnClickedBtnDel)
END_MESSAGE_MAP()


// CUIDlg message handlers

BOOL CUIDlg::OnInitDialog()
{
  CDialog::OnInitDialog();
 
  // init m_cbxProfiles combobox with mames of the profiles 
  TW_InitilizeProfiles(&m_cbxProfiles);
  return TRUE;
}


void CUIDlg::OnBnClickedOk()
{
  m_bFeeder = false;
  //Get root item property storage - CUSTOM_ROOT_PROP_ID1 contains name of the temporary profile file  
  IWiaPropertyStorage * pPropertyRootStorage = NULL;
  HRESULT hr = m_pDeviceDlgData->pIWiaItemRoot->QueryInterface( IID_IWiaPropertyStorage, (LPVOID *)&pPropertyRootStorage);
  
  if (SUCCEEDED(hr) )
  { 
    //Get child item - feeder ol flatbed
    if(TW_Get_FeederEnabled(&m_cbxProfiles,&m_bFeeder))
    {
      // store selected profile in file specified in CUSTOM_ROOT_PROP_ID1
      WIA_SelectProfile(&m_cbxProfiles,pPropertyRootStorage);
    }
    else
    {
      // not possible -> use default UI
      m_cbxProfiles.SetCurSel(0);
    }
    pPropertyRootStorage->Release(); 
  }

  if(SUCCEEDED(hr))
  {
    CString strProfileName;
    int nIndex = m_cbxProfiles.GetCurSel();
    m_cbxProfiles.GetWindowText(strProfileName);
    if(nIndex<0 || strProfileName.IsEmpty())//use default UI?
    {
      CDialog::EndDialog( IDC_SHOW_DEF_UI );
    }
    else
    {
      CDialog::EndDialog( IDOK );
    }
  }
}

void CUIDlg::OnBnClickedOk2()
{
  //select empty profile name
  m_cbxProfiles.SetCurSel(0);

  //Get root item property storage - CUSTOM_ROOT_PROP_ID1 contains name of the temporary profile file  
  IWiaPropertyStorage * pPropertyRootStorage = NULL;
  HRESULT hr = m_pDeviceDlgData->pIWiaItemRoot->QueryInterface( IID_IWiaPropertyStorage, (LPVOID *)&pPropertyRootStorage);
  
  if (SUCCEEDED(hr) )
  {      
    // delete context of the file specified in CUSTOM_ROOT_PROP_ID1
    WIA_SelectProfile(&m_cbxProfiles,pPropertyRootStorage);
    pPropertyRootStorage->Release(); 
  }
  
  if(SUCCEEDED(hr))
  {
    CDialog::EndDialog( IDC_SHOW_DEF_UI );
  }
}
void CUIDlg::OnBnClickedBtnNew()
{
  //Creates new profile using CustomDSData from TWAIN DS. It shows TWAIN DS UI
  EnableWindow(0);
  TW_NewProfile(&m_cbxProfiles);
  EnableWindow(1);
}

void CUIDlg::OnBnClickedBtnEdit()
{
  //Edit existing profile using TWAIN DS UI
  EnableWindow(0);
  TW_EditProfile(&m_cbxProfiles);
  EnableWindow(1);
}

void CUIDlg::OnBnClickedBtnDel()
{
  //delete profile from the disk and from combobox
  TW_DeleteProfile(&m_cbxProfiles);
}

