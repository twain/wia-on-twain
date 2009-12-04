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
 * @file PropPage.cpp
 * CPropPage implementation file
 * @author TWAIN Working Group
 * @date October 2009
 */


#include "stdafx.h"
#include "PropPage.h"
#include "WIA_UI.h"
#include "TWAIN_API.h"
#include "utilities.h"

// CPropPage dialog

IMPLEMENT_DYNAMIC(CPropPage, CPropertyPage)

CPropPage::CPropPage(CComPtr<IWiaItem> pItem , IUnknown *pUIext)
	: CPropertyPage(CPropPage::IDD)
{
  m_pItem = pItem;
  m_pUIext=pUIext;//This class will decrease reference counter of pUIext on exit
}

CPropPage::~CPropPage()
{
}

void CPropPage::DoDataExchange(CDataExchange* pDX)
{
  CPropertyPage::DoDataExchange(pDX);
  DDX_Control(pDX, IDC_CBX_PROF, m_cbxProfiles);
}


BEGIN_MESSAGE_MAP(CPropPage, CPropertyPage)
  ON_WM_DESTROY()
  ON_BN_CLICKED(IDC_BTN_NEW, &CPropPage::OnBnClickedBtnNew)
  ON_BN_CLICKED(IDC_BTN_EDIT, &CPropPage::OnBnClickedBtnEdit)
  ON_BN_CLICKED(IDC_BTN_DEL, &CPropPage::OnBnClickedBtnDel)
END_MESSAGE_MAP()

BOOL CPropPage::OnInitDialog()
{
  CPropertyPage::OnInitDialog();
  
  // init m_cbxProfiles combobox with mames of the profiles 
  TW_InitilizeProfiles(&m_cbxProfiles);

  return TRUE;
}

void CPropPage::OnOK()
{
  IWiaPropertyStorage * pPropertyRootStorage = NULL;
  IWiaItem *pIWiaItemRoot;
  
  HRESULT hr = m_pItem->GetRootItem(&pIWiaItemRoot);
  //Get root item property storage - CUSTOM_ROOT_PROP_ID1 contains name of the temporary profile file
  if (SUCCEEDED(hr) )
  {
    hr = pIWiaItemRoot->QueryInterface( IID_IWiaPropertyStorage, (LPVOID *)&pPropertyRootStorage);
  }

  if (SUCCEEDED(hr) )
  {      
    // store selected profile in file specified in CUSTOM_ROOT_PROP_ID1
    WIA_SelectProfile(&m_cbxProfiles,pPropertyRootStorage);
    pPropertyRootStorage->Release(); 
  }

  CPropertyPage::OnOK();
}

void CPropPage::OnCancel()
{
  CPropertyPage::OnCancel();
}

void CPropPage::OnDestroy()
{
  CPropertyPage::OnDestroy();
  m_pUIext->Release();//Free m_pUIext 
}

void CPropPage::OnBnClickedBtnNew()
{
  //Creates new profile using CustomDSData from TWAIN DS. It shows TWAIN DS UI
  GetParent()->EnableWindow(0);
  TW_NewProfile(&m_cbxProfiles);
  GetParent()->EnableWindow(1);
}

void CPropPage::OnBnClickedBtnEdit()
{
  //Edit existing profile using TWAIN DS UI
  GetParent()->EnableWindow(0);
  TW_EditProfile(&m_cbxProfiles);
  GetParent()->EnableWindow(1);
}

void CPropPage::OnBnClickedBtnDel()
{
  //delete profile from the disk and from combobox
  TW_DeleteProfile(&m_cbxProfiles);
}

