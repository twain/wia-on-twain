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
 * @file UIDlg.h
 * CPropPage class - UI extension for WIA driver in case of WIA 2.0 applications. It replace existing UI
 * @author TWAIN Working Group
 * @date October 2009
 */
#pragma once
#include "afxwin.h"

// CUIDlg dialog

class CUIDlg : public CDialog
{
	DECLARE_DYNAMIC(CUIDlg)

public:
	CUIDlg( CWnd* pParent, PDEVICEDIALOGDATA2 pDeviceDialogData );
	virtual ~CUIDlg();

// Dialog Data
	enum { IDD = IDD_UI_DIALOG };
	
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support
	DECLARE_MESSAGE_MAP()

	PDEVICEDIALOGDATA2 m_pDeviceDlgData;
				
public:
  virtual BOOL OnInitDialog();
  afx_msg void OnBnClickedOk();
  afx_msg void OnBnClickedOk2();
  afx_msg void OnBnClickedBtnNew();
  afx_msg void OnBnClickedBtnEdit();
  afx_msg void OnBnClickedBtnDel();

  BOOL m_bFeeder;           /**< true if feeder item will be used for current image transfer*/
  CComboBox m_cbxProfiles;  /**< combobox contains list of existing profiles */
};
