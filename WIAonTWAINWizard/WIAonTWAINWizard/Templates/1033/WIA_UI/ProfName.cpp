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
 * @file ProfName.cpp
 * CProfName implementation file
 * @author TWAIN Working Group
 * @date October 2009
 */

#include "stdafx.h"
#include "ProfName.h"


// CProfName dialog

IMPLEMENT_DYNAMIC(CProfName, CDialog)

CProfName::CProfName(CWnd* pParent /*=NULL*/)
	: CDialog(CProfName::IDD, pParent)
  , m_edtProfName(_T(""))
{

}

CProfName::~CProfName()
{
}

void CProfName::DoDataExchange(CDataExchange* pDX)
{
  CDialog::DoDataExchange(pDX);
  DDX_Control(pDX, IDOK, m_btnOK);
  DDX_Text(pDX, IDC_EDT_PROF_NAME, m_edtProfName);
}


BEGIN_MESSAGE_MAP(CProfName, CDialog)
  ON_EN_CHANGE(IDC_EDT_PROF_NAME, &CProfName::OnEnChangeEdtProfName)
END_MESSAGE_MAP()


// CProfName message handlers

void CProfName::OnEnChangeEdtProfName()
{
  UpdateData();
  //enable OK button is edit box is not empty
  m_btnOK.EnableWindow(!m_edtProfName.IsEmpty());
}
