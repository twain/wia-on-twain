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
 * @file utilities.h
 * Utility functions common for CUIDlg and CPropPage
 * @author TWAIN Working Group
 * @date October 2009
 */

#pragma once
/**
* Read string property from WIA property storage
* @param[in] pIPS a pointer IWiaPropertyStorage. The Property storage
* @param[in] propID a PROPID. ID of the property
* @param[out] bstrVal a CComBSTR. Property value
* @return S_OK on success
*/
HRESULT WIA_ReadPropBSTR( IWiaPropertyStorage* pIPS, PROPID propID, CComBSTR& bstrVal );

/**
* Set CustomDsData, display TWAIN DS UI, Get CustomDsData
* @param[in] hWnd a HWND. The Property storage
* @param[in] phData a pointer to HGLOBAL. Handle to data to be set if not NULL. On exit contains Handle to CustomDsData get from TWAIN DS
* @param[in] pdwDataSize a pointer to DWORD. Specifies size of the data pointed by phData
* @return true on success
*/
bool TW_Get_Set_DS_Data(HWND hWnd, HGLOBAL *phData, DWORD *pdwDataSize);

/**
* Set CustomDsData and get CAP_FEEDERENABLED - finds which child item has to be used during the transfer
* @param[in] pcbxProfiles a pointer to CComboBox. Combobox contains profile name
* @param[out] pbFeederEnabled a pointer to BOOL. true if feeder item ahve to be used during the transfer
* @return true on success
*/
bool TW_Get_FeederEnabled(CComboBox *pcbxProfiles, BOOL *pbFeederEnabled);

/**
* Initialize combobox with profile names
* @param[in] pcbxProfiles a pointer to CComboBox. 
* @return true on success
*/
void TW_InitilizeProfiles(CComboBox *pcbxProfiles);
/**
* Save CustomDSdata to file
* @param[in] strFileName a CString. Name of the file 
* @param[in] hData a HGLOBAL. Handle to CustomDSdata
* @param[in] dwDataSize a hData. Size of CustomDSdata
* @return true on success
*/
bool TW_SaveProfileToFile(CString strFileName, HGLOBAL hData, DWORD dwDataSize);
/**
* Load CustomDSdata from file
* @param[in] strFileName a CString. Name of the file 
* @param[out] hData a pointer to HGLOBAL. Handle to CustomDSdata
* @param[out] dwDataSize a pointer to hData. Size of CustomDSdata
* @return true on success
*/
bool TW_LoadProfileFromFile(CString strFileName, HGLOBAL *phData, DWORD *pdwDataSize);
/**
* Create new profile. Displays dialog box for creating the name of the profile, displays TWAIN UI, gets CustomDSdata and store it to profile
* @param[in] pcbxProfiles a pointer to CComboBox. 
* @return true on success
*/
bool TW_NewProfile(CComboBox *pcbxProfiles);
/**
* Delete selected profile from the disk
* @param[in] pcbxProfiles a pointer to CComboBox. 
* @return true on success
*/
bool TW_DeleteProfile(CComboBox *pcbxProfiles);
/**
* Loads profile, sets CustomDSdata, displays TWAIN UI, gets CustomDSdata and store it into the same profile
* @param[in] pcbxProfiles a pointer to CComboBox. 
* @return true on success
*/
bool TW_EditProfile(CComboBox *pcbxProfiles);
/**
* Loads profile and store it in temporary profile created by WIA driver
* @param[in] pcbxProfiles a pointer to CComboBox. 
* @param[in] pIPS a pointer to IWiaPropertyStorage. Root item property storage contains temporary profile name property
* @return true on success
*/
bool WIA_SelectProfile(CComboBox *pcbxProfiles, IWiaPropertyStorage* pIPS);
