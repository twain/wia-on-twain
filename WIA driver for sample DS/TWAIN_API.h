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
 * @file TWAIN_API.h
 * CTWAIN_API class for communicating with TWAIN DS
 * LL class unifies TWAIN capabilities 
 * @author TWAIN Working Group
 * @date October 2009
 */

#pragma once
#include "twain.h"

#pragma warning(disable:4995)
#include <string>
#include <list>
#include <map>
using namespace std;

/**
* This class is unified representation of TWAIN capability value.
*/
class LL
{
  //Constructors
private:
  LL(long long llVal,bool bFix32){m_bFix32=bFix32;m_val=llVal;};
public:
  long long m_val;
  bool m_bFix32;
  LL(){m_bFix32=false;m_val=0;};
  LL(long lVal){m_bFix32=false;m_val=lVal;};
  LL(DWORD ulVal){m_bFix32=false;m_val=ulVal;};
  LL(int lVal){m_bFix32=false;m_val=lVal;};
  LL(bool bVal){m_bFix32=false;m_val = bVal?1:0;};
  LL(TW_FIX32 fx32Val){m_bFix32=true;m_val = ((long)fx32Val.Whole<<16) +  fx32Val.Frac;};
  
  //cast operators
  operator const long long(){return m_val;};
  operator const WORD(){if(m_bFix32)return (WORD)HIWORD(m_val+0x8000); else return (WORD)m_val;};
  operator const int(){if(m_bFix32)return (int)((m_val+0x8000)>>16); else return (int)m_val;};
  operator const long(){if(m_bFix32)return (long)((m_val+0x8000)>>16); else return (long)m_val;};
  operator const char(){if(m_bFix32)return (char)HIWORD(m_val+0x8000); else return (char)m_val;};
  operator const BYTE(){if(m_bFix32)return (BYTE)HIWORD(m_val+0x8000); else return (BYTE)m_val;};
  operator const short(){if(m_bFix32)return (short)HIWORD(m_val+0x8000); else return (short)m_val;};
  operator const bool(){return m_val!=0;};
  operator const DWORD(){if(m_bFix32)return (DWORD)((m_val+0x8000)>>16); else return (DWORD)m_val;};
  operator const TW_FIX32(){TW_FIX32 ret; if(m_bFix32){ret.Frac=LOWORD(m_val);ret.Whole=HIWORD(m_val);}else{ret.Frac=0;ret.Whole=LOWORD(m_val);}return ret;};//TODO neg
  
  //Assign operators
  LL &operator=( long long llVal){m_bFix32=false;m_val = llVal;return *this;};
  LL &operator=( char chVal){m_bFix32=false;m_val = chVal;return *this;};
  LL &operator=( short sVal){m_bFix32=false;m_val = sVal;return *this;};
  LL &operator=( long lVal){m_bFix32=false;m_val = lVal;return *this;};
  LL &operator=( BYTE byVal){m_bFix32=false;m_val = byVal;return *this;};
  LL &operator=( WORD wVal){m_bFix32=false;m_val = wVal;return *this;};
  LL &operator=( DWORD dwVal){m_bFix32=false;m_val = dwVal;return *this;};
  LL &operator=( bool bVal){m_bFix32=false;m_val = bVal?1:0;return *this;};
  LL &operator=( TW_FIX32 fx32Val){m_bFix32=true;m_val = MAKELONG(fx32Val.Frac,fx32Val.Whole);return *this;};
  
  // Arithmetic operations
  LL operator +=( LL llVal){if(!llVal.m_bFix32&&m_bFix32){llVal.m_val=llVal.m_val<<16;}if(llVal.m_bFix32&&!m_bFix32){m_val=m_val<<16;m_bFix32=true;}m_val+=llVal.m_val;return LL(m_val,m_bFix32);};
  LL operator +( LL llVal){long long llTemp =m_val; if(!llVal.m_bFix32&&m_bFix32){llVal.m_val=llVal.m_val<<16;}if(llVal.m_bFix32&&!m_bFix32){llTemp=llTemp<<16;}llTemp+=llVal.m_val;return LL(llTemp,llVal.m_bFix32||m_bFix32);};
  LL operator -( LL llVal){long long llTemp =m_val; if(!llVal.m_bFix32&&m_bFix32){llVal.m_val=llVal.m_val<<16;}if(llVal.m_bFix32&&!m_bFix32){llTemp=llTemp<<16;}llTemp-=llVal.m_val;return LL(llTemp,llVal.m_bFix32||m_bFix32);};
  LL operator *( LL llVal){long long llTemp =m_val*llVal.m_val; if(m_bFix32&&llVal.m_bFix32){llTemp=llTemp>>16;} return LL(llTemp,llVal.m_bFix32||m_bFix32);};
  LL operator *( int lVal){long long llTemp =m_val*lVal; return LL(llTemp,m_bFix32);};
  LL operator /( LL llVal){long long llTemp =m_val; if(llVal.m_bFix32){llTemp=llTemp<<16;if(!m_bFix32){llTemp=llTemp<<16;}} return LL(llTemp/llVal.m_val,true);};
  LL operator /( int lVal){long long llTemp =m_val; if(!m_bFix32){llTemp=llTemp<<16;}return LL(llTemp/lVal,true);};
  LL operator /( long lVal){long long llTemp =m_val; if(!m_bFix32){llTemp=llTemp<<16;}return LL(llTemp/lVal,true);};
  bool operator >( LL llVal){return m_val>llVal.m_val;};
  bool operator <( LL llVal){return m_val<llVal.m_val;};
  bool operator >=( LL llVal){return m_val>=llVal.m_val;};
  bool operator <=( LL llVal){return m_val<=llVal.m_val;};
  bool operator ==( LL llVal){return m_val==llVal.m_val;};
  bool operator >=( int lVal){return m_val>=lVal;};
  bool operator <=( int lVal){return m_val<=lVal;};
  bool operator ==( int lVal){return m_val==lVal;};
  bool operator !=( LL llVal){return m_val!=llVal.m_val;};
 };

class LL_ARRAY : public list<LL>
{
public:
  bool IfExist(LL);
};

#define FAILURE(a) MAKELONG(TWRC_FAILURE,a)

/**
* possible States of the TWAIN APP.
*/
typedef enum _TWSTATE
{
  PRE_SESSION    = 0x01,/**< DSM not  loaded*/
  DSM_LOADED     = 0x02,/**< DSM loaded*/
  DSM_OPENED     = 0x03,/**< DSM loaded and opened. List of DS is available */
  DS_OPENED      = 0x04,/**< Source is open, and ready to: List & Negotiate Capabilities, Request the Acquisition of data, and Close. */
  DS_ENABLED     = 0x05,/**< If UI is being used it is displayed. */
  TRANSFER_READY = 0x06,/**< Transfers are ready. */
  TRANSFERRING   = 0x07 /**< Transfering data. */
}TWSTATE;

/**
* This is a TWAIN compliant class. It contains all of the capabilities
* and functions that are required to be a TWAIN compliant application.
*/
class CTWAIN_API
{
public:
  /**
  * Constructor
  * @param[in] hWindow a HWND of parent window
  */
  CTWAIN_API(HWND hWindow=NULL);
  ~CTWAIN_API();

  /**
  * It loads and opens DSM
  * @param[in] AppID a TW_IDENTITY.
  * @param[in] nID a unique ID
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD OpenDSM(TW_IDENTITY AppID, int nID);
  /**
  * It closes and unloads DSM
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD CloseDSM();
  /**
  * It loads and opens DS
  * @param[in] pDS_ID a pointer to DS TW_IDENTITY.
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD OpenDS(TW_IDENTITY *pDS_ID=NULL);
  /**
  * It closes and unloads DS
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD CloseDS();
  /**
  * Get Last condition code from DSM
  * @param[in] err a TW_UINT16. Last TWAIN return code
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetDSMConditionCode(TW_UINT16 err);
  /**
  * Get Last condition code from DS
  * @param[in] err a TW_UINT16. Last TWAIN return code
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetDSConditionCode(TW_UINT16 err);
  /**
  * Enables DS in UIless mode
  * @param[out] phEvent a pointer to event HANDLE. Event which will be signaled in case of DS event
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD EnableDS(HANDLE * phEvent);
  /**
  * Disables DS
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD DisableDS();
  /**
  * Gets Image Info for last scanned image. 
  * @param[out] pImageInfo a pointer to TW_IMAGEINFO.
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetImageInfo(TW_IMAGEINFO *pImageInfo);
  /**
  * Gets current Image Layout. 
  * @param[out] pImageLayout a pointer to TW_ IMAGELAYOUT.
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetImageLayout(TW_IMAGELAYOUT *pImageLayout);
  /**
  * Sets current Image Layout. 
  * @param[in] ImageLayout a TW_ IMAGELAYOUT.
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD SetImageLayout(TW_IMAGELAYOUT ImageLayout);
  /**
  * Scans one image in DIB format 
  * @param[out] phBMP a pointer to DIB HANDLE.
  * @param[out] pdwSize a pointer to DWORD. Size of the memory pointed by phBMP
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD ScanNative(HANDLE *phBMP, DWORD *pdwSize);
  /**
  * Gets memory transfer buffer size 
  * @param[out] pSetupMemTransfer a pointer to TW_SETUPMEMXFER.
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetMemTransferCfg(TW_SETUPMEMXFER *pSetupMemTransfer);
  /**
  * Ends current image transfer 
  * @param[out] pbMoreImages a pointer to bool. If true there are more images waiting to be transferred
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD EndTransfer(bool *pbMoreImages);
  /**
  * Ends the image transfer
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD ResetTransfer();
  /**
  * Sends message for processing by DS
  * @param[in] pMsg a pointer to MSG. Message to be processed
  * @param[out] pTwainMsg a pointer to WORD. TWAIN message , valid when pbDSeventis true
  * @param[out] pbDSevent a pointer to bool. If true then the message is for this DS
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD ProcessEvent(MSG *pMsg, WORD *pTwainMsg, bool *pbDSevent);
  /**
  * Get supported operation for a TWAIN capability
  * @param[in] wCapID an WORD. TWAIN capability ID
  * @param[out] pdwSupport a pointer to DWORD. Supported operations by this capability
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD QueryCapSupport(WORD wCapID, DWORD * pdwSupport);
  /**
  * Get current value of a TWAIN capability
  * @param[in] wCapID an WORD. TWAIN capability ID
  * @param[out] pllValue a pointer to LL. Current value
  * @param[out] pwType a pointer to WORD. Type of returned value
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetCapCurrentValue(WORD wCapID, LL *pllValue, WORD *pwType=0);
  /**
  * Get default value of a TWAIN capability
  * @param[in] wCapID an WORD. TWAIN capability ID
  * @param[out] pllValue a pointer to LL. Default value
  * @param[out] pwType a pointer to WORD. Type of returned value
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetCapDefaultValue(WORD wCapID, LL *pllValue, WORD *pwType=0);
  /**
  * Get all currently available values of a TWAIN capability
  * @param[in] wCapID an WORD. TWAIN capability ID
  * @param[out] pllValue a pointer to LL_ARRAY. All currently available values
  * @param[out] pwType a pointer to WORD. Type of returned value
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetCapConstrainedValues(WORD wCapID, LL_ARRAY *pdwList, WORD *pwType=0);
  /**
  * Get Min and Max possible values of a TWAIN capability
  * @param[in] wCapID an WORD. TWAIN capability ID
  * @param[out] pdwMin a pointer to LL. Min possible value
  * @param[out] pdwMax a pointer to LL. Max possible value
  * @param[out] pdwStep a pointer to LL. Increment if available or else 0
  * @param[out] pwType a pointer to WORD. Type of returned value
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetCapMinMaxValues(WORD wCapID, LL *pdwMin, LL *pdwMax, LL *pdwStep, WORD *pwType=0);
  /**
  * Set current values of a TWAIN capability
  * @param[in] wCapID an WORD. TWAIN capability ID
  * @param[in] dwValue a  LL. New value
  * @param[in] pwType a WORD. Type of the value
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD SetCapability(WORD wCapID, LL dwValue, WORD wType);
  /**
  * Get all currently available values of a BOOL TWAIN capability.
  * @param[in] wCapID an WORD. TWAIN capability ID
  * @param[out] pllValue a pointer to LL_ARRAY. All currently available values
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetBoolCapConstrainedValues(WORD wCapID, LL_ARRAY *pdwList);
  /**
  * Get all currently available values of a CAP_XFERCOUNT.
  * @param[out] pllValue a pointer to LL_ARRAY. All currently available values
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetXFERCOUNTConstrainedValues(LL_ARRAY *pdwList);
  /**
  * Enables DS in setup mode
  * @param[out] phEvent a pointer to event HANDLE. Event which will be signaled in case of DS event
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD EnableDSOnly(HANDLE *phEvent);
  /**
  * Sets custom DS
  * @param[in] Data a TW_CUSTOMDSDATA. Custom DS data to be set
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD SetDSCustomData(TW_CUSTOMDSDATA Data);
  /**
  * Gets custom DS
  * @param[out] pData a pointer to TW_CUSTOMDSDATA. Read Custom DS data
  * @return a valid DWORD return code. LOWORD TWAIN return code, HIWORD TWAIN condition code
  */
  DWORD GetDSCustomData(TW_CUSTOMDSDATA *pData);
  
  /**
  * Allocates memory 
  * @param[in] dwSize a pointer to TW_UINT32. Size of the memory
  * @return a valid TW_HANDLE. Handle to memory
  */
  TW_HANDLE DSM_Alloc(TW_UINT32 dwSize);
  /**
  * Frees allocated memory 
  * @param[in] hMemory a pointer TW_HANDLE. Handle to memory
  */
  void DSM_Free(TW_HANDLE hMemory);
  /**
  * Locks allocated memory 
  * @param[in] hMemory a pointer TW_HANDLE. Handle to memory
  * @return a valid TW_MEMREF. Pointer to memory
  */
  TW_MEMREF DSM_LockMemory(TW_HANDLE hMemory);
  /**
  * Unlocks allocated memory 
  * @param[in] hMemory a pointer TW_HANDLE. Handle to memory
  */
  void DSM_UnlockMemory(TW_MEMREF hMemory);

  /**
  * Converts value from current TWAIN DS units to inches 
  * @param[in] llVal a LL. Value to be converted
  * @return a LL. Converted value
  */
  LL ConvertToInch(LL llVal);
  /**
  * Converts value from inches to current TWAIN DS units
  * @param[in] llVal a LL. Value to be converted
  * @return a LL. Converted value
  */
  LL ConvertFromInch(LL llVal);
  /**
  * Returns last received message from DS
  * @return a TW_UINT16. TWAIN DS message
  */
  TW_UINT16 GetLastMsg(){return m_msgLast;}

private:
  int m_nID;                  /**< Unique ID*/
  WORD m_wDSunits;            /**< Current TWAIN DS units*/ 
  HWND m_hWindow;             /**< Handle to window*/
  bool m_bOwnWindow;          /**< true if this caller class supplies the m_hWindow */
  HMODULE m_hDSM;             /**< hadle to DSM library*/
  DSMENTRYPROC m_fnDSM_Entry; /**< pointer to DSM_ENTRY entry point of DSM*/
  TW_IDENTITY m_AppID;        /**< Application ID*/
  TW_IDENTITY m_DS;           /**< DS ID*/
  TWSTATE m_twState;          /**< Current TWAIN state*/
  TW_ENTRYPOINT m_DSM_Entry;  /**< structure with DSM entry points*/
  HANDLE m_hDSMEvent;         /**< handle of event signals caller class for TWAIN DS event */
  TW_UINT16 m_msgLast;        /**< Last received TWAIN message*/
  CString m_strClassName;     /**< class name of the window created by this class*/

  /**
  * Callback function. Called by DSM when DS send a message
  * @param[in] pOrigin a pointer to TW_IDENTITY. TWAIN DS ID 
  * @param[in] pDest a pointer to TW_IDENTITY. Application ID 
  * @param[in] DG a TW_UINT32. Data group 
  * @param[in] DAT a TW_UINT32. Data argument type
  * @param[in] MSG a TW_UINT32. Message ID
  * @param[in] pData a TW_MEMREF. Pointer to data
  * @return a TW_UINT16. TWAIN return code
  */
  static TW_UINT16 FAR PASCAL TWAIN_callback(pTW_IDENTITY pOrigin,pTW_IDENTITY pDest,TW_UINT32 DG,
                           TW_UINT16 DAT,TW_UINT16 MSG,TW_MEMREF pData);
};

