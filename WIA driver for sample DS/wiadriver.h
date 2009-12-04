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
 * @file wiadriver.h
 * @author TWAIN Working Group
 * @date October 2009
 */


#pragma once
#include "TWAIN_API.h"
#define DEFINE_WIA_PROPID_TO_NAME //enables table g_wiaPropIdToName
///////////////////////////////////////////////////////////////////////////////
// WIA driver core headers
//
#include <sti.h>                // STI defines
#include <stiusd.h>             // IStiUsd interface
#include <wiamindr.h>           // IWiaMinidrv interface 
#include <wiamdef.h>
#include "wiapropertymanager.h"     // WIA driver property manager class
#include "wiacapabilitymanager.h"   // WIA driver capability manager class

#define WIA_DRIVER_ROOT_NAME            L"Root"    // THIS SHOULD NOT BE LOCALIZED
#define WIA_DRIVER_FLATBED_NAME         L"Flatbed" // THIS SHOULD NOT BE LOCALIZED
#define WIA_DRIVER_FEEDER_NAME          L"Feeder"  // THIS SHOULD NOT BE LOCALIZED

#define DEFAULT_LOCK_TIMEOUT            1000  //timeout for locking STI device
#define DEFAULT_NUM_DRIVER_FORMATS      1     // number of supported formats - WIA 2.0
#define DEFAULT_NUM_DRIVER_FORMATS_WIA1 2     // number of supported formats - WIA 1.0

class INonDelegatingUnknown 
{
public:
  virtual STDMETHODIMP NonDelegatingQueryInterface(REFIID riid,LPVOID *ppvObj) = 0;
  virtual STDMETHODIMP_(ULONG) NonDelegatingAddRef() = 0;
  virtual STDMETHODIMP_(ULONG) NonDelegatingRelease() = 0;
};

class CWIADriver : public INonDelegatingUnknown, // NonDelegatingUnknown
                   public IStiUSD,               // STI USD interface
                   public IWiaMiniDrv            // WIA Minidriver interface
{
public:

  ///////////////////////////////////////////////////////////////////////////
  // Construction/Destruction Section
  ///////////////////////////////////////////////////////////////////////////

  CWIADriver(__in_opt LPUNKNOWN punkOuter);
  ~CWIADriver();

private:
  ///////////////////////////////////////////////////////////////////////////
  // WIA driver internals
  ///////////////////////////////////////////////////////////////////////////

  LONG                    m_cRef;                     /**< Device object reference count. */
  LPUNKNOWN               m_punkOuter;                /**< Pointer to outer unknown. */
  IStiDevice             *m_pIStiDevice;              /**< STI device interface for locking */
  IWiaDrvItem            *m_pIDrvItemRoot;            /**< WIA root item */
  LONG                    m_lClientsConnected;        /**< number of applications connected */
  CWIACapabilityManager   m_CapabilityManager;        /**< WIA driver capabilities */
  WIA_FORMAT_INFO        *m_pFormats;                 /**< WIA format information */
  ULONG                   m_ulNumFormats;             /**< number of data formats */
  BSTR                    m_bstrDeviceID;             /**< WIA device ID; */
  BSTR                    m_bstrRootFullItemName;     /**< WIA root item (full item name) */
  bool                    m_bHasFlatbed;              /**< True if scanner has flatbed*/
  bool                    m_bHasFeeder;               /**< True if scanner has ADF*/
  bool                    m_bDuplex;                  /**< True if scanner supports duplex*/
  bool                    m_bInitialized;             /**< Driver initialized*/
  bool                    m_bIsWindowsVista;          /**< True on Windows Vista and later*/
  bool                    m_bUnint;                   /**< A flag. If true drvUnInitializeWia has been called before drvUnLockWiaDevice*/
  int                     m_nLockCounter;             /**< Lock counter*/
  CString                 m_strProfilesPath;          /**< Profiles path*/
public:

  ///////////////////////////////////////////////////////////////////////////
  // Standard COM Section
  ///////////////////////////////////////////////////////////////////////////

  STDMETHODIMP QueryInterface(__in REFIID riid, __out LPVOID * ppvObj);

  STDMETHODIMP_(ULONG) AddRef();

  STDMETHODIMP_(ULONG) Release();

  ///////////////////////////////////////////////////////////////////////////
  // IStiUSD Interface Section (for all WIA drivers)
  ///////////////////////////////////////////////////////////////////////////

  STDMETHOD(Initialize)(THIS_
                        __in  PSTIDEVICECONTROL pHelDcb,
                              DWORD             dwStiVersion,
                        __in  HKEY              hParametersKey);

  STDMETHOD(GetCapabilities)(THIS_ __out PSTI_USD_CAPS pDevCaps);

  STDMETHOD(GetStatus)(THIS_ __inout PSTI_DEVICE_STATUS pDevStatus);

  STDMETHOD(DeviceReset)(THIS);

  STDMETHOD(Diagnostic)(THIS_ __out LPDIAG pBuffer);

  STDMETHOD(Escape)(THIS_
                                                  STI_RAW_CONTROL_CODE EscapeFunction,
                    __in_bcount(cbInDataSize)     LPVOID               lpInData,
                                                  DWORD                cbInDataSize,
                    __out_bcount(dwOutDataSize)   LPVOID               pOutData,
                                                  DWORD                dwOutDataSize,
                    __out                         LPDWORD              pdwActualData);

  STDMETHOD(GetLastError)(THIS_ __out LPDWORD pdwLastDeviceError);

  STDMETHOD(LockDevice)();

  STDMETHOD(UnLockDevice)();

  STDMETHOD(RawReadData)(THIS_
                         __out_bcount(*lpdwNumberOfBytes)   LPVOID       lpBuffer,
                         __out                              LPDWORD      lpdwNumberOfBytes,
                         __out                              LPOVERLAPPED lpOverlapped);

  STDMETHOD(RawWriteData)(THIS_
                          __in_bcount(dwNumberOfBytes) LPVOID       lpBuffer,
                                                       DWORD        dwNumberOfBytes,
                          __out                        LPOVERLAPPED lpOverlapped);

  STDMETHOD(RawReadCommand)(THIS_
                            __out_bcount(*lpdwNumberOfBytes)    LPVOID       lpBuffer,
                            __out                               LPDWORD      lpdwNumberOfBytes,
                            __out                               LPOVERLAPPED lpOverlapped);

  STDMETHOD(RawWriteCommand)(THIS_
                             __in_bcount(dwNumberOfBytes)  LPVOID       lpBuffer,
                                                           DWORD        dwNumberOfBytes,
                             __out                         LPOVERLAPPED lpOverlapped);

  STDMETHOD(SetNotificationHandle)(THIS_ __in HANDLE hEvent);

  STDMETHOD(GetNotificationData)(THIS_ __in LPSTINOTIFY lpNotify);

  STDMETHOD(GetLastErrorInfo)(THIS_ STI_ERROR_INFO *pLastErrorInfo);

  /////////////////////////////////////////////////////////////////////////
  // IWiaMiniDrv Interface Section (for all WIA drivers)                 //
  /////////////////////////////////////////////////////////////////////////

  STDMETHOD(drvInitializeWia)(THIS_
                              __inout BYTE        *pWiasContext,
                                      LONG        lFlags,
                              __in    BSTR        bstrDeviceID,
                              __in    BSTR        bstrRootFullItemName,
                              __in    IUnknown    *pStiDevice,
                              __in    IUnknown    *pIUnknownOuter,
                              __out   IWiaDrvItem **ppIDrvItemRoot,
                              __out   IUnknown    **ppIUnknownInner,
                              __out   LONG        *plDevErrVal);

  STDMETHOD(drvAcquireItemData)(THIS_
                                __in      BYTE                      *pWiasContext,
                                          LONG                      lFlags,
                                __in      PMINIDRV_TRANSFER_CONTEXT pmdtc,
                                __out     LONG                      *plDevErrVal);

  STDMETHOD(drvInitItemProperties)(THIS_
                                   __inout    BYTE *pWiasContext,
                                              LONG lFlags,
                                   __out      LONG *plDevErrVal);

  STDMETHOD(drvValidateItemProperties)(THIS_
                                       __inout    BYTE           *pWiasContext,
                                                  LONG           lFlags,
                                                  ULONG          nPropSpec,
                                       __in       const PROPSPEC *pPropSpec,
                                       __out      LONG           *plDevErrVal);

  STDMETHOD(drvWriteItemProperties)(THIS_
                                    __inout   BYTE                      *pWiasContext,
                                              LONG                      lFlags,
                                    __in      PMINIDRV_TRANSFER_CONTEXT pmdtc,
                                    __out     LONG                      *plDevErrVal);

  STDMETHOD(drvReadItemProperties)(THIS_
                                   __in       BYTE           *pWiasContext,
                                              LONG           lFlags,
                                              ULONG          nPropSpec,
                                   __in       const PROPSPEC *pPropSpec,
                                   __out      LONG           *plDevErrVal);

  STDMETHOD(drvLockWiaDevice)(THIS_
                              __in    BYTE *pWiasContext,
                                      LONG lFlags,
                              __out   LONG *plDevErrVal);

  STDMETHOD(drvUnLockWiaDevice)(THIS_
                                __in      BYTE *pWiasContext,
                                          LONG lFlags,
                                __out     LONG *plDevErrVal);

  STDMETHOD(drvAnalyzeItem)(THIS_
                            __in      BYTE *pWiasContext,
                                      LONG lFlags,
                            __out     LONG *plDevErrVal);

  STDMETHOD(drvGetDeviceErrorStr)(THIS_
                                            LONG     lFlags,
                                            LONG     lDevErrVal,
                                  __out     LPOLESTR *ppszDevErrStr,
                                  __out     LONG     *plDevErr);

  STDMETHOD(drvDeviceCommand)(THIS_
                              __inout     BYTE            *pWiasContext,
                                          LONG            lFlags,
                              __in        const GUID      *plCommand,
                              __out       IWiaDrvItem     **ppWiaDrvItem,
                              __out       LONG            *plDevErrVal);

  STDMETHOD(drvGetCapabilities)(THIS_
                                __in      BYTE            *pWiasContext,
                                          LONG            ulFlags,
                                __out     LONG            *pcelt,
                                __out     WIA_DEV_CAP_DRV **ppCapabilities,
                                __out     LONG            *plDevErrVal);

  STDMETHOD(drvDeleteItem)(THIS_
                           __inout    BYTE *pWiasContext,
                                      LONG lFlags,
                           __out      LONG *plDevErrVal);

  STDMETHOD(drvFreeDrvItemContext)(THIS_
                                   LONG lFlags,
                                   __in       BYTE *pSpecContext,
                                   __out      LONG *plDevErrVal);

  STDMETHOD(drvGetWiaFormatInfo)(THIS_
                                 __in     BYTE            *pWiasContext,
                                          LONG            lFlags,
                                 __out    LONG            *pcelt,
                                 __out    WIA_FORMAT_INFO **ppwfi,
                                 __out    LONG            *plDevErrVal);

  STDMETHOD(drvNotifyPnpEvent)(THIS_
                               __in   const GUID *pEventGUID,
                               __in   BSTR       bstrDeviceID,
                                      ULONG      ulReserved);

  STDMETHOD(drvUnInitializeWia)(THIS_ __inout BYTE *pWiasContext);

public:

  /////////////////////////////////////////////////////////////////////////
  // INonDelegating Interface Section (for all WIA drivers)              //
  /////////////////////////////////////////////////////////////////////////

  STDMETHODIMP NonDelegatingQueryInterface(REFIID  riid,LPVOID  *ppvObj);
  STDMETHODIMP_(ULONG) NonDelegatingAddRef();
  STDMETHODIMP_(ULONG) NonDelegatingRelease();

private:

  /////////////////////////////////////////////////////////////////////////
  // Minidriver private methods specific Section                         //
  /////////////////////////////////////////////////////////////////////////

  /**
  * Transfer image to IStream
  * @param[in] lFlags a Long. Is currently unused
  * @param[in] pWiasContext a pointer to BYTE. Pointer to a WIA item context
  * @param[in] pmdtc a PMINIDRV_TRANSFER_CONTEXT. Points to a MINIDRV_TRANSFER_CONTEXT structure containing the device transfer context
  * @param[in] pTransferCallback a pointer to IWiaMiniDrvTransferCallback. 
  * @param[out] plDevErrVal a pointer to LONG. Points to a memory location that will receive a status code for this method
  * @return S_OK on success
  */
  HRESULT DownloadToStream(           LONG                           lFlags,
                           __in       BYTE                           *pWiasContext,
                           __in       PMINIDRV_TRANSFER_CONTEXT      pmdtc,
                           __callback IWiaMiniDrvTransferCallback    *pTransferCallback,
                           __out      LONG                           *plDevErrVal);
  /**
  * drvAcquireItemData in case of WIA 1.0
  * @param[in] lFlags a Long. Is currently unused
  * @param[in] pWiasContext a pointer to BYTE. Pointer to a WIA item context
  * @param[in] pmdtc a PMINIDRV_TRANSFER_CONTEXT. Points to a MINIDRV_TRANSFER_CONTEXT structure containing the device transfer context
  * @param[out] plDevErrVal a pointer to LONG. Points to a memory location that will receive a status code for this method
  * @return S_OK on success
  */
  HRESULT drvAcquireItemDataWIA1(     LONG                           lFlags,
                           __in       BYTE                           *pWiasContext,
                           __in       PMINIDRV_TRANSFER_CONTEXT      pmdtc,
                           __out      LONG                           *plDevErrVal);

  /**
  * Builds WIA items tree
  * @return S_OK on success
  */
  HRESULT BuildDriverItemTree();

  /**
  * Destroy WIA items tree
  * @return S_OK on success
  */
  HRESULT DestroyDriverItemTree();

  /**
  * Execute WIA_CMD_SYNCHRONIZE
  * @param[in] pWiasContext a pointer to BYTE. Pointer to a WIA item context
  * @return S_OK on success
  */ 
  HRESULT DoSynchronizeCommand(__inout BYTE *pWiasContext);

  /**
  * This function initializes any root item properties
  * needed for this WIA driver.
  *
  * @param pWiasContext. Pointer to the WIA item context
  * @return S_OK on success
  */
  HRESULT InitializeRootItemProperties(
      __in    BYTE        *pWiasContext);

  /**
  * This function initializes child item properties
  * needed for this WIA driver.  The uiResourceID parameter
  * determines what image properties will be used.
  *
  * @param pWiasContext. Pointer to the WIA item context
  * @return S_OK on success
  */ 
  HRESULT InitializeWIAItemProperties(
      __in    BYTE        *pWiasContext,
              GUID guidItemCategory);

  /**
  * This function queues a WIA event using the passed in
  * WIA item context.
  *
  * @param pWiasContext
  *         Pointer to the WIA item context
  * @param guidWIAEvent
  *         WIA event to queue
  * @return S_OK on success
  */
  void QueueWIAEvent(
      __in    BYTE        *pWiasContext,
              const GUID  &guidWIAEvent);

  /**
  * This helper function attempts to grab a IWiaMiniDrvTransferCallback interface
  * from a PMINIDRV_TRANSFER_CONTEXT structure.
  * If successful, caller must Release.
  *
  * @param pmdtc  The PMINIDRV_TRANSFER_CONTEXT handed in during drvAcquireItemData.
  * @param ppIWiaMiniDrvTransferCallback. Address of a interface pointer which receives the callback.
  * @return HRESULT return value.
  */
  HRESULT GetTransferCallback(
      __in        PMINIDRV_TRANSFER_CONTEXT       pmdtc,
      __callback  IWiaMiniDrvTransferCallback     **ppIWiaMiniDrvTransferCallback);

  /**
  * This function creates a WIA child item
  *
  * @param wszItemName. Item name
  * @param pIWiaMiniDrv. WIA minidriver interface
  * @param pParent Parent's WIA driver item interface
  * @param lItemFlags Item flags
  * @param guidItemCategory Item category
  * @param ppChild Pointer to the newly created child item
  * @param wszStoragePath. Storage data path
  * @return HRESULT return value.
  */
  HRESULT CreateWIAChildItem(
      __in            LPOLESTR    wszItemName,
      __in            IWiaMiniDrv *pIWiaMiniDrv,
      __in            IWiaDrvItem *pParent,
                      LONG        lItemFlags,
                      GUID        guidItemCategory,
     __out_opt        IWiaDrvItem **ppChild       = NULL,
     __in_opt         const WCHAR *wszStoragePath = NULL);

  /**
  * This function creates a full WIA item name
  * from a given WIA item name.
  * The new full item name is created by concatinating
  * the WIA item name with the parent's full item name.
  * (e.g. 0000\Root + MyItem = 0000\Root\MyItem)
  *
  * @param pParent IWiaDrvItem interface of the parent WIA driver item
  * @param bstrItemName. Name of the WIA item
  * @param pbstrFullItemName. Returned full item name.  This parameter cannot be NULL.
  * @return S_OK - if successful
  */
  HRESULT MakeFullItemName(
      __in    IWiaDrvItem *pParent,
      __in    BSTR        bstrItemName,
      __out   BSTR        *pbstrFullItemName);

  /**
  * Gets scanner configuration
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetConfigParams(CTWAIN_API *pTwainApi);
  /**
  * Gets DS associated with WIA context
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[out] pTwainApi a pointer to pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetDS(BYTE *pWiasContext, CTWAIN_API** pTwainApi);
  /**
  * Deletes (close adn unload) DS associated with WIA context
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @return S_OK on success
  */   
  HRESULT DeleteDS(BYTE *pRootWiasContext);
  /**
  * Associates (create) DS associated with WIA context
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[out] pnAppID a pointer to int. Unique ID
  * @param[out] pTwainApi a pointer to pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT AddDS(BYTE *pRootWiasContext, int *pnAppID, CTWAIN_API** pTwainApi);
  /**
  * Validates WIA properties by setting TWAIN DS properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] llLst a LL_ARRAY. List of WIA propertie for validation
  * @return S_OK on success
  */  
  HRESULT ValidateThroughTWAINDS(BYTE *pWiasContext,LL_ARRAY llLst);
  /**
  * Set all TWAIN DS capabilities based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @return S_OK on success
  */  
  HRESULT SetAllTWAIN_Caps(BYTE *pWiasContext);
  /**
  * Get all TWAIN DS capabilities  and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @return S_OK on success
  */ 
  HRESULT GetAllTWAIN_Caps(BYTE *pWiasContext);
  /**
  * Get current image source
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pbFeeder a pointer to bool. True if it ADF is current image source
  * @return S_OK on success
  */ 
  HRESULT IsFeederItem(BYTE *pWiasContext, bool *pbFeeder);

  /**
  * Update WIA interdependent properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @return S_OK on success
  */  
  HRESULT UpdateWIAPropDepend(BYTE *pWiasContext);

  /**
  * Update WIA interdependent properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pContext a pointer to WIA_PROPERTY_CONTEXT. Pointer to WIA property context
  * @return S_OK on success
  */  
  HRESULT UpdateIntent(BYTE* pWiasContext, WIA_PROPERTY_CONTEXT *pContext);

  /**
  * Translate TWAIN error to WIA error
  * @param[in] dwError a DWORD. TWAIN error
  * @return translated error
  */  
  HRESULT TWAINtoWIAerror(DWORD dwError);
  
  /**
  * Loads profile and gets Custom DS data from it
  * @param[in] dwError a DWORD. TWAIN error
  * @param[out] phData a poiter to HGLOBAL. Pointer to memory contined Custom DS data
  * @param[out] pdwDataSize a poiter to DWORD. Size of the Custom DS data
  * @return S_OK on success
  */  
  HRESULT LoadProfileFromFile(BYTE *pWiasContext, HGLOBAL *phData, DWORD *pdwDataSize);

public:
  /**
  * Gets CAP_FEEDERENABLED and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetFEEDERENABLED(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets CAP_FEEDERALIGNMENT and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetFEEDERALIGNMENT(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets CAP_AUTOFEED and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetAUTOFEED(BYTE *pWiasContext, CTWAIN_API *pTwainApi){return S_OK;};
  /**
  * Gets CAP_DUPLEXENABLED and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetDUPLEXENABLED(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets ICAP_PIXELTYPE and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetPIXELTYPE(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets ICAP_BITDEPTH and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetBITDEPTH(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets ICAP_XRESOLUTION and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetXRESOLUTION(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets ICAP_YRESOLUTION and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetYRESOLUTION(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets ICAP_PIXELFLAVOR and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetPIXELFLAVOR(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets ICAP_THRESHOLD and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetTHRESHOLD(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets ICAP_BRIGHTNESS and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetBRIGHTNESS(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets ICAP_CONTRAST and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetCONTRAST(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets DAT_IMAGELAYOUT and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetDAT_IMAGELAYOUT(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Gets CAP_XFERCOUNT and update the WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */   
  HRESULT GetXFERCOUNT(BYTE *pWiasContext, CTWAIN_API *pTwainApi);

  /**
  * Sets CAP_FEEDERENABLED based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetFEEDERENABLED(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets CAP_AUTOFEED based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetAUTOFEED(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets CAP_DUPLEXENABLED based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetDUPLEXENABLED(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets ICAP_PIXELTYPE based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetPIXELTYPE(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets ICAP_BITDEPTH based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetBITDEPTH(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets ICAP_XRESOLUTION based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetXRESOLUTION(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets ICAP_YRESOLUTION based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetYRESOLUTION(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets ICAP_PIXELFLAVOR based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetPIXELFLAVOR(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets ICAP_THRESHOLD based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetTHRESHOLD(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets ICAP_BRIGHTNESS based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetBRIGHTNESS(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets ICAP_CONTRAST based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetCONTRAST(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets DAT_IMAGELAYOUT based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */  
  HRESULT SetDAT_IMAGELAYOUT(BYTE *pWiasContext, CTWAIN_API *pTwainApi);
  /**
  * Sets CAP_XFERCOUNT based on current WIA properties
  * @param[in] pWiasContext a pointer to BYTE. Pointer to WIA context
  * @param[in] pTwainApi a pointer to CTWAIN_API. 
  * @return S_OK on success
  */    
  HRESULT SetXFERCOUNT(BYTE *pWiasContext, CTWAIN_API *pTwainApi);

};
