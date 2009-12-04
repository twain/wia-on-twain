/*****************************************************************************
 *
 *  wiacapabilitymanager.h
 *
 *  Copyright (c) 2003 Microsoft Corporation.  All Rights Reserved.
 *
 *  DESCRIPTION:
 *
 *  Contains class declaration for CWIACapabilityManager
 *  
 *******************************************************************************/

#pragma once

#define MAX_CAPABILITY_STRING_SIZE_BYTES (sizeof(WCHAR) * MAX_PATH)

class CWIACapabilityManager {
public:
    CWIACapabilityManager();
    ~CWIACapabilityManager();
public:
    HRESULT Initialize(__in HINSTANCE hInstance);
    void Destroy();
    HRESULT AddCapability(const GUID    guidCapability,
                          UINT          uiNameResourceID,
                          UINT          uiDescriptionResourceID,
                          ULONG         ulFlags,
                          __in LPWSTR   wszIcon);
    HRESULT DeleteCapability(const GUID guidCapability,ULONG ulFlags);
    HRESULT AllocateCapability(__out WIA_DEV_CAP_DRV **ppWIADeviceCapability);
    void FreeCapability(__in WIA_DEV_CAP_DRV *pWIADeviceCapability, BOOL bFreeCapabilityContentOnly = FALSE);
    LONG GetNumCapabilities();
    LONG GetNumCommands();
    LONG GetNumEvents();

    WIA_DEV_CAP_DRV* GetCapabilities();
    WIA_DEV_CAP_DRV* GetCommands();
    WIA_DEV_CAP_DRV* GetEvents();
private:
    HINSTANCE                            m_hInstance;
    CBasicDynamicArray<WIA_DEV_CAP_DRV> m_CapabilityArray;
};

