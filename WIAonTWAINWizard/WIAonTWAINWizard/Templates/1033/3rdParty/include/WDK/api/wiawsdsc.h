/**************************************************************************
*
*  Copyright (c) Microsoft Corporation
*
*  File: wiawsdsc.h
*
*  Version: 1.0 
*
*  Description: contains custom WIA definitions for the WSD scan class driver 
*
***************************************************************************/

#ifndef _WIAWSDSC_
#define _WIAWSDSC_

#ifndef _WIADEF_
#include <wiadef.h>
#endif

#if (_WIN32_WINNT >= 0x0600)

//
// Custom WIA_IPA_ITEM_CATEGORY value for the Auto WIA item:
//
// WIA_CATEGORY_AUTO {DEFE5FD8-6C97-4dde-B11E-CB509B270E11}
//
// The Auto item represents a device auto-configured input source:
//
// - Implemented as a custom WIA item directly off the Root item. 
// - Coexists on the WIA Item Tree with the Flatbed and/or Feeder 
//   programmable input source items, never alone.
// - Implemented for scanner devices with one input source to allow 
//   automatic configuration of the scan parameters for this source.
// - Implemented for scanners with multiple input sources to allow 
//   automatic selection of the input source -plus- automatic source 
//   configuration.
// - The only programmable scan setting is the transfer file format  
//   (including the single/multi-page differentiator and compression).
//   The file format programmable setting is provided for application
//   compatibility (the application can choose the format it needs).
//   It is not recommended to break the UI-less rule for automatic
//   push-scan by displaying an user prompt at run-time asking the 
//   user to select a transfer format - this selection shall be made
//   programmatically (without UI input required) from the application. 
// - Stream transfer enabled similar to the Flatbed and Feeder items.
// - As a semi-programmable data source uses the same item flags 
//   the Flatbed and Feeder items use, minus WiaItemTypeImage: 
//   WiaItemTypeTransfer | WiaItemTypeFile | WiaItemTypeProgrammableDataSource  
// - Has no children items (e.g. does not support segmentation).
// - Not translatatble by the WIA Compatibility Layer, 
//   accessible from Version 2.0 WIA applications only. 
// 
// The Auto item implements the following minimal set of WIA properties, 
// along with regular WIA Service managed properties (WIA_IPA_ITEM_NAME, 
// WIA_IPA_FULL_ITEM_NAME, WIA_IPA_ITEM_FLAGS, WIA_IPA_ACCESS_RIGHTS, 
// WIA_IPA_COLOR_PROFILE), WIA_IPA_ITEM_CATEGORY (set to WIA_CATEGORY_AUTO), 
// and WIA_IPA_ITEM_SIZE (set to 0), all following the WIA requirements. 
// The default values at the beginning of a new session are the WIA 
// required values: WiaImgFmt_BMP, TYMED_FILE, WIA_COMPRESSION_NONE.
// For WIA_IPA_PREFERRED_FORMAT the driver reports WiaImgFmt_PNG if
// the device supports the PNG format, WiaImgFmt_BMP otherwise.
//
// WIA_IPA_TYMED
// WIA_IPA_FORMAT
// WIA_IPA_COMPRESSION
// WIA_IPA_FILENAME_EXTENSION
// WIA_IPA_PREFERRED_FORMAT
//
// These properties allow the WIA application client to enumerate the valid
// transfer file formats and select the file format to receive data scanned 
// from the Auto item, all other scan settings being selected by the device.
//
DEFINE_GUID(WIA_CATEGORY_AUTO, 0xdefe5fd8, 0x6c97, 0x4dde, 0xb1, 0x1e, 0xcb, 0x50, 0x9b, 0x27, 0x0e, 0x11);

#endif //#if (_WIN32_WINNT >= 0x0600)

//
// Custom WIA_DPS_DOCUMENT_HANDLING_CAPABILITIES flag value 
// describing when the Auto input source item is supported:
//

#define AUTO_SOURCE 0x8000

//
// Custom WIA property IDs (see wiadef.h) 
//
// These custom properties describe PnP-X device properties 
// read at run time from Function Discovery, along with:
// 
// WIA_DPS_SERVICE_ID      
// WIA_DPS_DEVICE_ID
// WIA_DPS_GLOBAL_IDENTITY 
// WIA_DPS_FIRMWARE_VERSION
// 
// All are read-only Root item properties maintained by the driver.
//
// Property Type: VT_BSTR
// Valid Values: WIA_PROP_NONE
// Access Rights: READONLY
//

#define WIA_WSD_MANUFACTURER             WIA_PRIVATE_DEVPROP
#define WIA_WSD_MANUFACTURER_STR         L"Device manufacturer"

#define WIA_WSD_MANUFACTURER_URL         (WIA_PRIVATE_DEVPROP + 1)
#define WIA_WSD_MANUFACTURER_URL_STR     L"Manufacurer URL"

#define WIA_WSD_MODEL_NAME               (WIA_PRIVATE_DEVPROP + 2)
#define WIA_WSD_MODEL_NAME_STR           L"Model name"

#define WIA_WSD_MODEL_NUMBER             (WIA_PRIVATE_DEVPROP + 3)
#define WIA_WSD_MODEL_NUMBER_STR         L"Model number"

#define WIA_WSD_MODEL_URL                (WIA_PRIVATE_DEVPROP + 4)
#define WIA_WSD_MODEL_URL_STR            L"Model URL"

#define WIA_WSD_PRESENTATION_URL         (WIA_PRIVATE_DEVPROP + 5)
#define WIA_WSD_PRESENTATION_URL_STR     L"Presentation URL"

#define WIA_WSD_FRIENDLY_NAME            (WIA_PRIVATE_DEVPROP + 6)
#define WIA_WSD_FRIENDLY_NAME_STR        L"Friendly name"

#define WIA_WSD_SERIAL_NUMBER            (WIA_PRIVATE_DEVPROP + 7)
#define WIA_WSD_SERIAL_NUMBER_STR        L"Serial number"

//
// Custom WIA property for automatic input-source selection
// during programmed push (device initiated) scanning:
//
// WIA_WSD_SCAN_AVAILABLE_ITEM
// 
// Read-only Root item property maintained by the driver.
//
// Property Type: VT_BSTR
// Valid Values: WIA_PROP_NONE
// Access Rights: READONLY
//
// Following a scan event the current property value is a 
// WIA item name (exactly as reported by WIA_IPA_ITEM_NAME) 
// describing the item where the scan job is available from
// if this information is known or an empty string otherwise.
// Immediately as a scan event is consumed (the device is no 
// longer in a scan available signaled state) the current value 
// is reset by the driver to an empty string.
//
// Note regarding backwards compatibility for v1.0 WIA:
//
// The WIA_WSD_SCAN_AVAILABLE_ITEM property is available 
// to be read from the Root item by legacy (Version 1.0)
// WIA applications. The application can translate the
// item name value to a WIA_DPS_DOCUMENT_HANDLING_SELECT
// flag and use it to select the right scan input source:
//
// WIA_WSD_SCAN_AVAILABLE_ITEM - WIA_DPS_DOCUMENT_HANDLING_SELECT
// --------------------------------------------------------------
// L"Flatbed"                  - FLATBED (defined in wiadef.h)
// L"Feeder"                   - FEEDER (defined in wiadef.h)
//

#define WIA_WSD_SCAN_AVAILABLE_ITEM      (WIA_PRIVATE_DEVPROP + 8)
#define WIA_WSD_SCAN_AVAILABLE_ITEM_STR  L"Scan available from"

#endif //_WIAWSDSC_
