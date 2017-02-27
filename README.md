# WIA on TWAIN
A WIA on TWAIN driver implementation designed to expose all necessary scanner functionality while 
adhering to the WIA 1.0 and WIA 2.0 specification and achieving compatibility with all WIA compliant applications.

## Features

 - Uses existing TWAIN datasource to communicate with scanner
 - Uses a MSVC wizard to generate the WIA driver
 - One binary for both WIA 1.0 and WIA 2.0
 - A Programmatic WIA interface that provides minimum operation required for obtaining Windows Logo.
 - Supports flatbed, ADF, and combo scanners
