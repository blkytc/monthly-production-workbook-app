# Windows XLS Conversion And Theme Design

## Goal

Make the Windows release open supported legacy `.xls` templates when LibreOffice, Microsoft Excel, or WPS Office is installed, and give the desktop application a consistent dark appearance on Windows and macOS.

## Conversion

`convert_xls` keeps LibreOffice as the first conversion method because it works on every supported operating system. On Windows, when LibreOffice is unavailable, it invokes a bundled PowerShell script that tries the installed spreadsheet COM applications in this order: Microsoft Excel, WPS `ket.Application`, and WPS `et.Application`. Each application opens the source without showing a window and saves an `.xlsx` copy using format code 51. The program verifies that the expected output exists before continuing.

If all available converters fail, the error explains that Excel, WPS, or LibreOffice must be installed and includes the converter failure detail. The application never edits the source `.xls` file.

## Theme

The application uses Tk's `clam` theme as a portable base and configures a restrained dark palette for frames, labels, entries, buttons, comboboxes, label frames, and tree views. Native file dialogs and message boxes remain operating-system controls. The theme setup is a standalone function so its palette and selected states can be tested without constructing the full application.

## Packaging And Validation

The PowerShell conversion script is packaged as application data by PyInstaller. Unit tests cover converter selection, Windows fallback invocation, failure messages, and theme configuration. GitHub Actions runs the full test suite and builds both platforms. Release `v1.0.3` contains only the Windows and macOS ZIP files; the Windows ZIP is downloaded and checked for the executable before delivery.
