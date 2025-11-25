# Purview Label Extractor
This PowerShell script extracts Microsoft Purview sensitivity label metadata stored inside Office files. The script reads the custom properties section of a DOCX, XLSX, or PPTX file and outputs all Purview label attributes, including label GUID, label name, tenant ID, method, and timestamp.

This information can be used by a DLP tool to identify and track the movement of files that contain matching metadata.

## Features
- Extracts all Purview sensitivity label metadata stored in the file
- Supports DOCX, XLSX, and PPTX

## Requirements
- Windows 10 or later
- PowerShell 5.1 or later on Windows
- .NET Framework 4.5 or later on Windows

## Limitations
The PowerShell script cannot extract metadata from encrypted Purview files.

## Usage
1. Clone or download this repository.
2. Run the script, e.g., `.\ExtractLabel.ps1 -FilePath "example.docx"`

## Example Output
```
name                                                        value
----                                                        -----
MSIP_Label_1b69da68-1579-44b2-a335-346afa06d22b_Name        Highly Confidential
MSIP_Label_1b69da68-1579-44b2-a335-346afa06d22b_SetDate     2025-06-18T15:44:58Z
MSIP_Label_1b69da68-1579-44b2-a335-346afa06d22b_Method      Privileged
MSIP_Label_1b69da68-1579-44b2-a335-346afa06d22b_SiteId      fa8def79-9770-4433-86fd-3db3dd80d4dc
```

## License
Licensed under MIT — free to use, modify, and share, with no warranty.
