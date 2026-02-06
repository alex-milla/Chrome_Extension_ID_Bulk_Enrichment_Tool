# Chrome Extension ID Bulk Enrichment Tool

PowerShell tool to enrich Chrome Extension IDs with their names and status by querying the Chrome Web Store.

## Requirements

- PowerShell 5.1 or higher
- Internet connection

## Usage

```powershell
.\Bulk-Extension-Enrichment.ps1
The script will prompt you for:

Input file path (local file or URL)

Example local: C:\Users\security\extensions.txt

Example URL: https://raw.githubusercontent.com/user/repo/main/extensions.txt

Output file name (optional, defaults to timestamped filename)

Input Format
One Extension ID per line (TXT or CSV):

text
mdaboflcmhejfihjcbmdiebgfchigjcf
gaoflciahikhligngeccdecgfjngejlh
fedimamkpgiemhacbdhkkaihgofncola
Output Format
CSV file with three columns:

ExtensionID	ExtensionName	Status
mdaboflcmhejfihjcbmdiebgfchigjcf	Blipshot: One-Click Full Page Screenshot	Active
gaoflciahikhligngeccdecgfjngejlh	Removed/NotFound	Removed
fedimamkpgiemhacbdhkkaihgofncola	WAToolkit	Active
Status values:

Active: Extension found in Chrome Web Store

Removed: Extension no longer available

Unknown: Extension exists but name unavailable

Error: Connection or query error

Features
Load extension IDs from local files or URLs

Real-time progress with percentage and color-coded output

Automatic rate limiting (500ms between requests)

Summary statistics upon completion

Timestamped output files

Example
powershell
.\Bulk-Extension-Enrichment.ps1

# Input
File path (local or URL): https://raw.githubusercontent.com/alex-milla/Chrome-Extensions/refs/heads/main/Malicious_Chrome_Extensions_ID.txt
Output file: [Press Enter for default]

# Output
Loading Extension IDs from URL...
Total Extension IDs found: 292
Starting extension enrichment...
[1/292] (0.34%) - mdaboflcmhejfihjcbmdiebgfchigjcf - Blipshot: One-Click Full Page Screenshot
...
Process completed successfully
Total processed: 292
Active: 45
Removed: 237
Unknown/Error: 10
Generated file: malicious_extensions_enriched_20260207_001234.csv
Use with Microsoft Sentinel
Load the generated CSV in KQL queries:

text
let MaliciousExtensions = externaldata(ExtensionID:string, ExtensionName:string, Status:string)
[@"https://raw.githubusercontent.com/user/repo/main/malicious_extensions.csv"]
with (format="csv", ignoreFirstRecord=true);
DeviceEvents
| where AdditionalFields has "extensionId"
| extend ExtensionData = parse_json(AdditionalFields)
| extend ExtensionID = tostring(ExtensionData.extensionId)
| join kind=inner (MaliciousExtensions) on ExtensionID
| project TimeGenerated=Timestamp, DeviceName, AccountName=InitiatingProcessAccountName, ExtensionID, ExtensionName, Status
Author
Alex Milla

alexmilla.net

cibersecurity.io

cyberintelligence.dev

License
MIT License
