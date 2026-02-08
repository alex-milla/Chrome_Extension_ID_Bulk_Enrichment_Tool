# Chrome Extension ID Bulk Enrichment Tool

PowerShell tool to enrich Chrome Extension IDs with their names and status by querying the Chrome Web Store.

## Requirements

- PowerShell 5.1 or higher
- Internet connection

## Usage

```powershell
.\Bulk-Extension-Enrichment.ps1
```

The script will prompt you for:

1. **Input file path** (local file or URL)
   - Example local: `C:\Users\security\extensions.txt`
   - Example URL: `https://raw.githubusercontent.com/user/repo/main/extensions.txt`

2. **Output file name** (optional, defaults to timestamped filename)

## Input Format

One Extension ID per line (TXT or CSV):

```text
mdaboflcmhejfihjcbmdiebgfchigjcf
gaoflciahikhligngeccdecgfjngejlh
fedimamkpgiemhacbdhkkaihgofncola
```

## Output Format

CSV file with four columns:

| ExtensionID | ExtensionName | Status | ChromeStoreURL |
|-------------|---------------|--------|----------------|
| mdaboflcmhejfihjcbmdiebgfchigjcf | Blipshot: One-Click Full Page Screenshot | Active | https://chrome.google.com/webstore/detail/mdaboflcmhejfihjcbmdiebgfchigjcf |
| gaoflciahikhligngeccdecgfjngejlh | Removed/NotFound | Removed | https://chrome.google.com/webstore/detail/gaoflciahikhligngeccdecgfjngejlh |
| fedimamkpgiemhacbdhkkaihgofncola | WAToolkit | Active | https://chrome.google.com/webstore/detail/fedimamkpgiemhacbdhkkaihgofncola |

**Status values:**
- `Active`: Extension found in Chrome Web Store
- `Removed`: Extension no longer available
- `Unknown`: Extension exists but name unavailable
- `Error`: Connection or query error

**Note:** All four columns are always generated for consistency, ensuring compatibility with CSV parsers and SIEM tools.

## Features

- Load extension IDs from local files or URLs
- Generates consistent 4-column CSV output
- Real-time progress with percentage and color-coded output
- Automatic rate limiting (500ms between requests)
- Summary statistics upon completion
- Timestamped output files
- Chrome Web Store URL included for quick validation

## Example

```powershell
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
Generated file: malicious_extensions_enriched_20260208_175300.csv
```

## Author

**Alex Milla**
- [alexmilla.dev](https://alexmilla.dev)

## License

MIT License
