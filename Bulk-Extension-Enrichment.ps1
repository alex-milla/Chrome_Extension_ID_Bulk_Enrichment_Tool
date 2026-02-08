# Request file path from user
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "  Extension ID Enrichment with Names" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Please enter the path to the file containing Extension IDs" -ForegroundColor Yellow
Write-Host "The file must contain one Extension ID per line (TXT or CSV format)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Examples:" -ForegroundColor Gray
Write-Host "  - C:\Users\alex\Desktop\extension_ids.txt" -ForegroundColor Gray
Write-Host "  - .\malicious_extensions.csv" -ForegroundColor Gray
Write-Host "  - https://raw.githubusercontent.com/user/repo/main/file.txt" -ForegroundColor Gray
Write-Host ""

$inputPath = Read-Host "File path (local or URL)"

# Validate if it's a URL or local file
if ($inputPath -match "^https?://") {
    Write-Host "`nLoading Extension IDs from URL..." -ForegroundColor Green
    try {
        $content = (Invoke-WebRequest -Uri $inputPath -UseBasicParsing).Content
        $extensionIds = $content -split "`n" | Where-Object { $_.Trim() -ne "" }
    } catch {
        Write-Host "Error loading URL: $_" -ForegroundColor Red
        exit
    }
} else {
    # Check if file exists
    if (-Not (Test-Path $inputPath)) {
        Write-Host "`nError: File does not exist at the specified path." -ForegroundColor Red
        exit
    }
    
    Write-Host "`nLoading Extension IDs from local file..." -ForegroundColor Green
    $extensionIds = Get-Content -Path $inputPath | Where-Object { $_.Trim() -ne "" }
}

Write-Host "Total Extension IDs found: $($extensionIds.Count)" -ForegroundColor Cyan
Write-Host ""

# Request output path
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$defaultOutput = "malicious_extensions_enriched_$timestamp.csv"
Write-Host "Enter output file name (press Enter to use: $defaultOutput)" -ForegroundColor Yellow
$outputPath = Read-Host "Output file"
if ([string]::IsNullOrWhiteSpace($outputPath)) {
    $outputPath = $defaultOutput
}

Write-Host "`nStarting extension enrichment..." -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""

$results = @()
$counter = 0

foreach ($id in $extensionIds) {
    $id = $id.Trim()
    $counter++
    
    # Initialize default values for all fields
    $extensionName = ""
    $status = ""
    $chromeStoreUrl = "https://chrome.google.com/webstore/detail/$id"
    
    try {
        $response = Invoke-WebRequest -Uri $chromeStoreUrl -UseBasicParsing -ErrorAction SilentlyContinue -TimeoutSec 10
        
        if ($response.StatusCode -eq 200 -and $response.Content -notmatch 'ItemNotFound') {
            # Extract extension title
            if ($response.Content -match '<title>(.*?)</title>') {
                $extensionName = $matches[1] -replace ' - Chrome Web Store', '' -replace ' - Chrome ウェブストア', ''
                $status = "Active"
            } else {
                $extensionName = 'Unknown'
                $status = "Unknown"
            }
        } else {
            $extensionName = 'Removed/NotFound'
            $status = "Removed"
        }
    } catch {
        $extensionName = 'Error/Unavailable'
        $status = "Error"
    }
    
    # Always create object with all 4 columns
    $results += [PSCustomObject]@{
        ExtensionID = $id
        ExtensionName = $extensionName
        Status = $status
        ChromeStoreURL = $chromeStoreUrl
    }
    
    # Show progress
    $percentage = [math]::Round(($counter / $extensionIds.Count) * 100, 2)
    Write-Host "[$counter/$($extensionIds.Count)] ($percentage%) - $id - $extensionName" -ForegroundColor $(if ($status -eq "Active") { "Green" } elseif ($status -eq "Removed") { "Red" } else { "Yellow" })
    
    # Rate limiting to avoid blocks
    Start-Sleep -Milliseconds 500
}

# Export results with explicit column order
$results | Select-Object ExtensionID, ExtensionName, Status, ChromeStoreURL | Export-Csv -Path $outputPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host "Process completed successfully" -ForegroundColor Green
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Summary:" -ForegroundColor Yellow
Write-Host "  - Total processed: $($results.Count)" -ForegroundColor White
Write-Host "  - Active: $(($results | Where-Object {$_.Status -eq 'Active'}).Count)" -ForegroundColor Green
Write-Host "  - Removed: $(($results | Where-Object {$_.Status -eq 'Removed'}).Count)" -ForegroundColor Red
Write-Host "  - Unknown/Error: $(($results | Where-Object {$_.Status -in @('Unknown','Error')}).Count)" -ForegroundColor Yellow
Write-Host ""
Write-Host "Generated file: $outputPath" -ForegroundColor Cyan
