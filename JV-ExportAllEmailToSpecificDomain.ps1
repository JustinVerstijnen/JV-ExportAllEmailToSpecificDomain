# Justin Verstijnen Export All Email Messages To Specific Domain script
# Github page: https://github.com/JustinVerstijnen/JV-ExportAllEmailToSpecificDomain
# Let's start!
Write-Host "Script made by..." -ForegroundColor DarkCyan
Write-Host "     _           _   _        __     __            _   _  _                  
    | |_   _ ___| |_(_)_ __   \ \   / /__ _ __ ___| |_(_)(_)_ __   ___ _ __  
 _  | | | | / __| __| | '_ \   \ \ / / _ \ '__/ __| __| || | '_ \ / _ \ '_ \ 
| |_| | |_| \__ \ |_| | | | |   \ V /  __/ |  \__ \ |_| || | | | |  __/ | | |
 \___/ \__,_|___/\__|_|_| |_|    \_/ \___|_|  |___/\__|_|/ |_| |_|\___|_| |_|
                                                       |__/                  " -ForegroundColor DarkCyan

# === PARAMETERS ===
$StartDate = (Get-Date).AddDays(-10) # 10 days is the limit for instant reporting
$EndDate = Get-Date
$Domain = "microsoft.com"
$ExportPad = Join-Path -Path $PSScriptRoot -ChildPath "ReportSentTo_$Domain.csv"
# === END PARAMETERS ===

# Step 1: Connecting to Exchange Online
Connect-ExchangeOnline -ShowBanner:$false

# Step 2: Quering and Exporting all needed data
Write-Host "Quering messages sent to $Domain between $StartDate and $EndDate..." -ForegroundColor Yellow
$traces = Get-MessageTraceV2 -StartDate $StartDate -EndDate $EndDate -ResultSize 5000
$filteredTraces = $traces | Where-Object {
    $_.RecipientAddress -like "*@$Domein"
}
# Check if there are results
if ($filteredTraces.Count -eq 0) {
    Write-Host "No sent email to $Domain in this period." -ForegroundColor Cyan
    [PSCustomObject]@{
        Received         = ""
        SenderAddress    = ""
        RecipientAddress = ""
        Subject          = "No sent email found"
        Status           = ""
    } | Export-Csv -Path $ExportPad -NoTypeInformation -Encoding UTF8 -Delimiter ";"
}
else {
    $filteredTraces |
        Sort-Object Received -Descending |
        Select-Object Received, SenderAddress, RecipientAddress, Subject, Status |
        Export-Csv -Path $ExportPad -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    Write-Host "Exported file to $ExportPad" -ForegroundColor Green
    Start-Sleep -Seconds 10
}
