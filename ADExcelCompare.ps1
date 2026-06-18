$EnvFilePath = ".\.env.local"

If (-not (Test-Path $EnvFilePath -PathType Leaf)) {
    Write-Error "Error: .env.local file not found at $($EnvFilePath)"
    Exit 1
}

$Env = @{}
$EnvLines = Get-Content $EnvFilePath

ForEach ($Line in $EnvLines) {
    $TrimmedLine = $Line.Trim()

    If ([string]::IsNullOrEmpty($TrimmedLine) -or $TrimmedLine.StartsWith('#')) {
        Continue
    }

    $Name, $Value = $TrimmedLine -split '=', 2
    $Env.Add($Name, $Value)
}

$ReportPath = $Env["REPORT_URL"]
$ExcelPackage = Open-ExcelPackage -Path $ReportPath
$SheetNames = $ExcelPackage.Workbook.Worksheets.Name
$ExcelSheetsComputerList = @()
$TrackedSheets = @($Env["TRACKED_SHEETS"].Split(','))

ForEach ($Sheet in $SheetNames) {
    If ($TrackedSheets -contains $Sheet) {
        $Names = Import-Excel -Path $ReportPath -WorksheetName $Sheet | Select-Object -ExpandProperty Name
        $ExcelSheetsComputerList += $Names 
    }   
}

$ADComputerList = @(Get-ADComputer -Filter * -SearchBase $Env["SEARCH_BASE"] | Select-Object -ExpandProperty Name)

ForEach ($ExcelComputer in $ExcelSheetsComputerList) {
    If (-not ($ADComputerList -contains $ExcelComputer)) {
        Write-Host $ExcelComputer "is present in Excel but not in Active Directory!" -BackgroundColor Green
    }
}

ForEach ($ADComputer in $ADComputerList) {
    If (-not ($ExcelSheetsComputerList -contains $ADComputer)) {
        Write-Host $ADComputer "is present in Active Directory but not in Excel!" -BackgroundColor DarkYellow
    }
}