﻿# Define the paths for the installers directories
$installersPath = "C:\Installers"
$pythonWorkshopPath = "C:\PythonWorkshop"
$wrhWorkshopPath = "C:\WRHWorkshop"

# Create the directories if they don't exist
if (-not (Test-Path $installersPath)) {
    New-Item -Path $installersPath -ItemType Directory | Out-Null
    Write-Host -ForegroundColor Green "Directory created: $installersPath"
} else {
    Write-Host -ForegroundColor Green "Directory already exists: $installersPath"
}

if (-not (Test-Path $pythonWorkshopPath)) {
    New-Item -Path $pythonWorkshopPath -ItemType Directory | Out-Null
    Write-Host -ForegroundColor Green "Directory created: $pythonWorkshopPath"
} else {
    Write-Host -ForegroundColor Green "Directory already exists: $pythonWorkshopPath"
}


if (-not (Test-Path $wrhWorkshopPath)) {
    New-Item -Path $wrhWorkshopPath -ItemType Directory | Out-Null
    Write-Host -ForegroundColor Green "Directory created: $wrhWorkshopPath"
} else {
    Write-Host -ForegroundColor Green "Directory already exists: $wrhWorkshopPath"
}

# Define the download URLs
$vsCommunityUrl = "https://aka.ms/vs/17/release/vs_community.exe"
$sevenZipUrl = "https://www.7-zip.org/a/7z2408-x64.exe"
$ssisUrl = "https://marketplace.visualstudio.com/_apis/public/gallery/publishers/SSIS/vsextensions/MicrosoftDataToolsIntegrationServices/1.5/vspackage"
$asVsixUrl = "https://marketplace.visualstudio.com/_apis/public/gallery/publishers/ProBITools/vsextensions/MicrosoftAnalysisServicesModelingProjects2022/3.0.16/vspackage"
$pythonUrl = "https://www.python.org/ftp/python/3.11.4/python-3.11.4-amd64.exe"
$vcRedistX86Url = "https://aka.ms/vs/17/release/vc_redist.x86.exe"
$vcRedistX64Url = "https://aka.ms/vs/17/release/vc_redist.x64.exe"
$msoledbSqlUrl = "https://go.microsoft.com/fwlink/?linkid=2278038"

# Define the local installer paths
$InstallerPath1 = "$installersPath\7z2408-x64.exe"
$InstallerPath2 = "$installersPath\vs_community.exe"
$InstallerPath3 = "$installersPath\Microsoft.DataTools.IntegrationServices.exe"
$InstallerPath4 = "$installersPath\Microsoft.DataTools.AnalysisServices.vsix"
$InstallerPath5 = "$installersPath\python-3.11.4-amd64.exe"
$InstallerPath6 = "$installersPath\vc_redist.x86.exe"
$InstallerPath7 = "$installersPath\vc_redist.x64.exe"
$InstallerPath8 = "$installersPath\msoledbsql.msi"

# Function to download files using Invoke-WebRequest in the background (parallel download)
function Start-Download {
    param (
        [string]$url,
        [string]$outputPath,
        [string]$name
    )

    Write-Host -ForegroundColor Yellow "Starting download for $name..."

    # Run Invoke-WebRequest in a background job
    Start-Job -ScriptBlock {
        param ($url, $outputPath, $name)
        try {
            Invoke-WebRequest -Uri $url -OutFile $outputPath -UseBasicParsing
            if (Test-Path $outputPath) {
                Write-Host -ForegroundColor Green "$name downloaded successfully."
            } else {
                Write-Host -ForegroundColor Red "$name download failed."
            }
        } catch {
            Write-Host -ForegroundColor Red "$name download failed with error: $_"
        }
    } -ArgumentList $url, $outputPath, $name
}

# Download installers with verification (in parallel)
Start-Download -url $vsCommunityUrl -outputPath $InstallerPath2 -name "vs_community.exe"
Start-Download -url $sevenZipUrl -outputPath $InstallerPath1 -name "7z2408-x64.exe"
Start-Download -url $ssisUrl -outputPath $InstallerPath3 -name "Microsoft.DataTools.IntegrationServices.exe"
Start-Download -url $asVsixUrl -outputPath $InstallerPath4 -name "Microsoft.DataTools.AnalysisServices.vsix"
Start-Download -url $pythonUrl -outputPath $InstallerPath5 -name "python-3.11.4-amd64.exe"
Start-Download -url $vcRedistX86Url -outputPath $InstallerPath6 -name "vc_redist.x86.exe"
Start-Download -url $vcRedistX64Url -outputPath $InstallerPath7 -name "vc_redist.x64.exe"
Start-Download -url $msoledbSqlUrl -outputPath $InstallerPath8 -name "msoledbsql.msi"

# Wait for all background jobs to finish
Get-Job | Wait-Job

# Removing completed jobs
Get-Job | Remove-Job

# Installers paths for installation
$vsixInstallerPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\VSIXInstaller.exe"

# Installing 7-Zip
Write-host -f Yellow "Installing 7-Zip..."
Start-Process -FilePath $InstallerPath1 -Args "/S" -Verb RunAs -Wait

# Installing Visual Studio
Write-host -f Yellow "Installing Visual Studio..."
Start-Process -FilePath $InstallerPath2 -ArgumentList "--add Microsoft.VisualStudio.Component.CoreEditor", "--add Microsoft.VisualStudio.Workload.Data", "--includeRecommended", "--includeOptional", "--quiet", "--downloadThenInstall" -Wait

# Installing Integration Services for Visual Studio
Write-host -f Yellow "Installing Integration Services for Visual Studio..."
Start-Process -FilePath $InstallerPath3 -ArgumentList '/S','/v','/qn' -Wait

# Installing Analysis Services for Visual Studio
Write-host -f Yellow "Installing Analysis Services for Visual Studio..."
Start-Process -FilePath $vsixInstallerPath -ArgumentList "/quiet", $InstallerPath4 -Wait

# Installing Python with Admin privileges and specific settings
Write-host -f Yellow "Installing Python..."
Start-Process -FilePath $InstallerPath5 -ArgumentList `
"/quiet InstallAllUsers=1 PrependPath=1 Include_doc=1 Include_pip=1 Include_tcltk=1 Include_test=1 Include_launcher=1 InstallLauncherAllUsers=1 InstallDir=""C:\Program Files\Python311"" Include_symbols=1 Include_debug=1 AssociateFiles=1 Shortcuts=1" -Wait

# Installing Visual C++ Redistributables
Write-host -f Yellow "Installing Visual C++ Redistributables (x86)..."
Start-Process -FilePath $InstallerPath6 -ArgumentList "/install", "/quiet", "/norestart" -Wait

Write-host -f Yellow "Installing Visual C++ Redistributables (x64)..."
Start-Process -FilePath $InstallerPath7 -ArgumentList "/install", "/quiet", "/norestart" -Wait

# Installing Microsoft OLE DB Driver for SQL Server
Write-host -f Yellow "Installing Microsoft OLE DB Driver for SQL Server..."
Start-Process -FilePath $InstallerPath8 -ArgumentList "/quiet", "/qn" -Wait

Write-Host -ForegroundColor Green "All installations and setups completed successfully."

# Define the sequence of commands
$commands = @"
SET PATH=%PATH%;C:\Program Files\Python311\Scripts;C:\Program Files\Python311\
cd C:\PythonWorkshop
python -m venv .venv
start "Environment setup" cmd /k "C:\PythonWorkshop\.venv\Scripts\activate.bat && pip install pandas && pip install pyodbc && pip install xlrd && pip install openpyxl && pip install datetime && pip install polars && pip install matplotlib && pip install jupyter"
cd C:\Users\sqladmin
pip install pandas pyodbc
"@

# Write the commands to a temporary batch file
$tempBatchFile = "$env:TEMP\run_python_setup.bat"
Set-Content -Path $tempBatchFile -Value $commands

# Run the batch file using cmd.exe in sequence
Start-Process -FilePath "cmd.exe" -ArgumentList "/k", $tempBatchFile -Wait

# Clean up - Remove the temporary batch file after execution
Remove-Item -Path $tempBatchFile

# Write the commands to a temporary batch file
$tempBatchFile = "$env:TEMP\run_python_setup.bat"
Set-Content -Path $tempBatchFile -Value $commands

# Run the batch file using cmd.exe in sequence
Start-Process -FilePath "cmd.exe" -ArgumentList "/k", $tempBatchFile -Wait

# Clean up - Remove the temporary batch file after execution
Remove-Item -Path $tempBatchFile

# -------------------------------------------
# 1. Create the directory for warehouse files
# -------------------------------------------
$directoryPath = "C:\WRHWorkshop\wsbwarehousefiles"
if (!(Test-Path -Path $directoryPath)) {
    New-Item -ItemType Directory -Path $directoryPath
    Write-Host "Created directory at $directoryPath."
} else {
    Write-Host "Directory $directoryPath already exists."
}

# -------------------------------------------
# 2. Download files from GitHub into the directory
# -------------------------------------------
$filesToDownload = @(
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/CustomerTransactions2016.txt"; destination="CustomerTransactions2016.txt"},
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/CustomerTransactions2016Load.py"; destination="CustomerTransactions2016Load.py"},
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/CustomerTransactions2019.txt"; destination="CustomerTransactions2019.txt"},
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/CustomerTransactions2019Load.py"; destination="CustomerTransactions2019Load.py"},
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/DataTableLoad.py"; destination="DataTableLoad.py"},
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/DateTable.txt"; destination="DateTable.txt"},
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/InvoiceLines2015.txt"; destination="InvoiceLines2015.txt"},
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/PurchaseOrderLines2020.txt"; destination="PurchaseOrderLines2020.txt"},
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/SupplierTransactions2021.txt"; destination="SupplierTransactions2021.txt"}
)


foreach ($file in $filesToDownload) {
    $targetFilePath = Join-Path $directoryPath $file.destination
    Write-Host "Downloading file from $($file.url) to $targetFilePath..."
    Invoke-WebRequest -Uri $file.url -OutFile $targetFilePath
}

# Download destination.bacpac and source.bacpac into C:\WRHWorkshop
$bacpacFilesToDownload = @(
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/destination.bacpac"; destination="C:\WRHWorkshop\destination.bacpac"},
    @{url="https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbWarehouseETL/refs/heads/master/source.bacpac"; destination="C:\WRHWorkshop\source.bacpac"}
)

foreach ($bacpac in $bacpacFilesToDownload) {
    Write-Host "Downloading .bacpac file from $($bacpac.url) to $($bacpac.destination)..."
    Invoke-WebRequest -Uri $bacpac.url -OutFile $bacpac.destination
}

# -------------------------------------------
# 3. Define the path to sqlpackage.exe for SQL Server 2022. Update this path if sqlpackage.exe is located elsewhere.
# -------------------------------------------
$sqlpackagePath = "C:\Program Files\Microsoft Visual Studio\2022\Community\Common7\IDE\Extensions\Microsoft\SQLDB\DAC\sqlpackage.exe"

# Parameters for the two .bacpac files
$destinationBacpacFilePath = "C:\WRHWorkshop\destination.bacpac"  # Path to the first .bacpac file (destination.bacpac)
$destinationDatabaseName = "DestinationDatabase"                  # Target database name for destination.bacpac

$sourceBacpacFilePath = "C:\WRHWorkshop\source.bacpac"            # Path to the second .bacpac file (source.bacpac)
$sourceDatabaseName = "SourceDatabase"                            # Target database name for source.bacpac

$serverName = "localhost"                                         # SQL Server instance (use "." for local)
$sqlAuthUsername = ""                                             # SQL Authentication username (leave empty for Windows Auth)
$sqlAuthPassword = ""                                             # SQL Authentication password (leave empty for Windows Auth)

# If using SQL Authentication, include the username and password in the command
$authOptions = @()
if ($sqlAuthUsername -ne "" -and $sqlAuthPassword -ne "") {
    $authOptions += "/TargetUser:$sqlAuthUsername"
    $authOptions += "/TargetPassword:$sqlAuthPassword"
}

# SSL and encryption options passed as individual arguments
$sslOptions = @(
    "/TargetTrustServerCertificate:True",  # Trust the server's certificate
    "/TargetEncryptConnection:False"       # Disable encryption
)

# -------------------------------------------
# 4. Import the destination.bacpac into the DestinationDatabase
# -------------------------------------------
$destinationArguments = @(
    "/Action:Import",
    "/SourceFile:$destinationBacpacFilePath",       # Use the destination.bacpac file
    "/TargetServerName:$serverName",
    "/TargetDatabaseName:$destinationDatabaseName"
) + $sslOptions + $authOptions

Write-Host "Importing the .bacpac file '$destinationBacpacFilePath' into the database '$destinationDatabaseName'..."
& $sqlpackagePath $destinationArguments

# Check if the import was successful
if ($LASTEXITCODE -eq 0) {
    Write-Host "Data-Tier Application import completed successfully from '$destinationBacpacFilePath' to '$destinationDatabaseName'."
} else {
    Write-Host "Data-Tier Application import failed with exit code $LASTEXITCODE."
    exit $LASTEXITCODE  # Exit if the first import fails
}

# -------------------------------------------
# 5. Import the source.bacpac into the SourceDatabase
# -------------------------------------------
$sourceArguments = @(
    "/Action:Import",
    "/SourceFile:$sourceBacpacFilePath",            # Use the source.bacpac file
    "/TargetServerName:$serverName",
    "/TargetDatabaseName:$sourceDatabaseName"
) + $sslOptions + $authOptions

Write-Host "Importing the .bacpac file '$sourceBacpacFilePath' into the database '$sourceDatabaseName'..."
& $sqlpackagePath $sourceArguments

# Check if the import was successful
if ($LASTEXITCODE -eq 0) {
    Write-Host "Data-Tier Application import completed successfully from '$sourceBacpacFilePath' to '$sourceDatabaseName'."
} else {
    Write-Host "Data-Tier Application import failed with exit code $LASTEXITCODE."
}
