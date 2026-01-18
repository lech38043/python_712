# -------------------------------------------

# 1. Create the directory for warehouse files

# -------------------------------------------

$directoryPath = "C:\PythonWorkshop\wsbpythonfiles"

if (!(Test-Path -Path $directoryPath)) {

    New-Item -ItemType Directory -Path $directoryPath

    Write-Host "Created directory at $directoryPath."

} else {

    Write-Host "Directory $directoryPath already exists."

}
 
# -------------------------------------------

# 2. Download files from GitHub into the directory using curl

# -------------------------------------------

$filesToDownload = @(

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.001",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.002",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.003",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.004",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.005",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.006",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.007",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.008",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.009",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.010",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/ZadaniaGitHub.7z.011",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - zadania kontekst danych.ipynb",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - wprowadzenie.pdf",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - wprowadzenie - Jupyter.ipynb",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - sredni - kontekst danych.ipynb",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - podstawy.pdf",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - podstawy zadania.pdf",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - podstawy zadania - Jupyter.ipynb",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Python - podstawy - Jupyter.ipynb",

    "https://raw.githubusercontent.com/MichalZycki/MichalZycki-WsbPython/main/Zajecia.7z"

)
 
foreach ($url in $filesToDownload) {

    $fileName = [System.IO.Path]::GetFileName($url)

    $destination = Join-Path $directoryPath $fileName

    Write-Host "Downloading $fileName to $destination using curl..."

    # Using curl to download the file

    curl -o $destination $url

}
 
# -------------------------------------------

# 3. Extract files using 7z

# -------------------------------------------

# Path to the 7z executable (assumes 7z is in PATH)

$sevenZipPath = "C:\Program Files\7-Zip\7z.exe"
 
# Extract ZadaniaGitHub.7z.001 and related parts

$zadaniaFilePath = Join-Path $directoryPath "ZadaniaGitHub.7z.001"

if (Test-Path -Path $zadaniaFilePath) {

    Write-Host "Extracting ZadaniaGitHub.7z.001 and related parts..."
& $sevenZipPath x $zadaniaFilePath -o"$directoryPath\Zadania" -y

    Write-Host "Extraction of ZadaniaGitHub completed."

} else {

    Write-Host "File ZadaniaGitHub.7z.001 not found. Skipping extraction."

}
 
# Extract Zajecia.7z

$zajeciaFilePath = Join-Path $directoryPath "Zajecia.7z"

if (Test-Path -Path $zajeciaFilePath) {

    Write-Host "Extracting Zajecia.7z..."
& $sevenZipPath x $zajeciaFilePath -o"$directoryPath\Zajecia" -y

    Write-Host "Extraction of Zajecia completed."

} else {

    Write-Host "File Zajecia.7z not found. Skipping extraction."

}
 