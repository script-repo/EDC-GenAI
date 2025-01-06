# Updated Powershell Script for Processing OBS Recordings with Fabric Post-Processing

# This script adds functionality to use 'fabric' (an open-source tool by Daniel Miessler) to perform additional actions after Whisper transcription.
# Specifically, it runs the pattern 'extract_wisdom' on the generated text file and outputs it to a new file named 'wisdom.txt'.
# Modify the $InputDirectory parameter to your own base directory where your mkv videos are.

param (
    [string]$InputDirectory = "C:\Users\UserName\Videos"
)

# Ensure InputDirectory is provided and is valid
if (-not (Test-Path $InputDirectory)) {
    Write-Error "Input directory does not exist. Please provide a valid directory path."
    exit 1
}

# Check if whisper-cli is installed
if (-not (Get-Command whisper -ErrorAction SilentlyContinue)) {
    Write-Error "Whisper CLI is not installed or not available in PATH. Please install it to proceed."
    exit 1
}

# Check if fabric is installed
if (-not (Get-Command fabric -ErrorAction SilentlyContinue)) {
    Write-Error "Fabric CLI is not installed or not available in PATH. Please install it to proceed."
    exit 1
}

# Get all MKV files in the specified directory
$InputFiles = Get-ChildItem -Path $InputDirectory -Filter "*.mkv"

foreach ($InputFile in $InputFiles) {
    $InputFilePath = $InputFile.FullName

    # Extract file name and directory information
    $InputFileName = [System.IO.Path]::GetFileNameWithoutExtension($InputFilePath)
    $InputFileExtension = [System.IO.Path]::GetExtension($InputFilePath)
    $SourceDirectory = [System.IO.Path]::GetDirectoryName($InputFilePath)

    # Check if the file name has a trailing space and remove it
    if ($InputFileName.TrimEnd() -ne $InputFileName) {
        $NewInputFileName = $InputFileName.TrimEnd()
        $NewInputFilePath = "$SourceDirectory\$NewInputFileName$InputFileExtension"
        try {
            Rename-Item -Path $InputFilePath -NewName "$NewInputFileName$InputFileExtension" -ErrorAction Stop
            $InputFilePath = $NewInputFilePath
            $InputFileName = $NewInputFileName
        } catch {
            Write-Error "Failed to rename file to remove trailing space: $InputFilePath. $_"
            continue
        }
    }

    # Step 1: Create a new folder with the same name as the modified file name but without the extension
    $NewFolderName = "$SourceDirectory\$InputFileName"
    if (-not (Test-Path $NewFolderName)) {
        try {
            New-Item -ItemType Directory -Path $NewFolderName -ErrorAction Stop
        } catch {
            Write-Error "Failed to create directory: $NewFolderName. $_"
            continue
        }
    }

    # Step 2: Move the source file into the new folder with error handling
    $NewFilePath = "$NewFolderName\$InputFileName$InputFileExtension"
    try {
        Move-Item -Path $InputFilePath -Destination $NewFilePath -ErrorAction Stop
    } catch {
        Write-Error "Failed to move file: $InputFilePath to $NewFilePath. $_"
        continue
    }

    # Step 3: Convert the MKV file to MP3 format using ffmpeg
    $MP3FilePath = "$NewFolderName\$InputFileName.mp3"
    try {
        Start-Process -FilePath "c:\Program Files (x86)\ffmpeg\bin\ffmpeg.exe" -ArgumentList "-i `"$NewFilePath`" -map 0:a:0 -b:a 320k `"$MP3FilePath`"" -NoNewWindow -Wait -ErrorAction Stop
    } catch {
        Write-Error "Failed to convert MKV to MP3 using ffmpeg for file: $NewFilePath. $_"
        continue
    }

    # Step 4: Transcribe the MP3 file to a text file using Whisper
    $TranscriptionFilePath = "$NewFolderName\$InputFileName.txt"
    try {
        Start-Process -FilePath "whisper" -ArgumentList "`"$MP3FilePath`" --model small --output_format txt --output_dir `"$NewFolderName`"" -NoNewWindow -Wait -ErrorAction Stop
    } catch {
        Write-Error "Failed to transcribe MP3 to text using Whisper for file: $MP3FilePath. $_"
        continue
    }

    # Step 5: Use Fabric to extract wisdom from the transcribed text file by echoing the content into fabric
    $WisdomFilePath = "$NewFolderName\$InputFileName-wisdom.txt"
    try {
        $Content = Get-Content -Path $TranscriptionFilePath | Out-String
        echo $Content | fabric --pattern extract_wisdom > "$WisdomFilePath"
    } catch {
        Write-Error "Failed to extract wisdom using Fabric for file: $TranscriptionFilePath. $_"
        continue
    }
}

Write-Host "Process complete. Files have been moved, converted, transcribed, and processed with Fabric." -ForegroundColor Green

# Script ends here
