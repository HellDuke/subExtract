param (
    [ValidateSet('list', 'directory', 'file')]
    [string]$mode = "list", # Specify the mode: file, directory or list
    # file - extract for a single file, path should be to a file
    # directory - path is pointing to directory, will extract subtitles for all files in direcotry and subdirectories
    # list - path is pointing to a text file, each new line has a new directory for which we do same as directory mode
    [string]$path = "", # Specify the path based on the chosen mode
    [switch]$log, # Enable log file output
    [bool]$isUpgrade = $false  # If it's an upgrade to the existing file. Deletes any subtitle files sharing the base file name
)

$ffmpegPath = "C:\Program Files\ffmpeg\bin\ffmpeg.exe"  # Adjust the path to ffmpeg.exe accordingly
$ffprobePath = "C:\Program Files\ffmpeg\bin\ffprobe.exe"  # Adjust the path to ffmpeg.exe accordingly

function SendNotification($title, $message) {
    $gotifyUrl = 'http://omv.home.arpa:84/message?token=AV5Bx3.Z3_DNOQ6'
    Invoke-RestMethod -Uri $gotifyUrl -Method POST -Body (@{
        title   = $title
        message = $message
      } | ConvertTo-Json -Compress) -ContentType 'application/json'
  }

function ExtractSubtitles($inputFile, $isUpgrade) {
    $inputFileName = (Get-Item -Path $inputFile).Name
    $inputDirectory = (Get-Item -Path $inputFile).DirectoryName
    $baseName = $inputFile.BaseName
    # Delete related subtitle files if it's an upgrade
    if ($isUpgrade) {
        $matchFiles = Get-ChildItem -Path $inputDirectory -Filter "$baseName*" -File
        foreach ($testFile in $matchFiles) {
            if ($testFile.Extension -match "srt|ass|ssa|vtt") {
                Write-Host "  ↳ Deleting subtitle file: $($testFile.Name)"
                Remove-Item $testFile.FullName
            }
        }
    }

    # Get subtitle stream information using ffprobe
    $streamInfo = & $ffprobePath -v quiet -print_format json -show_streams -select_streams s "$inputFile" | ConvertFrom-Json

    # Extract subtitles for each stream
    foreach ($stream in $streamInfo.streams) {
        $languageTag = $stream.tags.language
        if ($stream.codec_type -eq "subtitle" -and $languageTag -match '^(en(-\w{2})?|eng|english)$') {
            $index = $stream.index
            $language = "en"
            $formatName = $stream.codec_name.ToLower()
            Write-Host "  ↳ Extracting subtitle from: $inputFileName"
            if ($formatName -eq "subrip") {
                Write-Host "      Stream index: $index"
                Write-Host "      Language: $language"
                Write-Host "      Format: $formatName"
                Write-Host "Skipping SRT subtitle extraction"
                continue
            }
            $outputExtension = switch ($formatName) {
                "substation alpha" { "ass" }
                "ass" { "ass" }
                "ssa" { "ssa" }
                default { "srt" }
            }

            # Begin building the file name
            $title = ($stream.tags.title -replace '[<>:"/\\|?*]', ' ').Trim()
            if ([string]::IsNullOrWhiteSpace($title)) {
                $title = "English"
            }
            $outputFileName = "{0}.{1}" -f $baseName, $title
            if ($stream.disposition.default -eq 1) {
                $outputFileName += ".default"
            }
            $outputFileName += ".$language"
            if ($stream.disposition.captions -eq 1) {
                $outputFileName += ".cc"
            }
            elseif ($strea.disposition.hearing_impaired -eq 1) {
                $outputFileName += ".hi"
            }
            else {
                $acronyms = @("sdh", "hi", "cc")
                $acronymMatch = $acronyms | Where-Object { $title.ToLower() -match $_ }

                if ($acronymMatch) {
                    $impairTag = $acronymMatch.ToLower()
                    $outputFileName += ".$impairTag"
                }
            }
            $outputFileName += ".$outputExtension"
            # End of building the file name, generate the absolute path
            $outputFile = Join-Path -Path $inputDirectory -ChildPath $outputFileName

            # Skip extraction if the subtitle file already exists
            if (Test-Path $outputFile) {
                Write-Host "  ↳ Skipped extracting subtitle from: $inputFileName"
                Write-Host "      Subtitle file already exists: $outputFileName"
                continue
            }

            Write-Host "      Stream index: $index"
            Write-Host "      Language: $language"
            Write-Host "      Format: $formatName"
            Write-Host "      Output File: $outputFileName"

            # Extract the subtitle stream
            & $ffmpegPath -i $inputFile -map 0:$index -c:s copy -y $outputFile -hide_banner -nostats -loglevel error 2>$null
        }
    }
}

# Recursive function to traverse directories and subdirectories
function ProcessDirectory($directoryPath, $isUpgrade) {
    $files = Get-ChildItem -Path $directoryPath -File
    foreach ($file in $files) {
        ExtractSubtitles $file $isUpgrade
    }

    $subdirectories = Get-ChildItem -Path $directoryPath -Directory
    foreach ($subdirectory in $subdirectories) {
        ProcessDirectory $subdirectory.FullName $isUpgrade
    }
}

# Start transcript if the switch is provided
if ($log) {
    $transcriptFileName = "SubtitleExtraction.log"
    Start-Transcript -Path $transcriptFileName -Force
}

# Logic to determine which action to perform based on the mode
switch ($mode) {
    "list" {
        # Process directories from the input file
        Write-Host "Working in list mode"
        $directories = Get-Content $path
        foreach ($directory in $directories) {
            ProcessDirectory $directory $isUpgrade
        }
        SendNotification "Finished manual list extraction" "Extraction complete for the following folders:`n$directories"
    }
    "directory" {
        # Process the single directory specified by the path
        Write-Host "Working in directory mode"
        ProcessDirectory $path $isUpgrade
        SendNotification "Finished manual directory extract" "Extraction complete on $path"
    }
    "file" {
        # Process the single file specified by the path
        Write-Host "Working in file mode"
        $fileObject = Get-Item -Path $path
        ExtractSubtitles $fileObject $isUpgrade
        SendNotification "Finished manual file extract" "Complete extraction on file $path"
    }
    Default {
        Write-Host "Invalid mode specified. Please choose 'list', 'directory', or 'file'."
        exit 1
    }
}

# Stop transcript if it was started
if ($log) {
    Stop-Transcript
    Write-Host "Transcript saved to: $transcriptFileName"
}