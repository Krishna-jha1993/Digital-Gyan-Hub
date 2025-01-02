# PDF Manipulation Script using PDFtk Server
# Requires PDFtk Server to be installed

<#
PDF Manipulation Script using PDFtk Server
Requires PDFtk Server to be installed

Owner: Murari Jha

Please like and subscribe to our YouTube channel:
Digital Gyan hub @DigitalGyanhubm
#>

function Test-PDFtkInstallation {
    $pdftk32Path = "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe"
    $pdftk64Path = "C:\Program Files\PDFtk Server\bin\pdftk.exe"
    
    if (Test-Path $pdftk32Path) {
        $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        $pdftkDir = "C:\Program Files (x86)\PDFtk Server\bin"
        if ($currentPath -notlike "*$pdftkDir*") {
            [Environment]::SetEnvironmentVariable("Path", "$currentPath;$pdftkDir", "Machine")
            $env:Path = "$env:Path;$pdftkDir"
        }
        return $true
    }
    elseif (Test-Path $pdftk64Path) {
        $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        $pdftkDir = "C:\Program Files\PDFtk Server\bin"
        if ($currentPath -notlike "*$pdftkDir*") {
            [Environment]::SetEnvironmentVariable("Path", "$currentPath;$pdftkDir", "Machine")
            $env:Path = "$env:Path;$pdftkDir"
        }
        return $true
    }
    else {
        Write-Host "PDFtk Server is not installed. Installing now..." -ForegroundColor Yellow
        try {
            $installerUrl = "https://www.pdflabs.com/tools/pdftk-the-pdf-toolkit/pdftk_server-2.02-win-setup.exe"
            $installerPath = Join-Path $env:TEMP "pdftk_installer.exe"
            
            if (Test-Path $installerPath) {
                Remove-Item $installerPath -Force
            }
            
            Write-Host "Downloading PDFtk Server installer..." -ForegroundColor Yellow
            try {
                $webClient = New-Object System.Net.WebClient
                $webClient.DownloadFile($installerUrl, $installerPath)
            } catch {
                Write-Host "Download failed: $_" -ForegroundColor Red
                exit 1
            }
            
            Write-Host "Running installer... Please wait..." -ForegroundColor Yellow
            $process = Start-Process -FilePath $installerPath -ArgumentList "/VERYSILENT" -Wait -PassThru
            
            if ($process.ExitCode -eq 0) {
                Write-Host "Installation completed. Please close this window and run the script again." -ForegroundColor Green
                Read-Host "Press Enter to exit"
                exit
            }
            else {
                throw "Installation failed with exit code: $($process.ExitCode)"
            }
        }
        catch {
            Write-Host "Error: Installation failed. Please install PDFtk Server manually from https://www.pdflabs.com/tools/pdftk-server/" -ForegroundColor Red
            Write-Host "Error details: $_" -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
    }
}

function Get-OutputDirectory {
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.Description = "Select output folder"
    if ($FolderBrowser.ShowDialog() -eq 'OK') {
        return $FolderBrowser.SelectedPath
    }
    return $PWD.Path
}

function Get-PDFFile {
    param (
        [string]$PromptMessage
    )
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowser.Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*"
    $FileBrowser.Title = $PromptMessage
    if ($FileBrowser.ShowDialog() -eq 'OK') {
        return $FileBrowser.FileName
    }
    return $null
}

function Get-FileSize {
    param (
        [string]$FilePath
    )
    if (Test-Path $FilePath) {
        return (Get-Item $FilePath).Length
    }
    return 0
}

function Get-PageCount {
    param (
        [string]$PDFPath
    )
    try {
        $pdftkPath = ""
        if (Test-Path "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe") {
            $pdftkPath = "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe"
        }
        elseif (Test-Path "C:\Program Files\PDFtk Server\bin\pdftk.exe") {
            $pdftkPath = "C:\Program Files\PDFtk Server\bin\pdftk.exe"
        }
        
        if ($pdftkPath -eq "") {
            throw "PDFtk executable not found"
        }
        
        $output = & $pdftkPath $PDFPath dump_data | Select-String "NumberOfPages"
        if ($output -match "NumberOfPages: (\d+)") {
            return [int]$matches[1]
        }
        throw "Unable to get page count"
    }
    catch {
        Write-Host "Error getting page count: $_" -ForegroundColor Red
        exit 1
    }
}

function Merge-PDFFiles {
    $pdfFiles = @()
    Write-Host "`nPDF Merger" -ForegroundColor Cyan
    $fileCount = Read-Host "How many PDF files do you want to merge?"
    
    if (-not ($fileCount -match '^\d+$') -or [int]$fileCount -lt 1) {
        Write-Host "Invalid number of files specified." -ForegroundColor Red
        return
    }
    
    for ($i = 1; $i -le [int]$fileCount; $i++) {
        $file = Get-PDFFile "Select PDF file #$i to merge"
        if ($file) {
            $pdfFiles += $file
        }
        else {
            Write-Host "File selection cancelled." -ForegroundColor Yellow
            return
        }
    }
    
    $outputDir = Get-OutputDirectory
    $outputName = Read-Host "Enter name for merged file (without .pdf extension)"
    $outputFile = Join-Path $outputDir "$outputName.pdf"
    
    try {
        $pdftkPath = ""
        if (Test-Path "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe") {
            $pdftkPath = "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe"
        }
        elseif (Test-Path "C:\Program Files\PDFtk Server\bin\pdftk.exe") {
            $pdftkPath = "C:\Program Files\PDFtk Server\bin\pdftk.exe"
        }
        
        & $pdftkPath $pdfFiles cat output $outputFile
        if (Test-Path $outputFile) {
            $size = [math]::Round((Get-FileSize $outputFile) / 1MB, 2)
            Write-Host "Files successfully merged to: $outputFile" -ForegroundColor Green
            Write-Host "Output file size: $size MB" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "Error merging files: $_" -ForegroundColor Red
    }
}

function Split-PDFFile {
    Write-Host "`nPDF Splitter" -ForegroundColor Cyan
    $inputFile = Get-PDFFile "Select PDF file to split"
    if (-not $inputFile) {
        Write-Host "File selection cancelled." -ForegroundColor Yellow
        return
    }
    
    if (-not (Test-Path $inputFile)) {
        Write-Host "Input file not found: $inputFile" -ForegroundColor Red
        return
    }
    
    $pageCount = Get-PageCount $inputFile
    $fileSize = [math]::Round((Get-FileSize $inputFile) / 1MB, 2)
    Write-Host "Total pages in PDF: $pageCount"
    Write-Host "File size: $fileSize MB"
    
    Write-Host "`nSplit options:"
    Write-Host "1. Split into specific number of files"
    Write-Host "2. Split by page count per file"
    Write-Host "3. Split into equal parts"
    Write-Host "4. Split every page into separate files"
    Write-Host "5. Split by file size (approximate)"
    
    $choice = Read-Host "`nSelect split option (1-5)"
    $outputDir = Get-OutputDirectory
    $outputPrefix = Read-Host "Enter prefix for split files (e.g., 'part' will create part_1.pdf, part_2.pdf, etc.)"
    
    $pdftkPath = ""
    if (Test-Path "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe") {
        $pdftkPath = "C:\Program Files (x86)\PDFtk Server\bin\pdftk.exe"
    }
    elseif (Test-Path "C:\Program Files\PDFtk Server\bin\pdftk.exe") {
        $pdftkPath = "C:\Program Files\PDFtk Server\bin\pdftk.exe"
    }
    
    switch ($choice) {
        "1" {
            $numFiles = [int](Read-Host "Enter number of files to split into")
            if ($numFiles -lt 1) {
                Write-Host "Invalid number of files." -ForegroundColor Red
                return
            }
            
            $pagesPerFile = [math]::Floor($pageCount / $numFiles)
            $remainder = $pageCount % $numFiles
            $start = 1
            
            for ($i = 1; $i -le $numFiles; $i++) {
                $extraPage = if ($i -le $remainder) { 1 } else { 0 }
                $currentPages = $pagesPerFile + $extraPage
                $end = [math]::Min($start + $currentPages - 1, $pageCount)
                
                if ($start -le $pageCount) {
                    $outputFile = Join-Path $outputDir "${outputPrefix}_${i}.pdf"
                    & $pdftkPath $inputFile cat $start-$end output $outputFile
                    if (Test-Path $outputFile) {
                        $size = [math]::Round((Get-FileSize $outputFile) / 1MB, 2)
                        Write-Host "Created: $outputFile (Size: $size MB)" -ForegroundColor Green
                    }
                }
                $start = $end + 1
            }
        }
        "2" {
            $pagesPerFile = [int](Read-Host "Enter number of pages per file")
            if ($pagesPerFile -lt 1) {
                Write-Host "Invalid number of pages." -ForegroundColor Red
                return
            }
            
            $numFiles = [math]::Ceiling($pageCount / $pagesPerFile)
            $start = 1
            
            for ($i = 1; $i -le $numFiles -and $start -le $pageCount; $i++) {
                $end = [math]::Min($start + $pagesPerFile - 1, $pageCount)
                $outputFile = Join-Path $outputDir "${outputPrefix}_${i}.pdf"
                & $pdftkPath $inputFile cat $start-$end output $outputFile
                if (Test-Path $outputFile) {
                    $size = [math]::Round((Get-FileSize $outputFile) / 1MB, 2)
                    Write-Host "Created: $outputFile (Size: $size MB)" -ForegroundColor Green
                }
                $start = $end + 1
            }
        }
        "3" {
            $parts = [int](Read-Host "Enter number of equal parts")
            if ($parts -lt 1) {
                Write-Host "Invalid number of parts." -ForegroundColor Red
                return
            }
            
            $pagesPerFile = [math]::Floor($pageCount / $parts)
            $remainder = $pageCount % $parts
            $start = 1
            
            for ($i = 1; $i -le $parts; $i++) {
                $extraPage = if ($i -le $remainder) { 1 } else { 0 }
                $currentPages = $pagesPerFile + $extraPage
                $end = [math]::Min($start + $currentPages - 1, $pageCount)
                
                if ($start -le $pageCount) {
                    $outputFile = Join-Path $outputDir "${outputPrefix}_${i}.pdf"
                    & $pdftkPath $inputFile cat $start-$end output $outputFile
                    if (Test-Path $outputFile) {
                        $size = [math]::Round((Get-FileSize $outputFile) / 1MB, 2)
                        Write-Host "Created: $outputFile (Size: $size MB)" -ForegroundColor Green
                    }
                }
                $start = $end + 1
            }
        }
        "4" {
            for ($i = 1; $i -le $pageCount; $i++) {
                $outputFile = Join-Path $outputDir "${outputPrefix}_${i}.pdf"
                & $pdftkPath $inputFile cat $i output $outputFile
                if (Test-Path $outputFile) {
                    $size = [math]::Round((Get-FileSize $outputFile) / 1MB, 2)
                    Write-Host "Created: $outputFile (Size: $size MB)" -ForegroundColor Green
                }
            }
        }
        "5" {
            $targetSize = [double](Read-Host "Enter target size per file in MB")
            if ($targetSize -le 0) {
                Write-Host "Invalid file size." -ForegroundColor Red
                return
            }
            
            $pagesPerFile = [math]::Ceiling($pageCount * ($targetSize / $fileSize))
            $start = 1
            $fileNum = 1
            
            while ($start -le $pageCount) {
                $end = [math]::Min($start + $pagesPerFile - 1, $pageCount)
                $outputFile = Join-Path $outputDir "${outputPrefix}_${fileNum}.pdf"
                & $pdftkPath $inputFile cat $start-$end output $outputFile
                if (Test-Path $outputFile) {
                    $size = [math]::Round((Get-FileSize $outputFile) / 1MB, 2)
                    Write-Host "Created: $outputFile (Size: $size MB)" -ForegroundColor Green
                }
                $start = $end + 1
                $fileNum++
            }
        }
        default {
            Write-Host "Invalid option selected." -ForegroundColor Red
            return
        }
    }
    
    Write-Host "`nPDF split operation completed successfully!" -ForegroundColor Green
    Write-Host "Output files are in: $outputDir" -ForegroundColor Green
}

# Main script

if (-not (Test-PDFtkInstallation)) {
    Read-Host "Press Enter to exit"
    exit 1
}

try {

    while ($true) {
    Write-Host "`nPDF Manipulation Tool" -ForegroundColor Cyan
    Write-Host "1. Merge PDF files"
    Write-Host "2. Split PDF file"
    Write-Host "3. Exit"

    $choice = Read-Host "`nSelect an option (1-3)"

    switch ($choice) {
        "1" { Merge-PDFFiles }
        "2" { Split-PDFFile }
        "3" { 
            Write-Host "Goodbye!" -ForegroundColor Green
            exit 0
        }
        default {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
       		 }
   	 }
	}
}

finally {
    Write-Host "`nThank you for using the PDF Manipulation Tool!" -ForegroundColor Cyan
    Write-Host "Please like and subscribe to our YouTube channel:" -ForegroundColor Yellow
    Write-Host "Digital Gyan hub @DigitalGyanhub" -ForegroundColor Yellow
}


