# ==========================================
# CONFIGURATION
# ==========================================

$InputFolder = $PSScriptRoot
$OutputFolder = Join-Path $InputFolder "Converted"

$Format = "docx"
$SkipExisting = $true
$ShowWord = $true

$ErrorActionPreference = "Stop"
$SaveFormat = if ($Format -eq "docx") { 16 } else { 0 }

$StartTime = Get-Date
$Success = 0
$Failed = 0
$Skipped = 0
$Warnings = 0

# ==========================================
# PREPARE FOLDERS
# ==========================================

New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

$ReportFile = Join-Path $OutputFolder (
    "Conversion_Report_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv"
)

$Files = @(Get-ChildItem -Path $InputFolder -File -Recurse |
    Where-Object {
        $_.Extension -in ".html", ".htm" -and
        $_.FullName -notlike "$OutputFolder\*"
    })

$Total = $Files.Count

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "HTML TO WORD BULK CONVERTER" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "Input folder: $InputFolder"
Write-Host "Output folder: $OutputFolder"
Write-Host "Files found: $Total"
Write-Host "Output format: $Format"

if ($Total -eq 0) {
    Write-Host "No HTML files found." -ForegroundColor Red
    return
}

# ==========================================
# START MICROSOFT WORD
# ==========================================

$Word = $null

try {
    Write-Host "Starting Microsoft Word..." -ForegroundColor Yellow

    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $ShowWord
    $Word.DisplayAlerts = -1

    Write-Host "Microsoft Word started." -ForegroundColor Green

    $Current = 0

    foreach ($File in $Files) {

        $Current++
        $Document = $null
        $MissingImages = @()
        $ExternalImages = @()
        $ImageMatches = @()
        $ImageCount = 0
        $Status = "Pending"
        $ErrorMessage = ""
        $FileStart = Get-Date

        $Percent = [math]::Round(($Current / $Total) * 100, 1)

        # Preserve original folder structure

        $RelativeFolder = $File.DirectoryName.Substring(
            $InputFolder.TrimEnd('\').Length
        ).TrimStart('\')

        $DestinationFolder = Join-Path $OutputFolder $RelativeFolder

        $OutputFile = Join-Path $DestinationFolder (
            $File.BaseName + "." + $Format
        )

        Write-Progress `
            -Activity "Converting HTML to Word" `
            -Status "$Current / $Total - $($File.Name)" `
            -PercentComplete $Percent

        Write-Host ""
        Write-Host "[$Current/$Total] [$Percent%] $($File.Name)" `
            -ForegroundColor Cyan

        try {

            New-Item -ItemType Directory `
                -Path $DestinationFolder -Force | Out-Null

            # ==========================================
            # SKIP EXISTING DOCUMENTS
            # ==========================================

            if (Test-Path -LiteralPath $OutputFile) {

                if ($SkipExisting) {
                    $Status = "Skipped"
                    $Skipped++

                    Write-Host "Document already exists." `
                        -ForegroundColor Yellow

                    continue
                }

                throw "Output file already exists. Remove or rename it first."
            }

            # ==========================================
            # CHECK IMAGE REFERENCES
            # ==========================================

            Write-Host "Checking image references..." `
                -ForegroundColor Yellow

            $HtmlContent = Get-Content `
                -LiteralPath $File.FullName -Raw

            $ImageMatches = [regex]::Matches(
                $HtmlContent,
                '<img\b[^>]*\bsrc\s*=\s*["'']([^"'']+)["'']',
                [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
            )

            foreach ($Match in $ImageMatches) {

                $ImageSource = [System.Net.WebUtility]::HtmlDecode(
                    $Match.Groups[1].Value
                )

                if ($ImageSource -match '^data:') {
                    continue
                }

                if ($ImageSource -match '^(https?:)?//') {
                    $ExternalImages += $ImageSource
                    continue
                }

                try {
                    if ($ImageSource -match '^file:') {
                        $ImagePath = ([uri]$ImageSource).LocalPath
                    }
                    else {
                        $ImageSource = [uri]::UnescapeDataString(
                            ($ImageSource -split '[?#]', 2)[0]
                        )

                        if ([System.IO.Path]::IsPathRooted($ImageSource)) {
                            $ImagePath = $ImageSource
                        }
                        else {
                            $ImagePath = Join-Path `
                                $File.DirectoryName $ImageSource
                        }
                    }

                    if (-not (Test-Path -LiteralPath $ImagePath -PathType Leaf)) {
                        $MissingImages += $ImageSource
                    }
                }
                catch {
                    $MissingImages += $ImageSource
                }
            }

            if ($MissingImages.Count -gt 0) {

                Write-Host "Missing images: $($MissingImages.Count)" `
                    -ForegroundColor Yellow

                foreach ($Image in $MissingImages) {
                    Write-Host "  $Image" -ForegroundColor Yellow
                }
            }

            # ==========================================
            # OPEN HTML
            # ==========================================

            Write-Host "Opening HTML..." -ForegroundColor Yellow

            $Document = $Word.Documents.Open(
                $File.FullName,
                $false,
                $true
            )

            # ==========================================
            # EMBED INLINE PICTURES
            # ==========================================

            Write-Host "Processing inline pictures..." `
                -ForegroundColor Yellow

            for ($i = $Document.InlineShapes.Count; $i -ge 1; $i--) {

                $Picture = $Document.InlineShapes.Item($i)

                if ($Picture.Type -in 3, 4) {
                    $ImageCount++
                }

                if ($Picture.Type -eq 4) {
                    $Picture.LinkFormat.SavePictureWithDocument = $true
                    $Picture.LinkFormat.BreakLink()
                }
            }

            # ==========================================
            # EMBED FLOATING PICTURES
            # ==========================================

            Write-Host "Processing floating pictures..." `
                -ForegroundColor Yellow

            for ($i = $Document.Shapes.Count; $i -ge 1; $i--) {

                $Picture = $Document.Shapes.Item($i)

                if ($Picture.Type -in 11, 13) {
                    $ImageCount++
                }

                if ($Picture.Type -eq 11) {
                    $Picture.LinkFormat.SavePictureWithDocument = $true
                    $Picture.LinkFormat.BreakLink()
                }
            }

            # ==========================================
            # PREVENT OVERWRITING
            # ==========================================

            if (Test-Path -LiteralPath $OutputFile) {
                throw "Output filename collision."
            }

            # ==========================================
            # SAVE DOCUMENT
            # ==========================================

            Write-Host "Saving document..." -ForegroundColor Yellow

            $Document.SaveAs2($OutputFile, $SaveFormat)

            $Status = "Success"
            $Success++

            if (
                $MissingImages.Count -gt 0 -or
                $ExternalImages.Count -gt 0 -or
                $ImageCount -lt $ImageMatches.Count
            ) {
                $Status = "Success with warnings"
                $Warnings++
            }

            Write-Host "$Status : $OutputFile" -ForegroundColor Green
            Write-Host "Expected images: $($ImageMatches.Count)"
            Write-Host "Word images: $ImageCount"
        }
        catch {

            $Status = "Failed"
            $Failed++
            $ErrorMessage = $_.Exception.Message

            Write-Host "ERROR: $ErrorMessage" -ForegroundColor Red
        }
        finally {

            if ($null -ne $Document) {

                try {
                    $Document.Close(0)
                }
                catch {
                    Write-Warning "Could not close document."
                }

                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                    $Document
                )
            }

            $Duration = [math]::Round(
                ((Get-Date) - $FileStart).TotalSeconds,
                2
            )

            # ==========================================
            # SAVE REPORT ENTRY
            # ==========================================

            $Entry = [PSCustomObject]@{
                Timestamp      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                FileName       = $File.Name
                SourcePath     = $File.FullName
                OutputPath     = $OutputFile
                Status         = $Status
                ExpectedImages = $ImageMatches.Count
                WordImages     = $ImageCount
                MissingImages  = ($MissingImages -join "; ")
                ExternalImages = ($ExternalImages -join "; ")
                Error          = $ErrorMessage
                DurationSec    = $Duration
            }

            $Entry | Export-Csv `
                -Path $ReportFile `
                -Append `
                -NoTypeInformation `
                -Encoding UTF8
        }
    }
}
finally {

    Write-Progress -Activity "Converting HTML to Word" -Completed

    if ($null -ne $Word) {

        Write-Host "Closing Microsoft Word..." -ForegroundColor Yellow

        try {
            $Word.Quit()
        }
        finally {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject(
                $Word
            )
        }
    }
}

# ==========================================
# FINAL SUMMARY
# ==========================================

$TotalTime = (Get-Date) - $StartTime

Write-Host ""
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "CONVERSION SUMMARY" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

Write-Host "Total files:          $Total"
Write-Host "Converted:            $Success" -ForegroundColor Green
Write-Host "With warnings:        $Warnings" -ForegroundColor Yellow
Write-Host "Skipped:              $Skipped" -ForegroundColor Yellow
Write-Host "Failed:               $Failed" -ForegroundColor Red
Write-Host "Total time:           $($TotalTime.ToString('hh\:mm\:ss'))"
Write-Host "Report:               $ReportFile"

Write-Host "==========================================" -ForegroundColor Cyan