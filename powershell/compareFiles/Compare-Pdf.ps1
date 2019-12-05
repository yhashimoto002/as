<#
.SYNOPSIS

Compare 2 PDFs and output the difference.

.PARAMETER beforeDir

Directory path including PDFs before sanitizing

.PARAMETER afterDir

Directory path including PDFs after sanitizing

.OUTPUTS

CSV file

.EXAMPLE

PS> .\Compare-Pdf.ps1 .\before .\after

#>

param(
    [parameter(mandatory)]
    [string]$beforeDir,
    [parameter(mandatory)]
    [string]$afterDir
)

## change if needed
# set the dpi of an image
$imDensity = "100"
# set the threshold of differency
# the smaller the difference, the value is close to 0.
$identifyThreshold = "1000"

## don't change
$outputDir = Join-Path $PSScriptRoot "output"
$outCsvFilePath = Join-Path $PSScriptRoot ("result_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv")
$outHtmlFilePath = Join-Path $PSScriptRoot ("result_NG_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".html")
$count = 0

function Convert-PdfToPng
{
    param(
        [parameter(mandatory)]
        [string]$Path,
        [parameter(mandatory)]
        [string]$OutDir
    )

    mkdir $OutDir -Force | Out-Null
    Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} converting {1} to image..." -f (Get-Date), $Path)
    # convert -quiet -density $imDensity -alpha off $Path (Join-Path $OutDir "image.png")
    magick convert -quiet -colorspace rgb -density $imDensity -alpha remove -background white $Path (Join-Path $OutDir "image.png")
    Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} converting {0} is finished." -f (Get-Date), $Path)
    Write-Host ""
}

function Compare-Pdf
{
    param(
        [parameter(Mandatory, ValueFromPipeline)]
        [string]$Pdf
    )

    process
    {
        # skip if target PDF doesn't exist in the opposite dir
        if (! (Test-Path (Join-Path $afterDir $Pdf))) { return }
        
        $before_dir = Join-Path $outputDir $Pdf | Join-Path -ChildPath "before"
        $after_dir = Join-Path $outputDir $Pdf | Join-Path -ChildPath "after"
        $diff_dir = Join-Path $outputDir $Pdf | Join-Path -ChildPath "diff"
        mkdir $diff_dir -Force | Out-Null

        # convert pdf to image
        Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1}" -f (Get-Date), ++$count)
        Convert-PdfToPng -Path (Join-Path $beforeDir $Pdf) -OutDir $before_dir
        Convert-PdfToPng -Path (Join-Path $afterDir $Pdf) -OutDir $after_dir

        # compare images and analyze the difference
        $arrayResult = @()
        $page = 0
        dir $before_dir | sort -Property LastWriteTime | foreach {
            $png = $_.Name
            magick composite -quiet -compose difference (Join-Path $before_dir $png) `
                (Join-Path $after_dir $png) (Join-Path $diff_dir $png)
            [float]$identify = magick identify -format "%[mean]" (Join-Path $diff_dir $png)
            
            # output result to csv
            $result = "NG"
            $imageBeforePath = Join-Path $before_dir $png
            $imageAfterPath = Join-Path $after_dir $png
            $imageDiffPath = Join-Path $diff_dir $png
            if ($identify -lt $identifyThreshold)
            {
                $result = "OK"
                $imageBeforePath = ""
                $imageAfterPath = ""
                $imageDiffPath = ""
            }
            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1}/{2}: {3}({4})" -f (Get-Date), $Pdf, $png, $result, $identify) 
            $objectOfEachRecord = [pscustomobject]@{
                "No."=$count
                FileName=$Pdf
                ImageName=$png
                Page=++$page
                Identify=$identify
                Result=$result
                "Image(diff)"=$imageDiffPath
                "Image(before)"=$imageBeforePath
                "Image(after)"=$imageAfterPath
            }
            $arrayResult += $objectOfEachRecord
        }
        $arrayResult | Export-Csv $outCsvFilePath  -Encoding UTF8 -NoTypeInformation -Append
        Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
    }
}


# main
$startTime = Get-Date
if (Test-Path $outCsvFilePath) { rm $outCsvFilePath -Force }
dir $beforeDir -Include *.pdf -Name | Compare-Pdf

Import-Csv $outCsvFilePath | ConvertTo-Html | ? {
        $_ -notmatch "<td>OK</td>"
    } | % {
        $_ -replace "<table>", "<table border=`"1`" style=`"border-collapse: collapse`">" `
           -replace "</td>", "</td>`n" `
           -replace "C:\\(\S+)`.png</td>", "<a href=`"C:\`$1`.png`"><img src=`"C:\`$1`.png`" width=`"300`"></a></td>" `
    } | Out-File $outHtmlFilePath -Encoding utf8

$csvObj = Import-Csv $outCsvFilePath
$csvObj | Select-Object * -ExcludeProperty Image* |
    Export-Csv $outCsvFilePath -Encoding UTF8 -NoTypeInformation

$endTime = Get-Date
Write-Host ("Start: {0}" -f $startTime)
Write-Host ("End: {0}" -f $endTime)
Write-Host ("Total: {0}" -f ($endTime - $startTime))


