<#
.SYNOPSIS

Compare 2 Images and output the difference.

.PARAMETER beforeDir

Directory path including Images before sanitizing

.PARAMETER afterDir

Directory path including Images after sanitizing

.OUTPUTS

CSV file

.EXAMPLE

PS> .\Compare-Image.ps1 .\before .\after

#>

param(
    [parameter(mandatory)]
    [string]$beforeDir,
    [parameter(mandatory)]
    [string]$afterDir
)

## change if needed
# set the dpi of an image
$imDensity = "50"
# set the resize target
$resize = "1000"
# set the threshold of differency
# the smaller the difference, the value is close to 0.
$identifyThreshold = "1500"

## don't change
$outputDir = Join-Path $PSScriptRoot "output"
$outCsvFilePath = Join-Path $PSScriptRoot ("result_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv")
$outHtmlFilePath = Join-Path $PSScriptRoot ("result_NG_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".html")
$count = 0

function Convert-Image
{
    param(
        [parameter(mandatory)]
        [string]$Path,
        [parameter(mandatory)]
        [string]$OutDir
    )

    mkdir $OutDir -Force | Out-Null
    Write-Host ("converting {0} to image..." -f $Path)
    $basename = (Get-ChildItem $Path).BaseName
    $extension = (Get-ChildItem $Path).Extension
    if (($extension -eq ".emf") -or ($extension -eq ".wmf"))
    {
        magick convert -quiet -colorspace rgb -resize ${resize}x${resize}! -alpha remove -background white $Path (Join-Path $OutDir ($basename + ".png"))
    }
    else
    {
        #$option = "-quiet -colorspace rgb -resize ${resize}x${resize} -alpha remove -background white"
        #magick convert ${option} $Path (Join-Path $OutDir (Split-Path -Leaf $Path))
        magick convert -quiet -colorspace rgb -resize ${resize}x${resize} -alpha remove -background white $Path (Join-Path $OutDir (Split-Path -Leaf $Path))
    }
    Write-Host ("converting {0} is finished." -f $Path)
    Write-Host ""
}


function Compare-Image
{
    param(
        [parameter(Mandatory, ValueFromPipeline)]
        [string]$Image
    )

    process
    {
        # skip if target Image doesn't exist in the opposite dir
        if (! (Test-Path (Join-Path $afterDir $Image))) { return }
        
        $before_dir = Join-Path $outputDir $Image | Join-Path -ChildPath "before"
        $after_dir = Join-Path $outputDir $Image | Join-Path -ChildPath "after"
        $diff_dir = Join-Path $outputDir $Image | Join-Path -ChildPath "diff"
        mkdir $diff_dir -Force | Out-Null

        # convert image
        Convert-Image -Path (Join-Path $beforeDir $Image) -OutDir $before_dir
        Convert-Image -Path (Join-Path $afterDir $Image) -OutDir $after_dir

        # compare images and analyze the difference
        $arrayResult = @()
        Write-Host ("{0}" -f ++$count)
        $basename = (Get-ChildItem (Join-Path $beforeDir $Image)).BaseName
        $extention = (Get-ChildItem (Join-Path $beforeDir $Image)).Extension
        if (($extention -eq ".emf") -or ($extention -eq ".wmf"))
        {
            magick composite -quiet -compose difference (Join-Path $before_dir ($basename + ".png")) `
                (Join-Path $after_dir ($basename + ".png")) (Join-Path $diff_dir ($basename + ".png"))
            [float]$identify = magick identify -format "%[mean]" (Join-Path $diff_dir ($basename + ".png"))
        }
        else
        {
            magick composite -quiet -compose difference (Join-Path $before_dir $Image) `
                (Join-Path $after_dir $Image) (Join-Path $diff_dir $Image)
            [float]$identify = magick identify -format "%[mean]" (Join-Path $diff_dir $Image)
        }
        
            
        # output result to csv
        $result = "NG"
        if (($extention -eq ".emf") -or ($extention -eq ".wmf"))
        {
            $imageBeforePath = Join-Path $before_dir ($basename + ".png")
            $imageAfterPath = Join-Path $after_dir ($basename + ".png")
            $imageDiffPath = Join-Path $diff_dir ($basename + ".png")
        }
        else
        {
            $imageBeforePath = Join-Path $before_dir $Image
            $imageAfterPath = Join-Path $after_dir $Image
            $imageDiffPath = Join-Path $diff_dir $Image
        }
        if ($identify -lt $identifyThreshold) { $result = "OK"; $imageBeforePath = ""; $imageAfterPath = ""; $imageDiffPath = "" }
        Write-Host ("{0}/{1}: {2}" -f $Image, $identify, $result) 
        $objectOfEachRecord = [pscustomobject]@{
            "No."=$count
            FileName=$Image
            Identify=$identify
            Result=$result
            "Image(diff)"=$imageDiffPath
            "Image(before)"=$imageBeforePath
            "Image(after)"=$imageAfterPath
        }
        $arrayResult += $objectOfEachRecord
        $arrayResult | Export-Csv $outCsvFilePath  -encoding UTF8 -NoTypeInformation -Append
        Write-Host "------------------------------"
    }
}


# main
$startTime = Get-Date
if (Test-Path $outCsvFilePath) { rm $outCsvFilePath -Force }
#dir $beforeDir -Include *.jpg, *.png, *.gif, *.tif, *.emf, *.wmf -Name | Compare-Image
dir $beforeDir -Include *.emf, *.wmf -Name | Compare-Image

Import-Csv $outCsvFilePath | ConvertTo-Html | ? {
        $_ -notmatch "<td>OK</td>"
    } | % {
        $_ -replace "<table>", "<table border=`"1`" style=`"border-collapse: collapse`">" `
           -replace "</td>", "</td>`n" `
           -replace "C:\\(\S+)`.(jpg|png|gif|tif|emf|wmf)</td>", "<a href=`"C:\`$1`.`$2`"><img src=`"C:\`$1`.`$2`" width=`"300`"></a></td>" `
    } | Out-File $outHtmlFilePath -Encoding utf8

$csvObj = Import-Csv $outCsvFilePath
$csvObj | Select-Object * -ExcludeProperty Image* |
    Export-Csv $outCsvFilePath -Encoding UTF8 -NoTypeInformation

$endTime = Get-Date
Write-Host ("Start: {0}" -f $startTime)
Write-Host ("End: {0}" -f $endTime)
Write-Host ("Total: {0}" -f ($endTime - $startTime))


