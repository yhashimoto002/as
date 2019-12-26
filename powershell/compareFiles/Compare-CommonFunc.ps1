
# don't change
$outputDir = Join-Path $PSScriptRoot "output"
$outLogFilePath = Join-Path $PSScriptRoot ("result_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".log")
$outCsvFilePath = Join-Path $PSScriptRoot ("result_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv")
$outHtmlFilePath = Join-Path $PSScriptRoot ("result_NG_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".html")
$script:count = 0


# load config
$conf = Get-Content (Join-Path $PSScriptRoot "settings.ini") | Where-Object { $_ -match "=" } | ConvertFrom-StringData
$imDensity = $conf.imDensity
$resize = $conf.resize
$identifyThreshold = $conf.identifyThreshold

# install check
function Test-InstalledIM
{
    try
    {
        Get-Command magick -ErrorAction Stop
        $true
    }
    catch
    {
        Write-Host "ImageMagick is not installed!`nThis script cannot be continued ..."
        $false
    }
}

function Convert-ToImage
{
    param(
        [parameter(mandatory)]
        [string]$Path,
        [parameter(mandatory)]
        [string]$OutDir
    )

    Add-Message ("converting {0} to img ..." -f (Resolve-Path $Path)) $outLogFilePath

    if ($Path -match "`.(emf|wmf)$")
    {
        magick convert -quiet -colorspace rgb -resize ${resize}x${resize}! -alpha remove -background white $Path (Join-Path $OutDir "image.jpg")
    }
    else
    {
        #magick convert -quiet -colorspace rgb -density $imDensity -alpha remove -background white $Path (Join-Path $OutDir (Split-Path -Leaf $Path))
        magick convert -quiet -colorspace rgb -density $imDensity -alpha remove -background white $Path (Join-Path $OutDir "image.jpg")
    }
    Add-Message ("converting {0} to finished." -f (Resolve-Path $Path)) $outLogFilePath
    Write-Host ""
}


function Compare-Image
{
    param(
        [parameter(Mandatory, ValueFromPipeline)]
        [string]$Image,
        [switch]$Word,
        [switch]$Excel,
        [switch]$PowerPoint
    )

    process
    {
        # skip if target Image doesn't exist in the opposite dir
        if (-not (Test-Path (Join-Path $afterDir $Image))) { return }
        
        $before_dir = Join-Path $outputDir $Image | Join-Path -ChildPath "before"
        $after_dir = Join-Path $outputDir $Image | Join-Path -ChildPath "after"
        $diff_dir = Join-Path $outputDir $Image | Join-Path -ChildPath "diff"
        foreach ($dir in @($before_dir, $after_dir, $diff_dir))
        {
            if (-not (Test-Path $dir)) { New-Item $dir -ItemType "Directory" -Force | Out-Null }
        }

        # convert image
        Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1}" -f (Get-Date), ++$script:count)
        if($PowerPoint)
        {
            Convert-PowerPointToPng -Path (Join-Path $beforeDir $Image) -OutDir $before_dir
            Convert-PowerPointToPng -Path (Join-Path $afterDir $Image) -OutDir $after_dir
        }
        else
        {
            Convert-ToImage -Path (Join-Path $beforeDir $Image) -OutDir $before_dir
            Convert-ToImage -Path (Join-Path $afterDir $Image) -OutDir $after_dir
        }

        # compare images and analyze the difference
        $arrayResult = @()
        $page = 0
        if($Word -or $Excel) { $Image = $Image -replace ".pdf$", "" }
        Add-Message ("comparing {0} ..." -f $Image) $outLogFilePath
        Get-ChildItem $before_dir | Sort-Object -Property LastWriteTime | ForEach-Object {
            $imageName = $_.Name
            $imageBeforePath = Join-Path $before_dir $imageName
            $imageAfterPath = Join-Path $after_dir $imageName
            $imageDiffPath = Join-Path $diff_dir $imageName
            magick composite -quiet -compose difference $imageBeforePath $imageAfterPath $imageDiffPath
            [float]$identify = magick identify -format "%[mean]" $imageDiffPath
                
            # output result to csv
            $result = "NG"
            if ($identify -lt $identifyThreshold)
            {
                $result = "OK"
                $imageBeforePath = ""
                $imageAfterPath = ""
                $imageDiffPath = ""
            }
            Add-Message ("`t{0}/{1}: {2}({3})" -f $Image, $imageName, $result, $identify) $outLogFilePath
            $objectOfEachRecord = [pscustomobject]@{
                "No."=$script:count
                FileName=$Image
                ImageName=$imageName
                Page=++$page
                Identify=$identify
                Result=$result
                "Image(diff)"=$imageDiffPath
                "Image(before)"=$imageBeforePath
                "Image(after)"=$imageAfterPath
            }
            $arrayResult += $objectOfEachRecord
        }
        $arrayResult | Export-Csv $outCsvFilePath  -encoding UTF8 -NoTypeInformation -Append
        Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
    }
}


function Add-Message
{
    param(
        [parameter(mandatory)]
        [string]$Message,
        [parameter(mandatory)]
        [string]$LogFilePath
    )

    "{0:yyyy/MM/dd HH:mm:ss.fff} {1}" -f (Get-Date), $Message | Tee-Object $LogFilePath -Append
}


function Get-HtmlReport
{
    Import-Csv $outCsvFilePath | ConvertTo-Html | Where-Object {
            $_ -notmatch "<td>OK</td>"
    } | ForEach-Object {
        $_ -replace "<table>", "<table border=`"1`" style=`"border-collapse: collapse`">" `
           -replace "</td>", "</td>`n" `
           -replace "C:\\([^<]+)</td>", "<a href=`"C:\`$1`"><img src=`"C:\`$1`" width=`"300`"></a></td>"
    } | Out-File $outHtmlFilePath -Encoding utf8

    $csvObj = Import-Csv $outCsvFilePath
    $csvObj | Select-Object * -ExcludeProperty "Image(*" |
        Export-Csv $outCsvFilePath -Encoding UTF8 -NoTypeInformation
}
