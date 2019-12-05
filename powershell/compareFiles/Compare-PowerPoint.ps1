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


function Convert-PowerPointToPng
{
    param(
        [parameter(Mandatory)]
        [string]$Path,
        [parameter(Mandatory)]
        [string]$OutDir
    )

    begin
    {
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} START converting PowerPoint to PDF" -f (Get-Date))
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
    }

    process
    {
        $powerpointFullPath = Resolve-Path $Path
        if (-not (Test-Path $OutDir)) { mkdir $OutDir -Force | Out-Null }
        try
        {
            try
            {
                $powerpointApplication = New-Object -ComObject PowerPoint.Application
            }
            catch
            {
                Write-Host ("cannot create com object: {0}" -f $_.Exception)
            }
            #$powerpointApplication.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} opening {1} ..." -f (Get-Date), $powerpointFullPath)
            # https://docs.microsoft.com/ja-jp/previous-versions/office/developer/office-2010/ff763759%28v%3doffice.14%29
            $password = "xxxxx"
            $presentations = $powerpointApplication.Presentations.Open([string]$powerpointFullPath+"::$password",
                                                                        [Microsoft.Office.Core.MsoTriState]::msoTrue,  # readonly
                                                                        [Type]::Missing,                               # untitled
                                                                        [Microsoft.Office.Core.MsoTriState]::msoFalse) # withwindow
            
            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} converting {1} to PNG ..." -f (Get-Date), $powerpointFullPath)
            # https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff762466%28v%3doffice.14%29
            $presentations.SaveAs($OutDir, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPNG)
            $presentations.Saved = $true
            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is successfully converted to PNG." -f (Get-Date), $powerpointFullPath)
        }
        catch
        {
            if ($_.Exception.Message -match "読み取りパスワードをもう一度入力してください")
            {
                $errMessage = "パスワード保護"
            }
            elseif ($_.Exception.Message -match "HRESULT")
            {
                $errMessage = "ファイル破損"
            }
            else
            {
                $errMessage = "Error: {0}" -f $_.Exception.Message
            }
            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is failed to convert. ({2})" -f (Get-Date), $powerpointFullPath, $errMessage)
        }
        finally
        {
            # closing
            if ($presentations) { $presentations.Close() }
            if ($powerpointApplication) { $powerpointApplication.Quit() }
            $presentations = $powerpointApplication = $null
            [void][GC]::Collect
            [void][GC]::WaitForPendingFinalizers()
            [void][GC]::Collect

            # export to csv
            $arrayResult = @()
            $objectOfEachRecord = [pscustomobject]@{
                FileName=$powerpointFullPath
                Result=$result
                Error=$errMessage
            }
            $arrayResult += $objectOfEachRecord
            $arrayResult | Export-Csv $outFilePathOfConvertOffice  -encoding Default -NoTypeInformation -Append
            Write-Host ""
        }
    }

    end
    {
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} FINISHED converting PowerPoint to PDF" -f (Get-Date))
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
    }
}


function Compare-PowerPoint
{
    param(
    [parameter(Mandatory, ValueFromPipeline)]
    [string]$PowerPoint
    )

    process
    {
        # skip if target PDF doesn't exist in the opposite dir
        if (! (Test-Path (Join-Path $afterDir $PowerPoint))) { return }
        
        $before_dir = Join-Path $outputDir $PowerPoint | Join-Path -ChildPath "before"
        $after_dir = Join-Path $outputDir $PowerPoint | Join-Path -ChildPath "after"
        $diff_dir = Join-Path $outputDir $PowerPoint | Join-Path -ChildPath "diff"
        mkdir $diff_dir -Force | Out-Null

        # convert pdf to image
        Write-Host ("{0}" -f ++$count)
        Convert-PowerPointToPng -Path (Join-Path $beforeDir $PowerPoint) -OutDir $before_dir
        Convert-PowerPointToPng -Path (Join-Path $afterDir $PowerPoint) -OutDir $after_dir

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
            Write-Host ("{0}/{1}: {2}({3})" -f $PowerPoint, $png, $result, $identify) 
            $objectOfEachRecord = [pscustomobject]@{
                "No."=$count
                FileName=$PowerPoint
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
        Write-Host "------------------------------"
    }
}


# main
$startTime = Get-Date

if (Test-Path $outCsvFilePath)
{
    try
    {
        Remove-Item $outCsvFilePath -ErrorAction Stop
    }
    catch
    {
        Write-Error ("Error: {0}" -f $_.Exception.Message)
        exit 1
    }
}

$outFilePathOfConvertOffice = Join-Path $PSScriptRoot ("result_convert_office_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv")
if (Test-Path $outFilePathOfConvertOffice)
{
    try
    {
        Remove-Item $outFilePathOfConvertOffice -ErrorAction Stop
    }
    catch
    {
        Write-Error ("Error: {0}" -f $_.Exception.Message)
        exit 1
    }
}

$powerpointRegex = "^.*`.(ppt|pptx|pptm|pot|potx|potm|pps|ppsx|ppsm)$"
dir $beforeDir | ? { $_.Name -match $powerpointRegex } | Compare-PowerPoint

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

[void][GC]::Collect
[void][GC]::WaitForPendingFinalizers()
[void][GC]::Collect



