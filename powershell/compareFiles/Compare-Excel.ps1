param(
    [parameter(mandatory)]
    [string]$beforeDir,
    [parameter(mandatory)]
    [string]$afterDir,
    [switch]$Office
)

## change if needed
# set the threshold of differency
# the smaller the difference, the value is close to 0.
#$identifyThreshold = "1000"


## don't change
$excelRegex = "^.*`.(xls|xlsx|xlsm|xlt|xltx|xltm)$"
$script:pdfArray = @()

if(-not $Office)
{
    $outCsvFilePath = Join-Path $PSScriptRoot ("result_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv")
    $outHtmlFilePath = Join-Path $PSScriptRoot ("result_NG_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".html")
    $outFilePathOfConvertOffice = Join-Path $PSScriptRoot ("result_convert_office_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv")
    $script:count = 0
}


function Convert-ExcelToPdf
{
    param(
        [parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
        [string]$Path,
        [parameter()]
        [string]$OutDir
    )
    
    begin
    {
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} START converting Word to PDF" -f (Get-Date))
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
    }
    
    process
    {
        $excelFilePath = [string](Resolve-Path $Path)
        $outDirFullPath = [string](Resolve-Path (Split-Path -Parent $Path))
        if ($OutDir)
        {
            $outDirFullPath = [string](Resolve-Path $outDir)
        }
        $pdfFilePath = [string]((Join-Path $outDirFullPath (Split-Path -Leaf $Path)) + ".pdf")
        $result = ""
        $errMessage = ""
        
        try
        {
            $excelApplication = New-Object -ComObject Excel.Application
            $excelApplication.Visible = $false
    
            # DEBUG
            # Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} opening {1} ..." -f (Get-Date), $excelFilePath)
            # https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbooks.open
            $workbooks = $excelApplication.Workbooks.Open($excelFilePath,    #FileName
                                                            $false,          #UpdateLinks
                                                            $true,           #ReadOnly
                                                            [Type]::Missing, #Format
                                                            "xxxxx")         #Password
    
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} converting {1} to PDF ..." -f (Get-Date), $excelFilePath)
            Add-Message ("converting {0} to PDF ..." -f $excelFilePath) $outLogFilePath
            # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.tools.excel.worksheet.exportasfixedformat?view=vsto-2017
            $workbooks.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.xlFixedFormatType]::xlTypePDF, $pdfFilePath)
            $workbooks.Saved = $true
            $script:pdfArray += $pdfFilePath
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} converting {1} is finished." -f (Get-Date), $excelFilePath)
            Add-Message ("converting {0} is finished." -f $excelFilePath) $outLogFilePath
            $result = "OK"
        }
        catch
        {
            if ($_.Exception.Message -match "入力したパスワードが間違っています")
            {
                $errMessage = "パスワード保護"
            }
            elseif ($_.Exception.Message -match "ファイルが壊れている可能性があります")
            {
                $errMessage = "ファイル破損"
            }
            elseif ($_.Exception.Message -match "ファイル形式がファイル拡張子と一致していない")
            {
                $errMessage = "拡張子不一致"
            }
            else
            {
                $errMessage = "Error: {0}" -f $_.Exception.Message
            }
            $result = "NG"
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is failed to convert. ({2})" -f (Get-Date), $excelFilePath, $errMessage)
            Add-Message ("{0} is failed to convert. ({1})" -f $excelFilePath, $errMessage) $outLogFilePath
        }
        finally
        {
            # closing
            if (Test-Path Variable:workbooks)
            {
                $workbooks.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
                $workbooks = $null
                Remove-Variable workbooks -ErrorAction SilentlyContinue
                [GC]::Collect | Out-Null
                [GC]::WaitForPendingFinalizers() | Out-Null
                [GC]::Collect | Out-Null
            }
    
            if (Test-Path Variable:excelApplication)
            {
                $excelApplication.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApplication) | Out-Null
                $excelApplication = $null
                Remove-Variable excelApplication -ErrorAction SilentlyContinue
                [GC]::Collect | Out-Null
                [GC]::WaitForPendingFinalizers() | Out-Null
                [GC]::Collect | Out-Null
            }
    
            # export to csv
            $arrayResult = @()
            $objectOfEachRecord = [pscustomobject]@{
                FileName=$excelFilePath
                Result=$result
                Error=$errMessage
            }
            $arrayResult += $objectOfEachRecord
            $arrayResult | Export-Csv $outFilePathOfConvertOffice -encoding Default -NoTypeInformation -Append
            Write-Host ""
        }
    }
    
    end
    {

        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} FINISHED converting Word to PDF" -f (Get-Date))
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
    }
    
}


# main
if(-not $Office)
{
    $startTime = Get-Date
}

Get-ChildItem $beforeDir | Where-Object { $_.Name -match $excelRegex } | Convert-ExcelToPdf
$beforeFiles = $script:pdfArray
$script:pdfArray = @()
Get-ChildItem $afterDir | Where-Object { $_.Name -match $excelRegex } | Convert-ExcelToPdf
$afterFiles = $script:pdfArray
. ".\Compare-Pdf.ps1" -beforeFiles $beforeFiles -afterFiles $afterFiles -Office


if(-not $Office)
{
    Import-Csv $outCsvFilePath | ConvertTo-Html | Where-Object {
        $_ -notmatch "<td>OK</td>"
    } | ForEach-Object {
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
}

