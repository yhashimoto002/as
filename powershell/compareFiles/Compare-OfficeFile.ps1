<#
.SYNOPSIS

Compare two Office files groups and output the difference.

.PARAMETER beforeDir

Directory path including Office files before sanitizing

.PARAMETER afterDir

Directory path including Office files after sanitizing

.EXAMPLE

PS> .\Compare-OfficeFiles.ps1 .\before .\after
#>

[CmdletBinding()]
param(
    [parameter(mandatory)]
    [string]$beforeDir,
    [parameter(mandatory)]
    [string]$afterDir,
    [switch]$Word,
    [switch]$Excel,
    [switch]$PowerPoint
)

## don't change
$docRegex = "^.*`.(doc|docx|docm|dot|dotx|dotm)$"
$excelRegex = "^.*`.(xls|xlsx|xlsm|xlt|xltx|xltm)$"
$powerpointRegex = "^.*`.(ppt|pptx|pptm|pot|potx|potm|pps|ppsx|ppsm)$"
$outFilePathOfConvertOffice = Join-Path $PSScriptRoot ("result_convert_office_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv")

# load function
. ".\Compare-CommonFunc.ps1"

# install check
if(-not (Test-InstalledIM)) { exit 1 }

function Convert-WordToPdf
{
    param(
        [parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $FullName
    )

    begin
    {
        $script:pdfArray = @()
    }

    process
    {
        $wordFilePath = $FullName
        $pdfFilePath = $wordFilePath + ".pdf"
        $result = ""
        $errMessage = ""

        try
        {
            $wordApplication = New-Object -ComObject Word.Application
            $wordApplication.Visible = $false

            # DEBUG
            # Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} opening {1} ..." -f (Get-Date), $wordFilePath)
            # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.word.documents.opennorepairdialog?view=word-pia
            $documents = $wordApplication.Documents.OpenNoRepairDialog($wordFilePath, #FileName
                                                                        $false,       #ConfirmConversions
                                                                        $true,        #ReadOnly
                                                                        $false,       #AddToRecentFiles
                                                                        "xxxxxx")     #PasswordDocument
           
            Add-Message ("converting {0} to Pdf ..." -f $wordFilePath)  $outLogFilePath
            # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.word._document.exportasfixedformat?view=word-pia
            $documents.ExportAsFixedFormat($pdfFilePath, [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF)
            $script:pdfArray += $pdfFilePath
            Add-Message ("converting {0} is finished." -f $wordFilePath)  $outLogFilePath
            $result = "OK"
        }
        catch
        {
            if ($_.Exception.Message -match "パスワードが正しくありません")
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
                $errMessage = "不明"
            }
            $result = "NG"
            Add-Message ("{0} is failed to convert. ({1})`nERROR: {2}" -f $wordFilePath, $errMessage, $_.Exception)  $outLogFilePath
        }
        finally
        {
            # closing
            if (Test-Path Variable:documents)
            {
                # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.word.documents.close?view=word-pia
                $documents.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($documents) | Out-Null
                $documents = $null
                Remove-Variable documents -ErrorAction SilentlyContinue
                [GC]::Collect | Out-Null
                [GC]::WaitForPendingFinalizers() | Out-Null
                [GC]::Collect | Out-Null
            }

            if (Test-Path Variable:wordApplication)
            {
                $wordApplication.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApplication) | Out-Null
                $wordApplication = $null
                Remove-Variable wordApplication -ErrorAction SilentlyContinue
                [GC]::Collect | Out-Null
                [GC]::WaitForPendingFinalizers() | Out-Null
                [GC]::Collect | Out-Null
            }

            # export to csv
            $arrayResult = @()
            $objectOfEachRecord = [pscustomobject]@{
                FileName=$wordFilePath
                Result=$result
                Error=$errMessage
            }
            $arrayResult += $objectOfEachRecord
            $arrayResult | Export-Csv $outFilePathOfConvertOffice  -encoding Default -NoTypeInformation -Append
            Write-Host ""
        }
    }

}


function Convert-ExcelToPdf
{
    param(
        [parameter(Mandatory, ValueFromPipelineByPropertyName)]
        $FullName
    )

    begin
    {
        $script:pdfArray = @()
    }

    process
    {
        $excelFilePath = $FullName
        $pdfFilePath = $excelFilePath + ".pdf"
        $result = ""
        $errMessage = ""
        
        try
        {
            $excelApplication = New-Object -ComObject Excel.Application
            $excelApplication.Visible = $false
            $excelApplication.DisplayAlerts = $false
    
            # DEBUG
            # Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} opening {1} ..." -f (Get-Date), $excelFilePath)
            # https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbooks.open
            $workbooks = $excelApplication.Workbooks
            $workbook = $workbooks.Open($excelFilePath,    #FileName
                                        $false,            #UpdateLinks
                                        $true,             #ReadOnly
                                        [Type]::Missing,   #Format
                                        "xxxxx")           #Password
    
            Add-Message ("converting {0} to PDF ..." -f $excelFilePath) $outLogFilePath
            # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.tools.excel.worksheet.exportasfixedformat?view=vsto-2017
            $workbook.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.xlFixedFormatType]::xlTypePDF, $pdfFilePath)                                                                      #OpenAfterPublish
            $workbook.Saved = $true
            $script:pdfArray += $pdfFilePath
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
                #$errMessage = "Error: {0}" -f $_.Exception.Message
                $errMessage = "不明"
            }
            $result = "NG"
            Add-Message ("{0} is failed to convert. ({1})`nERROR: {2}" -f $excelFilePath, $errMessage, $_.Exception) $outLogFilePath
        }
        finally
        {
            # closing
            if (Test-Path Variable:workbook)
            {
                $workbook.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                $workbook = $null
                Remove-Variable workbook -ErrorAction SilentlyContinue
                [GC]::Collect | Out-Null
                [GC]::WaitForPendingFinalizers() | Out-Null
                [GC]::Collect | Out-Null
            }
    
            if (Test-Path Variable:workbooks)
            {
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
    
}


function Convert-PowerPointToPng
{
    param(
        [parameter(Mandatory)]
        [string]$Path,
        [parameter(Mandatory)]
        [string]$OutDir
    )

    $powerpointFullPath = Resolve-Path $Path
    $result = ""
    $errMessage = ""

    try
    {
        $powerpointApplication = New-Object -ComObject PowerPoint.Application

        # DEBUG
        # Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} opening {1} ..." -f (Get-Date), $powerpointFullPath)
        # https://docs.microsoft.com/ja-jp/previous-versions/office/developer/office-2010/ff763759%28v%3doffice.14%29
        $password = "xxxxx"
        $presentations = $powerpointApplication.Presentations
        $presentation = $presentations.Open([string]$powerpointFullPath+"::$password",
                                                                    [Microsoft.Office.Core.MsoTriState]::msoTrue,  # readonly
                                                                    [Type]::Missing,                               # untitled
                                                                    [Microsoft.Office.Core.MsoTriState]::msoFalse) # withwindow
        
        Add-Message ("converting {0} to PNG ..." -f $powerpointFullPath) $outLogFilePath
        # https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff762466%28v%3doffice.14%29
        $presentation.SaveAs($OutDir, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPNG)
        $presentation.Saved = [Microsoft.Office.Core.MsoTriState]::msoTrue
        Add-Message ("converting {0} is finished." -f $powerpointFullPath) $outLogFilePath
        $result = "OK"
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
            $errMessage = "不明"
        }
        $result = "NG"
        Add-Message ("{0} is failed to convert. ({1})`nERROR: {2}" -f $powerpointFullPath, $errMessage, $_.Exception) $outLogFilePath

    }
    finally
    {
        # closing
        # https://qiita.com/mima_ita/items/aa811423d8c4410eca71
        if (Test-Path Variable:presentation)
        {
            $presentation.Close()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
            $presentation = $null
            Remove-Variable presentation -ErrorAction SilentlyContinue
            [GC]::Collect | Out-Null
            [GC]::WaitForPendingFinalizers() | Out-Null
            [GC]::Collect | Out-Null
        }

        if (Test-Path Variable:presentations)
        {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentations) | Out-Null
            $presentations = $null
            Remove-Variable presentations -ErrorAction SilentlyContinue
            [GC]::Collect | Out-Null
            [GC]::WaitForPendingFinalizers() | Out-Null
            [GC]::Collect | Out-Null
        }

        if (Test-Path Variable:powerpointApplication)
        {
            $powerpointApplication.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpointApplication) | Out-Null
            $powerpointApplication = $null
            Remove-Variable powerpointApplication -ErrorAction SilentlyContinue
            [GC]::Collect | Out-Null
            [GC]::WaitForPendingFinalizers() | Out-Null
            [GC]::Collect | Out-Null
        }

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

# main
$startTime = Get-Date
if($Word -or (-not $Word -and -not $Excel -and -not $PowerPoint))
{
    Get-ChildItem $beforeDir | Where-Object { $_.Name -match $docRegex } | Convert-WordToPdf
    $beforeFiles = $script:pdfArray
    if($beforeFiles)
    {
        Get-ChildItem $afterDir | Where-Object { $_.Name -match $docRegex } | Convert-WordToPdf
        Get-ChildItem $beforeFiles -Name | Compare-Image -Word
    }
}
if($Excel -or (-not $Word -and -not $Excel -and -not $PowerPoint))
{
    Get-ChildItem $beforeDir | Where-Object { $_.Name -match $excelRegex } | Convert-ExcelToPdf
    $beforeFiles = $script:pdfArray
    if($beforeFiles)
    {
        Get-ChildItem $afterDir | Where-Object { $_.Name -match $excelRegex } | Convert-ExcelToPdf
        Get-ChildItem $beforeFiles -Name | Compare-Image -Excel
    }
}
if($PowerPoint -or (-not $Word -and -not $Excel -and -not $PowerPoint))
{
    Get-ChildItem $beforeDir | Where-Object { $_.Name -match $powerpointRegex } | Compare-Image -PowerPoint
}

# report
Get-HtmlReport

$endTime = Get-Date
Write-Host ("StartTime: {0}" -f $startTime)
Write-Host ("EndTime: {0}" -f $endTime)
Write-Host ("TotalTime: {0}" -f ($endTime - $startTime))
Write-Host ("TotalCount: {0}" -f $script:count)
