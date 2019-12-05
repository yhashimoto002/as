param(
    [parameter(mandatory)]
    [string]$beforeDir,
    [parameter(mandatory)]
    [string]$afterDir
)


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

            #$workbooks = $excelApplication.Workbooks.OpenNoRepairDialog($excelFilePath)
            # https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbooks.open
            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} opening {1} ..." -f (Get-Date), $excelFilePath)
            $workbooks = $excelApplication.Workbooks.Open($excelFilePath,    #FileName
                                                            $false,          #UpdateLinks
                                                            $true,           #ReadOnly
                                                            [Type]::Missing, #Format
                                                            "xxxxx")         #Password

            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} converting {1} to PDF ..." -f (Get-Date), $excelFilePath)
            # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.tools.excel.worksheet.exportasfixedformat?view=vsto-2017
            $workbooks.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.xlFixedFormatType]::xlTypePDF, $pdfFilePath)
            $workbooks.Saved = $true
            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is successfully converted to PDF." -f (Get-Date), $excelFilePath)
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
            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is failed to convert. ({2})" -f (Get-Date), $excelFilePath, $errMessage)
        }
        finally
        {
            # closing
            if ($workbooks) { $workbooks.Close($false) }
            $excelApplication.Quit()
            $workbooks = $excelApplication = $null
            [GC]::Collect()

            # export to csv
            $arrayResult = @()
            $objectOfEachRecord = [pscustomobject]@{
                FileName=$excelFilePath
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
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} FINISHED converting Word to PDF" -f (Get-Date))
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
    }

}


# main
$startTime = Get-Date

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

$excelRegex = "^.*`.(xls|xlsx|xlsm|xlt|xltx|xltm)$"
dir $beforeDir | ? { $_.FullName -match $excelRegex } | Convert-ExcelToPdf
dir $afterDir | ? { $_.FullName -match $excelRegex } | Convert-ExcelToPdf


# compare images and analyze the difference
. ".\Compare-Pdf.ps1" -beforeDir $beforeDir -afterDir $afterDir

$endTime = Get-Date
Write-Host ("Start: {0}" -f $startTime)
Write-Host ("End: {0}" -f $endTime)
Write-Host ("Total: {0}" -f ($endTime - $startTime))

