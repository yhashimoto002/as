param(
    [parameter(mandatory)]
    [string]$beforeDir,
    [parameter(mandatory)]
    [string]$afterDir
)


function Convert-WordToPdf
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
        $wordFilePath = [string](Resolve-Path $Path)
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
            $wordApplication = New-Object -ComObject Word.Application
            $wordApplication.Visible = $false

            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} opening {1} ..." -f (Get-Date), $wordFilePath)
            # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.word.documents.opennorepairdialog?view=word-pia
            $documents = $wordApplication.Documents.OpenNoRepairDialog($wordFilePath, #FileName
                                                                        $false,       #ConfirmConversions
                                                                        $true,        #ReadOnly
                                                                        $false,       #AddToRecentFiles
                                                                        "xxxxxx")     #PasswordDocument
           
            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} converting {1} to Pdf ..." -f (Get-Date), $wordFilePath)
            # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.word._document.exportasfixedformat?view=word-pia
            $documents.ExportAsFixedFormat($pdfFilePath, [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF)
            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is successfully converted to PDF." -f (Get-Date), $wordFilePath)
            $result = "OK"
        }
        catch
        {
            #Write-Error ("Error: {0}" -f $_.Exception.Message)
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
                $errMessage = "Error: {0}" -f $_.Exception.Message
            }
            $result = "NG"
            Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is failed to convert. ({2})" -f (Get-Date), $wordFilePath, $errMessage)
        }
        finally
        {
            # closing
            #if ($documents) { $documents.Close() }
            # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.word.documents.close?view=word-pia
            if ($documents) { $documents.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges) }
            $wordApplication.Quit()
            $documents = $wordApplication = $null
            [GC]::Collect()

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
        Remove-Item $outFilePath -ErrorAction Stop
    }
    catch
    {
        Write-Error ("Error: {0}" -f $_.Exception.Message)
        exit 1
    }
}

$docRegex = "^.*`.(doc|docx|docm|dot|dotx|dotm)$"
dir $beforeDir | ? { $_.FullName -match $docRegex } | Convert-WordToPdf
dir $afterDir | ? { $_.FullName -match $docRegex } | Convert-WordToPdf


# compare images and analyze the difference
. ".\Compare-Pdf.ps1" -beforeDir $beforeDir -afterDir $afterDir

$endTime = Get-Date
Write-Host ("Start: {0}" -f $startTime)
Write-Host ("End: {0}" -f $endTime)
Write-Host ("Total: {0}" -f ($endTime - $startTime))

