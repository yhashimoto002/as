<#
.SYNOPSIS 

This script converts Word, Excel, and PowerPoint files to PDF.

.PARAMETER SourceFiles 

Target files you want to convert
 
.PARAMETER Word

If this switch is add, the target is limited to Word files.

.PARAMETER Excel

If this switch is add, the target is limited to Excel files.

.PARAMETER PowerPoint

If this switch is add, the target is limited to PowerPoint files.

.PARAMETER All

If this switch is add, all office files (Word, Excel, and PowerPoint) are targeted.

.PARAMETER outputDirectory

Set a output directory. If this parameter is not set, use the same directory as of target files.

.EXAMPLE

C:\PS> Confert-Dxp2Pdf C:\Work\sample.xlsx

.EXAMPLE

C:\PS> Confert-Dxp2Pdf .\sample.xlsx .\output

.EXAMPLE

C:\PS> Confert-Dxp2Pdf C:\Work\* -Word

.EXAMPLE

C:\PS> .\* | Confert-Dxp2Pdf -All


#>


#https://gallery.technet.microsoft.com/office/d8cebc7d-73a7-42a0-aabb-0e73b1e26ac1
[CmdletBinding()] 
param(
    [parameter(mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName,position=0)]
    [Alias('FullName')]
    [Object[]]$sourceFiles,
    [parameter(mandatory=$false,position=1)]
    [string]$outputDirectory,
    [switch]$Word,
    [switch]$Excel,
    [switch]$PowerPoint,
    [switch]$All
)

begin
{
    $startTime = Get-Date

    # do not change
    $docRegex = "^.*`.(doc|docx|docm|dot|dotx|dotm)$"
    $excelRegex = "^.*`.(xls|xlsx|xlsm|xlt|xltx|xltm)$"
    $powerpointRegex = "^.*`.(ppt|pptx|pptm|pot|potx|potm|pps|ppsx|ppsm)$"

    $outFileName = "result_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv"
    $outFilePath = Join-Path $PSScriptRoot $outFileName
    if (Test-Path $outFilePath) { rm $outFilePath -Force }

}


process
{
    # if specified no $outputDirectory, output the same directory as input
    $outputDirectoryFullPath = Resolve-Path (Split-Path -Parent $sourceFiles)
    if ($outputDirectory)
    {
        $outputDirectoryFullPath = Resolve-Path $outputDirectory
    }


    function Convert-Word
    {
        param(
            [parameter(Mandatory, ValueFromPipeline)]
            $Object
        )

        begin
        {
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} START converting Word to PDF" -f (Get-Date))
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
        }

        process
        {
            $docFilePath = $Object.FullName
            $pdfFilePath = (Join-Path $outputDirectoryFullPath $Object.Name) + ".pdf"
            $result = ""
            $errMessage = ""

            try
            {
                $wordApplication = New-Object -ComObject Word.Application
                $wordApplication.Visible = $false

                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} opening {1} ..." -f (Get-Date), $docFilePath)
                # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.word.documents.opennorepairdialog?view=word-pia
                $documents = $wordApplication.Documents.OpenNoRepairDialog($docFilePath, #FileName
                                                                           $false,       #ConfirmConversions
                                                                           $true,        #ReadOnly
                                                                           $false,       #AddToRecentFiles
                                                                           "xxxxxx")     #PasswordDocument
           
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} converting {1} to PDF ..." -f (Get-Date), $docFilePath)
                # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.interop.word._document.exportasfixedformat?view=word-pia
                $documents.ExportAsFixedFormat($pdfFilePath, [Microsoft.Office.Interop.Word.WdExportFormat]::wdExportFormatPDF)
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is successfully converted to PDF." -f (Get-Date), $docFilePath)
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
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is failed to convert. ({2})" -f (Get-Date), $docFilePath, $errMessage)
            }
            finally
            {
                # closing
                if ($documents) { $documents.Close() }
                $wordApplication.Quit()
                $documents = $wordApplication = $null
                [GC]::Collect()

                # export to csv
                $arrayResult = @()
                $objectOfEachRecord = [pscustomobject]@{
                    FileName=$docFilePath
                    Result=$result
                    Error=$errMessage
                }
                $arrayResult += $objectOfEachRecord
                $arrayResult | Export-Csv $outFilePath  -encoding Default -NoTypeInformation -Append
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
            }
        }

    end
    {
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} FINISHED converting Word to PDF" -f (Get-Date))
        #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
    }

}


    function Convert-Excel
    {
        param(
            [parameter(Mandatory, ValueFromPipeline)]
            $Object
        )

        begin
        {
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} START converting Excel to PDF" -f (Get-Date))
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
        }

        process
        {
            $xlsFilePath = $Object.FullName
            $pdfFilePath = (Join-Path $outputDirectoryFullPath $Object.Name) + ".pdf"
            if (Test-Path $pdfFilePath) { rm $pdfFilePath -Force }
            $result = ""
            $errMessage = ""

            try
            {
                #Add-Type -AssemblyName Microsoft.Office.Interop.Excel
                #$excelApplication = New-Object Microsoft.Office.Interop.Excel.ApplicationClass
                $excelApplication = New-Object -ComObject Excel.Application
                $excelApplication.Visible = $false

                #$workbooks = $excelApplication.Workbooks.OpenNoRepairDialog($xlsFilePath)
                # https://docs.microsoft.com/ja-jp/office/vba/api/excel.workbooks.open
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} opening {1} ..." -f (Get-Date), $xlsFilePath)
                $workbooks = $excelApplication.Workbooks.Open($xlsFilePath,    #FileName
                                                              $false,          #UpdateLinks
                                                              $true,           #ReadOnly
                                                              [Type]::Missing, #Format
                                                              "xxxxx")         #Password

                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} converting {1} to PDF ..." -f (Get-Date), $xlsFilePath)
                # https://docs.microsoft.com/ja-jp/dotnet/api/microsoft.office.tools.excel.worksheet.exportasfixedformat?view=vsto-2017
                $workbooks.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.xlFixedFormatType]::xlTypePDF, $pdfFilePath)
                $workbooks.Saved = $true
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is successfully converted to PDF." -f (Get-Date), $xlsFilePath)
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
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is failed to convert. ({2})" -f (Get-Date), $xlsFilePath, $errMessage)
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
                    FileName=$xlsFilePath
                    Result=$result
                    Error=$errMessage
                }
                $arrayResult += $objectOfEachRecord
                $arrayResult | Export-Csv $outFilePath  -encoding Default -NoTypeInformation -Append
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
            }
        }
    
        end
        {
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} FINISHED converting Excel to PDF" -f (Get-Date))
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
        }
    }


    function Convert-PowerPoint
    {
        param(
            [parameter(Mandatory, ValueFromPipeline)]
            $Object
        )

        begin
        {
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} START converting PowerPoint to PDF" -f (Get-Date))
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
        }

        process
        {
            $pptFilePath = $Object.FullName
            $pdfFilePath = (Join-Path $outputDirectoryFullPath $Object.Name) + ".pdf"
            if (Test-Path $pdfFilePath) { rm $pdfFilePath -Force }
            $result = ""
            $errMessage = ""

            try
            {
                #Add-Type -AssemblyName Microsoft.Office.Interop.PowerPoint
                #$excelApplication = New-Object Microsoft.Office.Interop.PowerPoint.ApplicationClass
                try
                {
                    $powerpointApplication = New-Object -ComObject PowerPoint.Application
                }
                catch
                {
                    Write-Host ("cannot create com object: {0}" -f $_.Exception)
                }
                #$powerpointApplication.Visible = [Microsoft.Office.Core.MsoTriState]::msoTrue

                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} opening {1} ..." -f (Get-Date), $pptFilePath)
                # https://docs.microsoft.com/ja-jp/previous-versions/office/developer/office-2010/ff763759%28v%3doffice.14%29
                $password = "xxxxx"
                $presentations = $powerpointApplication.Presentations.Open($pptFilePath+"::$password",
                                                                           [Microsoft.Office.Core.MsoTriState]::msoTrue,  # readonly
                                                                           [Type]::Missing,                               # untitled
                                                                           [Microsoft.Office.Core.MsoTriState]::msoFalse) # withwindow
            
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} converting {1} to PDF ..." -f (Get-Date), $pptFilePath)
                # https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff761136(v%3Doffice.14)
                #$presentations.ExportAsFixedFormat($pdfFilePath, [Microsoft.Office.Interop.PowerPoint.PpFixedFormatType]::ppFixedFormatTypePDF)
                # https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/ff762466%28v%3doffice.14%29
                $presentations.SaveAs($pdfFilePath, [Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType]::ppSaveAsPDF)
                $presentations.Saved = $true
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is successfully converted to PDF." -f (Get-Date), $pptFilePath)
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
                    $errMessage = "Error: {0}" -f $_.Exception.Message
                }
                $result = "NG"
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} {1} is failed to convert. ({2})" -f (Get-Date), $pptFilePath, $errMessage)
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
                    FileName=$pptFilePath
                    Result=$result
                    Error=$errMessage
                }
                $arrayResult += $objectOfEachRecord
                $arrayResult | Export-Csv $outFilePath  -encoding Default -NoTypeInformation -Append
                Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
            }
        }

        end
        {
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} FINISHED converting PowerPoint to PDF" -f (Get-Date))
            #Write-Host ("{0:yyyy/MM/dd HH:mm:ss.fff} ------------------------------" -f (Get-Date))
        }
    }


    if($All -or (!$All -and !$Word -and !$Excel -and !$PowerPoint))
    {
        Get-ChildItem $sourceFiles | ? { $_.FullName -match $docRegex } | Convert-Word
        Get-ChildItem $sourceFiles | ? { $_.FullName -match $excelRegex } | Convert-Excel
        Get-ChildItem $sourceFiles | ? { $_.FullName -match $powerpointRegex } | Convert-PowerPoint
    }
    if($Word)
    {
        Get-ChildItem $sourceFiles | ? { $_.FullName -match $docRegex } | Convert-Word
    }
    if($Excel)
    {
        Get-ChildItem $sourceFiles | ? { $_.FullName -match $excelRegex } | Convert-Excel
    }
    if($PowerPoint)
    {
        Get-ChildItem $sourceFiles | ? { $_.FullName -match $powerpointRegex } | Convert-PowerPoint
    }

    [void][GC]::Collect
    [void][GC]::WaitForPendingFinalizers()
    [void][GC]::Collect

}


end
{
    $endTime = Get-Date
    #Write-Host "------------------------------"
    Write-Host ("Start: {0}" -f $startTime)
    Write-Host ("End: {0}" -f $endTime)
    Write-Host ("Total: {0}" -f ($endTime - $startTime))
}

