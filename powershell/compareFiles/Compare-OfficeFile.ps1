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
        [parameter(Mandatory, ValueFromPipelineByPropertyName)]
        [Alias('FullName')]
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
        Get-ChildItem $before_dir | Sort-Object -Property LastWriteTime | ForEach-Object {
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
$excelRegex = "^.*`.(xls|xlsx|xlsm|xlt|xltx|xltm)$"
$powerpointRegex = "^.*`.(ppt|pptx|pptm|pot|potx|potm|pps|ppsx|ppsm)$"
#Get-ChildItem $beforeDir | ? { $_.FullName -match $docRegex } | Convert-WordToPdf
#Get-ChildItem $beforeDir | ? { $_.FullName -match $excelRegex } | Convert-ExcelToPdf
#Get-ChildItem $afterDir | ? { $_.FullName -match $docRegex } | Convert-WordToPdf
#Get-ChildItem $afterDir | ? { $_.FullName -match $excelRegex } | Convert-ExcelToPdf
#Get-ChildItem $beforeDir | ? { $_.Name -match $powerpointRegex } | Compare-PowerPoint


foreach($file in Get-ChildItem $beforeDir)
{
    if($file.FullName -match $docRegex) { Convert-WordToPdf $file.FullName }
    if($file.FullName -match $excelRegex) { Convert-ExcelToPdf $file.FullName }
    if($file.FullName -match $powerpointRegex) { Compare-PowerPoint $file.FullName }
}

foreach($file in Get-ChildItem $afterDir)
{
    if($file.FullName -match $docRegex) { Convert-WordToPdf $file.FullName }
    if($file.FullName -match $excelRegex) { Convert-ExcelToPdf $file.FullName }
    if($file.FullName -match $powerpointRegex) { Compare-PowerPoint $file.FullName }
}

# compare images and analyze the difference
. ".\Compare-Pdf.ps1" -beforeDir $beforeDir -afterDir $afterDir

$endTime = Get-Date
Write-Host ("Start: {0}" -f $startTime)
Write-Host ("End: {0}" -f $endTime)
Write-Host ("Total: {0}" -f ($endTime - $startTime))

