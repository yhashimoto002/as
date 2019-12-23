<#
.SYNOPSIS

Compare two PowerPoint files groups and output the difference.

.PARAMETER beforeDir

Directory path including PowerPoint files before sanitizing

.PARAMETER afterDir

Directory path including PowerPoint files after sanitizing

.EXAMPLE

PS> .\Compare-PowerPoint.ps1 .\before .\after
#>

[CmdletBinding()]
param(
    [parameter(mandatory)]
    [string]$beforeDir,
    [parameter(mandatory)]
    [string]$afterDir,
    [switch]$Office
)

## don't change
$powerpointRegex = "^.*`.(ppt|pptx|pptm|pot|potx|potm|pps|ppsx|ppsm)$"

if(-not $Office)
{
    $outFilePathOfConvertOffice = Join-Path $PSScriptRoot ("result_convert_office_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv")
    . ".\Add-Message.ps1"
    . ".\Compare-CommonFunc.ps1"
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
    if (-not (Test-Path $OutDir)) { New-Item $OutDir -ItemType "Directory" -Force | Out-Null }
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
if(-not $Office)
{
    $startTime = Get-Date
}

# compare
Get-ChildItem $beforeDir | Where-Object { $_.Name -match $powerpointRegex } | Compare-Image -PowerPoint

if(-not $Office)
{
    # report
    Get-HtmlReport

    # measurement
    $endTime = Get-Date
    Write-Host ("Start: {0}" -f $startTime)
    Write-Host ("End: {0}" -f $endTime)
    Write-Host ("Total: {0}" -f ($endTime - $startTime))
    Write-Host ("TotalCount: {0}" -f $script:count)
}

