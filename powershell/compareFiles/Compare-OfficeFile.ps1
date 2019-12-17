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
    [string]$afterDir
)

## don't change
$script:outputDir = Join-Path $PSScriptRoot "output"
$script:outCsvFilePath = Join-Path $PSScriptRoot ("result_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv")
$script:outLogFilePath = Join-Path $PSScriptRoot ("result_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".log")
$script:outHtmlFilePath = Join-Path $PSScriptRoot ("result_NG_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".html")
$script:outFilePathOfConvertOffice = Join-Path $PSScriptRoot ("result_convert_office_" + (Get-Date -Format "yyyy-MM-dd_HHmmss") + ".csv")
$script:count = 0

# load function
. ".\Add-Message.ps1"


# main
$startTime = Get-Date


# compare
. ".\Compare-Word.ps1" -beforeDir  $beforeDir -afterDir $afterDir -Office
. ".\Compare-Excel.ps1" -beforeDir  $beforeDir -afterDir $afterDir -Office
. ".\Compare-PowerPoint.ps1" -beforeDir $beforeDir -afterDir $afterDir -Office


# report
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
Write-Host ("StartTime: {0}" -f $startTime)
Write-Host ("EndTime: {0}" -f $endTime)
Write-Host ("TotalTime: {0}" -f ($endTime - $startTime))
Write-Host ("TotalCount: {0}" -f $script:count)
