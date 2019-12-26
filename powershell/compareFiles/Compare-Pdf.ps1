<#
.SYNOPSIS

Compare two PDF files groups and output the difference.

.PARAMETER beforeDir

Directory path including PDF files before sanitizing

.PARAMETER afterDir

Directory path including PDF files after sanitizing

.EXAMPLE

PS> .\Compare-Pdf.ps1 .\before .\after
#>

[CmdletBinding()]
param(
    [parameter(mandatory)]
    [string]$beforeDir,
    [parameter(mandatory)]
    [string]$afterDir
)


# load function
. ".\Compare-CommonFunc.ps1"

# install check
if(-not (Test-InstalledIM)) { exit 1 }

# main
$startTime = Get-Date

# compare
Get-ChildItem $beforeDir | Where-Object { $_.Name -like "*.pdf" } | Compare-Image

# report
Get-HtmlReport

# measurement
$endTime = Get-Date
Write-Host ("Start: {0}" -f $startTime)
Write-Host ("End: {0}" -f $endTime)
Write-Host ("Total: {0}" -f ($endTime - $startTime))
Write-Host ("TotalCount: {0}" -f $script:count)
