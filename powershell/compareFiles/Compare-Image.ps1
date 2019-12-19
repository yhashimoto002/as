<#
.SYNOPSIS

Compare 2 Images and output the difference.

.PARAMETER beforeDir

Directory path including Images before sanitizing

.PARAMETER afterDir

Directory path including Images after sanitizing

.OUTPUTS

CSV file

.EXAMPLE

PS> .\Compare-Image.ps1 .\before .\after

#>

param(
    [parameter(mandatory)]
    [string]$beforeDir,
    [parameter(mandatory)]
    [string]$afterDir
)

## don't change
$imageRegex = "^.*`.(jpg|jpeg|jp2|png|gif|tif|tiff|emf|wmf)$"

# load function
. ".\Compare-CommonFunc.ps1"

# main
$startTime = Get-Date
Get-ChildItem $beforeDir | Where-Object { $_.Name -match $imageRegex } | Compare-Image

# report
Get-HtmlReport

# measurement
$endTime = Get-Date
Write-Host ("Start: {0}" -f $startTime)
Write-Host ("End: {0}" -f $endTime)
Write-Host ("Total: {0}" -f ($endTime - $startTime))
Write-Host ("TotalCount: {0}" -f $script:count)
