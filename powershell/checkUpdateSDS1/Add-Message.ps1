function Add-Message
{
    param(
        [string]$Message,
        [string]$LogFilePath
    )

    "{0} {1}" -f (Get-Date), $Message | Out-File $LogFilePath -Append
}
    