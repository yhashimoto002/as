function Add-Message
{
    param(
        [parameter(mandatory)]
        [string]$Message,
        [parameter(mandatory)]
        [string]$LogFilePath
    )

    "{0:yyyy/MM/dd HH:mm:ss.fff} {1}" -f (Get-Date), $Message | Tee-Object $LogFilePath -Append
}
    