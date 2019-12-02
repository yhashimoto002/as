function Invoke-WebrequestToUpdateServer
{
    param (
        [parameter(mandatory)][string]$UserName,
        [string]$StationName,
        [string]$Method = "GET",
        [string]$File,
        [string]$OutFile
    )

    $userNameForCred = "asgent"
    $passwordForCred = "gusestA7"
    $cred = New-Object PSCredential $userNameForCred, (ConvertTo-SecureString $passwordForCred -AsPlainText -Force)
    $updateServerUrl = "https://updates.votiro.com/"
    $fAvSmd = "av-smd.bin"
    $fAvSmdSig2 = "av-smd.bin.sig2"
    $fActions = "actions.bin"
    $fActionsSig2 = "actions.bin.sig2"
    $fCsmd = "c-smd.bin"
    $fCsmdSig2 = "c-smd.bin.sig2"
    $fLicense = "license.bin"
    $fLicenseSig2 = "license.bin.sig2"
    $fConfig = "config.bin"
    $fConfigSig2 = "config.bin.sig2"
    $fPsmd = "p-smd.bin"
    $fPsmdSig2 = "p-smd.bin.sig2"
    $fIso = $StationName + "_System.iso"

    # URL
    # av-smd.bin (.sig2): /username/xxx.bin(.sig2)
    # av-smd.bin (.sig2) 以外: /username/station/xxx.bin(.sig2)
    if (($File -eq $fAvSmd) -Or ($File -eq $fAvSmdSig2))
    {
        $url = $updateServerUrl + $UserName + "/" + $File
    }
    else
    {
        $url = $updateServerUrl + $UserName + "/" + $StationName + "/" + $File
    }

    # 出力ファイル
    if (! $outFile)
    {
        $outFile = $File
    }
 

    # Web リクエスト
    try
    {
        $ProgressPreference = "SilentlyContinue"
        if ( $Method -eq "HEAD" )
        {
            Invoke-WebRequest -Uri $url -Method $method -Credential $cred
        }
        else
        {
            Invoke-WebRequest -Uri $url -Method $method -Credential $cred -OutFile $outFile
        }
    }
    catch [System.Net.WebException]
    {
        $_.Exception
        $exceptionDetails = $_.Exception
    }
}