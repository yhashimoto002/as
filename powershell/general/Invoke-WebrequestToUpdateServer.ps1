$curl = Join-Path $PSScriptRoot "curl.exe"

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
    if ( $Method -eq "HEAD" )
    {
        & $curl -I -sS -f --connect-timeout 5 --user ${userNameForCred}:$passwordForCred $url 2>&1
    }
    else
    {
        & $curl -f -D - --connect-timeout 5 --user ${userNameForCred}:$passwordForCred -o $outFile $url
    }
}