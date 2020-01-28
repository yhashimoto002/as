################################################################################
# パラメータ
################################################################################
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
$userlist = Join-Path $PSScriptRoot "user.txt"
$outFileName = "result_ALL_" + (Get-Date -Format "yyyyMMdd") + ".csv"
$outFilePath = Join-Path $PSScriptRoot $outFileName
$arrayResult = @()


################################################################################
# 関数読み込み
################################################################################
Join-Path $PSScriptRoot "general\*.ps1" | dir | foreach { . $_ }


################################################################################
# 処理
################################################################################
# チェックするファイルリストの生成
$arrayOfDownloadFiles = New-Object System.Collections.Generic.List[string]

$arrayOfDownloadFiles.Add($fLicense)          # license.bin 
$arrayOfDownloadFiles.Add($fLicenseSig2)      # license.bin.sig2
$arrayOfDownloadFiles.Add($fConfig)           # config.bin
$arrayOfDownloadFiles.Add($fConfigSig2)       # config.bin.sig2
$arrayOfDownloadFiles.Add($fPsmd)             # p-smd.bin
$arrayOfDownloadFiles.Add($fPsmdSig2)         # p-smd.bin.sig2
$arrayOfDownloadFiles.Add($fAvSmd)            # av-smd.bin
$arrayOfDownloadFiles.Add($fAvSmdSig2)        # av-smd.bin.sig2
$arrayOfDownloadFiles.Add($fActions)          # action.bin
$arrayOfDownloadFiles.Add($fActionsSig2)      # action.bin.sig2
$arrayOfDownloadFiles.Add($fCsmd)             # c-smd.bin
$arrayOfDownloadFiles.Add($fCsmdSig2)         # c-smd.bin.sig2

# 各ユーザごとのファイルのタイムスタンプを取得し、結果を配列に格納
Get-Content $userlist | foreach {
    $UserName = $_.Trim()
    $response = Invoke-WebrequestToUpdateServer -UserName $UserName
    $stationName = [string[]]($response.Links | ? { $_.href -match "^/[\w-]+/[\w-]+/$" }).innerText
    foreach ($s in $stationName)
    {
        foreach ($f in $arrayOfDownloadFiles)
        {
            $response = Invoke-WebrequestToUpdateServer -UserName $UserName -StationName $s -Method "HEAD" -File $f
            $lastModifiedDate = [datetime]$response.Headers["Last-Modified"]
            $nowDate = (Get-Date)
            # 結果を配列に格納
            $objectOfEachRecord = [pscustomobject]@{
                User=$UserName
                Station=$s
                File=$f
                TimeStamp=$lastModifiedDate
                CheckDate=$nowDate
            }
            $script:arrayResult += $objectOfEachRecord

            # 標準出力
            $objectOfEachRecord

        }
    }
}

# csv に出力
$arrayResult | Export-Csv $outFilePath -Delimiter `t -NoTypeInformation -Append











