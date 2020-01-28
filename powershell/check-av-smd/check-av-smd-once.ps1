################################################################################
# パラメータ
################################################################################
$fAvSmd = "av-smd.bin"
$fAvSmdSig2 = "av-smd.bin.sig2"
$userlist = Join-Path $PSScriptRoot "user.txt"
$outFileName = "result_AV_" + (Get-Date -Format "yyyyMMdd") + ".csv"
$outFilePath = Join-Path $PSScriptRoot $outFileName
$arrayResult = @()


################################################################################
# 関数読み込み
################################################################################
Join-Path $PSScriptRoot "general\*.ps1" | dir | foreach { . $_ }


################################################################################
# 設定読み込み
################################################################################
$conf = Get-Content (Join-Path $PSScriptRoot "settings.ini") | ? { $_ -match "=" } | ConvertFrom-StringData
$checkDate = $conf.checkDate
$reportDayOfWeek = [string[]]@($conf.reportDayOfWeek -split "," | % { $_.trim() })
$reportHour = [string[]]@($conf.reportHour -split "," | % { $_.trim() })
$mailToInNG = [string[]]@($conf.mailToInNG -split "," | % { $_.trim() })
$mailSubjectInNG = $conf.mailSubjectInNG -f $checkDate
$mailBodyInNG = $conf.mailBodyInNG
$mailToInOK = [string[]]@($conf.mailToInOK -split "," | % { $_.trim() })
$mailSubjectInOK = $conf.mailSubjectInOK -f (Get-Date -Format "yyyy/MM/dd HH:mm")
$mailBodyInOK = $conf.mailBodyInOK


################################################################################
# 処理
################################################################################
# チェックするファイルリストの生成
$arrayOfDownloadFiles = New-Object System.Collections.Generic.List[string]

$arrayOfDownloadFiles.Add($fAvSmd)            # av-smd.bin
$arrayOfDownloadFiles.Add($fAvSmdSig2)        # av-smd.bin.sig2

# 各ユーザごとのファイルのタイムスタンプを取得し、結果を配列に格納
Get-Content $userlist | foreach {
    $UserName = $_.Trim()
    foreach ($f in $arrayOfDownloadFiles)
    {
        $response = Invoke-WebrequestToUpdateServer -UserName $UserName -Method "HEAD" -File $f
        $lastModifiedDate = [datetime]$response.Headers["Last-Modified"]
        $nowDate = (Get-Date)
        # 現在の日付との差が $checkDate 以内なら OK
        if (($nowDate - $lastModifiedDate).totalDays -ge $checkDate)
        {
            $result = "NG"
        }
        else
        {
            $result = "OK"
        }

        # 結果を配列に格納
        $objectOfEachRecord = [pscustomobject]@{
            User=$UserName
            File=$f
            TimeStamp=$lastModifiedDate
            CheckDate=$nowDate
            Result=$result
        }
        $script:arrayResult += $objectOfEachRecord

        # 標準出力
        $objectOfEachRecord
    }
}

# csv に出力
$arrayResult | Export-Csv $outFilePath -Delimiter `t -NoTypeInformation -Append

# Result が OK 以外ならメールを送る
# 仮に Update サーバが落ちていれば Result は Null になるので、その場合もメールが飛ぶ
if ($arrayResult | ? { $_.Result -ne "OK"})
{
    Send-MailMessage-Net -To $mailToInNG -Subject $mailSubjectInNG -Body $mailBodyInNG
}
# Result が OK でも現在日時が $reportDayOfWeek と $reportHour に一致すればメールを送る
elseif (($reportDayOfWeek -contains [int](Get-Date).DayOfWeek) -And ($reportHour -contains [int](Get-Date).Hour))
{
    Send-MailMessage-Net -To $mailToInOK -Subject $mailSubjectInOK -Body $mailBodyInOK
}


