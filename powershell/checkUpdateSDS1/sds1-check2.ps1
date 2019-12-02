################################################################################
# 
# SDS1 バージョンアップ チェックツール
# 
# [変更履歴]
# 2018/04/27 新規作成
# 2018/05/15 ダウンロード対象ファイルに av-smd.bin、action.bin、c-smd.bin、*.sig2 を追加
#            ISO 内の networkadapters.ini の中身を表示
# 2018/08/27 StationName を指定しない場合はすべての StationName をチェックするように修正
# 2019/01/23 c-smd.bin 内のパスもチェックするように修正
################################################################################


################################################################################
# パラメータ
################################################################################
param (
    [parameter(mandatory)][string]$UserName,
    [string[]]$StationName
)

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
$fIso = "_System.iso"
$7zexe = $PSScriptRoot + "\7z.exe"


################################################################################
# 関数読み込み
################################################################################
Join-Path $PSScriptRoot "general\*.ps1" | dir | foreach { . $_ }


################################################################################
# 各ファイルの DL を実施するかどうか
# av-smd.bin と ISO は DL に時間がかかるので "y" を入力しない限り DL しない
################################################################################
$doDownloadLicenseFile = $true  # license.bin (.sig2)
$doDownloadConfigFile = $true  # config.bin (.sig2), c-smd.bin (.sig2)
$doDownloadPsmdFile = $true  # p-smd.bin (.sig2)
$doDownloadOtherFile = Read-Host "Download All .BIN files? [y/n]"
    # av-smd.bin (.sig2), action.bin (.sig2), c-smd.bin (.sig2)
$doDownloadIso = Read-Host "Download ISO file? [y/n]"
    # config.bin, ISO


################################################################################
# ダウンロードファイルリスト
################################################################################
# List<string> の生成
$arrayOfDownloadFiles = New-Object System.Collections.Generic.List[string]

# ダウンロードするファイルを List に追加
if ($doDownloadLicenseFile)
{
    $arrayOfDownloadFiles.Add($fLicense)          # license.bin 
    $arrayOfDownloadFiles.Add($fLicenseSig2)      # license.bin.sig2
}

if ($doDownloadConfigFile)
{
    $arrayOfDownloadFiles.Add($fConfig)           # config.bin
    $arrayOfDownloadFiles.Add($fConfigSig2)       # config.bin.sig2
    $arrayOfDownloadFiles.Add($fCsmd)             # c-smd.bin
    $arrayOfDownloadFiles.Add($fCsmdSig2)         # c-smd.bin.sig2
}

if ($doDownloadPsmdFile)
{
    $arrayOfDownloadFiles.Add($fPsmd)             # p-smd.bin
    $arrayOfDownloadFiles.Add($fPsmdSig2)         # p-smd.bin.sig2
}

if ($doDownloadOtherFile -match "y|Y|[yY][eE][sS]")
{
    $arrayOfDownloadFiles.Add($fAvSmd)            # av-smd.bin
    $arrayOfDownloadFiles.Add($fAvSmdSig2)        # av-smd.bin.sig2
    $arrayOfDownloadFiles.Add($fActions)          # action.bin
    $arrayOfDownloadFiles.Add($fActionsSig2)      # action.bin.sig2
}
if ($doDownloadIso -match "y|Y|[yY][eE][sS]")
{
    $arrayOfDownloadFiles.Add($fIso)              # xxx_System.iso
}


################################################################################
# 事前準備
################################################################################
# 現在時刻 (yyyyMMddHHmmss) を作業ディレクトリにする
$workDir = Join-Path $PSScriptRoot (Get-Date -Format "yyyyMMddHHmmss")
New-Item $workDir -ItemType Directory | Out-Null
$mountDir = Join-Path "c:\votiro_iso" (Get-Date -Format "yyyyMMddHHmmss")
New-Item $mountDir -ItemType Directory | Out-Null


# 出力用メッセージ
$output = @"
################################### result #####################################
"@ | Out-File $workDir\result

# メッセージ出力用関数
# @param　出力メッセージ
function addMessage()
{
    param([string]$message)

    #Invoke-Expression "`"$message`" | Out-File $workDir\result -Append"
    $message | Out-File $workDir\result -Append
}


################################################################################
# ファイルダウンロード
################################################################################
# 出力用メッセージ
#addMessage "[downloaded files]"

# Station 名が指定されていなければ取得
if (! $StationName)
{
    $response = Invoke-WebrequestToUpdateServer -UserName $UserName
    $stationName = [string[]]($response.Links | ? { $_.href -match "^/[\w-]+/[\w-]+/$" }).innerText
}
else
{
    $stationName = $StationName
}

# Update サーバから各ファイルをダウンロードする
foreach ($s in $stationName)
{
    $outDir = Join-Path $workDir $s
    New-Item $outDir -ItemType Directory | Out-Null
    addMessage "[$s/file list]"
    addMessage ("{0,-20}{1}" -f "FileName", "LastModifiedDate")
    addMessage "---------------------------------------"
    foreach ($f in $arrayOfDownloadFiles)
    {
        if ($f -eq $fIso) { $f = $s + $fIso }
        $outFile = Join-Path $outDir $f
        $response = Invoke-WebrequestToUpdateServer -UserName $UserName -StationName $s -Method "HEAD" -File $f
        $lastModifiedDate = [datetime]$response.Headers["Last-Modified"]
        Write-Host ("Downloading {0} ..." -f $f)
        Invoke-WebrequestToUpdateServer -UserName $UserName -StationName $s -File $f -OutFile $outFile

        #addMessage $f`t`t$lastModifiedDate
        addMessage ("{0,-20}{1}" -f $f, $lastModifiedDate)
    }
    addMessage ""
}


################################################################################
# Check 1 : ExpirationData from license.bin
################################################################################

$CheckLicense = {
    Write-Host "### Check 1 : ExpirationData from license.bin ###"
    Write-Host "Check 1 started."
    foreach ($s in $stationName)
    {
        # license.bin から license.xml を取り出す
        $outDir = Join-Path $workDir $s
        Expand-Archive-7zip -File (Join-Path $outDir $fLicense) -OutDir $outDir
    
        # license.xml から CustomerName と ExpirationDate の値を取り出す
        $xml = [xml](Get-Content $outDir"\license.xml")
        $customerName = $xml.MobileTickXML.CustomerName
        $expirationDate = $xml.MobileTickXML.License.ExpirationDate

        # <ExpirationDate> が <ExpirationDate evaluation="Evaluation"> になっている場合
        if ($xml.MobileTickXML.License.ExpirationDate.GetType().Name -eq "XmlElement")
        {
            $expirationDate = $xml.MobileTickXML.License.ExpirationDate.'#text'
        }

        # 出力用メッセージ
        addMessage "[$s/license.bin]"
        addMessage "CustomerName: $customerName"
        addMessage "ExpirationDate: $expirationDate"
        addMessage ""
    }
    Write-Host "Check 1 finished."
}


################################################################################
# Check 2 : RemotePath from config.bin & c-smd.bin
################################################################################
$CheckConfig = {
    Write-Host "### Check 2 : RemotePath from config.bin & c-smd.bin ###"
    Write-Host "Check 2 started."
    foreach ($s in $stationName)
    {
        # config.bin から config.xml を取り出す
        $outDir = Join-Path $workDir $s
        Expand-Archive-7zip -FilePath (Join-Path $outDir $fConfig) -OutDir $outDir

        # config.xml から RemotePath の値を取り出す
        $xml = [xml](Get-Content $outDir\config.xml)
        $remotePathOfPsmd = $xml.MobileTickDLP.ProgFiles.RemotePath[0]
        $remotePathOfCsmd = $xml.MobileTickDLP.ProgFiles.RemotePath[1]

        # 出力用メッセージ
        addMessage "[$s/config.bin]"
        addMessage "RemotePath: "
        addMessage $remotePathOfPsmd
        addMessage $remotePathOfCsmd
        addMessage ""

        # c-smd.bin から config.xml を取り出す
        $outDir = Join-Path $workDir $s
        Expand-Archive-7zip -FilePath (Join-Path $outDir $fCsmd) -OutDir $outDir

        # config.xml から RemotePath の値を取り出す
        $xml = [xml](Get-Content $outDir\config.xml)
        $remotePathOfPsmd = $xml.MobileTickDLP.ProgFiles.RemotePath[0]
        $remotePathOfCsmd = $xml.MobileTickDLP.ProgFiles.RemotePath[1]

        # customer.xml から progLicense, progExternal, progInternal の値を取り出す
        $xml = [xml](Get-Content $outDir\customer.xml)
        $remotePathOfLicense = $xml.MobileTickDLP.progLicense
        $remotePathOfAvSmd = $xml.MobileTickDLP.progExternal
        $remotePathOfInternalFtpRoot = $xml.MobileTickDLP.progInternal

        # 出力用メッセージ
        addMessage "[$s/c-smd.bin]"
        addMessage "RemotePath: "
        addMessage $remotePathOfPsmd
        addMessage $remotePathOfCsmd
        addMessage "progLicense: "
        addMessage $remotePathOfLicense
        addMessage "Anti-Virus Signatures Location: "
        addMessage $remotePathOfAvSmd
        addMessage "Internal FTP Root: "
        addMessage $remotePathOfInternalFtpRoot
        addMessage ""

    }
    Write-Host "Check 2 finished."
}


################################################################################
# Check 3 : SDS version from p-smd.bin
################################################################################
$CheckPsmd = {
    Write-Host "### Check 3 : SDS version from p-smd.bin ###"
    Write-Host "Check 3 started."
    foreach ($s in $stationName)
    {
        # p-smd.bin の SHA256 ハッシュ値を算出する
        $outDir = Join-Path $workDir $s
        $sha256OfPsmd = (Get-FileHash -Algorithm SHA256 $outDir\$fPsmd).Hash

        # p-smd.bin の歴代 Ver と比較する
        switch($sha256OfPsmd)
        {
            "C3A105CA584B319109C9646FBC1DAA57ED7C224196B644C39620F3E0FBA1D8D3" { $SdsVer = "v7.5.0.145"; break }
            "50E4A85297A94770C65938B5A0425691D75713F87A973C78AA1279B0E6D641D4" { $SdsVer = "v7.3.4.1"; break }
            "81CA6D97F57FF7C5A36E77C26656DB28A0388E145EDB2E17F47D4FDC9EC3282C" { $SdsVer = "v7.3.3.1"; break }
            "F137F769B175C679864959C331C399758082D39B977C9B2775AFB05952E533C6" { $SdsVer = "v7.3.2.5"; break }
            "677C4315AEC30186AE6E383A1C65FB3DD36C9BD197AA7C9C6ED67EC6708831BA" { $SdsVer = "v7.2.1.24"; break }
            "D7C477A3BCE9FB154C3073E0157CCA0DC6C32FA565D6FFD0E227289C98BAC7D8" { $SdsVer = "v7.2.0.369"; break }
            "78007ADED143B7B75C384F4695F2223326759CF086017056A002A60B195C80F8" { $SdsVer = "v7.1.4.1"; break }
            "6B29381DB021895DC4A4EB23338A2BECC87DAC418ACA4C94E0F42D779D2B3374" { $SdsVer = "v7.1.2.27"; break }
            "F42BB7781D11C4DD46258B8E6E43729CB019749F00A31D5C07F922970C48FDD2" { $SdsVer = "v7.1.1.16"; break }
            "DC48BC0FBAAAEBA02D3CCEFE2F8D327C718BCF961481209C11196A181C4E2652" { $SdsVer = "v7.0.2.2"; break }
            "EDA449A37F9C17B819CA63B210BCB078A5BA10A1073290499369ABBC9C3436A9" { $SdsVer = "v7.0.0.97"; break }
            "E853B5B0D5DB00C5704115CC2491570103C08CDF9DFE1A625F7BD01FF63B113B" { $SdsVer = "v6.0.1.6"; break }
            "E4117F08790C27708E9239C9C57D00F4753EE4376DAFCDAA741017F351E1DADA" { $SdsVer = "v6.0.0.174"; break }
            default { $SdsVer = "不明なバージョンです" }
        }

        # 出力用メッセージ
        addMessage "[$s/p-smd.bin]"
        addMessage "SHA256: $sha256OfPsmd"
        addMessage "SDS Ver: $SdsVer"
        addMessage ""
    }
    Write-Host "Check 3 finished."
}


################################################################################
# Check 4 : config.bin & networkadapters.ini & timezone & codepage from ISO
################################################################################
$CheckIso = {
    Write-Host "### config.bin & networkadapters.ini & timezone & codepage from ISO ###"
    Write-Host "Check 4 started."
    foreach ($s in $stationName)
    {
        # ISO の SHA256 ハッシュ値を算出する
        $outDir = Join-Path $workDir $s
        $fIso2 = $s + $fIso
        $isoPath = Join-Path $outDir $fIso2

        $sha256OfIso = (Get-FileHash -Algorithm SHA256 $isoPath).Hash

        # Update サーバの config.bin と分けるため、$outDir に iso ディレクトリを別途作成
        $workDirForIso = Join-Path $outDir "\iso"
        New-Item $workDirForIso -ItemType Directory | Out-Null
        # ISO から config.bin を取り出す
        Write-Host ("Extracting {0} ...") -f $fIso2
        Expand-Archive-7zip -FilePath $isoPath -OutDir $workDirForIso

        # config.bin から config.xml を取り出す
        Expand-Archive-7zip -FilePath (Join-Path $workDirForIso $fConfig) -OutDir $workDirForIso

        # config.xml から RemotePath を取り出す
        $xml = [xml](Get-Content $workDirForIso\config.xml)
        $remotePathOfPsmd = $xml.MobileTickDLP.ProgFiles.RemotePath[0]
        $remotePathOfCsmd = $xml.MobileTickDLP.ProgFiles.RemotePath[1]

        # ISOVERxx.x.TXT から ISO Version と ISO Date を取り出す
        $isoVer = "ISO Version: 不明"
        if(Test-Path $workDirForIso\ISOVER*.TXT)
        {
            $isoVer = Get-Content $workDirForIso\ISOVER*.TXT | Select-String "^ISO Version"
        }

        $isoDate = "ISO Date: 不明"
        if(Test-Path $outDir\iso\ISOVER*.TXT)
        {
            $isoDate = Get-Content $workDirForIso\ISOVER*.TXT | Select-String "^ISO Date"
        }

        # networkadapters.ini からネットワーク情報を取り出す
        $networkInfo = Get-Content $workDirForIso\networkadapters.ini

        # ISO を c:\votiro_iso\<Station名> にコピー
        $mountDir_s = Join-Path $mountDir $s
        New-Item $mountDir_s -ItemType Directory | Out-Null
        Write-Host ("Copying {0} to {1}" -f $fIso2, $mountDir_s)
        Copy-Item $isoPath -Destination $mountDir_s -Force
        $newIsoPath = Join-Path $mountDir_s $fIso2

        # mount 場所の作成
        $workDirForMount = Join-Path $mountDir_s "\mount"
        New-Item $workDirForMount -ItemType Directory | Out-Null

        # ISO のマウント
        Write-Host ("Mounting {0} ..." -f $newIsoPath)
        $mountResult = Mount-DiskImage -ImagePath $newIsoPath -PassThru
        $wimPath = ($mountResult | Get-Volume).DriveLetter + ":\SOURCES\BOOT.WIM"

        # Windows イメージのマウント
        Write-Host "Mounting Windows image ..."
        $ProgressPreference = "SilentlyContinue"
        Mount-WindowsImage -ImagePath $wimPath -Index 1 -Path $workDirForMount -ReadOnly | Out-Null

        # タイムゾーン情報の取得
        $timeZone = dism /image:$workDirForMount /get-intl |
            Select-String -Pattern "^Default time zone"

        # 文字コード情報を取得
        $regRoot = "HKLM\VOTIROISO"
        $regFile = Join-Path $workDirForMount "Windows\system32\config\SYSTEM"
        $regCodePage = Join-Path $RegRoot "ControlSet001\Control\Nls\CodePage"
        Write-Host ("Loading registry {0} ..." -f $regRoot)
        reg load $RegRoot $RegFile | Out-Null
        $getReg = Get-Item -Path "Registry::${RegCodePage}"
        $codeACP = ""
        switch ($getReg.GetValue("ACP"))
        {
            "862" { $codeDesc = "OEM Hebrew; Hebrew (DOS)" ; break}
            "1255" { $codeDesc = "ANSI Hebrew; Hebrew (Windows)" ; break}
            "932" { $codeDesc = "ANSI/OEM Japanese; Japanese (Shift-JIS)" ; break}
            default { $codeDesc = "Unknown" }
        }
        $codeACP = "{0}({1})" -f $getReg.GetValue("ACP"), $codeDesc
        $codeOEMCP = ""
        switch ($getReg.GetValue("OEMCP"))
        {
            "862" { $codeDesc = "OEM Hebrew; Hebrew (DOS)" ; break}
            "1255" { $codeDesc = "ANSI Hebrew; Hebrew (Windows)" ; break}
            "932" { $codeDesc = "ANSI/OEM Japanese; Japanese (Shift-JIS)" ; break}
            default { $codeDesc = "Unknown" }
        }
        $codeOEMCP = "{0}({1})" -f $getReg.GetValue("OEMCP"), $codeDesc
        $getReg.close()
        [gc]::Collect()
        Write-Host ("Unloading registry {0} ..." -f $regRoot)
        reg unload $RegRoot | Out-Null

        # Windows イメージと ISO のアンマウント
        Write-Host "Unmounting Windows image ..."
        Dismount-WindowsImage -Path $workDirForMount -Discard
        Write-Host ("Unmounting {0} ..." -f $newIsoPath)
        DisMount-DiskImage $IsoPath | Out-Null

        # 出力用メッセージ
        addMessage "[[$s/ISO]]"
        addMessage "SHA256: $sha256OfIso"
        addMessage ""
        addMessage "[ISO Ver]"
        addMessage $isoVer
        addMessage $isoDate
        addMessage ""
        addMessage "[config.bin in ISO]"
        addMessage "RemotePath: "
        addMessage $remotePathOfPsmd
        addMessage $remotePathOfCsmd
        addMessage ""
        addMessage "[networkadapters.ini]"
        foreach ($l in $networkInfo)
        {
            addMessage $l
        }
        addMessage ""
        addMessage "[TimeZone]"
        addMessage $timeZone
        addMessage ""
        addMessage "[CodePage]"
        addMessage "ACP: ${codeACP}"
        addMessage "OEMCP: ${codeOEMCP}"
        addMessage ""
    }
    Write-Host "Check 4 finished."
}


################################################################################
# 実行
################################################################################
if ($doDownloadLicenseFile) { & $CheckLicense }
if ($doDownloadConfigFile) { & $CheckConfig }
if ($doDownloadPsmdFile) { & $CheckPsmd }
if ($doDownloadIso -match "y|Y|[yY][eE][sS]") { & $CheckIso }


################################################################################
# 出力
################################################################################
Get-Content $workDir\result


################################################################################
# 後始末
################################################################################
Remove-Item $workDir\*\* -Recurse -Exclude $arrayOfDownloadFiles
