
■説明

-- check-av-smd-once.ps1

ユーザの av-smd.bin と av-smd.bin.sig2 のタイムスタンプをチェックします。

スクリプトと同じフォルダに result_AV_yyyyMMdd.csv という名前でチェック結果が
出力されていきます。

"User"	"File"	"TimeStamp"	"CheckDate"	"Result"	"Error"
"adachicit"	"av-smd.bin"			"NG"	"curl: (22) The requested URL returned error: 404 Not Found"
"adachicit"	"av-smd.bin.sig2"			"NG"	"curl: (22) The requested URL returned error: 404 Not Found"
"AioiCity"	"av-smd.bin"	"2020/01/28 7:18:54"	"2020/01/28 16:38:46"	"OK"	""
"AioiCity"	"av-smd.bin.sig2"	"2020/01/28 7:19:21"	"2020/01/28 16:38:47"	"OK"	""
(...)

指定日数より過去のタイムスタンプが検出されたら maillist_NG.txt に記載された
メールアドレス宛てにメールが通知されます。

さらに、結果が OK の場合でも毎週月曜日の 8:00 に通知します。
定期通知する曜日と時間は複数指定可能です


■使い方

-- check-av-smd-once.ps1

1. 任意のフォルダに以下のようにファイルとフォルダを配置します

	\- check-av-smd-once.ps1 ファイル
	\- settings.ini ファイル
	\- user.txt ファイル
	\- general フォルダ
		\- Invoke-WebrequestToUpdateServer.ps1
		\- Send-MailMessage-Net.ps1

2. settings.ini をテキストエディタで開き、必要に応じて設定を変更します
　 メールの通知先 (mailToInNG、mailToInOK) には注意してください

3. 1回だけ実行する場合は PowerShell を立ち上げてスクリプトを実行します。

	PS> .\check-av-smd-once.ps1

4. 定期実行する場合は Windows のタスクスケジューラにスクリプトを登録します

	新規でタスクを登録するには タスクスケジューラ > タスクの作成 を開き、
	必要な設定を登録します。

	1時間置きに実行し続ける場合の トリガー と 操作 の設定例のスクリーンショットを
	taskscheduler_sample01.png と taskscheduler_sample02.png に撮っていますので、
	参考にしてください。
	
	もしくは タスクのインポート より AV signature update check.xml を指定して
	インポートします。
	この場合は、C:\work\check-av-smd フォルダを作成して、そこにスクリプト等を
	配置してください。


■ 注意

・PowerShell 4.0 未満だと正しく動作しません。Windows 7 だとバージョンアップしていなければ 2.0 のままです。

PowerShell を開いて以下のコマンドを実行して「2」が表示される場合、PowerShell 4.0 以上を
インストールするか、Windows Server 2012 R2 などの環境で実行してください。

PS> $PSVersionTable.PSVersion.Major

PowerShell のアップグレード方法は以下の URL が参考になります。

Windows PowerShell のインストール - 既存の Windows PowerShell をアップグレードする
https://docs.microsoft.com/ja-jp/powershell/scripting/setup/installing-windows-powershell?view=powershell-6#upgrading-existing-windows-powershell


・スクリプトの実行ポリシーが Restricted の場合は、スクリプトが実行できません。
その場合は、以下のコマンドで RemoteSigned に変更してから実行してください。

> Get-ExecutionPolicy
Restricted

> Set-ExecutionPolicy RemoteSigned

> Get-ExecutionPolicy
RemoteSigned

・実行許可が毎回求められる場合は以下のコマンドを実行してください。

> Unblock-File .\*.ps1


■履歴
2018/6/8 橋本 新規作成
2018/7/6 橋本 リニューアル
2020/1/21 橋本 check-all-once.ps1 は sds1-check.ps1 で代用できることに気付いたので削除

