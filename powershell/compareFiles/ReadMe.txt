
■ 説明

2 つのフォルダにある同じ名前のファイルの見た目を比較し、結果を csv で出力します。
csv には以下のカラムが出力されます。

- No.: ファイルのナンバリング
- FileName: ファイル名
- Image: イメージ名
- Page: ページ
- Identify: 差分 (0 に近ければ差がなく、値が大きいほど差が大きい)
- Result: NG か OK か

またスクリプトが置いてあるフォルダに output フォルダが作成され、
以下のように比較過程で生成した画像が配置されます。

output\
│
├ sample1.pdf
│ ├ before
│ │ ├ image-0.png
│ │ ├ image-1.png
│ │ └ ...
│ ├ after
│ │ ├ image-0.png
│ │ ├ image-1.png
│ │ └ ...
│ └ diff
│    ├ image-0.png
│    ├ image-1.png
│    └ ...
├ sample1.pdf
│ ├ before
│ │ ├ image-0.png
│ │ ├ image-1.png
│ │ └ ...
│ ├ after
│ │ ├ image-0.png
│ │ ├ image-1.png
│ │ └ ...
│ └ diff
│    ├ image-0.png
│    ├ image-1.png
│    └ ...


■ 使い方

1. compareFiles フォルダごと適当な場所にコピーします。

2. ImageMagick-7.0.8-14-Q16-x64-static.exe を実行して、ImageMagick をインストールします。

3. gs925w64.exe を実行して、ghostscript をインストールします。

4. Windows + R キーで「ファイル名を指定して実行」を表示して「powershell」と入力します。

5. スクリプトが置いてあるパスに移動します。以下は c:\work に置いた場合です。

PS> cd c:\work

6. 比較したいファイルに合わせて、以下のようにファイルが保存されているフォルダをそれぞれ引数で指定して実行します。

[PDFファイルの場合]
PS> .\Compare-Pdf.ps1 .\before .\after

[Officeファイルの場合]
PS> .\Compare-OfficeFile.ps1 .\before .\after

[画像ファイルの場合]
PS> .\Compare-Image.ps1 .\before .\after

※対象ファイルだけを選んで比較するので、対象ファイル以外のファイルが混ざっていても特に問題はありません。

7. 処理が終わるとスクリプトが置いてあるフォルダに result_NG_日付.html が出力されます。

8. result_NG_日付.html をブラウザで開き、目視でチェックを行ってください。


■ 注意

・Office ファイルをチェックするときは、Office がインストールされている環境で実施してください。
　Office がインストールされていないと失敗します。

・PowerShell 4.0 未満ではテストしていません。うまく動作しなければ PowerShell 4.0 以上の
環境で実行してみてください。

PowerShell を開いて以下のコマンドを実行して「2」が表示される場合、PowerShell 4.0 以上を
インストールするか、Windows Server 2012 R2 などの環境で実行してください。

PS> $PSVersionTable.PSVersion.Major

PowerShell のアップグレード方法は以下の URL が参考になります。

Windows PowerShell のインストール - 既存の Windows PowerShell をアップグレードする
https://docs.microsoft.com/ja-jp/powershell/scripting/setup/installing-windows-powershell?view=powershell-6#upgrading-existing-windows-powershell


・スクリプトの実行ポリシーが Restricted の場合は、スクリプトが実行できません。
その場合は、以下のコマンドで RemoteSigned に変更してから実行してください。

PS> Get-ExecutionPolicy
Restricted

PS> Set-ExecutionPolicy RemoteSigned

PS> Get-ExecutionPolicy
RemoteSigned

・それでも実行できない場合は以下のコマンドを実行してください。

PS> Unblock-File スクリプトファイル

・呼び出している magick.exe がなんらかの理由で停止すると、「xxx は動作を停止しました」という
エラー画面がポップアップし、「プログラムを終了します」をクリックしないと処理が進まなくなります。

一つの回避策として、以下のレジストリキーを登録することで、プログラムが停止してもエラー画面が
ポップアップせずに自動終了するようになり、次の処理に自動で進めるようになります。

キー：
HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\Windows Error Reporting

名前：Disabled
種類：REG_DWORD
データ：1

名前：DontShowUI
種類：REG_DWORD
データ：1

参考：
WER Settings
https://docs.microsoft.com/en-us/windows/desktop/wer/wer-settings

・かなりメモリを食います。少なくとも 2GB 程度は空きがある状態で実行した方がよいでしょう。





