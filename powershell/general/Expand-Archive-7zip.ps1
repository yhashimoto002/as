function Expand-Archive-7zip
{
    param(
        [parameter(mandatory)][string]$FilePath,
        [string]$OutDir = $PSScriptRoot
    )

    $7zexe = $PSScriptRoot + "\7z.exe"

    # パスワードの指定
    $fileName = Split-Path $FilePath -Leaf
    $password = ""
    if ($fileName -eq "license.bin")
    {
        $password = "nuchhkyhe"
    }
    elseif ($fileName -eq "config.bin")
    {
        $password = "vdsru,ngrf,@2@"
    }
    elseif (($fileName -eq "actions.bin") -or ($fileName -eq "c-smd.bin"))
    {
        $password = "aslktybn"
    }

    # 7z.exe に渡すパラメータ
    if ($password)
    {
        $arg = "x -y -p$password -o$OutDir $FilePath"
    }
    else
    {
        $arg = "x -y -o$OutDir $FilePath"
    }

    # 展開
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $7zexe
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $false
    $psi.CreateNoWindow = $true
    $psi.UseShellExecute = $false
    $psi.Arguments = $arg

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    $process.Start() | Out-Null
    $process.WaitForExit()
}



