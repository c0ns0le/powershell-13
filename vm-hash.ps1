function generatePwd($string){
    $p = Start-Process -FilePath .\runCrypto.bat -ArgumentList $string -Wait -passthru -NoNewWindow;
    $hash = $(cat CryptPwd).Split('\n') | where { $_.Substring(0,3) -eq "pwd" } | % { $_.Split('=')[1] }
    return $hash
}