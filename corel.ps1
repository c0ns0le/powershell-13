#brute force attempt to verify software installation
100..600 | % {
        $computer =  "AT1P" + $_.ToString().PadLeft(3, '0')
        if ((Test-Connection -ComputerName $computer -Quiet)) {
            $path = "\\" + $computer + "\C$\Program Files (x86)\Corel\CorelDRAW Graphics Suite 13\Programs\CorelDRW.exe"
            if (test-path $path) {
                write-host $computer
            }
        }
}