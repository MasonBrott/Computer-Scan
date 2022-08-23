foreach($comp in $computers){
    icm -cn $comp -ScriptBlock{
    #T1547.005
    Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\lsass.exe' | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_registry'.csv -Append
    Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\Security packages' | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_registry'.csv -Append
    Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa\OSConfig\security packages' | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_registry'.csv -Append
    #T1003.004
    Get-ItemProperty HKLM:\SECURITY\policy\secrets | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_registry'.csv -Append
}
}