foreach($comp in $computers){
    icm -cn $comp -ScriptBlock{
    Get-WmiObject -Class win32_process | select name, processid, parentprocessid, executablepath, commandline | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_processes'.csv -Append
}
}