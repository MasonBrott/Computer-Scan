foreach($comp in $computers){
    icm -cn $comp -ScriptBlock{
    #T1207 / T1570
    Get-NetTCPConnection | select localaddress, localport, remoteaddress, remoteport, state, owningprocess | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_netconnections'.csv -Append
}
}