foreach ($comp in $computers){
    icm -cn $comp -ScriptBlock{
    #T1098 Monitor for changes to account object/permissions
    Get-EventLog -LogName Security | ? {$_.EventID -eq 4738 -or $_.EventID -eq 4728 -or $_.EventID -eq 4670} | select timegenerated, instanceid, message | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_events'.csv -Append
    #T1550.002/003 request of new ticket granting ticket or service ticket
    Get-EventLog -LogName Security | ? {$_.EventID -eq 4768 -or $_.EventID -eq 4769} | select timegenerated, instanceid, message  | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_events'.csv -Append
    #T1059.001 Abusing powershell commands and scripts / WMIC
    Get-EventLog -LogName Security | ? {$_.EventID -eq 400 -or $_.EventID -eq 403 -or $_.EventID -eq 4104 -or $_.EventID -eq 4688} | select timegenerated, instanceid, message | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_events'.csv -Append
    #T1199 New logons 
    Get-EventLog -LogName Security | ? {$_.EventID -eq 4624} | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_events'.csv -Append
    #T1070.004 File deletion
    Get-EventLog -LogName Security | ? {$_.EventID -eq 4460 -or $_.EventID -eq 4463} | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_events'.csv -Append
    #T1053.005 New scheduled jobs
    Get-EventLog -LogName Security | ? {$_.EventID -eq 4698} | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_events'.csv -Append
    #T1569.002 Changes to registry keys or values
    Get-EventLog -LogName Security | ? {$_.EventID -eq 4697 -or $_.EventID -eq 4657} | Export-Csv -Path $env:USERPROFILE\desktop\output\$comp'_events'.csv -Append
    }
}