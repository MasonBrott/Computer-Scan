cd $env:USERPROFILE\desktop\scripts\
$computers = gc $env:USERPROFILE\desktop\computers.txt
.\get_registry.ps1
.\processes.ps1
.\Get_events.ps1
.\get_prefetch.ps1
.\get_netconnections.ps1
.\get_files.ps1

foreach ($comp in $computers)
{
# Create Excel object and load first csv
$xl = new-object -ComObject excel.application
$xl.visible = $true
$workbook = $xl.workbooks.open("$env:USERPROFILE\desktop\output\$comp'_registry'.csv")

#Import second sheet and format properties to load on second sheet
$processes = import-csv $env:USERPROFILE\desktop\output\$comp'_processes'.csv
$i = 2
$last = $xl.Worksheets | select -Last 1
$ws = $xl.Worksheets.add($last)
$ws.name = "processes"
$last.move($ws)
foreach ($proc in $processes)
{
    $ws.cells.item($i,1) = $proc.name
    $ws.cells.item($i,2) = $proc.processid
    $ws.cells.item($i,3) = $proc.parentprocessid
    $ws.cells.item($i,4) = $proc.executablepath
    $ws.cells.item($i,5) = $proc.commandline
    $i++
}
#Import third sheet
$events = import-csv $env:USERPROFILE\desktop\output\$comp'_events'.csv
$i = 2
$last = $xl.Worksheets | select -Last 1
$ws = $xl.Worksheets.add($last)
$ws.name = "events"
$last.move($ws)
foreach ($event in $events)
{
    $ws.cells.item($i,1) = $event.timegenerated
    $ws.cells.item($i,2) = $event.instanceid
    $ws.cells.item($i,3) = $event.message
    $i++
}
#import fourth sheet
$prefetch = import-csv $env:USERPROFILE\desktop\output\$comp'_prefetch'.csv
$i = 2
$last = $xl.Worksheets | select -Last 1
$ws = $xl.Worksheets.add($last)
$ws.name = "prefetch"
$last.move($ws)
foreach ($entry in $prefetch)
{
    $ws.cells.item($i,1) = $entry.name
    $ws.cells.item($i,2) = $entry.creationtime
    $ws.cells.item($i,3) = $entry.lastwritetime
    $i++
}
#import fifth sheet
$connections = import-csv $env:USERPROFILE\desktop\output\$comp'_netconnections'.csv
$i = 2
$last = $xl.Worksheets | select -Last 1
$ws = $xl.Worksheets.add($last)
$ws.name = "net-connections"
$last.move($ws)
foreach ($connection in $connections)
{
    $ws.cells.item($i,1) = $connection.localaddress
    $ws.cells.item($i,2) = $connection.localport
    $ws.cells.item($i,3) = $connection.remoteaddress
    $ws.cells.item($i,4) = $connection.remoteport
    $ws.cells.item($i,5) = $connection.state
    $ws.cells.item($i,6) = $connection.owningprocess
    $i++
}
#import last sheet
$files = import-csv $env:USERPROFILE\desktop\output\$comp'_files'.csv
$i = 2
$last = $xl.Worksheets | select -Last 1
$ws = $xl.Worksheets.add($last)
$ws.name = "files"
$last.move($ws)
foreach ($file in $files)
{
    $ws.cells.item($i,1) = $file.pspath
    $ws.cells.item($i,2) = $file.readcount
    $i++
}
#saving excel doc and moving on
$workbook.saveas("$env:USERPROFILE\desktop\reports\$comp'_report'.csv")
}