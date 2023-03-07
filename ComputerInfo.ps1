$HostName = Read-Host "Enter the host name / IP address"

$SysArray = (
    "Host Name",
    "OS Version",
    "OS Manufacturer",
    "Original Install Date" ,
    "System Manufacturer",
    "System Model",
    "System Boot Time",
    "System Type",
    "BIOS Version",
    "Total Physical Memory",
    "Domain",
    "Logon Server"
    )
    $Count = 0
    $SysArray.Length

while($true)
{
    Write-Host "What would you like to see?"
    
    foreach($item in $SysArray)
    {
        Write-Host "$($Count). $($SysArray[$Count])"
        $Count++
    }

    $Input = Read-Host "Enter"

    clear

    systeminfo /s $HostName | find $SysArray[$Input]
    ''
    $Count = 0
}

