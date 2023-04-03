$HostName = Read-Host "Enter the host name / IP address"
Clear-Host

$SysArray = (
    "All Info",
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
    "Logon Server",
    "Change Host",
    "Exit"
    )
    $Count = 0
    $SysArray.Length

while($true)
{
    Write-Host "What would you like to see with $($HostName)?"
    
    foreach($item in $SysArray)
    {
        Write-Host "$($Count). $($SysArray[$Count])"
        $Count++
    }

    $UserInput = Read-Host "Enter"
    
    Clear-Host

    if ($SysArray[$UserInput] -eq "Exit")
    {
        break;
    }
    elseif ($SysArray[$UserInput] -eq "Change Host")
    {
        $HostName = Read-Host "Enter the host name / IP address"
    }
    elseif ($SysArray[$UserInput] -eq "All Info")
    {
        systeminfo /s $HostName
    }
    else 
    {
        <# Action when all if and elseif conditions are false #>
        systeminfo /s $HostName | find $SysArray[$UserInput]
    }
    ''
    $Count = 0
}
