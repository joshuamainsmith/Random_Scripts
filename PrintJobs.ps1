param([switch]$Elevated)

function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) {
        # tried to elevate, did not work, aborting
    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    exit
}

#Log File Locations
$ErrorLogs = "C:\PrintLogs\PrintScriptStoppedLog.txt"
$SpoolerLogs = "C:\PrintLogs\PrintSpoolerLog.txt"
$JobLogs = "C:\PrintLogs\PrintJobLog.txt"

#$PrinterName = Read-Host "Enter printer name: "
New-CimSession -ComputerName psc-app-01,yak-app-01,wen-app-01
$Servers = "psc-app-01","yak-app-01","wen-app-01"
$PrinterName = "psc-copy","yak-adm","wen-copy"
$FirstPrintJob = -1
$PrintJobsById = -1
$RemoveFirstJob = $false
$Count = 0

try
{
    while ($true)
    {
        if ($Count % 3 -eq 0)
        {
            $Count = 0
        }

        Write-Host "Checked $($Servers[$Count]) at $(Get-Date)"        

        $s = Get-Service Spooler -ComputerName $Servers[$Count]
        if ($s.Status -eq "Stopped")
        {
            'Spooler Stopped'
            '========================================== Log Start ==========================================' >> $SpoolerLogs

            Get-Date >> $SpoolerLogs
            $Servers[$Count] >> $SpoolerLogs
            Start-Service -InputObject $s -PassThru | Format-List >> $SpoolerLogs
            '' >> $SpoolerLogs
        }

        $Printer = Get-Printer -CimSession $Servers[$Count] -Name $PrinterName[$Count]

        $PrintJobsById = Get-PrintJob -PrinterObject $Printer | select -ExpandProperty "Id"

        if ($PrintJobsById)
        {
            if ($PrintJobsById[0] -eq $FirstPrintJob)
            {
                'Print Jobs Stuck'
                '========================================== Log Start ==========================================' >> $JobLogs
                Get-Date >> $JobLogs
                $Printer >> $JobLogs
                '' >> $JobLogs
                '########## Print Job ID ##########' >> $JobLogs
                $PrintJobsById >> $JobLogs
                if ($RemoveFirstJob)
                {
                    Remove-PrintJob -PrinterObject $Printer -ID $PrintJobsById[0]
                }
                Restart-Service -InputObject $s -PassThru | Format-List >> $JobLogs               
                '' >> $JobLogs
                $RemoveFirstJob = $true
            }
            else
            {
                $RemoveFirstJob = $false
                $FirstPrintJob = $PrintJobsById[0]
            }            
        }

        $Count++

        Start-Sleep -Seconds 300
    }
}
finally
{
    write-host "Script Stopped"

    '========================================== Log Start ==========================================' >> $ErrorLogs

    Get-Date >> $ErrorLogs

    '########## Printer ##########' >> $ErrorLogs
    $Printer >> $ErrorLogs

    '########## Error ##########' >> $ErrorLogs
    $Error >> $ErrorLogs

    '########## Server ##########' >> $ErrorLogs
    $Servers[$Count] >> $ErrorLogs

    '' >> $ErrorLogs
}
