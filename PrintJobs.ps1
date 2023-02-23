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
$ErrorLogs = "C:\PrintLogs\"
$LogFile = "C:\PrintLogs\"

#$PrinterName = Read-Host "Enter printer name: "
New-CimSession -ComputerName psc-app-01,yak-app-01,wen-app-01,fif-app-01,van-app-01,lac-app-01,lnc-app-01,cyn-app-01,oxr-app-01,fmn-app-01,mso-app-01,bil-app-01,anc-app-01,was-app-01
$Servers = "psc-app-01","yak-app-01","wen-app-01","fif-app-01","van-app-01","lac-app-01","lnc-app-01","cyn-app-01","oxr-app-01",
    "fmn-app-01","mso-app-01","bil-app-01","was-app-01","anc-app-01","anc-app-01","anc-app-01","anc-app-01"
$PrinterName = "psc-copy","yak-adm","wen-copy","fif-adm","van-adm","lac-adm","lnc-admn","cyn-admn","oxr-copy",
    "fmn-copy-01","MSO-ADM-KM454","BIL-ADM-KM454","WAS-KM458E","anc-adms","ANC-ANX","ANC-EDU","ANC-EP"
$FirstPrintJob = -1
$PrintJobsById = -1
$RemoveFirstJob = $false
$Count = 0

try
{
    while ($true)
    {
        if ($Count % $PrinterName.Length -eq 0)
        {
            $Count = 0
        }

        Write-Host "Checking $($Servers[$Count]) - $(Get-Date)"        

        $s = Get-Service Spooler -ComputerName $Servers[$Count]
        if ($s.Status -eq "Stopped")
        {
            'Spooler Stopped'
            $LogFile = "$($ErrorLogs)$($Servers[$Count])\PrintSpoolerLog_$($PrinterName[$Count]).txt"
            '========================================== Log Start ==========================================' >> $LogFile

            Get-Date >> $LogFile
            $Servers[$Count] >> $LogFile
            Start-Service -InputObject $s -PassThru | Format-List >> $LogFile
            '' >> $LogFile
        }

        $Printer = Get-Printer -CimSession $Servers[$Count] -Name $PrinterName[$Count]

        $PrintJobsById = Get-PrintJob -PrinterObject $Printer | select -ExpandProperty "Id"

        if ($PrintJobsById)
        {
            if ($PrintJobsById[0] -eq $FirstPrintJob)
            {
                'Print Jobs Stuck'
                $LogFile = "$($ErrorLogs)$($Servers[$Count])\PrintJobLog_$($PrinterName[$Count]).txt"
                '========================================== Log Start ==========================================' >> $LogFile
                Get-Date >> $LogFile
                $Printer >> $LogFile
                '' >> $LogFile
                '########## Print Job ID ##########' >> $LogFile
                $PrintJobsById >> $LogFile
                if ($RemoveFirstJob)
                {
                    Remove-PrintJob -PrinterObject $Printer -ID $PrintJobsById[0]
                }
                Restart-Service -InputObject $s -PassThru | Format-List >> $LogFile               
                '' >> $LogFile
                $RemoveFirstJob = $true
            }
            else
            {
                $RemoveFirstJob = $false
                $FirstPrintJob = $PrintJobsById[0]
            }            
        }        

        Start-Sleep -Seconds 53
        $Count++
    }
}
finally
{
    write-host "Script Stopped"
    $LogFile = "$($ErrorLogs)$($Servers[$Count])\PrintScriptStoppedLog_$($PrinterName[$Count]).txt"

    '========================================== Log Start ==========================================' >> $LogFile

    Get-Date >> $LogFile

    if ($Printer)
    {
        '########## Printer ##########' >> $LogFile
        $Printer >> $LogFile
    }

    if ($Error)
    {
        '########## Error ##########' >> $LogFile
        $Error >> $LogFile
    }

    '########## Server ##########' >> $LogFile
    $Servers[$Count] >> $LogFile

    '' >> $LogFile
}
