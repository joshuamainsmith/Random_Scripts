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
New-CimSession -ComputerName psc-app-01,yak-app-01,wen-app-01,fif-app-01,van-app-01,lac-app-01,lnc-app-01,cyn-app-01,oxr-app-01,
    fmn-app-01,mso-app-01,bil-app-01,anc-app-01,was-app-01,rno-app-01

$Servers = ("psc-app-01", "psc-copy"),
    ("yak-app-01", "yak-adm"),
    ("wen-app-01", "wen-copy"),
    ("fif-app-01","fif-adm"),
    ("van-app-01","van-adm"),
    ("lac-app-01","lac-adm"),
    ("lnc-app-01","lnc-admn"),
    ("cyn-app-01","cyn-admn"),
    ("oxr-app-01","oxr-copy"),
    ("fmn-app-01","fmn-copy-01"),
    ("mso-app-01","MSO-ADM-KM454"),
    ("bil-app-01","BIL-ADM-KM454"),
    ("was-app-01","WAS-KM458E"),
    ("anc-app-01","anc-adms","ANC-ANX","ANC-EDU","ANC-EP")#,
    #("rno-app-01","RNO-REG","RNO-FA","RNO-ACCT")

$FirstPrintJob = New-Object 'object[]' 20
$PrintJobsById = -1
$RemoveFirstJob = $false
$Count = 0
$CurrentServer = ""

try
{

    while ($true)
    {        

        foreach ($Server in $Servers)
        {
            write-host "Count $($Count)"
            Write-Host "Checking $($Server[0]) - $(Get-Date)"
            $CurrentServer = $Server[0]

            $p = Test-Connection -ComputerName $CurrentServer -Count 1

            if ($p)
            {
                $s = Get-Service Spooler -ComputerName $Server[0]
                if ($s.Status -eq "Stopped")
                {
                    ''
                    'Spooler Stopped'
                    ''
                    $LogFile = "$($ErrorLogs)$($Server[0])\PrintSpoolerLog.txt"
                    '========================================== Log Start ==========================================' >> $LogFile

                    Get-Date >> $LogFile
                    "Server: $($Server[0])" >> $LogFile
                    Start-Service -InputObject $s -PassThru | Format-List >> $LogFile
                    '' >> $LogFile
                    "Spooler: $(Get-Date) $($CurrentServer)" >> "$($ErrorLogs)update.txt"
                }            

                foreach ($Printer in $Server)
                {
                    if ($Printer -notlike "*app-01")
                    {
                        $PrinterObj = Get-Printer -CimSession $Server[0] -Name $Printer
                        $PrintJobs = Get-PrintJob -PrinterObject $PrinterObj
                        $PrintJobsById = $PrintJobs | select -ExpandProperty "Id"

                        write-host "Checking $($Printer)"

                        if ($PrintJobsById)
                        {
                            Write-Host "PrintJobsById: $($PrintJobsById[0]) Count: $($FirstPrintJob[$Count])"
                            if ($PrintJobsById[0] -eq $FirstPrintJob[$Count])
                            {
                                ''
                                'Print Jobs Stuck'
                                ''

                                $LogFile = "$($ErrorLogs)$($Server[0])\PrintJobLog_$($Printer).txt"
                                '========================================== Log Start ==========================================' >> $LogFile

                                Get-Date >> $LogFile
                                $PrinterObj | Format-List >> $LogFile
                                $PrintJobs | Format-List >> $LogFile

                                if ($RemoveFirstJob)
                                {
                                    # Remove first print job - commented for now in case this method is too risky
                                    # Remove-PrintJob -PrinterObject $PrinterObj -ID $PrintJobsById[0]
                                }

                                Restart-Service -InputObject $s -PassThru | Format-List >> $LogFile               
                                '' >> $LogFile
                                "Print Jobs: $(Get-Date) $($CurrentServer)" >> "$($ErrorLogs)update.txt"
                                $RemoveFirstJob = $true
                            }
                            else
                            {
                                $RemoveFirstJob = $false
                                $FirstPrintJob[$Count] = $PrintJobsById[0]
                            }            
                        }
                        Start-Sleep -Seconds 45
                        $Count++ 
                    }            
                }                
            }
            else
            {
                "Connection Error: $(Get-Date) $($CurrentServer)" >> "$($ErrorLogs)update.txt"
            }                    
        }
        $Count = 0
    }
}
finally
{
    write-host "Script Stopped"
    $LogFile = "$($ErrorLogs)$($CurrentServer)\PrintScriptStoppedLog.txt"

    '========================================== Log Start ==========================================' >> $LogFile

    Get-Date >> $LogFile

    if ($Printer)
    {
        '########## Printer ##########' >> $LogFile
        $Printer | Format-List >> $LogFile
    }

    if ($Error)
    {
        '########## Error ##########' >> $LogFile
        $Error >> $LogFile
    }

    '########## Server ##########' >> $LogFile
    $CurrentServer >> $LogFile

    '' >> $LogFile

    "Stopped: $(Get-Date) $($CurrentServer)" >> "$($ErrorLogs)update.txt"
}
