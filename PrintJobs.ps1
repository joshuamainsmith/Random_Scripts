param([switch]$Elevated)

function Test-Admin{
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false) {
    if ($elevated) {
        # not elevated
    } else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($MyInvocation.MyCommand.Definition))
    }
    exit
}

$Printer = Read-Host "Enter printer name: "
$FirstPrintJob = -1

try
{
    while ($true)
    {
        Write-Host "Checked $(Get-Date)"

        $s = Get-Service Spooler
        if ($s.Status -eq "Stopped")
        {
            'Spooler Stopped'
            #Log
            Start-Service -InputObject $s -PassThru
        }

        $PrinterJobs = Get-Printer -Name $Printer
        $PrintJobsById = Get-PrintJob -PrinterObject $PrinterJobs | select -ExpandProperty "Id"

        if ($PrintJobsById[0] -eq $FirstPrintJob)
        {
            Restart-Service -InputObject $s -PassThru
        }

        $FirstPrintJob = $PrintJobsById[0]

        Start-Sleep -Seconds 900
    }
}
finally
{
    'Done'
}
