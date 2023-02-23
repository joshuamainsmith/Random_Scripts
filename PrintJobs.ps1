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

$PrinterName = Read-Host "Enter printer name: "
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
            Get-Date >> C:\PrintSpoolerLog.txt
            Start-Service -InputObject $s -PassThru | Format-List >> C:\PrintSpoolerLog.txt
        }

        $Printer = Get-Printer -Name $PrinterName
        $PrintJobsById = Get-PrintJob -PrinterObject $Printer | select -ExpandProperty "Id"

        if ($PrintJobsById[0] -eq $FirstPrintJob)
        {
            'Print Jobs Stuck'
            Get-Date >> C:\PrintJobLog.txt
            Restart-Service -InputObject $s -PassThru | Format-List >> C:\PrintJobLog.txt
        }

        $FirstPrintJob = $PrintJobsById[0]

        Start-Sleep -Seconds 900
    }
}
finally
{
    write-host "Script Stopped"
    Get-Date >> C:\PrintScriptStoppedLog.txt
    $Printer >> C:\PrintScriptStoppedLog.txt
}
