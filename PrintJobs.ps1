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

        $Printer = Get-Printer -Name "HP LaserJet Professional M1212nf MFP"
        $PrintJobsById = Get-PrintJob -PrinterObject $Printer | select -ExpandProperty "Id"
        write-host $PrintJobsById.Length

        Start-Sleep -Seconds 600
    }
}
finally
{
    'Done'
}