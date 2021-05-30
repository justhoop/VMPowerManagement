param ([string]$viserver, [string]$vihost)

Function Write-Log
{
    Param([string]$output)
    $now = Get-Date  
    $log = $now.ToString() + " $output"
    Write-Host $log
    Send-Email $log
    $log | Out-File -filepath ".\log.txt" -Append
}

function Get-PreReqs {
    if (-not(Get-Command "connect-viserver")) {
        Write-Log -output "powercli not installed"
        exit
    }
    elseif (-not(Test-Path "~/.secrets/EmailServiceAccount.xml")) {
        Write-Log -output "Email account credentials missing"
        exit
    }
    elseif (-not(Test-Path "~/.secrets/VMServiceAccount.xml")) {
        Write-Log -output "VM Service account missing"
        exit
    }
}
function Send-Email {
    param (
        [string]$message
    )
    $email = Import-Clixml "~/.secrets/EmailServiceAccount.xml"
    $Subject = "Power Source Change"
    $SMTPServer = "smtp.gmail.com"
    $SMTPPort = "587"
    Send-MailMessage -From $email.UserName -to $email.UserName -Subject $Subject `
    -Body $message -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
    -Credential $email
}

function Start-VMShutdown {
    [CmdletBinding()]
    param (
        [string]$viserver,
        [string]$vihost
    )
    $serviceaccount = Import-Clixml "~/.secrets/VMServiceAccount.xml"
    connect-viserver -server $vihost -user $serviceaccount.UserName -password $serviceaccount.Password
    $vms = get-vm | where-object {$_.powerstate -eq "PoweredOn"} | Where-Object {$_.Name -ne $viserver}
    foreach ($vm in $vms) {
        $name = $vm.name
        Write-Log "Shutting down $name"
        Shutdown-VMGuest $vm -Confirm:$false
    }
    while ((get-vm | where-object {$_.powerstate -eq "PoweredOn"} | Where-Object {$_.Name -ne $viserver}).count -gt 0) {
        Start-Sleep -Seconds 30
    }
    Write-Log "Shutting down"$viserver
    Shutdown-VMGuest $viserver -Confirm:$false
    Start-Sleep -Seconds 120
    Write-Log "Shutting down"$vihost
    Stop-VMHost -Server $vihost -Force
}

Get-PreReqs
while($True){
    $ups = Get-WmiObject win32_battery
    if ($ups.BatteryStatus -eq 1) {
        $pct = $ups.EstimatedChargeRemaining.toint()
        Write-Log "Running on battery. $pct percent remaining."
        if ($ups.EstimatedChargeRemaining -le 9) {
            Start-VMShutdown
        }
        Start-Sleep -Seconds 60
    }
    if ($ups.BatteryStatus -eq 2) {
        Write-Log "Running on AC power"
        break
    }
}