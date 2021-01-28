
function Test-Admin {
    $wid = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $prp = New-Object System.Security.Principal.WindowsPrincipal($wid)
    $adm = [System.Security.Principal.WindowsBuiltInRole]::Administrator
    $prp.IsInRole($adm)
}


if ([int]$psversiontable.psversion.major -lt 6) {
    Write-Host "PowerShell version need 6 or later" -BackgroundColor Red -ForegroundColor White
}
else {
    Write-Host "PowerShell version is Fit" -BackgroundColor Green -ForegroundColor White
}


if ((Test-Admin) -eq $false) {
    Write-Host 'You need Administrator privileges to run this.' -BackgroundColor Red -ForegroundColor White
    # Abort the script
    # this will work only if you are actually running a script
    # if you did not save your script, the ISE editor runs it as a series
    # of individual commands, so break will not break then.
    return
}

[Object[]]$inputdata=$null
$inputdata=get-WindowsCapability -Online | Where-Object -Property Name -Match "^OpenSSH.*"
[Microsoft.Dism.Commands.BasicCapabilityObject]$server=$null
$server=$inputdata| Where-Object -Property Name -Match ".*Server.*"
[Microsoft.Dism.Commands.BasicCapabilityObject]$client=$null
$client=$inputdata| Where-Object -Property Name -Match ".*Client.*"

if ($server.State -ne 'Installed') {
    # Install the OpenSSH Client
    Add-WindowsCapability -Online -Name $server.Name
}
if ($client.State -ne 'Installed') {
    # Install the OpenSSH Server
    Add-WindowsCapability -Online -Name $client.Name
}


Start-Service sshd
# OPTIONAL but recommended:
Set-Service -Name sshd -StartupType 'Automatic'
# Confirm the Firewall rule is configured. It should be created automatically by setup. 
Get-NetFirewallRule -Name *ssh*
# There should be a firewall rule named "OpenSSH-Server-In-TCP", which should be enabled
# If the firewall does not exist, create one
$inputdata=Get-NetFirewallRule |Where-Object -Property name -Match ^sshd$
if($null -eq $inputdata){
    New-NetFirewallRule -Name sshd -DisplayName 'OpenSSH Server (sshd)' -Enabled True -Direction Inbound -Protocol TCP -Action Allow -LocalPort 22
}

# Ensure LocalAccountTokenFilterPolicy is set to 1
# https://github.com/ansible/ansible/issues/42978
$token_path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
$token_prop_name = "LocalAccountTokenFilterPolicy"
$token_key = Get-Item -Path $token_path
$token_value = $token_key.GetValue($token_prop_name, $null)
if ($token_value -ne 1) {
    Write-Verbose "Setting LocalAccountTOkenFilterPolicy to 1"
    if ($null -ne $token_value) {
        Remove-ItemProperty -Path $token_path -Name $token_prop_name
    }
    New-ItemProperty -Path $token_path -Name $token_prop_name -Value 1 -PropertyType DWORD > $null
}


Write-Host "Finish!!"