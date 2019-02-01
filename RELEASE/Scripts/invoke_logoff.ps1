#Pass parameters from scripot
param([string]$Username,[string]$Password, [string]$Domain, [string]$ComputerName)
$UsernameWithDomain = $Domain + "\" + $Username
$txtPassword = ConvertTo-SecureString -AsPlainText $Password -Force; 
$Creds = New-Object System.Management.Automation.PSCredential -ArgumentList $UsernameWithDomain, $txtPassword

#Log user out
Start-Process Powershell.exe -ArgumentList "-Command `"& {`$SessionID = ((quser /SERVER:$ComputerName | Where-Object -FilterScript  {`$_ -match '$Username'}) -split ' +')[2]; logoff `$SessionID /SERVER:$ComputerName}`"" -Credential $Creds -WindowStyle Hidden


