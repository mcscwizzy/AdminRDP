#Pass parameters from script
param([string]$Username,[string]$Password, [string]$Domain, [string]$ComputerName)
$UsernameWithDomain = $Domain + "\" + $Username
$txtPassword = ConvertTo-SecureString -AsPlainText $Password -Force; 
$Creds = New-Object System.Management.Automation.PSCredential -ArgumentList $UsernameWithDomain, $txtPassword

#enter psession
Enter-PSSession -ComputerName $ComputerName -Credential $Creds


