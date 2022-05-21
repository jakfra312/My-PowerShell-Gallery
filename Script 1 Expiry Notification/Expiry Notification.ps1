#Author: Jakub Frąckowiak
#This script requires Active Directory Module
if(Get-Module -ListAvailable -Name ActiveDirectory){
$ADUsers = Get-ADUser -Filter * -Properties * | Where-Object {($_.Enabled -eq $true) -and ($_.PasswordNeverExpires -eq $false) -and -not ($_.accountExpires -eq 0)}
$MaxDays = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
$Serv = (Get-ADDomainController).Forest
$Info = $ADUsers | select Name, mail, SamAccountName, PasswordLastSet 
Foreach($User in $Info){
if(-not $User.mail){
   Write-Host "User: $($User.Name) doesn't have an assigned E-mail address"
} else {
    $LastSet = $User.PasswordLastSet
    $PP = Get-ADUserResultantPasswordPolicy -Identity $User.SamAccountName -Server $Serv 
    if($PP){
        $MaxDays = $PP.maxPasswordAge}
    else{    
        $ExpDate = $LastSet + $MaxDays
        $Today = [datetime]::Today
        $ExpDays = (New-TimeSpan -Start $Today -End $ExpDate).Days
        if($ExpDays -lt 20){
            $Msg = "Your password will expire in $ExpDays days"
            $SMTPServer = 'smtp.gmail.com'
            $SMTPPort = '587'
            $From = 'testpass1@gmail.com'
            $To = $User.mail
            $Pass = Get-Content C:\Test\passexample.txt
            $Pass = ConvertTo-SecureString -String $Pass -Force -AsPlainText 
            $creden = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $From, $Pass
            Send-MailMessage -From $From -To $To -Subject 'Password Notification' -Body $Msg -SmtpServer $SMTPServer -Port $SMTPPort -Credential $creden -UseSsl
        }
    
}}}}
else{
Write-Host 'No ActiveDirectory Module'
}