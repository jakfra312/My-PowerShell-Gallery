# this script requires BurntToast Module
# Author: Jakub Frąckowiak
$FreeSpace = Get-PSDrive c | % {[Math]::Truncate(($_.free/($_.used+$_.free))*100)}
if (Get-Module -ListAvailable -Name BurntToast) {
    if ($FreeSpace -lt 40) {
        New-BurntToastNotification -Text "WARNING !!!", 'Free disk space is less than 40%' 
    } else {
        New-BurntToastNotification -Text "Notification",'Free disk space is abouve 40%'    
    }
} else {
    Write-Host 'No BurnToast Module'  
}