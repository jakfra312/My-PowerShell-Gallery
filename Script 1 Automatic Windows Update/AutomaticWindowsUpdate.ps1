if (Get-Module -ListAvailable -Name PSWindowsUpdate){
   Import-Module PSWindowsUpdate -Force
   if ((Get-WindowsUpdate).Title -eq "") {
        exit   
   }
   else {
    Install-WindowsUpdate -AcceptAll -AutoReboot
   }
} else {
    Install-Module -Name PSWindowsUpdate -Force
}