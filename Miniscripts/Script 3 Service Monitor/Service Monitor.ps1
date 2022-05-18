$Services = @("ALG","WerSvc") #<--- List of services to be monitored | This is an example list only
foreach ($Service in $Services) {
    $SV = (Get-Service -Name $Service).Status
    if ($SV -eq 'Stopped') {
        try {
         Start-Service -Name $Service
        }
        catch {
            Write-Host $_.Exception.Message
        }
    }
}