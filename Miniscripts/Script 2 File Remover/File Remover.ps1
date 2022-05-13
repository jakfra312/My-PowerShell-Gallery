#Author of script: Jakub Frąckowiak
$P = 'C:\Users\User01\Downloads\'
$Today = Get-Date
$14Ago = $Today.AddDays(-14)
Get-ChildItem -Path $P -Recurse | Where-Object {($_.LastWriteTime -lt $14Ago)} | Remove-Item