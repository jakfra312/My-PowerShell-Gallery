#Author : Jakub Frąckowiak
#This script requires Active Directory Module
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
Add-Type -Assembly System.Windows.forms
Add-Type -AssemblyName System.Web
$XAMLDesign = @'
<Window x:Class="WpfApp9.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp9"
        mc:Ignorable="d"
        Title="AD Account Password Reset" Height="450" Width="800">
    <Grid>
        <ListView x:Name="ListViewOne" Margin="34,64,432,47">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Parameters" DisplayMemberBinding = "{Binding Parameters}" Width="330"/>
                </GridView>
            </ListView.View>
        </ListView>
        <TextBox x:Name="TextBoxOne" HorizontalAlignment="Left" Margin="34,25,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="334" RenderTransformOrigin="0.233,0.392"/>
        <Button x:Name="BtnOne" Content="Search" HorizontalAlignment="Left" Margin="391,23,0,0" VerticalAlignment="Top" Width="62"/>
        <Label x:Name="LabelOne" Content="E-Mail:" HorizontalAlignment="Left" Margin="472,64,0,0" VerticalAlignment="Top" RenderTransformOrigin="1.105,-0.692"/>
        <Label x:Name="LabelTwo" Content="Department:" HorizontalAlignment="Left" Margin="472,107,0,0" VerticalAlignment="Top"/>
        <Label x:Name="LabelThree" Content="Position:" HorizontalAlignment="Left" Margin="472,156,0,0" VerticalAlignment="Top"/>
        <Label x:Name="LabelFour" Content="" HorizontalAlignment="Left" Margin="523,64,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.079,-0.076" Width="267"/>
        <Label x:Name="LabelFive" Content="" HorizontalAlignment="Left" Margin="553,107,0,0" VerticalAlignment="Top" RenderTransformOrigin="0,-0.114" Width="201"/>
        <Label x:Name="LabelSix" Content="" HorizontalAlignment="Left" Margin="532,156,0,0" VerticalAlignment="Top" Width="247"/>
        <Button x:Name="BtnTwo" Content="Reset Password" HorizontalAlignment="Left" Margin="510,226,0,0" VerticalAlignment="Top" Width="128" IsEnabled='False'/>

    </Grid>
</Window>
'@
$XAMLDesign=$XAMLDesign -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
[XML]$XAML=$XAMLDesign 
$Reader = (New-Object System.Xml.XmlNodeReader $XAML)
$GUI = [Windows.Markup.XamlReader]::Load($Reader)
$XAML.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $GUI.FindName($_.Name) }
if (Get-Module -ListAvailable -Name ActiveDirectory) {
$BtnOne.Add_Click{
    $Fil = $TextBoxOne.Text
    if($Fil -eq ''){
       [System.Windows.MessageBox]::Show("This field cannot be empty","Error","OK") 
    } else{
    $All = Get-ADUser -Filter * -Properties * | select Name,mail | Where-Object {($_.Name -like $Fil) -and (-not $_.mail -eq '')}
    if($All.Count -eq 0){
    $ListViewOne.Items.Clear()
    $ListViewOne.Items.Add([PSCustomObject]@{'Parameters'='Object not found'})
    } else{
    $ListViewOne.Items.Clear()
    Foreach($User in $All){
    $ListViewOne.Items.Add([PSCustomObject]@{'Parameters'="$($User.Name)"})
}}}}
$ListViewOne.Add_SelectionChanged{
$Hash = $ListViewOne.Items
$In = $ListViewOne.SelectedIndex
$Info = Get-ADUser -Filter * -Properties * | select Name, Title, Department, EmailAddress | Where-Object {$_.Name -eq $Hash[$In].Parameters}
$LabelFour.Content = $Info.EmailAddress
$LabelFive.Content = $Info.Department
$LabelSix.Content = $Info.Title
$BtnTwo.IsEnabled = 'True'
}
$BtnTwo.Add_Click{
$LabelValue = $LabelFour.Content
$Find = Get-ADUser -Filter * -Properties * | select Name, mail | Where-Object {$_.mail -eq $LabelValue}
$Pass = [System.Web.Security.Membership]::GeneratePassword(15,3)
$Encrypted = ConvertTo-SecureString -String $Pass -AsPlainText -Force
Set-ADAccountPassword -Identity $Find.Name -Reset -NewPassword $Encrypted
Set-ADUser -Identity $Find.Name -ChangePasswordAtLogon $True
$From = 'pstest@gmail.com'
$MailPass = Get-Content C:\Users\Administrator\Documents\passtest.txt 
$MailPass = ConvertTo-SecureString -String $MailPass -Force -AsPlainText
$Msg = "Welcome $($Find.Name). Here you have your new password:$Pass"
$creden = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $From, $MailPass
Send-MailMessage -From $From -To $Find.mail -Subject 'New Password' -Body $Msg -SmtpServer 'smtp.gmail.com' -Port '587' -Credential $creden -UseSsl
[System.Windows.MessageBox]::Show("User password changed","Notification","OK")
}
} else {
    [System.Windows.MessageBox]::Show("There is no ActiveDirectory Module","Error","OK")
    exit
}
$GUI.ShowDialog() | Out-Null
