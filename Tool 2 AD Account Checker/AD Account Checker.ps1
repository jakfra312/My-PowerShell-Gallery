[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
$XAMLDesign = @'
<Window x:Class="WpfApp4.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp4"
        mc:Ignorable="d"
        Title="ADAccountChecker" Height="450" Width="508">
    <Grid HorizontalAlignment="Left" Width="790">
        <ListView x:Name="ListViewOne" Margin="184,12,332,11">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Users" DisplayMemberBinding = "{Binding Users}" Width="265"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Button x:Name="BtnOne" Content="Show Users" HorizontalAlignment="Left" Margin="19,117,0,0" VerticalAlignment="Top" Height="26" Width="89"/>
        <TextBox x:Name="TextBoxOne" HorizontalAlignment="Left" Margin="19,68,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Label Content="Enter Number Of Days" HorizontalAlignment="Left" Margin="19,20,0,0" VerticalAlignment="Top" Width="130"/>
    </Grid>
</Window>
'@
$XAMLDesign=$XAMLDesign -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
[XML]$XAML=$XAMLDesign 
$Reader = (New-Object System.Xml.XmlNodeReader $XAML)
$GUI = [Windows.Markup.XamlReader]::Load($Reader)
$XAML.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $GUI.FindName($_.Name) }
if(-Not (Get-Module -ListAvailable -Name ActiveDirectory)){
    [System.Windows.MessageBox]::Show("There isn't Active Directory Module","Error","OK")
    exit
}
$BtnOne.Add_Click{
[Int32]$TxtVar = $TextBoxOne.Text
if ($TxtVar.GetType().Name -eq 'Int32' -and $TxtVar -gt 0) {
$Day = Get-Date
$Ago = $Day.AddDays(-1*$TxtVar)
$ListViewOne.Items.Clear()
$Users = (Get-ADUser -Filter "(Enabled -eq 'True') -and (PasswordLastSet -lt $Ago)").Name
Foreach($User in $Users){
$ListViewOne.Items.Add([PSCustomObject]@{'Users'="$User"})
}

} else {
    [System.Windows.MessageBox]::Show("Value in TextBox must be a number and greater than zero","Error","OK")
}
}
$GUI.ShowDialog() | Out-Null
#Author of script: Jakub Frąckowiak
