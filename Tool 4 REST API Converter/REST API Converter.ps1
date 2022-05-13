#Author of script: Jakub Frąckowiak
#This app requires Excel module
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
$XAMLDesign = @'
<Window x:Class="WpfApp7.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp7"
        mc:Ignorable="d"
        Title="REST API Converter" Height="176" Width="800">
    <Grid>
        <Label x:Name="LabelOne" Content="Paste the API link" HorizontalAlignment="Left" Margin="18,26,0,0" VerticalAlignment="Top"/>
        <TextBox x:Name="TextBoxOne" HorizontalAlignment="Left" Margin="141,30,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="574"/>
        <Button x:Name="BtnOne" Content="Show Content" HorizontalAlignment="Left" Margin="27,72,0,0" VerticalAlignment="Top" Width="83"/>

    </Grid>
</Window>
'@
$XAMLDesign=$XAMLDesign -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
[XML]$XAML=$XAMLDesign 
$Reader = (New-Object System.Xml.XmlNodeReader $XAML)
$GUI = [Windows.Markup.XamlReader]::Load($Reader)
$XAML.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $GUI.FindName($_.Name) }
$BtnOne.Add_Click{
    if ($TextBoxOne.Text.Contains("http") -or $TextBoxOne.Text.Contains("https")) {
    $Txt = $TextBoxOne.Text
    $Result = Invoke-RestMethod -Uri "$Txt"
    #GuiTwo
    $XAMLDesignTwo = @'
    <Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:local="clr-namespace:WpfApp7"
    mc:Ignorable="d"
    Title="REST API Converter" Height="312" Width="590">
<Grid>
    <ListView x:Name="ListViewOne" Margin="27,10,240,10">
        <ListView.View>
            <GridView>
                <GridViewColumn Header="Content" DisplayMemberBinding = "{Binding Content}" Width="320"/>
            </GridView>
        </ListView.View>
    </ListView>
    <Label x:Name="LabelTwo" Content="Choose Data Format" HorizontalAlignment="Left" Margin="355,10,0,0" VerticalAlignment="Top"/>
    <RadioButton x:Name="RBOne" Content="CSV" HorizontalAlignment="Left" Margin="360,60,0,0" VerticalAlignment="Top" GroupName="On"/>
    <RadioButton x:Name="RBTwo" Content="JSON" HorizontalAlignment="Left" Margin="426,60,0,0" VerticalAlignment="Top" GroupName="On" RenderTransformOrigin="0.182,0.537"/>
    <RadioButton x:Name="RBThree" Content="XLSX" HorizontalAlignment="Left" Margin="495,60,0,0" VerticalAlignment="Top" GroupName="On"/>
    <Button x:Name="BtnTwo" Content="Export" HorizontalAlignment="Left" Margin="361,107,0,0" VerticalAlignment="Top" Height="27" Width="89"/>
</Grid>
</Window>
'@
$XAMLDesignTwo=$XAMLDesignTwo -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
[XML]$XAMLTwo=$XAMLDesignTwo 
$ReaderTwo = (New-Object System.Xml.XmlNodeReader $XAMLTwo)
$GUITwo = [Windows.Markup.XamlReader]::Load($ReaderTwo)
$XAMLTwo.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $GUITwo.FindName($_.Name) }
$ListViewOne.Items.Clear()
$PropertiesPSCustomObject = ($Result | Get-Member -MemberType NoteProperty | select *).Name
$D = ''
foreach ($Tab in $Result) {
    foreach($Pr in $PropertiesPSCustomObject){
        $D = $Tab.$Pr
        $ListViewOne.Items.Add([PSCustomObject]@{'Content'="$Pr : $D"})
    }
}
$BtnTwo.Add_Click{
if ($RBOne.IsChecked){
    foreach ($Tab in $Result) {
        foreach($Pr in $PropertiesPSCustomObject){
            $D = $Tab.$Pr
            $ExCSV = @{
                $Pr = $D
            }
            $ExCSV.GetEnumerator() | Select-Object -Property Key,Value | Export-Csv -Path './Result.csv' -Append -Force -NoTypeInformation
        }
    } 
}
elseif($RBTwo.IsChecked){
    $ExJSON = @{}
    foreach ($Tab in $Result) {
        $i = 1;
        foreach($Pr in $PropertiesPSCustomObject){
            $D = $Tab.$Pr
            $ExJSON.Add($Pr,$D)
            if ($i -eq 5){
              $ExJSON | ConvertTo-Json -Depth 10 | Out-File './Result.json' -Append -Force
              $ExJSON = @{}
            }
            $i = $i + 1
        }
    } 
}
elseif($RBThree.IsChecked){
    if (Get-Module -ListAvailable -Name ImportExcel) { 
        foreach ($Tab in $Result) {
            foreach($Pr in $PropertiesPSCustomObject){
                $D = $Tab.$Pr
                $ExXLSX = @{
                    $Pr = $D
                }
                $ExXLSX.GetEnumerator() | Select-Object -Property Key,Value | Export-Excel -Path './Result.xlsx' -Append   
            }
        } 
    }
    else{
        [System.Windows.MessageBox]::Show("No ImportExcel Module","Error","OK")  
    }
}
else {
    [System.Windows.MessageBox]::Show("You didn't check any option","Notification","OK")
}
}
$GUITwo.ShowDialog() | Out-Null
}
else {
    [System.Windows.MessageBox]::Show("Wrong api link","Error","OK")
}
} 
$GUI.ShowDialog() | Out-Null
