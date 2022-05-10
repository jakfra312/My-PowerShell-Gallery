[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
$XAMLDesign = @'
<Window x:Class="WpfApp3.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp3"
        mc:Ignorable="d"
        Title="HS Info" Height="339" Width="600">
    <Grid Height="405" VerticalAlignment="Top">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="108*"/>
            <ColumnDefinition Width="291*"/>
            <ColumnDefinition Width="77*"/>
            <ColumnDefinition Width="22*"/>
            <ColumnDefinition Width="102*"/>
        </Grid.ColumnDefinitions>
        <Button x:Name="BtnGet" Content="Get Details" HorizontalAlignment="Left" Margin="40,20,0,0" VerticalAlignment="Top" Grid.ColumnSpan="2" Height="29" Width="92" RenderTransformOrigin="0.471,-0.481"/>
        <Button x:Name="BtnCsv" Content="Export To CSV" HorizontalAlignment="Left" Margin="40,76,0,0" VerticalAlignment="Top" Grid.ColumnSpan="2" Height="29" Width="92" IsEnabled="False"/>
        <ListView x:Name="InfoList" Grid.Column="1" Grid.ColumnSpan="4" Margin="93,18,37,138"  ScrollViewer.VerticalScrollBarVisibility="Visible">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Parameters" DisplayMemberBinding ="{Binding Parameters}" Width="340"/>
                </GridView>
            </ListView.View>
        </ListView>
    </Grid>
</Window>
'@
$XAMLDesign=$XAMLDesign -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
[XML]$XAML=$XAMLDesign 
$Reader = (New-Object System.Xml.XmlNodeReader $XAML)
$GUI = [Windows.Markup.XamlReader]::Load($Reader)
$XAML.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $GUI.FindName($_.Name) }
Write-Host "Author of Script:Jakub Frąckowiak"
$Info = @{
'Computername' = $env:COMPUTERNAME
'AddressIP' =  (Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = 'True'").IPAddress[0]
'Diskspace (GB)' = [Math]::Round((Get-PSDrive -Name C).Free/1GB)
'RegisteredUser' = (Get-CimInstance -Class Win32_OperatingSystem).RegisteredUser
'OSVersion' = (Get-CimInstance -Class Win32_OperatingSystem).Version 
'Memory (GB)' = (Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum /1gb
'Manufacturer' = (Get-CimInstance -Class Win32_PhysicalMemory).Manufacturer
'Processor' = (Get-Ciminstance -Class Win32_Processor).Name
'Graphic' = (Get-Ciminstance -Class Win32_VideoController).Description
}
$BtnGet.Add_Click{
    try {
        $InfoList.Items.Clear()
        $Info.GetEnumerator() | ForEach-Object {
            $InfoList.Items.Add([PSCustomObject]@{'Parameters'="$($_.Key): $($_.Value)"})
        }
        if ($BtnCsv.IsEnabled = 'False'){
            $BtnCsv.IsEnabled = 'True'
        }   
    }
    catch {
     $_.Exception.Message 
    }
}
$BtnCsv.Add_Click{
    try {
        $Info.GetEnumerator() | Select-Object -Property Name,value | Export-Csv -Path .\HSInfo.csv -NoTypeInformation
        [System.Windows.MessageBox]::Show("Succesfully Exported to CSV!","Notification","OK")   
    }
    catch {
     $_.Exception.Message   
    }
}
$GUI.ShowDialog() | Out-Null
#Author of script: Jakub Frąckowiak