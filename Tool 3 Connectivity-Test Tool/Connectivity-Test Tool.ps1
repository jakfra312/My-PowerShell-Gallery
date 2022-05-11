[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
Add-Type -Assembly System.Windows.forms
$XAMLDesign = @'
<Window x:Class="WpfApp5.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp5"
        mc:Ignorable="d"
        Title="Connectivity-Test Tool" Height="191" Width="800">
    <Grid Height="310" VerticalAlignment="Top">
        <Label x:Name="LabelOne" Content="Input Data" HorizontalAlignment="Left" Margin="31,23,0,0" VerticalAlignment="Top" Height="25" Width="66"/>
        <Button x:Name="BtnOne" Content="Pick File" HorizontalAlignment="Left" Margin="31,61,0,0" VerticalAlignment="Top" Width="66"/>
        <TextBox x:Name="TextBoxOne" HorizontalAlignment="Left" Margin="119,61,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="635" Height="20" FontSize="13"/>
        <Label x:Name="LabelTwo" Content="Pick Format" HorizontalAlignment="Left" Margin="31,109,0,0" VerticalAlignment="Top"/>
        <RadioButton x:Name="RadioTwo" Content="JSON" HorizontalAlignment="Left" Margin="128,114,0,0" VerticalAlignment="Top" GroupName="RadioBtn"/>
        <RadioButton x:Name="RadioOne" Content="CSV" HorizontalAlignment="Left" Margin="197,114,0,0" VerticalAlignment="Top" RenderTransformOrigin="0.017,0.448" GroupName="RadioBtn"/>
        <Button x:Name="BtnTwo" Content="Check Connection" HorizontalAlignment="Left" Margin="265,112,0,0" VerticalAlignment="Top" Width="108" IsEnabled="True"/>

    </Grid>
</Window>
'@
$XAMLDesign=$XAMLDesign -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
[XML]$XAML=$XAMLDesign 
$Reader = (New-Object System.Xml.XmlNodeReader $XAML)
$GUI = [Windows.Markup.XamlReader]::Load($Reader)
$XAML.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $GUI.FindName($_.Name) }
$BtnOne.Add_Click{
    $PickFile = New-Object System.Windows.Forms.OpenFileDialog
    $PickFile.ShowDialog()
    $TextBoxOne.Text = $PickFile.FileName
}
$BtnTwo.Add_Click{
    if($TextBoxOne.Text -eq ''){
        [System.Windows.MessageBox]::Show("You didn't choose a file","Error","OK") 
    }
    elseif ((-not $RadioOne.IsChecked) -and (-not $RadioTwo.IsChecked)){
        [System.Windows.MessageBox]::Show("You didn't pick a data format","Error","OK") 
    }
    else {
        if(($RadioOne.IsChecked) -and ($TextBoxOne.Text.SubString($TextBoxOne.Text.Length-3) -eq 'csv')){
              $FileCSV = $TextBoxOne.Text
              $RowsCSV = Import-Csv -Path $FileCSV
              Write-Host 'IsConnected | Computername | IPAddress'
              foreach ($row in $RowsCSV) {
                  try{
                     $Info = @{
                         Hostname = $row.Hostname
                         IPAddress = $row.IPAddress
                         Connected = 'No'
                     }
                     if(Test-Connection -ComputerName $row.IPAddress -Count 1 -Quiet){
                         $Info.Connected = 'Yes'
                     } 
                     Write-Host $Info.Values
                  }
                  catch {
                    [System.Windows.MessageBox]::Show("Something went wrong","Error","OK")
                  }
              }
              
        }
        elseif(($RadioTwo.IsChecked) -and ($TextBoxOne.Text.Substring($TextBoxOne.Text.Length-4)) -eq 'json'){
            $FileJSON = $TextBoxOne.Text
            $RowsJSON = (Get-Content -Path $FileJSON -Raw | ConvertFrom-Json).Comps
            Write-Host 'IsConnected | Computername | IPAddress'
            foreach ($rowJ in $RowsJSON) {
                try{
                   $InfoJ = @{
                       IPAddress = $rowJ.IPAddress
                       Hostname = $rowJ.Hostname
                       Connected = 'No'
                   }
                   if(Test-Connection -ComputerName $rowJ.IPAddress -Count 1 -Quiet){
                       $InfoJ.Connected = 'Yes'
                   } 
                   Write-Host $InfoJ.Values
                }
                catch {
                  [System.Windows.MessageBox]::Show("Something went wrong","Error","OK")
                }
            }
        }
        else{
            [System.Windows.MessageBox]::Show("Wrong Data Format","Error","OK")
        }
    }
}
$GUI.ShowDialog() | Out-Null
#Author of scripts: Jakub Frąckowiak