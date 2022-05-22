#Author: Jakub Fr¹ckowiak
#This tool requires ActiveDirectory and ImportExcel Module
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
Add-Type -Assembly System.Windows.forms
Add-Type -AssemblyName System.Web
$XAMLDesign = @'
<Window x:Class="WpfApp8.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp8"
        mc:Ignorable="d"
        Title="Automatic AD Account Creator" Height="557" Width="1036">
    <Grid Margin="0,0,0,-16">
        <Label x:Name="LabelOne" Content="Available Formats: XLSX CSV JSON" HorizontalAlignment="Left" Margin="25,21,0,0" VerticalAlignment="Top"/>
        <Button x:Name="BTNOne" Content="Pick a file" HorizontalAlignment="Left" Margin="25,60,0,0" VerticalAlignment="Top" Width="83"/>
        <TextBox x:Name="TextBoxOne" HorizontalAlignment="Left" Margin="146,60,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="859" Height="20"/>
        <ListView x:Name="ListViewOne" Margin="25,159,6,87" IsEnabled="False">
            <ListView.View>
                <GridView>
                    <GridViewColumn Header="Name" DisplayMemberBinding ="{Binding 'Name'}" Width="150"/>
                    <GridViewColumn Header="Surname" DisplayMemberBinding ="{Binding 'Surname'}" Width="150"/>
                    <GridViewColumn Header="E-mail" DisplayMemberBinding ="{Binding 'Email'}" Width="250"/>
                    <GridViewColumn Header="Telephone" DisplayMemberBinding ="{Binding 'Telephone'}" Width="100"/>
                    <GridViewColumn Header="Position" DisplayMemberBinding ="{Binding 'Position'}" Width="150"/>
                    <GridViewColumn Header="Department" DisplayMemberBinding ="{Binding 'Department'}" Width="200"/>
                </GridView>
            </ListView.View>
        </ListView>
        <Label x:Name="LabelTwo" Content="Choose Group Scope" HorizontalAlignment="Left" Margin="38,457,0,0" VerticalAlignment="Top"/>
        <ComboBox x:Name="ComboBoxOne" HorizontalAlignment="Left" Margin="176,461,0,0" VerticalAlignment="Top" Width="173" RenderTransformOrigin="0.142,1.095" IsEnabled="False">
         <ComboBoxItem>Universal</ComboBoxItem>
         <ComboBoxItem>Global</ComboBoxItem>
         <ComboBoxItem>DomainLocal</ComboBoxItem>
        </ComboBox>
        <Label x:Name="LabelThree" Content="Choose Group" HorizontalAlignment="Left" Margin="363,457,0,0" VerticalAlignment="Top"/>
        <ComboBox x:Name="ComboBoxTwo" HorizontalAlignment="Left" Margin="468,461,0,0" VerticalAlignment="Top" Width="209" IsEnabled="False"/>
        <Button x:Name="BTNTwo" Content="Create And Add Users" HorizontalAlignment="Left" Margin="743,460,0,0" VerticalAlignment="Top" Width="136" Height="23" IsEnabled="False"/>
        <Button x:Name="BTNThree" Content="Load Content" HorizontalAlignment="Left" Margin="25,105,0,0" VerticalAlignment="Top" Width="83" IsEnabled="False"/>
    </Grid>
</Window>
'@
$XAMLDesign=$XAMLDesign -replace 'mc:Ignorable="d"','' -replace "x:N","N" -replace '^<Win.*','<Window'
[XML]$XAML=$XAMLDesign 
$Reader = (New-Object System.Xml.XmlNodeReader $XAML)
$GUI = [Windows.Markup.XamlReader]::Load($Reader)
$XAML.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $GUI.FindName($_.Name)}

function ComboBoxChoose {
$ComboBoxOne.Add_SelectionChanged({
    $ComboBoxTwo.IsEnabled = 'True'
    $Scope = $ComboBoxOne.SelectedItem.Content
    $Grp = Get-ADGroup -Filter * | select GroupScope, Name | Where-Object {$_.GroupScope -eq $Scope}
    $ComboBoxTwo.Items.Clear()
    Foreach($Group in $Grp){
    $ComboBoxTwo.Items.Add($Group.Name)
    }})}

function AddToListView {
$ListViewOne.Items.Clear()
    foreach($Row in $DFile){
         $ListViewOne.Items.Add([PSCustomObject]@{'Name'="$($Row.Name)";'Surname'="$($Row.Surname)";'Email'="$($Row.Email)";'Telephone'="$($Row.Telephone)";'Position'="$($Row.Position)";'Department'="$($Row.Department)"}) 
    }
}

function ComboBoxEnable {
$ComboBoxTwo.Add_SelectionChanged({
    $BTNTwo.IsEnabled = 'True'
   })
}

function AddToAD {
$BTNTwo.Add_Click{
    $ComboBoxTwoSelectedItem = $ComboBoxTwo.Items[$ComboBoxTwo.SelectedIndex]
    $Forest = (Get-ADDomain).DNSRoot
    $FromListView = $ListViewOne.Items
    $MailPass = Get-Content C:\Users\Administrator\Documents\passtest.txt 
    $MailPass = ConvertTo-SecureString -String $MailPass -Force -AsPlainText
    $From = 'bagienko.jankos@gmail.com' 
    $creden = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $From, $MailPass
    Foreach($Item in $FromListView){
        $FullName = "$($Item.Name) $($Item.Surname)"
        $PName = "$($Item.Name).$($Item.Surname)@$Forest"
        if(Get-ADUser -Filter "Name -eq '$FullName'"){
        Write-Host "There is already a user with this name: $FullName"}
        else{
        $Pass = [System.Web.Security.Membership]::GeneratePassword(15,3)
        $Encrypted = ConvertTo-SecureString -String $Pass -AsPlainText -Force
        New-ADUser -Name $FullName -Surname $Item.Surname -Department $Item.Department -MobilePhone $Item.Telephone -EmailAddress $Item.Email -Title $Item.Position -GivenName $Item.Name -UserPrincipalName $PName -AccountPassword $Encrypted
        Enable-ADAccount -Identity $FullName
        Set-ADUser -Identity $FullName -ChangePasswordAtLogon $True
        Write-Host "User Added: $FullName"
        Add-ADGroupMember -Members $FullName -Identity $ComboBoxTwoSelectedItem
        $Msg = "Welcome $FullName in our company. Here you have your login:$PName and your password:$Pass"
        Send-MailMessage -From $From -To $Item.Email -Subject 'Login Data' -Body $Msg -SmtpServer 'smtp.gmail.com' -Port '587' -Credential $creden -UseSsl
        }}}
}

if ((Get-Module -ListAvailable -Name ActiveDirectory) -and (Get-Module -ListAvailable -Name ImportExcel)){
    
    $BTNOne.Add_Click{
    $PickFile = New-Object System.Windows.Forms.OpenFileDialog
    $PickFile.ShowDialog()
    $TextBoxOne.Text = $PickFile.FileName
    $BTNThree.IsEnabled = 'True'}

    $BTNThree.Add_Click{
    if($TextBoxOne.Text.Contains('xlsx') -or $TextBoxOne.Text.Contains('json') -or $TextBoxOne.Text.Contains('csv')){
    $ListViewOne.IsEnabled = 'True'
    $ComboBoxOne.IsEnabled = 'True'
    if($TextBoxOne.Text.Contains('xlsx')){
    $DFile = Import-Excel $TextBoxOne.Text
    AddToListView
    ComboBoxChoose
    ComboBoxEnable
    AddToAD
    }

    elseif($TextBoxOne.Text.Contains('csv')){
    $DFile = Import-Csv -Path $TextBoxOne.Text
    AddToListView
    ComboBoxChoose
    ComboBoxEnable
    AddToAD}

    else {
    $DFile = (Get-Content -Path $TextBoxOne.Text -Raw | ConvertFrom-Json).Users
    AddToListView
    ComboBoxChoose
    ComboBoxEnable
    AddToAD
    }

     }
     else {
       [System.Windows.MessageBox]::Show("Wrong data format","Error","OK")
   }}} 
    else {
    [System.Windows.MessageBox]::Show("There is not ActiveDirectory or ImportExcel Module","Error","OK")
    exit
}
$GUI.ShowDialog() | Out-Null
