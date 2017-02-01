# mechhvac
# Tieto
# some change

Add-Type -AssemblyName presentationframework

#region GUI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShellGUI"
        Title="Simple calendar rights GUI" Height="391" Width="543">
    <Grid>
        <TextBox x:Name="textBox" HorizontalAlignment="Left" Height="21" Margin="35,16,0,0" TextWrapping="Wrap" Text="mechhvac@corp.contoso.com" VerticalAlignment="Top" Width="202"/>
        <TextBox x:Name="textBoxUser" HorizontalAlignment="Left" Height="21" Margin="35,46,0,0" TextWrapping="Wrap" Text="user" VerticalAlignment="Top" Width="202"/>
        <TextBox x:Name="outputTextBox" HorizontalAlignment="Left" Height="187" Margin="35,118,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="319" />
        <Button x:Name="buttonGet" Content="Get!" HorizontalAlignment="Left" Height="20" Margin="270,16,0,0" VerticalAlignment="Top" Width="78"/>
        <Button x:Name="buttonSet" Content="Set!" HorizontalAlignment="Left" Height="20" Margin="270,46,0,0" VerticalAlignment="Top" Width="78"/>
        <ComboBox x:Name="comboBox" HorizontalAlignment="Left" Margin="35,81,0,0" VerticalAlignment="Top" Width="120"/>
        <Label x:Name="labelCalendarPath" Content="[calendar path]" HorizontalAlignment="Left" Height="30" Margin="35,321,0,0" VerticalAlignment="Top" Width="481"/>
    </Grid>
</Window>
"@
$reader=(New-Object System.Xml.XmlNodeReader $xaml)

$Window=[Windows.Markup.XamlReader]::Load($reader)

#endregion GUI

#Connect to Controls 

$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object {
    New-Variable  -Name $_.Name -Value $window.FindName($_.Name) -Force
}

# data

$listItemSource = New-Object System.Collections.ObjectModel.ObservableCollection[string]
$listItemSource.Add("Reviewer")
$listItemSource.Add("AvailabilityOnly")
$listItemSource.Add("Owner")
$listItemSource.Add("Author")
$comboBox.ItemsSource = $listItemSource
$comboBox.SelectedIndex = 0

# end date





#region Events
$buttonGet.Add_Click({
    try{
        if (-NOT ([string]::IsNullOrEmpty($textBox.Text))){
            Write-Host "Getting the calendar permissions.."
            $outputTextBox.Text = getCalendarData
            $labelCalendarPath.Content = getCalendarFolder
        }
    }catch{
        Write-Warning $_
    }
})


$buttonSet.Add_Click({
    try{
        if (-NOT ([string]::IsNullOrEmpty($textBoxUser.Text))){
            Write-Host "Setting the calendar permissions.."
            $calendarData = getCalendarData
            Write-Host $calendarData
            Write-Host $textBoxUser.Text
            [bool]$contains = $calendarData.contains($textBoxUser.Text)
            
            Write-Host $($contains)
            
            if($contains){
                $folder = getCalendarFolder
                Set-MailboxFolderPermission $($folder) -User $textBoxUser.Text -AccessRights $combobox.SelectedItem
                $outputTextBox.Text = getCalendarData
                Write-Host "Setting..."
            }else{
                $folder = getCalendarFolder
                Add-MailboxFolderPermission $($folder) -User $textBoxUser.Text -AccessRights $combobox.SelectedItem
                $outputTextBox.Text = getCalendarData
                Write-Host "Adding..."
            }
        }else {
            Write-Host "Booo!"
        }
    }catch{
        Write-Warning $_
    }
})

#endregion Events

function getCalendarFolder{
    $mailbox = $textbox.Text
    $calendarFolder = ((Get-MailboxFolderStatistics $mailbox | Where-Object {$_.FolderType -like "Calendar"}).Folderpath).Replace('/','\')
    $result = "$($mailbox):$calendarFolder"
    [string]$result
}

function getCalendarData{
    $folder = getCalendarFolder
    $calendarData = Get-MailboxFolderPermission $($folder) | Out-String
    $calendarData
}

$Null = $window.ShowDialog()