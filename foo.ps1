# mechhvac
# Tieto

# add support for WPF
Add-Type -AssemblyName presentationframework

#region GUI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShellGUI"
        Title="Simple calendar rights GUI, mechhvac" Height="433" Width="603">
    <Grid Margin="10,0,0,0" RenderTransformOrigin="0.531,0.507" HorizontalAlignment="Left" Width="563">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <TextBox x:Name="textBox" HorizontalAlignment="Left" Height="21" Margin="73,31,0,0" TextWrapping="Wrap" Text="mechhvac@corp.contoso.com" VerticalAlignment="Top" Width="202"/>
        <TextBox x:Name="textBoxUser" HorizontalAlignment="Left" Height="21" Margin="73,78,0,0" TextWrapping="Wrap" Text="user" VerticalAlignment="Top" Width="202"/>
        <Button x:Name="buttonGet" Content="Get!" HorizontalAlignment="Left" Height="20" Margin="335,32,0,0" VerticalAlignment="Top" Width="78"/>
        <Button x:Name="buttonSet" Content="Set!" HorizontalAlignment="Left" Height="20" Margin="335,79,0,0" VerticalAlignment="Top" Width="78" RenderTransformOrigin="0.577,0"/>
        <ComboBox x:Name="comboBox" HorizontalAlignment="Left" Margin="73,125,0,0" VerticalAlignment="Top" Width="148" Height="22"/>
        <Label x:Name="labelCalendarPath" Content="[calendar path]" HorizontalAlignment="Left" Height="30" Margin="73,362,0,0" VerticalAlignment="Top" Width="481"/>
        <Label x:Name="labelMailbox" Content="Mailbox:" HorizontalAlignment="Left" Margin="73,2,0,0" VerticalAlignment="Top" RenderTransformOrigin="0,3.731" Width="58" Height="35"/>
        <Label x:Name="labelUser" Content="User:" HorizontalAlignment="Left" Margin="73,52,0,0" VerticalAlignment="Top" Height="26" Width="36"/>
        <Label x:Name="labelAccessRights" Content="Access rights:" HorizontalAlignment="Left" Height="31" Margin="73,99,0,0" VerticalAlignment="Top" Width="120"/>
        <DataGrid x:Name="dataGrid" SelectionMode="Single" HorizontalAlignment="Left" Height="187" Margin="73,170,0,0" VerticalAlignment="Top" Width="440" AutoGenerateColumns="False" ColumnWidth="Auto">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Folder name" Binding="{Binding FolderName}"/>
                <DataGridTextColumn Header="User" Binding="{Binding User}"/>
                <DataGridTextColumn Header="Access rights" Binding="{Binding AccessRights}"/>
            </DataGrid.Columns>
        </DataGrid>
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

# fill the combo box with Access rights data
$listItemSource = New-Object System.Collections.ObjectModel.ObservableCollection[string]
$listItemSource.Add("Reviewer")
$listItemSource.Add("AvailabilityOnly")
$listItemSource.Add("Owner")
$listItemSource.Add("Author")
$comboBox.ItemsSource = $listItemSource
$comboBox.SelectedIndex = 0
# 

#region Events
$buttonGet.Add_Click({
    try{
        if (-NOT ([string]::IsNullOrEmpty($textBox.Text))){
            Write-Host "Getting the calendar permissions.."
            #$outputTextBox.Text = getCalendarData
            refreshDataGrid            
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
            $calendarData = getCalendarData | Out-String
            Write-Host $calendarData
            Write-Host $textBoxUser.Text
            [bool]$contains = $calendarData.contains($textBoxUser.Text)            
            Write-Host $($contains)
            
            if($contains){
                $folder = getCalendarFolder
                Set-MailboxFolderPermission $($folder) -User $textBoxUser.Text -AccessRights $combobox.SelectedItem
                refreshDataGrid
                Write-Host "Setting..."
            }else{
                $folder = getCalendarFolder
                Add-MailboxFolderPermission $($folder) -User $textBoxUser.Text -AccessRights $combobox.SelectedItem
                refreshDataGrid
                Write-Host "Adding..."
            }
        }else {
            Write-Host "wtf"
        }
    }catch{
        Write-Warning $_
    }
})

$dataGrid.Add_MouseDoubleClick({
        $selectedItem = $dataGrid.SelectedItems
        $textBoxUser.Text = $selectedItem.User    
})

#endregion Events


# supporting functions
function getCalendarFolder{
    $mailbox = $textbox.Text
    $calendarFolder = ((Get-MailboxFolderStatistics $mailbox | Where-Object {$_.FolderType -like "Calendar"}).Folderpath).Replace('/','\')
    $result = "$($mailbox):$calendarFolder"
    [string]$result
}

function getCalendarData{
    $folder = getCalendarFolder
    $calendarData = Get-MailboxFolderPermission $($folder) | Select-Object foldername, user, @{name="AccessRights";expression={ [string]::join(",",@($_.accessrights)) }}
    $calendarData
}

function refreshDataGrid{
    $dataGrid.clear()
    $dataGrid.ItemsSource = getCalendarData
    $dataGrid.IsReadOnly = $true
}
#

$Null = $window.ShowDialog()