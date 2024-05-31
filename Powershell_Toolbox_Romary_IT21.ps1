# Funktsioonide defineerimine

# Suured failid
function Get-LargeFiles {
    param (
        [string]$Path,
        [int]$MinSizeMB = 100
    )

    Get-ChildItem -Path $Path -Recurse -File | Where-Object { $_.Length -gt ($MinSizeMB * 1MB) } | 
    Select-Object Name, @{Name="SizeMB";Expression={[math]::round($_.Length / 1MB, 2)}}, FullName | 
    Sort-Object SizeMB -Descending
}

# Varundamine
function New-Backup {
    param (
        [string]$SourcePath,
        [string]$BackupPath
    )

    $originalFileName = (Get-Item $SourcePath).Name
    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm"
    $backupFileName = "$originalFileName" + "_" + "$timestamp"
    $destination = Join-Path -Path $BackupPath -ChildPath $backupFileName

    Copy-Item -Path $SourcePath -Destination $destination -Recurse -Force
    Write-Output "Failid on edukalt varundatud asukohta: $destination"
}

# Süsteemi informatsioon
function Get-SystemInfo {
    $os = Get-CimInstance Win32_OperatingSystem
    $cpu = Get-CimInstance Win32_Processor
    $memory = Get-CimInstance Win32_PhysicalMemory
    $disk = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }

    [PSCustomObject]@{
        OS        = $os.Caption
        OSVersion = $os.Version
        CPU       = $cpu.Name
        MemoryGB  = [math]::round(($memory.Capacity | Measure-Object -Sum).Sum / 1GB, 2)
        DiskUsage = $disk | Select-Object DeviceID, @{Name="UsedGB";Expression={[math]::round(($_.Size - $_.FreeSpace) / 1GB, 2)}}, @{Name="TotalGB";Expression={[math]::round($_.Size / 1GB, 2)}}
    }
}

# Võrguühenduse kontroll
function Test-InternetConnection {
    $pingResults = Test-Connection -ComputerName "www.google.com" -Count 5

    # Andmepakettide (packet loss) kaotuse protsendi arvutamine
    $packetLossCount = ($pingResults | Measure-Object -Property StatusCode -Maximum).Maximum
    $packetLossPercentage = if ($packetLossCount) { ($packetLossCount / ($pingResults.Count * 5)) * 100 } else { 0 }

    $summary = "Andmepakketide kaotus (packet loss): $packetLossPercentage%"

    # Tulemuse kuvamine
    $pingDetails = foreach ($result in $pingResults) {
        $address = "Aadress: $($result.Address)"
        $responseTime = "Reageerimise aeg: $($result.ResponseTime)ms"
        $padding = " " * (30 - $address.Length)
        "$address$padding$responseTime`n"
    }

    if ($packetLossPercentage -eq 0) {
        $resultMessage = "Edukas"
    } else {
        $resultMessage = "Ebaõnnestunud"
    }

    $resultMessage += "`n`n" + $summary

    [System.Windows.MessageBox]::Show("Võrguühenduse test`nTulemus: $resultMessage`n`n$pingDetails", "Võrguühenduse test", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
}

function Get-EventLogReport {
    param (
        [string]$LogName = "System",
        [int]$LastDays = 7
    )

    $events = Get-EventLog -LogName $LogName -After (Get-Date).AddDays(-$LastDays)
    $report = $events | Group-Object Source | Select-Object Name, @{Name="EventCount";Expression={$_.Count}}, @{Name="RecentEvents";Expression={$_.Group | Select-Object TimeGenerated, Message -First 5}}

    $report
}

# Funktsioon failihalduri (file explorer) avamiseks
function Show-FolderBrowserDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $result = $folderBrowser.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    } else {
        return $null
    }
}

# Funktsioon sisestuskasti kuvamiseks
function Show-InputBox {
    param (
        [string]$message,
        [string]$title = "Sisend on vajalik"
    )
    Add-Type -AssemblyName Microsoft.VisualBasic
    [Microsoft.VisualBasic.Interaction]::InputBox($message, $title)
}

# WPF GUI loomine
Add-Type -AssemblyName PresentationFramework

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Romary Powershell Toolbox"
        Height="315" Width="400"
        ResizeMode="NoResize" WindowStyle="SingleBorderWindow" WindowStartupLocation="CenterScreen">
    <StackPanel Background="LightGray">
        <Button Name="GetLargeFilesButton" Content="Mahukate failide otsing" Padding="5" Margin="10" FontSize="16" Background="SkyBlue"/>
        <Button Name="NewBackupButton" Content="Failide varundamine" Padding="5" Margin="10" FontSize="16" Background="SkyBlue"/>
        <Button Name="GetSystemInfoButton" Content="Süsteemi-info" Padding="5" Margin="10" FontSize="16" Background="SkyBlue"/>
        <Button Name="TestInternetConnectionButton" Content="Võrguühenduse kontroll" Padding="5" Margin="10" FontSize="16" Background="SkyBlue"/>
        <Button Name="GetEventLogReportButton" Content="Sündmuste logi" Padding="5" Margin="10" FontSize="16" Background="SkyBlue"/>
    </StackPanel>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

$window.FindName("GetLargeFilesButton").Add_Click({
    [System.Windows.MessageBox]::Show("Valige kaust", "Valige kaust", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    $path = Show-FolderBrowserDialog
    if ($path) {
        $minSizeMB = Show-InputBox -message "Sisestage minimaalne faili suurus megabaitides" -title "Minimaalne faili suurus"
        if ($minSizeMB -match '\d+') {
            Get-LargeFiles -Path $path -MinSizeMB ([int]$minSizeMB) | Out-GridView
        } else {
            Write-Host "Vigane sisend. Palun sisestage kehtiv täisarv."
        }
    }
})

$window.FindName("NewBackupButton").Add_Click({
    # Kasutajalt varundatava kausta küsimine
    [System.Windows.MessageBox]::Show("Valige varunduse allikas", "Varunduse allika valimine", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    $sourcePath = Show-FolderBrowserDialog
    if ($sourcePath) {
        [System.Windows.MessageBox]::Show("Varunduse allikas valitud edukalt.", "Varunduse allikas valitud", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        
        # Kasutajalt varunduse sihtkoha küsimine
        [System.Windows.MessageBox]::Show("Valige varunduse sihtkoht", "Varunduse sihtkoha valimine", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        $backupPath = Show-FolderBrowserDialog
        if ($backupPath) {
            [System.Windows.MessageBox]::Show("Varunduse sihtkoht valitud edukalt.", "Varunduse sihtkoht valitud", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)

            # Varundamise käivitamine
            New-Backup -SourcePath $sourcePath -BackupPath $backupPath

            # Kuva edukuse teade
            [System.Windows.MessageBox]::Show("Varundamine lõpetatud. Failid varundati siia: $backupPath", "Varundamine lõpetatud", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
    }
})


$window.FindName("GetSystemInfoButton").Add_Click({
    Get-SystemInfo | Out-GridView
})

$window.FindName("TestInternetConnectionButton").Add_Click({
    Test-InternetConnection
})

$window.FindName("GetEventLogReportButton").Add_Click({
    try {
        # Loetelu hankimine saadaolevatest sündmuste logidest
        $availableEventLogs = Get-EventLog -List | Where-Object { $_.Entries.Count -gt 1 } | Select-Object -ExpandProperty LogDisplayName

        if ($availableEventLogs.Count -eq 0) {
            throw "Sündmuste logisid rohkem kui ühe kirjega ei leitud."
        }

        # Palu kasutajal valida sündmuste logi
        $logName = Show-InputBox -message "Saadaolevad sündmuste logid:`n`n$($availableEventLogs -join "`n")" -title "Sisesta sündmuste logi nimi"

        if ([string]::IsNullOrWhiteSpace($logName)) {
            Write-Host "Sündmuste logi nime ei sisestatud."
            return
        }

        # Kontrolli, kas sisestatud sündmuste logi nimi eksisteerib saadaolevate sündmuste logide hulgas
        if ($availableEventLogs -notcontains $logName) {
            throw "Vigane sündmuste logi nimi sisestatud. Palun proovi uuesti."
        }

        # Palu kasutajal sisestada päevade arv
        $lastDays = Show-InputBox -message "Sisesta tagasiulatuvate päevade arv" -title "Sisesta päevade arv"
        if (![int]::TryParse($lastDays, [ref]$null)) {
            throw "Vigane sisend päevade arvu jaoks. Palun sisesta kehtiv täisarv."
        }

        # Hangi ja kuva sündmuste logi aruanne
        Get-EventLogReport -LogName $logName -LastDays $lastDays | Out-GridView
    }
    catch {
        [System.Windows.MessageBox]::Show("Viga: $_`nPalun proovi uuesti.", "Viga", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

$window.ShowDialog() | Out-Null