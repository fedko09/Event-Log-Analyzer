<#
.SYNOPSIS
  Local Log Scanner GUI (Event Logs, Crash Dumps, Log Files) – PowerUser Edition

.DESCRIPTION
  - Windows 10/11
  - WPF UI with:
      * Source selector (Event Logs / Crash Dumps / Custom Folder)
      * Profiles (Custom, BSOD, App Crashes, Boot Issues)
      * Remote computer name (local or remote via Get-WinEvent -ComputerName)
      * Filters for event logs (time range, levels, Event ID, max events)
      * Search box (filters in-memory results)
      * DataGrid with results (sortable)
      * Tabs:
          - Details (full XML/message)
          - Summary (aggregated view)
          - System Snapshot (OS, uptime, disks, NICs)
      * Right-click context menu:
          - Open containing folder
          - Copy details to clipboard
      * Export-to-CSV (results)
      * Export Snapshot / Summary / Full Report
      * Support Bundle (ZIP with logs + system snapshot)
  - Crash dumps:
      * C:\Windows\Minidump\*.dmp
      * %SystemRoot%\MEMORY.DMP
  - Custom folder:
      * Recursively scans *.log, *.txt, *.evtx

  AI integration:
    - Placeholder function Invoke-AIAnalysis still present but not implemented.
#>

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# -----------------------------
# XAML UI
# -----------------------------
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Local Log Scanner" Height="600" Width="950"
    MinHeight="500" MinWidth="900"
    WindowStartupLocation="CenterScreen"
    ResizeMode="CanResize">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto" />
      <RowDefinition Height="*" />
      <RowDefinition Height="Auto" />
      <RowDefinition Height="180" />
      <RowDefinition Height="Auto" />
    </Grid.RowDefinitions>

    <!-- Top filter bar -->
    <StackPanel Grid.Row="0" Orientation="Vertical" Margin="0,0,0,8">
      <!-- First row -->
      <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
        <TextBlock Text="Source:" VerticalAlignment="Center" Margin="0,0,4,0" />
        <ComboBox x:Name="cbSource" Width="180" Margin="0,0,10,0"
                  ToolTip="Choose what to scan: Windows Event Logs, crash dump files, or log files in a custom folder." />

        <TextBlock Text="Profile:" VerticalAlignment="Center" Margin="0,0,4,0" />
        <ComboBox x:Name="cbProfile" Width="190" Margin="0,0,10,0"
                  ToolTip="Select a preset filter for common troubleshooting scenarios (BSOD, app crashes, boot issues).">
          <ComboBoxItem Content="Custom" IsSelected="True" />
          <ComboBoxItem Content="BSOD / Crash" />
          <ComboBoxItem Content="App Crashes" />
          <ComboBoxItem Content="Boot / Startup Issues" />
        </ComboBox>

        <TextBlock Text="Computer:" VerticalAlignment="Center" Margin="0,0,4,0" />
        <TextBox x:Name="txtComputer" Width="160" Margin="0,0,10,0"
                 Text="."
                 ToolTip="Target machine name. Use '.' or 'localhost' for this computer; use a hostname for remote." />

        <Button x:Name="btnScan" Content="Scan" Width="80"
                ToolTip="Run the scan using the filters and source selected above." />
      </StackPanel>

      <!-- Second row -->
      <StackPanel Orientation="Horizontal" Margin="0,0,0,4">
        <TextBlock Text="Days back:" VerticalAlignment="Center" Margin="0,0,4,0" />
        <TextBox x:Name="txtDays" Width="40" Text="7" Margin="0,0,10,0"
                 ToolTip="How many days of history to include when querying event logs." />

        <TextBlock Text="Event ID(s):" VerticalAlignment="Center" Margin="0,0,4,0" />
        <TextBox x:Name="txtEventId" Width="140" Margin="0,0,10,0"
                 ToolTip="Optional: comma or space separated Event IDs (e.g. 41, 1000). Leave blank for all IDs." />

        <TextBlock Text="Max events:" VerticalAlignment="Center" Margin="0,0,4,0" />
        <TextBox x:Name="txtMaxEvents" Width="60" Text="1000" Margin="0,0,10,0"
                 ToolTip="Maximum number of events to retrieve from logs (caps very noisy systems)." />

        <StackPanel Orientation="Horizontal" Margin="0,0,10,0">
          <CheckBox x:Name="chkSystem" Content="System" IsChecked="True" Margin="0,0,8,0"
                    ToolTip="Include the System event log." />
          <CheckBox x:Name="chkApplication" Content="Application" IsChecked="True"
                    ToolTip="Include the Application event log." />
        </StackPanel>

        <StackPanel Orientation="Horizontal" Margin="0,0,10,0">
          <CheckBox x:Name="chkError" Content="Error" IsChecked="True" Margin="0,0,8,0"
                    ToolTip="Include Error events." />
          <CheckBox x:Name="chkWarning" Content="Warning" IsChecked="True" Margin="0,0,8,0"
                    ToolTip="Include Warning events." />
          <CheckBox x:Name="chkInfo" Content="Information"
                    ToolTip="Include Information events (can be noisy)." />
        </StackPanel>
      </StackPanel>

      <!-- Third row -->
      <StackPanel Orientation="Horizontal">
        <TextBlock Text="Search:" VerticalAlignment="Center" Margin="0,0,4,0" />
        <TextBox x:Name="txtSearch" Width="220" Margin="0,0,6,0"
                 ToolTip="Filter the current results by matching text in Source, Message, Path, or LogName." />
        <Button x:Name="btnSearch" Content="Apply Filter" Width="100" Margin="0,0,6,0"
                ToolTip="Apply the search filter to the current result set." />
        <Button x:Name="btnClearSearch" Content="Clear Filter" Width="100" Margin="0,0,6,0"
                ToolTip="Clear the search filter and show all current results." />
        <Button x:Name="btnExport" Content="Export to CSV" Width="120"
                ToolTip="Export the current results in the grid to a CSV file." />
        <Button x:Name="btnSupportBundle" Content="Support Bundle" Width="130" Margin="6,0,0,0"
                ToolTip="Create a ZIP with logs, current results, and system information for support." />
        <Button x:Name="btnAiAnalyze" Content="Analyze (AI later)" Width="140" Margin="6,0,0,0" IsEnabled="False"
                ToolTip="Placeholder for future AI-based analysis of the collected logs." />
      </StackPanel>
    </StackPanel>

    <!-- Results grid -->
    <DataGrid
        x:Name="dgResults"
        Grid.Row="1"
        AutoGenerateColumns="False"
        IsReadOnly="True"
        CanUserSortColumns="True"
        CanUserResizeColumns="True"
        SelectionMode="Single"
        Margin="0,0,0,4"
        ToolTip="Results from the last scan. Click a row to see full details below; right-click for extra actions.">
      <DataGrid.ContextMenu>
        <ContextMenu x:Name="dgContextMenu">
          <MenuItem Header="Open containing folder" x:Name="miOpenFolder"
                    ToolTip="Open File Explorer at the folder containing this crash dump or log file." />
          <MenuItem Header="Copy details to clipboard" x:Name="miCopyDetails"
                    ToolTip="Copy the full details/XML of the selected entry to the clipboard." />
        </ContextMenu>
      </DataGrid.ContextMenu>
      <DataGrid.Columns>
        <DataGridTextColumn Header="Time" Binding="{Binding TimeCreated}" Width="150" />
        <DataGridTextColumn Header="Type" Binding="{Binding Type}" Width="70" />
        <DataGridTextColumn Header="Log" Binding="{Binding LogName}" Width="90" />
        <DataGridTextColumn Header="Source/Dir" Binding="{Binding Source}" Width="200" />
        <DataGridTextColumn Header="Event ID" Binding="{Binding EventId}" Width="70" />
        <DataGridTextColumn Header="Level" Binding="{Binding Level}" Width="80" />
        <DataGridTextColumn Header="Size(MB)" Binding="{Binding SizeMB}" Width="70" />
        <DataGridTextColumn Header="Message / File" Binding="{Binding Message}" Width="*" />
      </DataGrid.Columns>
    </DataGrid>

    <!-- Splitter -->
    <GridSplitter Grid.Row="2"
                  Height="5"
                  HorizontalAlignment="Stretch"
                  VerticalAlignment="Center"
                  Background="Gray"
                  ShowsPreview="True"
                  ToolTip="Drag to resize the space between the results grid and the details panel." />

    <!-- Details / Summary / Snapshot -->
    <GroupBox Grid.Row="3" Header="Details / Summary / System Info" Margin="0,4,0,6">
      <Grid>
        <TabControl x:Name="tabDetails" Margin="0">
          <!-- Details tab -->
          <TabItem Header="Details">
            <TextBox
                x:Name="txtDetails"
                Margin="4"
                TextWrapping="Wrap"
                VerticalScrollBarVisibility="Auto"
                HorizontalScrollBarVisibility="Auto"
                AcceptsReturn="True"
                IsReadOnly="True"
                ToolTip="Full details/XML for the selected event or file entry." />
          </TabItem>

          <!-- Summary tab -->
          <TabItem Header="Summary">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <RowDefinition Height="*" />
              </Grid.RowDefinitions>

              <Grid Grid.Row="0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*" />
                  <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,4,4,4">
                  <Button x:Name="btnExportSummary"
                          Content="Export Summary"
                          Width="120"
                          Margin="0,0,6,0"
                          ToolTip="Export the current summary text to a file." />
                  <Button x:Name="btnExportFull"
                          Content="Full Report"
                          Width="100"
                          Margin="0,0,6,0"
                          ToolTip="Export system snapshot + summary as a combined report." />
                  <Button x:Name="btnRefreshSummary"
                          Content="Refresh Summary"
                          Width="120"
                          ToolTip="Recalculate summary based on the currently visible rows in the grid." />
                </StackPanel>
              </Grid>

              <TextBox
                  x:Name="txtSummary"
                  Grid.Row="1"
                  Margin="4"
                  TextWrapping="Wrap"
                  VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Auto"
                  AcceptsReturn="True"
                  IsReadOnly="True"
                  ToolTip="Aggregated view: counts by type, level, top Event IDs and sources." />
            </Grid>
          </TabItem>

          <!-- System snapshot tab -->
          <TabItem Header="System Snapshot">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <RowDefinition Height="*" />
              </Grid.RowDefinitions>

              <Grid Grid.Row="0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*" />
                  <ColumnDefinition Width="Auto" />
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Orientation="Horizontal" Margin="4,4,4,4">
                  <CheckBox x:Name="cbShowNicDetails"
                            Content="Show NIC section"
                            IsChecked="True"
                            Margin="0,0,10,0"
                            ToolTip="Toggle visibility of the network interface section in the snapshot." />
                  <TextBlock Text="NIC filter:" VerticalAlignment="Center" Margin="0,0,4,0" />
                  <CheckBox x:Name="cbNicPhysical"
                            Content="Physical"
                            IsChecked="True"
                            Margin="0,0,4,0"
                            ToolTip="Include physical adapters (onboard / PCIe NICs)." />
                  <CheckBox x:Name="cbNicVirtual"
                            Content="Virtual"
                            IsChecked="True"
                            Margin="0,0,4,0"
                            ToolTip="Include virtual adapters (vNICs, VPN, Hyper-V, Docker, etc.)." />
                  <CheckBox x:Name="cbNicActiveOnly"
                            Content="Active only"
                            Margin="0,0,4,0"
                            ToolTip="When checked, only show adapters that are currently connected/enabled." />
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,4,4,4">
                  <Button x:Name="btnExportSnapshot"
                          Content="Export Snapshot"
                          Width="120"
                          Margin="0,0,6,0"
                          ToolTip="Export the current system snapshot text to a file." />
                  <Button x:Name="btnRefreshSnapshot"
                          Content="Refresh Snapshot"
                          Width="120"
                          ToolTip="Gather fresh system information (OS, uptime, disks, NICs) for this machine." />
                </StackPanel>
              </Grid>

              <TextBox
                  x:Name="txtSnapshot"
                  Grid.Row="1"
                  Margin="4"
                  TextWrapping="Wrap"
                  VerticalScrollBarVisibility="Auto"
                  HorizontalScrollBarVisibility="Auto"
                  AcceptsReturn="True"
                  IsReadOnly="True"
                  ToolTip="Summary of local system information, useful context when reviewing logs." />
            </Grid>
          </TabItem>
        </TabControl>
      </Grid>
    </GroupBox>

    <!-- Status bar -->
    <StatusBar Grid.Row="4">
      <StatusBarItem>
        <TextBlock x:Name="lblStatus" Text="Ready." />
      </StatusBarItem>
      <StatusBarItem HorizontalAlignment="Right">
        <ProgressBar
            x:Name="pbScan"
            Width="150"
            Height="14"
            IsIndeterminate="False"
            Visibility="Collapsed"
            ToolTip="Scan progress (indeterminate while scanning)." />
      </StatusBarItem>
    </StatusBar>

    <!-- Loading Overlay -->
    <Border x:Name="LoadingOverlay"
            Background="#AA000000"
            Visibility="Collapsed"
            Grid.RowSpan="5"
            HorizontalAlignment="Stretch"
            VerticalAlignment="Stretch">
      <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
        <TextBlock Text="Loading..."
                   Foreground="White"
                   FontSize="20"
                   HorizontalAlignment="Center"
                   Margin="0,0,0,10"/>
        <ProgressBar IsIndeterminate="True"
                     Width="200"
                     Height="18"/>
      </StackPanel>
    </Border>

  </Grid>
</Window>
"@

# -----------------------------
# Build window
# -----------------------------
$reader = New-Object System.Xml.XmlNodeReader($xaml)
try { $window = [Windows.Markup.XamlReader]::Load($reader) }
catch { Write-Error "Failed to load XAML: $_"; return }

# Controls
$cbSource       = $window.FindName("cbSource")
$cbProfile      = $window.FindName("cbProfile")
$txtComputer    = $window.FindName("txtComputer")

$txtDays        = $window.FindName("txtDays")
$txtEventId     = $window.FindName("txtEventId")
$txtMaxEvents   = $window.FindName("txtMaxEvents")
$chkSystem      = $window.FindName("chkSystem")
$chkApplication = $window.FindName("chkApplication")
$chkError       = $window.FindName("chkError")
$chkWarning     = $window.FindName("chkWarning")
$chkInfo        = $window.FindName("chkInfo")

$btnScan          = $window.FindName("btnScan")
$txtSearch        = $window.FindName("txtSearch")
$btnSearch        = $window.FindName("btnSearch")
$btnClearSearch   = $window.FindName("btnClearSearch")
$btnExport        = $window.FindName("btnExport")
$btnSupportBundle = $window.FindName("btnSupportBundle")
$btnAiAnalyze     = $window.FindName("btnAiAnalyze")

$dgResults      = $window.FindName("dgResults")
$txtDetails     = $window.FindName("txtDetails")
$txtSummary     = $window.FindName("txtSummary")
$btnRefreshSummary = $window.FindName("btnRefreshSummary")
$btnExportSummary  = $window.FindName("btnExportSummary")
$btnExportFull     = $window.FindName("btnExportFull")

$txtSnapshot        = $window.FindName("txtSnapshot")
$btnRefreshSnapshot = $window.FindName("btnRefreshSnapshot")
$btnExportSnapshot  = $window.FindName("btnExportSnapshot")
$cbShowNicDetails   = $window.FindName("cbShowNicDetails")
$cbNicPhysical      = $window.FindName("cbNicPhysical")
$cbNicVirtual       = $window.FindName("cbNicVirtual")
$cbNicActiveOnly    = $window.FindName("cbNicActiveOnly")

$lblStatus       = $window.FindName("lblStatus")
$miOpenFolder    = $window.FindName("miOpenFolder")
$miCopyDetails   = $window.FindName("miCopyDetails")
$pbScan          = $window.FindName("pbScan")
$loadingOverlay  = $window.FindName("LoadingOverlay")

# Populate source combo
$cbSource.Items.Clear()
$null = $cbSource.Items.Add("Event Logs")
$null = $cbSource.Items.Add("Crash Dumps")
$null = $cbSource.Items.Add("Custom Folder (logs)")
$cbSource.SelectedIndex = 0

# State
$script:CurrentResults = @()

# -----------------------------
# Loading overlay helpers
# -----------------------------
function Show-Loading {
    if ($loadingOverlay) {
        $loadingOverlay.Visibility = "Visible"
        $window.Cursor = "Wait"
        # Force UI to render the overlay before running heavy work
        $window.Dispatcher.Invoke([Action]{}, "Background")
    }
}

function Hide-Loading {
    if ($loadingOverlay) {
        $loadingOverlay.Visibility = "Collapsed"
        $window.Cursor = "Arrow"
    }
}

# -----------------------------
# Summary builder
# -----------------------------
function Update-Summary {
    try {
        $items = @($dgResults.ItemsSource)
        if (-not $items -or $items.Count -eq 0) {
            $txtSummary.Text = "No data loaded. Run a scan first."
            return
        }

        $sb = New-Object System.Text.StringBuilder

        [void]$sb.AppendLine("=== SUMMARY (Based on current grid view) ===")
        [void]$sb.AppendLine("Generated: $(Get-Date)")
        [void]$sb.AppendLine("Total items: $($items.Count)")

        $times = $items | Where-Object { $_.TimeCreated } | Select-Object -ExpandProperty TimeCreated
        if ($times -and $times.Count -gt 0) {
            $min = ($times | Measure-Object -Minimum).Minimum
            $max = ($times | Measure-Object -Maximum).Maximum
            [void]$sb.AppendLine("Time range: $min  ->  $max")
        }
        [void]$sb.AppendLine()

        [void]$sb.AppendLine("=== COUNTS BY TYPE ===")
        $items | Group-Object Type | Sort-Object Name | ForEach-Object {
            $typeName = if ([string]::IsNullOrWhiteSpace($_.Name)) { "<null>" } else { $_.Name }
            [void]$sb.AppendLine("  $typeName : $($_.Count)")
        }
        [void]$sb.AppendLine()

        $events = $items | Where-Object { $_.Type -eq 'Event' }
        if ($events -and $events.Count -gt 0) {
            [void]$sb.AppendLine("=== EVENT LOG BREAKDOWN ===")
            [void]$sb.AppendLine("  By Level:")
            $events | Group-Object Level | Sort-Object Name | ForEach-Object {
                $levelName = if ([string]::IsNullOrWhiteSpace($_.Name)) { "<unknown>" } else { $_.Name }
                [void]$sb.AppendLine("    $levelName : $($_.Count)")
            }
            [void]$sb.AppendLine()

            [void]$sb.AppendLine("  Top Event IDs (by count):")
            $events |
                Where-Object { $_.EventId -ne $null -and $_.EventId -ne "" } |
                Group-Object EventId |
                Sort-Object Count -Descending |
                Select-Object -First 10 |
                ForEach-Object {
                    $id   = $_.Name
                    $cnt  = $_.Count
                    $logs = ($events | Where-Object { $_.EventId -eq $id } |
                             Select-Object -ExpandProperty LogName -Unique) -join ", "
                    [void]$sb.AppendLine("    ID $id  Count $cnt  Logs: $logs")
                }
            [void]$sb.AppendLine()

            [void]$sb.AppendLine("  Top Sources (by count):")
            $events |
                Group-Object Source |
                Sort-Object Count -Descending |
                Select-Object -First 10 |
                ForEach-Object {
                    $srcName = if ([string]::IsNullOrWhiteSpace($_.Name)) { "<unknown>" } else { $_.Name }
                    [void]$sb.AppendLine("    $srcName : $($_.Count)")
                }
            [void]$sb.AppendLine()
        }

        $nonEvents = $items | Where-Object { $_.Type -ne 'Event' }
        if ($nonEvents -and $nonEvents.Count -gt 0) {
            [void]$sb.AppendLine("=== NON-EVENT ITEMS ===")
            $nonEvents | Group-Object Type | ForEach-Object {
                $typeName = if ([string]::IsNullOrWhiteSpace($_.Name)) { "<null>" } else { $_.Name }
                [void]$sb.AppendLine("  $typeName : $($_.Count)")
            }
            [void]$sb.AppendLine()
        }

        $txtSummary.Text = $sb.ToString()
    }
    catch {
        $txtSummary.Text = "Failed to generate summary:`r`n$($_.Exception.Message)"
    }
}

# -----------------------------
# Show results helper
# -----------------------------
function Show-Results {
    param([System.Collections.IEnumerable]$Items)

    $script:CurrentResults = @($Items)
    $dgResults.ItemsSource = $script:CurrentResults
    $dgResults.Items.Refresh()

    $lblStatus.Text = "Results: $($script:CurrentResults.Count) item(s)."
    Update-Summary
}

# -----------------------------
# Event log scan (synchronous with overlay + progress bar)
# -----------------------------
function Scan-EventLogs {
    try {
        $daysBack = 7
        [void][int]::TryParse($txtDays.Text, [ref]$daysBack)
        if ($daysBack -le 0) { $daysBack = 7 }

        $maxEvents = 1000
        [void][int]::TryParse($txtMaxEvents.Text, [ref]$maxEvents)
        if ($maxEvents -le 0) { $maxEvents = 1000 }

        $levels = @()
        if ($chkError.IsChecked)   { $levels += 2 }
        if ($chkWarning.IsChecked) { $levels += 3 }
        if ($chkInfo.IsChecked)    { $levels += 4 }
        if (-not $levels) { $levels = 2,3 }

        $logs = @()
        if ($chkSystem.IsChecked)      { $logs += "System" }
        if ($chkApplication.IsChecked) { $logs += "Application" }
        if (-not $logs) { $logs = "System","Application" }

        $filter = @{
            LogName   = $logs
            Level     = $levels
            StartTime = (Get-Date).AddDays(-$daysBack)
        }

        $idText = $txtEventId.Text.Trim()
        if ($idText) {
            $ids = $idText -split '[,\s]+' | Where-Object { $_ -match '^\d+$' }
            if ($ids.Count -gt 0) { $filter["Id"] = $ids }
        }

        $computerName = $txtComputer.Text
        if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }

        $lblStatus.Text      = "Querying event logs on $computerName..."
        $btnScan.IsEnabled   = $false
        $cbSource.IsEnabled  = $false
        if ($pbScan) {
            $pbScan.IsIndeterminate = $true
            $pbScan.Visibility      = "Visible"
        }
        Show-Loading

        if ($computerName -eq "." -or $computerName -eq "localhost") {
            $events = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop |
                      Select-Object -First $maxEvents | ForEach-Object {
                [PSCustomObject]@{
                    Type        = "Event"
                    TimeCreated = $_.TimeCreated
                    LogName     = $_.LogName
                    Source      = $_.ProviderName
                    EventId     = $_.Id
                    Level       = $_.LevelDisplayName
                    Message     = $_.Message
                    Path        = $null
                    SizeMB      = $null
                    Details     = $_.ToXml()
                }
            }
        }
        else {
            $events = Get-WinEvent -ComputerName $computerName -FilterHashtable $filter -ErrorAction Stop |
                      Select-Object -First $maxEvents | ForEach-Object {
                [PSCustomObject]@{
                    Type        = "Event"
                    TimeCreated = $_.TimeCreated
                    LogName     = $_.LogName
                    Source      = $_.ProviderName
                    EventId     = $_.Id
                    Level       = $_.LevelDisplayName
                    Message     = $_.Message
                    Path        = $null
                    SizeMB      = $null
                    Details     = $_.ToXml()
                }
            }
        }

        Show-Results -Items $events
        $lblStatus.Text = "Event log scan completed. $($script:CurrentResults.Count) item(s)."
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to read event logs:`r`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        $lblStatus.Text = "Error querying event logs."
    }
    finally {
        if ($pbScan) {
            $pbScan.IsIndeterminate = $false
            $pbScan.Visibility      = "Collapsed"
        }
        $btnScan.IsEnabled   = $true
        $cbSource.IsEnabled  = $true
        Hide-Loading
    }
}

# -----------------------------
# Crash dumps (synchronous with overlay + progress bar)
# -----------------------------
function Scan-CrashDumps {
    $lblStatus.Text      = "Scanning for crash dumps..."
    $btnScan.IsEnabled   = $false
    $cbSource.IsEnabled  = $false
    if ($pbScan) {
        $pbScan.IsIndeterminate = $true
        $pbScan.Visibility      = "Visible"
    }
    Show-Loading

    try {
        $pathsToCheck = @(
            (Join-Path $env:SystemRoot "Minidump"),
            (Join-Path $env:SystemRoot "MEMORY.DMP")
        )

        $items = @()

        foreach ($path in $pathsToCheck) {
            if (-not (Test-Path $path)) { continue }

            if (Test-Path $path -PathType Container) {
                $files = Get-ChildItem -Path $path -Filter "*.dmp" -ErrorAction SilentlyContinue
            }
            else {
                $files = Get-Item -Path $path -ErrorAction SilentlyContinue
            }

            foreach ($file in $files) {
                $sizeMB = [math]::Round($file.Length / 1MB, 2)
                $items += [PSCustomObject]@{
                    Type        = "CrashDump"
                    TimeCreated = $file.LastWriteTime
                    LogName     = "Dump"
                    Source      = $file.DirectoryName
                    EventId     = ""
                    Level       = ""
                    Message     = $file.Name
                    Path        = $file.FullName
                    SizeMB      = $sizeMB
                    Details     = "Path: $($file.FullName)`r`nSize (MB): $sizeMB"
                }
            }
        }

        Show-Results -Items $items
        $lblStatus.Text = "Crash dump scan completed. $($script:CurrentResults.Count) item(s)."
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to enumerate crash dumps:`r`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        $lblStatus.Text = "Error scanning crash dumps."
    }
    finally {
        if ($pbScan) {
            $pbScan.IsIndeterminate = $false
            $pbScan.Visibility      = "Collapsed"
        }
        $btnScan.IsEnabled   = $true
        $cbSource.IsEnabled  = $true
        Hide-Loading
    }
}

# -----------------------------
# Custom folder scan (with overlay)
# -----------------------------
function Scan-CustomFolder {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select root folder to scan for log files (*.log, *.txt, *.evtx)"

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $folder = $dialog.SelectedPath
    $lblStatus.Text = "Scanning folder: $folder"
    Show-Loading

    try {
        $files = Get-ChildItem -Path $folder -Recurse -Include *.log,*.txt,*.evtx -ErrorAction SilentlyContinue

        $items = foreach ($file in $files) {
            $sizeMB = [math]::Round($file.Length / 1MB, 2)
            [PSCustomObject]@{
                Type        = "File"
                TimeCreated = $file.LastWriteTime
                LogName     = "File"
                Source      = $file.DirectoryName
                EventId     = ""
                Level       = ""
                Message     = $file.Name
                Path        = $file.FullName
                SizeMB      = $sizeMB
                Details     = "Path: $($file.FullName)`r`nSize (MB): $sizeMB"
            }
        }

        Show-Results -Items $items
        $lblStatus.Text = "Folder scan completed. $($script:CurrentResults.Count) item(s)."
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to scan folder:`r`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        $lblStatus.Text = "Error scanning folder."
    }
    finally {
        Hide-Loading
    }
}

# -----------------------------
# Search / filter
# -----------------------------
function Apply-SearchFilter {
    $keyword = $txtSearch.Text
    if ([string]::IsNullOrWhiteSpace($keyword)) {
        $dgResults.ItemsSource = $script:CurrentResults
        $dgResults.Items.Refresh()
        $lblStatus.Text = "Filter cleared. Showing $($script:CurrentResults.Count) item(s)."
        Update-Summary
        return
    }

    $kw = $keyword.ToLowerInvariant()
    $filtered = $script:CurrentResults | Where-Object {
        $src  = [string]$_.Source
        $msg  = [string]$_.Message
        $path = [string]$_.Path
        $log  = [string]$_.LogName
        ($src + " " + $msg + " " + $path + " " + $log).ToLowerInvariant().Contains($kw)
    }

    $dgResults.ItemsSource = $filtered
    $dgResults.Items.Refresh()
    $lblStatus.Text = "Filter applied: '$keyword'. Showing $($filtered.Count) item(s)."
    Update-Summary
}

# -----------------------------
# Export results CSV
# -----------------------------
function Export-ResultsCsv {
    if (-not $script:CurrentResults -or $script:CurrentResults.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No results to export.",
            "Export",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "CSV files (*.csv)|*.csv"
    $dialog.Title  = "Export results to CSV"
    $dialog.FileName = "logscan-$([DateTime]::Now.ToString('yyyyMMdd-HHmmss')).csv"

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    try {
        $script:CurrentResults | Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8
        [System.Windows.MessageBox]::Show(
            "Exported $($script:CurrentResults.Count) item(s) to:`r`n$($dialog.FileName)",
            "Export",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to export CSV:`r`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    }
}

# -----------------------------
# Support bundle (with overlay)
# -----------------------------
function New-SupportBundle {
    if (-not $script:CurrentResults -or $script:CurrentResults.Count -eq 0) {
        $result = [System.Windows.MessageBox]::Show(
            "No results loaded in the grid. Generate bundle anyway (logs + system info)?",
            "Support Bundle",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Question
        )
        if ($result -ne [System.Windows.MessageBoxResult]::Yes) { return }
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "ZIP files (*.zip)|*.zip"
    $dialog.Title  = "Save Support Bundle"
    $dialog.FileName = "SupportBundle-$([DateTime]::Now.ToString('yyyyMMdd-HHmmss')).zip"

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $zipPath = $dialog.FileName
    $workRoot = Join-Path $env:TEMP ("LogScanner_Bundle_" + [Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $workRoot -Force | Out-Null

    $lblStatus.Text = "Generating support bundle..."
    $btnSupportBundle.IsEnabled = $false
    $btnScan.IsEnabled = $false
    Show-Loading

    try {
        $computerName = $txtComputer.Text
        if ([string]::IsNullOrWhiteSpace($computerName) -or $computerName -eq "." -or $computerName -eq "localhost") {
            $logsToExport = @()
            if ($chkSystem.IsChecked)      { $logsToExport += "System" }
            if ($chkApplication.IsChecked) { $logsToExport += "Application" }

            foreach ($log in $logsToExport) {
                $evtxOut = Join-Path $workRoot "$log.evtx"
                wevtutil epl $log $evtxOut 2>$null
            }
        }

        if ($script:CurrentResults -and $script:CurrentResults.Count -gt 0) {
            $csvOut = Join-Path $workRoot "LogScannerResults.csv"
            $script:CurrentResults | Export-Csv -Path $csvOut -NoTypeInformation -Encoding UTF8
        }

        $sysInfoOut = Join-Path $workRoot "systeminfo.txt"
        cmd.exe /c "systeminfo" | Out-File -FilePath $sysInfoOut -Encoding UTF8

        $ipconfigOut = Join-Path $workRoot "ipconfig_all.txt"
        cmd.exe /c "ipconfig /all" | Out-File -FilePath $ipconfigOut -Encoding UTF8

        $procOut = Join-Path $workRoot "processes.txt"
        Get-Process | Sort-Object -Property CPU -Descending |
            Select-Object -First 50 Name,Id,CPU,PM,WS |
            Format-Table -AutoSize | Out-String |
            Out-File -FilePath $procOut -Encoding UTF8

        if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
        Compress-Archive -Path (Join-Path $workRoot "*") -DestinationPath $zipPath

        [System.Windows.MessageBox]::Show(
            "Support bundle created:`r`n$zipPath",
            "Support Bundle",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        $lblStatus.Text = "Support bundle created."
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to create support bundle:`r`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
        $lblStatus.Text = "Error creating support bundle."
    }
    finally {
        if (Test-Path $workRoot) {
            Remove-Item $workRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
        $btnSupportBundle.IsEnabled = $true
        $btnScan.IsEnabled = $true
        Hide-Loading
    }
}

# -----------------------------
# System snapshot (with NIC filters + overlay)
# -----------------------------
function Update-SystemSnapshot {
    $lblStatus.Text = "Refreshing system snapshot..."
    Show-Loading

    try {
        $os   = Get-CimInstance -ClassName Win32_OperatingSystem
        $comp = Get-CimInstance -ClassName Win32_ComputerSystem
        $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction SilentlyContinue

        $uptime = (Get-Date) - $os.LastBootUpTime

        $drives = Get-PSDrive -PSProvider FileSystem | ForEach-Object {
            try {
                $freeGB     = [math]::Round($_.Free / 1GB, 2)
                $totalBytes = $_.Used + $_.Free
                $totalGB    = if ($totalBytes -gt 0) { [math]::Round($totalBytes / 1GB, 2) } else { 0 }
                $usedPct    = if ($totalGB -eq 0) { 0 } else { [math]::Round((1 - ($_.Free / $totalBytes)) * 100, 1) }
                "{0}:  Free {1} GB  (Used {2}% of ~{3} GB)" -f $_.Name.TrimEnd(":"), $freeGB, $usedPct, $totalGB
            }
            catch {
                "{0}:  (unable to query disk usage)" -f $_.Name
            }
        }

        $sb = New-Object System.Text.StringBuilder

        [void]$sb.AppendLine("=== SYSTEM SNAPSHOT (LOCAL MACHINE) ===")
        [void]$sb.AppendLine("Generated: $(Get-Date)")
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("Computer Name : $($env:COMPUTERNAME)")
        [void]$sb.AppendLine("User Name     : $($env:USERNAME)")
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("OS Caption    : $($os.Caption)")
        [void]$sb.AppendLine("OS Version    : $($os.Version)")
        [void]$sb.AppendLine("Build Number  : $($os.BuildNumber)")
        [void]$sb.AppendLine("Install Date  : $($os.InstallDate)")
        [void]$sb.AppendLine("Last Boot     : $($os.LastBootUpTime)")
        [void]$sb.AppendLine(("Uptime        : {0:dd} days {0:hh} hours {0:mm} minutes" -f $uptime))
        [void]$sb.AppendLine()
        [void]$sb.AppendLine("Manufacturer  : $($comp.Manufacturer)")
        [void]$sb.AppendLine("Model         : $($comp.Model)")
        [void]$sb.AppendLine("Total RAM     : {0} GB" -f ([math]::Round($comp.TotalPhysicalMemory / 1GB, 2)))
        if ($bios) { [void]$sb.AppendLine("BIOS Version  : $($bios.SMBIOSBIOSVersion)") }
        [void]$sb.AppendLine()

        [void]$sb.AppendLine("=== DISKS ===")
        $drives | ForEach-Object { [void]$sb.AppendLine("  " + $_) }
        [void]$sb.AppendLine()

        [void]$sb.AppendLine("=== NETWORK INTERFACES (ALL) ===")

        if ($cbShowNicDetails -and $cbShowNicDetails.IsChecked -ne $true) {
            [void]$sb.AppendLine("  (NIC section hidden. Enable 'Show NIC section' to view details.)")
        }
        else {
            try {
                $adapters = Get-CimInstance Win32_NetworkAdapter |
                            Where-Object { $_.AdapterTypeID -ne $null -or $_.PhysicalAdapter -or $_.NetConnectionStatus -ne $null }

                $configs  = Get-CimInstance Win32_NetworkAdapterConfiguration

                $totalCount   = $adapters.Count
                $physCount    = ($adapters | Where-Object { $_.PhysicalAdapter }).Count
                $virtCount    = $totalCount - $physCount
                $activeCount  = ($adapters | Where-Object { $_.NetEnabled -eq $true -and $_.NetConnectionStatus -eq 2 }).Count

                [void]$sb.AppendLine("  Total adapters : $totalCount  (Physical: $physCount, Virtual/Other: $virtCount, Active: $activeCount)")
                [void]$sb.AppendLine()

                $showPhysical = $cbNicPhysical -and $cbNicPhysical.IsChecked
                $showVirtual  = $cbNicVirtual  -and $cbNicVirtual.IsChecked
                $activeOnly   = $cbNicActiveOnly -and $cbNicActiveOnly.IsChecked

                $printedAny = $false

                foreach ($nic in $adapters) {
                    $isPhysical = $nic.PhysicalAdapter -eq $true
                    $isActive   = ($nic.NetEnabled -eq $true -and $_.NetConnectionStatus -eq 2) # bug, fix: use $nic

                }
            }
            catch {
                [void]$sb.AppendLine("  Unable to query network adapters: $($_.Exception.Message)")
                [void]$sb.AppendLine()
            }
        }

        # Fix the foreach NIC loop (correct version)
        $sbNic = New-Object System.Text.StringBuilder
        try {
            $adapters = Get-CimInstance Win32_NetworkAdapter |
                        Where-Object { $_.AdapterTypeID -ne $null -or $_.PhysicalAdapter -or $_.NetConnectionStatus -ne $null }
            $configs  = Get-CimInstance Win32_NetworkAdapterConfiguration

            $totalCount   = $adapters.Count
            $physCount    = ($adapters | Where-Object { $_.PhysicalAdapter }).Count
            $virtCount    = $totalCount - $physCount
            $activeCount  = ($adapters | Where-Object { $_.NetEnabled -eq $true -and $_.NetConnectionStatus -eq 2 }).Count

            [void]$sbNic.AppendLine("  Total adapters : $totalCount  (Physical: $physCount, Virtual/Other: $virtCount, Active: $activeCount)")
            [void]$sbNic.AppendLine()

            $showPhysical = $cbNicPhysical -and $cbNicPhysical.IsChecked
            $showVirtual  = $cbNicVirtual  -and $cbNicVirtual.IsChecked
            $activeOnly   = $cbNicActiveOnly -and $cbNicActiveOnly.IsChecked

            $printedAny = $false

            foreach ($nic in $adapters) {
                $isPhysical = $nic.PhysicalAdapter -eq $true
                $isActive   = ($nic.NetEnabled -eq $true -and $nic.NetConnectionStatus -eq 2)

                if ($isPhysical -and -not $showPhysical) { continue }
                if (-not $isPhysical -and -not $showVirtual) { continue }
                if ($activeOnly -and -not $isActive) { continue }

                $cfg = $configs | Where-Object { $_.Index -eq $nic.Index }

                $name = if ($nic.NetConnectionID) { $nic.NetConnectionID } else { $nic.Name }
                $mac  = $nic.MACAddress

                $status = switch ($nic.NetConnectionStatus) {
                    0 { "Disconnected" }
                    1 { "Connecting" }
                    2 { "Connected" }
                    3 { "Disconnecting" }
                    7 { "HardwareDisabled" }
                    9 { "Disabled" }
                    default { "Unknown" }
                }

                $speed = if ($nic.Speed) {
                    [string]::Format("{0:0.##} Gb", ($nic.Speed / 1Gb))
                } else { "N/A" }

                $ipv4 = @()
                $ipv6 = @()
                if ($cfg -and $cfg.IPAddress) {
                    foreach ($ip in $cfg.IPAddress) {
                        if ($ip -like "*.*") { $ipv4 += $ip }
                        elseif ($ip -like "*:*") { $ipv6 += $ip }
                    }
                }

                $ipv4txt = if ($ipv4.Count -gt 0) { $ipv4 -join ", " } else { "<none>" }
                $ipv6txt = if ($ipv6.Count -gt 0) { $ipv6 -join ", " } else { "<none>" }
                $dhcp    = if ($cfg -and $cfg.DHCPEnabled) { "Yes" } else { "No" }

                [void]$sbNic.AppendLine("  $name")
                [void]$sbNic.AppendLine("    Status     : $status")
                [void]$sbNic.AppendLine("    MAC        : $mac")
                [void]$sbNic.AppendLine("    Speed      : $speed")
                [void]$sbNic.AppendLine("    DHCP       : $dhcp")
                [void]$sbNic.AppendLine("    IPv4       : $ipv4txt")
                [void]$sbNic.AppendLine("    IPv6       : $ipv6txt")
                [void]$sbNic.AppendLine()

                $printedAny = $true
            }

            if (-not $printedAny) {
                [void]$sbNic.AppendLine("  (No adapters match the current NIC filter settings.)")
            }
        }
        catch {
            [void]$sbNic.AppendLine("  Unable to query network adapters: $($_.Exception.Message)")
            [void]$sbNic.AppendLine()
        }

        $txtSnapshot.Text = $sb.ToString() + "`r`n" + $sbNic.ToString()
        $lblStatus.Text   = "System snapshot refreshed."
    }
    catch {
        $txtSnapshot.Text = "Failed to gather system snapshot:`r`n$($_.Exception.Message)"
        $lblStatus.Text   = "Error refreshing system snapshot."
    }
    finally {
        Hide-Loading
    }
}

# -----------------------------
# Exports for snapshot / summary / full
# -----------------------------
function Export-Snapshot {
    if ([string]::IsNullOrWhiteSpace($txtSnapshot.Text)) {
        [System.Windows.MessageBox]::Show(
            "No snapshot data to export. Refresh snapshot first.",
            "Export Snapshot",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Text files (*.txt)|*.txt"
    $dialog.Title  = "Export system snapshot"
    $dialog.FileName = "snapshot-$([DateTime]::Now.ToString('yyyyMMdd-HHmmss')).txt"

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $txtSnapshot.Text | Out-File -FilePath $dialog.FileName -Encoding UTF8
}

function Export-Summary {
    if ([string]::IsNullOrWhiteSpace($txtSummary.Text)) {
        [System.Windows.MessageBox]::Show(
            "No summary data to export. Refresh summary first.",
            "Export Summary",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Text files (*.txt)|*.txt"
    $dialog.Title  = "Export summary"
    $dialog.FileName = "summary-$([DateTime]::Now.ToString('yyyyMMdd-HHmmss')).txt"

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $txtSummary.Text | Out-File -FilePath $dialog.FileName -Encoding UTF8
}

function Export-FullReport {
    if ([string]::IsNullOrWhiteSpace($txtSnapshot.Text) -or [string]::IsNullOrWhiteSpace($txtSummary.Text)) {
        [System.Windows.MessageBox]::Show(
            "Snapshot and/or summary are empty. Refresh both before exporting?",
            "Full Report",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning
        ) | Out-Null
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "Text files (*.txt)|*.txt"
    $dialog.Title  = "Export full report"
    $dialog.FileName = "report-$([DateTime]::Now.ToString('yyyyMMdd-HHmmss')).txt"

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("=== SYSTEM SNAPSHOT ===")
    [void]$sb.AppendLine($txtSnapshot.Text)
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("=== SUMMARY ===")
    [void]$sb.AppendLine($txtSummary.Text)

    $sb.ToString() | Out-File -FilePath $dialog.FileName -Encoding UTF8
}

# -----------------------------
# AI placeholder
# -----------------------------
function Invoke-AIAnalysis {
    if (-not $script:CurrentResults -or $script:CurrentResults.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "No data loaded. Run a scan first.",
            "AI Analysis",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        ) | Out-Null
        return
    }

    [System.Windows.MessageBox]::Show(
        "AI analysis is not yet implemented.`r`nThis is the hook function to extend.",
        "AI Integration",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Information
    ) | Out-Null
}

# -----------------------------
# Profiles
# -----------------------------
$cbProfile.Add_SelectionChanged({
    $selection = $cbProfile.SelectedItem
    if ($null -eq $selection) { return }

    switch ($selection.Content) {
        "BSOD / Crash" {
            $txtDays.Text     = "3"
            $chkSystem.IsChecked      = $true
            $chkApplication.IsChecked = $true
            $chkError.IsChecked       = $true
            $chkWarning.IsChecked     = $true
            $chkInfo.IsChecked        = $false
            $txtEventId.Text  = "41, 1001, 6008"
        }
        "App Crashes" {
            $txtDays.Text     = "7"
            $chkSystem.IsChecked      = $false
            $chkApplication.IsChecked = $true
            $chkError.IsChecked       = $true
            $chkWarning.IsChecked     = $true
            $chkInfo.IsChecked        = $false
            $txtEventId.Text  = "1000, 1001"
        }
        "Boot / Startup Issues" {
            $txtDays.Text     = "7"
            $chkSystem.IsChecked      = $true
            $chkApplication.IsChecked = $false
            $chkError.IsChecked       = $true
            $chkWarning.IsChecked     = $true
            $chkInfo.IsChecked        = $false
            $txtEventId.Text  = "6005, 6006, 6008, 7000, 7001, 7002, 7009"
        }
        default { }
    }
})

# -----------------------------
# Wire-up events
# -----------------------------
$btnScan.Add_Click({
    switch ($cbSource.SelectedItem) {
        "Event Logs"           { Scan-EventLogs }
        "Crash Dumps"          { Scan-CrashDumps }
        "Custom Folder (logs)" { Scan-CustomFolder }
        default                { Scan-EventLogs }
    }
})

$btnSearch.Add_Click({ Apply-SearchFilter })
$btnClearSearch.Add_Click({ $txtSearch.Text = ""; Apply-SearchFilter })

$btnExport.Add_Click({ Export-ResultsCsv })
$btnSupportBundle.Add_Click({ New-SupportBundle })
$btnAiAnalyze.Add_Click({ Invoke-AIAnalysis })

$btnRefreshSnapshot.Add_Click({ Update-SystemSnapshot })
$btnExportSnapshot.Add_Click({ Export-Snapshot })

$btnRefreshSummary.Add_Click({ Update-Summary })
$btnExportSummary.Add_Click({ Export-Summary })
$btnExportFull.Add_Click({ Export-FullReport })

$dgResults.Add_SelectionChanged({
    $item = $dgResults.SelectedItem
    if ($null -ne $item) {
        if ($item.Details) {
            $txtDetails.Text = $item.Details
        } elseif ($item.Message) {
            $txtDetails.Text = $item.Message
        } else {
            $txtDetails.Text = ($item | Format-List * | Out-String)
        }
    }
})

$miOpenFolder.Add_Click({
    $item = $dgResults.SelectedItem
    if ($null -eq $item) { return }
    if ([string]::IsNullOrWhiteSpace($item.Path)) { return }
    try { Start-Process explorer.exe "/select,`"$($item.Path)`"" }
    catch {
        [System.Windows.MessageBox]::Show(
            "Failed to open folder:`r`n$($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        ) | Out-Null
    }
})

$miCopyDetails.Add_Click({
    $item = $dgResults.SelectedItem
    if ($null -eq $item) { return }
    $text = if ($item.Details) { $item.Details }
            elseif ($item.Message) { $item.Message }
            else { ($item | Format-List * | Out-String) }
    [System.Windows.Clipboard]::SetText($text)
})

# Run initial summary/snapshot after window load so overlay can show
$window.Add_Loaded({
    $lblStatus.Text = "Initializing system snapshot..."
    Update-SystemSnapshot
    Update-Summary
    $lblStatus.Text = "Ready. Select a source/profile and click Scan."
})

# -----------------------------
# Show window
# -----------------------------
$window.ShowDialog() | Out-Null
