Local Log Scanner – PowerUser Edition
-------------------------------------

Overview
--------
Local Log Scanner is a Windows diagnostic and triage utility designed for
technicians, power users, and system administrators who need fast visibility
into system health. The tool provides a unified interface for gathering and
reviewing Windows Event Logs, crash dumps, custom log folders, and detailed
system information — all from a clean WPF GUI.

Key Features
------------
• Event Log Viewer  
  - Query System and Application logs  
  - Filter by days, severity, Event IDs, and source  
  - Search across message text, paths, and metadata  
  - Profiles for Crash, AppIssues, and BootTroubleshooting

• Crash Dump Scanner  
  - Automatically detects Minidumps and MEMORY.DMP  
  - Displays size, timestamp, and file metadata

• Custom Folder Log Scan  
  - Recursively scans *.log, *.txt, *.evtx  
  - Shows timestamps, sizes, and full paths

• Data Grid
  - Sortable columns  
  - Context menu: open file folder, copy full details  
  - Double-click exposes XML or message body

• Summary Engine  
  - Counts by log type  
  - Level breakdown (Error, Warning, Info)  
  - Top Event IDs  
  - Top Sources  
  - Derived exclusively from what’s currently visible in the grid

• System Snapshot  
  - OS version, build, uptime  
  - Disk usage summaries  
  - Full network adapter breakdown (physical/virtual, speeds, IPs)  
  - NIC filtering (physical / virtual / active-only)

• Export Options  
  - Export current results to CSV  
  - Export Summary or Snapshot  
  - Full Report (Snapshot + Summary)  
  - Technician Support Bundle (ZIP containing logs, systeminfo, ipconfig, results)

• UI/UX Enhancements  
  - Loading overlay + progress indicator during scans  
  - Non-blocking updates and status line messages  
  - Resizable interface with splitter  
  - Dynamic rendering of NIC sections

Requirements
-----------
• Windows 10 or Windows 11  
• PowerShell 5.1 or later  
• Administrator privileges recommended for full visibility  
• .NET Framework + WPF support (default on Win10/11)

Usage
-----
Run the script with PowerShell:

    powershell.exe -ExecutionPolicy Bypass -File LogScanner.ps1

Select your source, apply filters, and click **Scan**.  
Results load into the main grid and drive the Summary/Snapshot tabs.

--------------------------------------------------------
Created for internal diagnostics and power-user workflows.
Bogdan Fedko
--------------------------------------------------------
