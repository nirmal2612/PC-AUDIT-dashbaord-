# LAN Device Explorer - PowerShell WinForms
# Works on Windows PowerShell 5.1
# Purpose: Scan LAN IP range, ping devices, read ARP table, show IP/MAC/Hostname/Vendor, export CSV/Excel CSV.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Hide PowerShell console window
try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32Window {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
    $consolePtr = [Win32Window]::GetConsoleWindow()
    if ($consolePtr -ne [IntPtr]::Zero) { [Win32Window]::ShowWindow($consolePtr, 0) | Out-Null }
} catch {}

$script:StopScan = $false
$script:Results = New-Object System.Collections.Generic.List[object]
$script:OuiEntriesLoaded = 0
$script:OuiExternalFiles = @()
$script:OuiSourceMode = 'Built-in common IT vendor OUI map'

# Built-in IT/OUI vendor map.
# Exact global OUI database is very large. This script supports both:
# 1) Built-in common IT vendors below for fast offline scan.
# 2) Optional complete IEEE/Wireshark OUI database beside this script: oui.csv / oui.txt / manuf.txt.
#    Put the full file beside this PS1 and the scanner will load 50,000+ OUI entries offline.
#    Supported formats: IEEE oui.txt, Wireshark manuf.txt, CSV/text: OUI,Vendor.
$VendorMap = @{
    '00:A8:59' = 'Fanvil'
    '0C:38:3E' = 'Fanvil'
    '00:09:6E' = 'Fanvil'
    '3C:E5:A6' = 'Fanvil'
    '80:5E:C0' = 'Fanvil'
    '00:17:61' = 'ZKTeco / ZKSoftware'
    '00:25:45' = 'ZKTeco / Access Control'
    'E8:65:49' = 'ZKTeco / Access Control'
    'FC:2F:40' = 'ZKTeco / Access Control'
    'B4:4C:3B' = 'ZKTeco / Access Control'
    '00:23:63' = 'Savior / Access Control'
    '00:13:8F' = 'Access Control Device'
    '00:14:22' = 'Dell'
    '00:18:8B' = 'Dell'
    '00:19:B9' = 'Dell'
    '00:1A:A0' = 'Dell'
    '00:21:70' = 'Dell'
    '00:23:AE' = 'Dell'
    '00:24:E8' = 'Dell'
    '00:26:B9' = 'Dell'
    '14:18:77' = 'Dell'
    '18:03:73' = 'Dell'
    '34:17:EB' = 'Dell'
    '3C:2C:30' = 'Dell'
    '44:A8:42' = 'Dell'
    '54:BF:64' = 'Dell'
    '5C:26:0A' = 'Dell'
    '74:86:E2' = 'Dell'
    '78:2B:CB' = 'Dell'
    '90:B1:1C' = 'Dell'
    'B8:2A:72' = 'Dell'
    'B8:AC:6F' = 'Dell'
    'D0:94:66' = 'Dell'
    'D4:BE:D9' = 'Dell'
    'E4:54:E8' = 'Dell'
    'F8:B1:56' = 'Dell'
    'F8:CA:B8' = 'Dell'
    '00:12:FE' = 'Lenovo'
    '00:15:58' = 'Lenovo'
    '00:16:D3' = 'Lenovo'
    '00:19:1D' = 'Lenovo'
    '00:1A:6B' = 'Lenovo'
    '00:1E:37' = 'Lenovo'
    '00:21:CC' = 'Lenovo'
    '00:24:7E' = 'Lenovo'
    '08:9E:01' = 'Lenovo'
    '10:7D:1A' = 'Lenovo'
    '14:9F:E8' = 'Lenovo'
    '1C:39:47' = 'Lenovo'
    '20:89:84' = 'Lenovo'
    '28:D2:44' = 'Lenovo'
    '34:97:F6' = 'Lenovo'
    '38:F3:AB' = 'Lenovo'
    '40:B0:34' = 'Lenovo'
    '50:7B:9D' = 'Lenovo'
    '54:EE:75' = 'Lenovo'
    '5C:F3:FC' = 'Lenovo'
    '60:CF:84' = 'Lenovo'
    '68:F7:28' = 'Lenovo'
    '6C:4B:90' = 'Lenovo'
    '70:1C:E7' = 'Lenovo'
    '80:FA:5B' = 'Lenovo'
    '8C:16:45' = 'Lenovo'
    '98:FA:9B' = 'Lenovo'
    'A4:1F:72' = 'Lenovo'
    'B8:8A:60' = 'Lenovo'
    'C8:5B:76' = 'Lenovo'
    'D8:71:57' = 'Lenovo'
    'E8:6A:64' = 'Lenovo'
    'F0:76:1C' = 'Lenovo'
    'F4:7B:09' = 'Lenovo'
    '00:10:83' = 'HP'
    '00:11:0A' = 'HP'
    '00:12:79' = 'HP'
    '00:14:38' = 'HP'
    '00:16:35' = 'HP'
    '00:17:08' = 'HP'
    '00:18:FE' = 'HP'
    '00:1A:4B' = 'HP'
    '00:1B:78' = 'HP'
    '00:1F:29' = 'HP'
    '00:21:5A' = 'HP'
    '00:23:7D' = 'HP'
    '00:24:81' = 'HP'
    '00:25:B3' = 'HP'
    '00:26:55' = 'HP'
    '08:2E:5F' = 'HP'
    '10:60:4B' = 'HP'
    '10:7B:44' = 'HP'
    '18:60:24' = 'HP'
    '2C:41:38' = 'HP'
    '2C:44:FD' = 'HP'
    '3C:52:82' = 'HP'
    '40:A8:F0' = 'HP'
    '48:0F:CF' = 'HP'
    '5C:B9:01' = 'HP'
    '64:31:50' = 'HP'
    '6C:3B:E5' = 'HP'
    '70:5A:0F' = 'HP'
    '80:C1:6E' = 'HP'
    '84:34:97' = 'HP'
    '90:1B:0E' = 'HP'
    '94:57:A5' = 'HP'
    '98:E7:F4' = 'HP'
    'A0:1D:48' = 'HP'
    'A4:5D:36' = 'HP'
    'AC:16:2D' = 'HP'
    'B0:5A:DA' = 'HP'
    'B4:99:BA' = 'HP'
    'C8:D9:D2' = 'HP'
    'D4:85:64' = 'HP'
    'D8:9D:67' = 'HP'
    'E0:07:1B' = 'HP'
    'E4:11:5B' = 'HP'
    'EC:8E:B5' = 'HP'
    'F0:92:1C' = 'HP'
    'F4:CE:46' = 'HP'
    'FC:15:B4' = 'HP'
    '00:1D:60' = 'ASUS'
    '08:60:6E' = 'ASUS'
    '10:BF:48' = 'ASUS'
    '14:DA:E9' = 'ASUS'
    '18:31:BF' = 'ASUS'
    '1C:87:2C' = 'ASUS'
    '2C:56:DC' = 'ASUS'
    '30:5A:3A' = 'ASUS'
    '38:D5:47' = 'ASUS'
    '40:16:7E' = 'ASUS'
    '54:04:A6' = 'ASUS'
    '60:45:CB' = 'ASUS'
    '70:8B:CD' = 'ASUS'
    '74:D0:2B' = 'ASUS'
    '88:D7:F6' = 'ASUS'
    '9C:5C:8E' = 'ASUS'
    'AC:22:0B' = 'ASUS'
    'B0:6E:BF' = 'ASUS'
    'C8:7F:54' = 'ASUS'
    'D8:50:E6' = 'ASUS'
    'E0:3F:49' = 'ASUS'
    'F0:2F:74' = 'ASUS'
    'FC:34:97' = 'ASUS'
    '00:13:77' = 'Acer'
    '00:16:36' = 'Acer'
    '00:1B:24' = 'Acer'
    '00:23:5A' = 'Acer'
    '00:26:2D' = 'Acer'
    '20:6A:8A' = 'Acer'
    '30:65:EC' = 'Acer'
    '54:AB:3A' = 'Acer'
    '70:1A:04' = 'Acer'
    '84:2A:FD' = 'Acer'
    '88:AE:1D' = 'Acer'
    'B8:70:F4' = 'Acer'
    'C8:FF:28' = 'Acer'
    'DC:85:DE' = 'Acer'
    'F8:A9:D0' = 'Acer'
    '00:19:DB' = 'MSI'
    '00:21:85' = 'MSI'
    '44:8A:5B' = 'MSI'
    '4C:CC:6A' = 'MSI'
    'D8:CB:8A' = 'MSI'
    'F8:0F:41' = 'MSI'
    '00:03:93' = 'Apple'
    '00:0A:95' = 'Apple'
    '00:16:CB' = 'Apple'
    '00:17:F2' = 'Apple'
    '00:19:E3' = 'Apple'
    '00:1B:63' = 'Apple'
    '00:1E:C2' = 'Apple'
    '00:21:E9' = 'Apple'
    '00:23:12' = 'Apple'
    '00:25:00' = 'Apple'
    '00:26:08' = 'Apple'
    '04:0C:CE' = 'Apple'
    '08:00:07' = 'Apple'
    '10:9A:DD' = 'Apple'
    '14:10:9F' = 'Apple'
    '18:65:90' = 'Apple'
    '20:78:F0' = 'Apple'
    '28:CF:E9' = 'Apple'
    '30:10:E4' = 'Apple'
    '34:36:3B' = 'Apple'
    '38:C9:86' = 'Apple'
    '3C:07:54' = 'Apple'
    '40:A6:D9' = 'Apple'
    '48:60:BC' = 'Apple'
    '54:26:96' = 'Apple'
    '58:B0:35' = 'Apple'
    '5C:F9:38' = 'Apple'
    '60:F8:1D' = 'Apple'
    '64:B9:E8' = 'Apple'
    '68:FE:F7' = 'Apple'
    '70:56:81' = 'Apple'
    '78:31:C1' = 'Apple'
    '7C:D1:C3' = 'Apple'
    '84:38:35' = 'Apple'
    '88:63:DF' = 'Apple'
    '8C:85:90' = 'Apple'
    '90:72:40' = 'Apple'
    '98:01:A7' = 'Apple'
    'A4:5E:60' = 'Apple'
    'A8:86:DD' = 'Apple'
    'AC:BC:32' = 'Apple'
    'B8:09:8A' = 'Apple'
    'BC:67:1C' = 'Apple'
    'C0:63:94' = 'Apple'
    'C8:2A:14' = 'Apple'
    'CC:08:E0' = 'Apple'
    'D0:23:DB' = 'Apple'
    'D8:30:62' = 'Apple'
    'DC:2B:2A' = 'Apple'
    'E0:B5:2D' = 'Apple'
    'E8:80:2E' = 'Apple'
    'F0:18:98' = 'Apple'
    'F8:1E:DF' = 'Apple'
    '00:15:5D' = 'Microsoft Hyper-V'
    '28:18:78' = 'Microsoft'
    '7C:1E:52' = 'Microsoft'
    '98:5F:D3' = 'Microsoft'
    'B4:AE:2B' = 'Microsoft'
    'C8:3F:26' = 'Microsoft'
    'DC:B4:C4' = 'Microsoft'
    '00:0C:29' = 'VMware'
    '00:05:69' = 'VMware'
    '00:1C:14' = 'VMware'
    '00:50:56' = 'VMware'
    '08:00:27' = 'VirtualBox'
    '52:54:00' = 'KVM/QEMU'
    '00:16:3E' = 'Xen Virtual Machine'
    '00:00:0C' = 'Cisco'
    '00:01:42' = 'Cisco'
    '00:01:43' = 'Cisco'
    '00:01:63' = 'Cisco'
    '00:02:16' = 'Cisco'
    '00:02:4A' = 'Cisco'
    '00:02:4B' = 'Cisco'
    '00:02:B9' = 'Cisco'
    '00:03:6B' = 'Cisco'
    '00:03:E3' = 'Cisco'
    '00:04:27' = 'Cisco'
    '00:05:31' = 'Cisco'
    '00:06:28' = 'Cisco'
    '00:07:0E' = 'Cisco'
    '00:07:4F' = 'Cisco'
    '00:08:20' = 'Cisco'
    '00:09:43' = 'Cisco'
    '00:0A:41' = 'Cisco'
    '00:0B:45' = 'Cisco'
    '00:0C:30' = 'Cisco'
    '00:0D:28' = 'Cisco'
    '00:0E:38' = 'Cisco'
    '00:0F:23' = 'Cisco'
    '00:10:7B' = 'Cisco'
    '00:11:20' = 'Cisco'
    '00:12:00' = 'Cisco'
    '00:13:19' = 'Cisco'
    '00:14:1B' = 'Cisco'
    '00:15:2B' = 'Cisco'
    '00:16:46' = 'Cisco'
    '00:17:0E' = 'Cisco'
    '00:18:18' = 'Cisco'
    '00:19:06' = 'Cisco'
    '00:1A:2F' = 'Cisco'
    '00:1B:0C' = 'Cisco'
    '00:1C:58' = 'Cisco'
    '00:1D:45' = 'Cisco'
    '00:1E:13' = 'Cisco'
    '00:1F:6C' = 'Cisco'
    '00:21:1B' = 'Cisco'
    '00:22:55' = 'Cisco'
    '00:23:04' = 'Cisco'
    '00:24:14' = 'Cisco'
    '00:26:0A' = 'Cisco'
    '04:4E:06' = 'Cisco'
    '08:17:35' = 'Cisco'
    '0C:27:24' = 'Cisco'
    '10:05:CA' = 'Cisco'
    '18:33:9D' = 'Cisco'
    '1C:6A:7A' = 'Cisco'
    '20:BB:C0' = 'Cisco'
    '24:01:C7' = 'Cisco'
    '2C:54:2D' = 'Cisco'
    '34:62:88' = 'Cisco'
    '38:ED:18' = 'Cisco'
    '44:03:A7' = 'Cisco'
    '4C:00:82' = 'Cisco'
    '50:06:04' = 'Cisco'
    '54:7F:EE' = 'Cisco'
    '58:97:BD' = 'Cisco'
    '5C:50:15' = 'Cisco'
    '64:00:F1' = 'Cisco'
    '68:BC:0C' = 'Cisco'
    '6C:20:56' = 'Cisco'
    '70:10:5C' = 'Cisco'
    '74:A2:E6' = 'Cisco'
    '78:DA:6E' = 'Cisco'
    '7C:0E:CE' = 'Cisco'
    '80:E0:1D' = 'Cisco'
    '84:B8:02' = 'Cisco'
    '88:90:8D' = 'Cisco'
    '8C:60:4F' = 'Cisco'
    '90:6C:AC' = 'Cisco'
    '94:B4:0F' = 'Cisco'
    '98:4B:E1' = 'Cisco'
    'A0:EC:F9' = 'Cisco'
    'A4:6C:2A' = 'Cisco'
    'A8:B4:56' = 'Cisco'
    'AC:7A:56' = 'Cisco'
    'B0:AA:77' = 'Cisco'
    'B4:14:89' = 'Cisco'
    'B8:38:61' = 'Cisco'
    'BC:16:65' = 'Cisco'
    'C0:25:5C' = 'Cisco'
    'C4:0A:CB' = 'Cisco'
    'C8:4C:75' = 'Cisco'
    'CC:46:D6' = 'Cisco'
    'D0:57:4C' = 'Cisco'
    'D4:8C:B5' = 'Cisco'
    'D8:B1:90' = 'Cisco'
    'DC:7B:94' = 'Cisco'
    'E0:2F:6D' = 'Cisco'
    'E4:48:C7' = 'Cisco'
    'E8:04:62' = 'Cisco'
    'EC:1D:8B' = 'Cisco'
    'F0:29:29' = 'Cisco'
    'F4:4E:05' = 'Cisco'
    'F8:0F:6F' = 'Cisco'
    'FC:99:47' = 'Cisco'
    '00:04:96' = 'Extreme Networks'
    '00:12:CF' = 'Extreme Networks'
    '00:1F:45' = 'Extreme Networks'
    '20:B3:99' = 'Extreme Networks'
    '5C:0E:8B' = 'Extreme Networks'
    '6C:9C:ED' = 'Extreme Networks'
    'A4:6A:A8' = 'Extreme Networks'
    'B4:8C:5F' = 'Extreme Networks'
    '00:04:0D' = 'Avaya / Nortel'
    '00:08:02' = 'Avaya'
    '00:0A:E8' = 'Avaya'
    '00:13:65' = 'Avaya'
    '00:1B:4F' = 'Avaya'
    '00:22:67' = 'Avaya'
    '10:0F:18' = 'Avaya'
    '24:6A:AB' = 'Avaya'
    '34:75:C7' = 'Avaya'
    '50:CD:32' = 'Avaya'
    '6C:94:F8' = 'Avaya'
    'B4:B0:17' = 'Avaya'
    '00:04:0F' = 'Brocade / Ruckus'
    '00:05:1E' = 'Brocade'
    '00:10:18' = 'Brocade'
    '00:27:F8' = 'Brocade'
    '08:17:F4' = 'Brocade'
    '2C:27:D7' = 'Ruckus'
    '40:70:09' = 'Ruckus'
    '54:78:1A' = 'Ruckus'
    '58:FB:84' = 'Ruckus'
    '78:8C:54' = 'Ruckus'
    '84:18:88' = 'Ruckus'
    'C0:8A:DE' = 'Ruckus'
    'D4:68:4D' = 'Ruckus'
    '00:0B:86' = 'Aruba / HP'
    '00:1A:1E' = 'Aruba / HP'
    '00:24:6C' = 'Aruba / HP'
    '04:BD:88' = 'Aruba / HP'
    '18:64:72' = 'Aruba / HP'
    '24:DE:C6' = 'Aruba / HP'
    '34:FC:B9' = 'Aruba / HP'
    '48:4A:E9' = 'Aruba / HP'
    '6C:F3:7F' = 'Aruba / HP'
    '70:3A:0E' = 'Aruba / HP'
    '94:18:82' = 'Aruba / HP'
    '9C:8C:D8' = 'Aruba / HP'
    'AC:A3:1E' = 'Aruba / HP'
    'B0:B8:67' = 'Aruba / HP'
    'D8:C7:C8' = 'Aruba / HP'
    'F0:5C:19' = 'Aruba / HP'
    '00:05:5D' = 'D-Link'
    '00:0D:88' = 'D-Link'
    '00:0F:3D' = 'D-Link'
    '00:11:95' = 'D-Link'
    '00:13:46' = 'D-Link'
    '00:15:E9' = 'D-Link'
    '00:17:9A' = 'D-Link'
    '00:19:5B' = 'D-Link'
    '00:1B:11' = 'D-Link'
    '00:1C:F0' = 'D-Link'
    '00:1E:58' = 'D-Link'
    '00:21:91' = 'D-Link'
    '00:22:B0' = 'D-Link'
    '00:24:01' = 'D-Link'
    '00:26:5A' = 'D-Link'
    '08:BE:AC' = 'D-Link'
    '14:D6:4D' = 'D-Link'
    '1C:7E:E5' = 'D-Link'
    '28:10:7B' = 'D-Link'
    '34:08:04' = 'D-Link'
    '54:B8:0A' = 'D-Link'
    '5C:D9:98' = 'D-Link'
    '78:54:2E' = 'D-Link'
    '84:C9:B2' = 'D-Link'
    '90:94:E4' = 'D-Link'
    'B8:A3:86' = 'D-Link'
    'BC:F6:85' = 'D-Link'
    'C0:A0:BB' = 'D-Link'
    'C4:12:F5' = 'D-Link'
    'C8:BE:19' = 'D-Link'
    'D8:FE:E3' = 'D-Link'
    'E4:6F:13' = 'D-Link'
    'F0:7D:68' = 'D-Link'
    'FC:75:16' = 'D-Link'
    '00:14:6C' = 'Netgear'
    '00:1B:2F' = 'Netgear'
    '00:1E:2A' = 'Netgear'
    '00:22:3F' = 'Netgear'
    '00:24:B2' = 'Netgear'
    '04:A1:51' = 'Netgear'
    '08:BD:43' = 'Netgear'
    '10:0D:7F' = 'Netgear'
    '20:4E:7F' = 'Netgear'
    '28:C6:8E' = 'Netgear'
    '2C:30:33' = 'Netgear'
    '30:46:9A' = 'Netgear'
    '38:94:ED' = 'Netgear'
    '44:94:FC' = 'Netgear'
    '50:6A:03' = 'Netgear'
    '6C:B0:CE' = 'Netgear'
    '74:44:01' = 'Netgear'
    '80:37:73' = 'Netgear'
    '84:1B:5E' = 'Netgear'
    '9C:3D:CF' = 'Netgear'
    'A0:04:60' = 'Netgear'
    'A0:40:A0' = 'Netgear'
    'B0:7F:B9' = 'Netgear'
    'C0:3F:0E' = 'Netgear'
    'C4:3D:C7' = 'Netgear'
    'CC:40:D0' = 'Netgear'
    'E0:46:9A' = 'Netgear'
    'E0:91:F5' = 'Netgear'
    'E4:F4:C6' = 'Netgear'
    'F8:32:E4' = 'Netgear'
    '00:1D:0F' = 'TP-Link'
    '14:CC:20' = 'TP-Link'
    '18:A6:F7' = 'TP-Link'
    '1C:3B:F3' = 'TP-Link'
    '30:B5:C2' = 'TP-Link'
    '50:C7:BF' = 'TP-Link'
    '54:A7:03' = 'TP-Link'
    '5C:63:BF' = 'TP-Link'
    '60:A4:B7' = 'TP-Link'
    '64:66:B3' = 'TP-Link'
    '6C:5A:B0' = 'TP-Link'
    '70:4F:57' = 'TP-Link'
    '74:DA:88' = 'TP-Link'
    '84:16:F9' = 'TP-Link'
    '90:9A:4A' = 'TP-Link'
    '98:DA:C4' = 'TP-Link'
    'A0:F3:C1' = 'TP-Link'
    'A4:2B:B0' = 'TP-Link'
    'B0:4E:26' = 'TP-Link'
    'C0:4A:00' = 'TP-Link'
    'C4:6E:1F' = 'TP-Link'
    'D8:0D:17' = 'TP-Link'
    'E8:48:B8' = 'TP-Link'
    'EC:08:6B' = 'TP-Link'
    'F4:EC:38' = 'TP-Link'
    'F8:1A:67' = 'TP-Link'
    '00:0C:42' = 'MikroTik'
    '4C:5E:0C' = 'MikroTik'
    '64:D1:54' = 'MikroTik'
    '6C:3B:6B' = 'MikroTik'
    '74:4D:28' = 'MikroTik'
    'B8:69:F4' = 'MikroTik'
    'CC:2D:E0' = 'MikroTik'
    'D4:CA:6D' = 'MikroTik'
    'E4:8D:8C' = 'MikroTik'
    '00:15:6D' = 'Ubiquiti'
    '04:18:D6' = 'Ubiquiti'
    '18:E8:29' = 'Ubiquiti'
    '24:A4:3C' = 'Ubiquiti'
    '44:D9:E7' = 'Ubiquiti'
    '68:72:51' = 'Ubiquiti'
    '78:8A:20' = 'Ubiquiti'
    '80:2A:A8' = 'Ubiquiti'
    '9C:05:D6' = 'Ubiquiti'
    'A0:63:91' = 'Ubiquiti'
    'AC:8B:A9' = 'Ubiquiti'
    'B4:FB:E4' = 'Ubiquiti'
    'DC:9F:DB' = 'Ubiquiti'
    'E0:63:DA' = 'Ubiquiti'
    'F0:9F:C2' = 'Ubiquiti'
    'F4:92:BF' = 'Ubiquiti'
    'FC:EC:DA' = 'Ubiquiti'
    '00:09:0F' = 'Fortinet'
    '00:0C:E6' = 'Fortinet'
    '00:17:09' = 'Fortinet'
    '04:D5:90' = 'Fortinet'
    '08:5B:0E' = 'Fortinet'
    '10:47:80' = 'Fortinet'
    '18:19:D6' = 'Fortinet'
    '1C:1B:0D' = 'Fortinet'
    '20:3A:07' = 'Fortinet'
    '28:8B:1C' = 'Fortinet'
    '2C:FD:A1' = 'Fortinet'
    '3C:61:04' = 'Fortinet'
    '40:AC:BF' = 'Fortinet'
    '58:96:71' = 'Fortinet'
    '70:4C:A5' = 'Fortinet'
    '94:FF:3C' = 'Fortinet'
    'A0:32:99' = 'Fortinet'
    'E8:1C:BA' = 'Fortinet'
    'F4:8E:38' = 'Fortinet'
    '00:1B:17' = 'Palo Alto Networks'
    '24:A9:37' = 'Palo Alto Networks'
    '5C:49:7D' = 'Palo Alto Networks'
    '78:72:5D' = 'Palo Alto Networks'
    'B4:0C:25' = 'Palo Alto Networks'
    'D4:F4:BE' = 'Palo Alto Networks'
    '00:1C:0E' = 'Check Point'
    '00:1A:8C' = 'Sophos'
    '00:25:13' = 'Sophos'
    '00:1D:4F' = 'Juniper'
    '00:05:85' = 'Juniper'
    '00:10:DB' = 'Juniper'
    '2C:6B:F5' = 'Juniper'
    '40:A6:77' = 'Juniper'
    '54:4B:8C' = 'Juniper'
    '78:19:F7' = 'Juniper'
    '84:B5:9C' = 'Juniper'
    '88:A2:D7' = 'Juniper'
    'B0:A8:6E' = 'Juniper'
    'C4:05:28' = 'Juniper'
    'D0:07:CA' = 'Juniper'
    'EC:13:DB' = 'Juniper'
    '00:1C:73' = 'Arista Networks'
    '24:8A:07' = 'Arista Networks'
    '28:99:3A' = 'Arista Networks'
    '44:4C:A8' = 'Arista Networks'
    '64:64:9B' = 'Arista Networks'
    '84:8A:8D' = 'Arista Networks'
    'A4:4C:C8' = 'Arista Networks'
    'C0:42:D0' = 'Arista Networks'
    'D4:AF:F7' = 'Arista Networks'
    'E4:43:4B' = 'Arista Networks'
    '00:1E:49' = 'H3C'
    '00:23:89' = 'H3C'
    '0C:DA:41' = 'H3C'
    '14:43:19' = 'H3C'
    '38:97:D6' = 'H3C'
    '70:BA:EF' = 'H3C'
    '74:EA:CB' = 'H3C'
    '88:5A:92' = 'H3C'
    'C8:CB:B8' = 'H3C'
    'E8:EA:6A' = 'H3C'
    'F4:83:CD' = 'H3C'
    '00:E0:FC' = 'Huawei'
    '04:BD:70' = 'Huawei'
    '08:19:A6' = 'Huawei'
    '10:1B:54' = 'Huawei'
    '14:B9:68' = 'Huawei'
    '18:CF:5E' = 'Huawei'
    '1C:1D:67' = 'Huawei'
    '20:F3:A3' = 'Huawei'
    '24:69:A5' = 'Huawei'
    '28:31:52' = 'Huawei'
    '2C:AB:00' = 'Huawei'
    '30:D1:7E' = 'Huawei'
    '34:00:A3' = 'Huawei'
    '38:BC:01' = 'Huawei'
    '3C:DF:BD' = 'Huawei'
    '40:4D:8E' = 'Huawei'
    '44:55:B1' = 'Huawei'
    '48:46:FB' = 'Huawei'
    '4C:1F:CC' = 'Huawei'
    '54:89:98' = 'Huawei'
    '58:1F:28' = 'Huawei'
    '5C:7D:5E' = 'Huawei'
    '60:DE:44' = 'Huawei'
    '64:16:F0' = 'Huawei'
    '68:A0:3E' = 'Huawei'
    '70:72:3C' = 'Huawei'
    '78:D7:52' = 'Huawei'
    '80:13:82' = 'Huawei'
    '84:A8:E4' = 'Huawei'
    '88:CE:FA' = 'Huawei'
    '8C:34:FD' = 'Huawei'
    '90:94:97' = 'Huawei'
    '94:04:9C' = 'Huawei'
    '98:E0:D9' = 'Huawei'
    'A0:8C:FD' = 'Huawei'
    'A4:C6:4F' = 'Huawei'
    'A8:CA:7B' = 'Huawei'
    'AC:4E:91' = 'Huawei'
    'B0:08:75' = 'Huawei'
    'B4:30:52' = 'Huawei'
    'B8:08:D7' = 'Huawei'
    'BC:76:70' = 'Huawei'
    'C0:70:09' = 'Huawei'
    'C4:47:3F' = 'Huawei'
    'C8:94:BB' = 'Huawei'
    'CC:96:A0' = 'Huawei'
    'D0:7A:B5' = 'Huawei'
    'D4:40:F0' = 'Huawei'
    'D8:49:0B' = 'Huawei'
    'DC:D2:FC' = 'Huawei'
    'E0:19:1D' = 'Huawei'
    'E4:A7:C5' = 'Huawei'
    'E8:4D:74' = 'Huawei'
    'EC:23:3D' = 'Huawei'
    'F0:25:72' = 'Huawei'
    'F4:4C:7F' = 'Huawei'
    'F8:4A:BF' = 'Huawei'
    'FC:48:EF' = 'Huawei'
    '00:80:77' = 'Brother'
    '30:05:5C' = 'Brother'
    '3C:2A:F4' = 'Brother'
    '74:5E:1C' = 'Brother'
    '80:77:29' = 'Brother'
    'A8:6B:AD' = 'Brother'
    'B0:9F:BA' = 'Brother'
    'C8:3A:35' = 'Brother'
    'E8:DA:00' = 'Brother'
    '00:00:85' = 'Canon'
    '00:1E:8F' = 'Canon'
    '00:1F:16' = 'Canon'
    '00:26:AB' = 'Canon'
    '34:9F:7B' = 'Canon'
    '40:8D:5C' = 'Canon'
    '88:87:17' = 'Canon'
    'BC:60:A7' = 'Canon'
    'F4:81:39' = 'Canon'
    '00:00:48' = 'Epson'
    '38:1A:52' = 'Epson'
    '44:D2:44' = 'Epson'
    '64:EB:8C' = 'Epson'
    '70:5A:9E' = 'Epson'
    '9C:AE:D3' = 'Epson'
    'AC:18:26' = 'Epson'
    'B0:E8:92' = 'Epson'
    'D0:3D:67' = 'Epson'
    '00:04:00' = 'Lexmark'
    '00:21:B7' = 'Lexmark'
    '00:26:59' = 'Lexmark'
    '10:1F:74' = 'Lexmark'
    '24:FD:52' = 'Lexmark'
    '40:F2:E9' = 'Lexmark'
    '54:13:79' = 'Lexmark'
    '70:5A:B6' = 'Lexmark'
    '84:2B:2B' = 'Lexmark'
    'B0:7B:25' = 'Lexmark'
    '00:80:91' = 'Ricoh'
    '00:26:73' = 'Ricoh'
    '58:38:79' = 'Ricoh'
    '58:52:8A' = 'Ricoh'
    '74:72:B0' = 'Ricoh'
    'A0:E4:CB' = 'Ricoh'
    'D8:49:2F' = 'Ricoh'
    '00:13:E8' = 'Intelbras'
    '00:19:3F' = 'Ruckus/Wireless'
    '00:18:E7' = 'Cameo / Network Device'
    'B8:27:EB' = 'Raspberry Pi'
    'DC:A6:32' = 'Raspberry Pi'
    'E4:5F:01' = 'Raspberry Pi'
    'BC:AD:28' = 'Hikvision'
    'C0:56:E3' = 'Hikvision'
    'D4:E8:53' = 'Hikvision'
    'E0:50:8B' = 'Hikvision'
    '44:19:B6' = 'Hikvision'
    '08:54:11' = 'Hikvision'
    '00:12:12' = 'Dahua'
    '24:52:6A' = 'Dahua'
    '3C:EF:8C' = 'Dahua'
    '4C:11:BF' = 'Dahua'
    '90:02:A9' = 'Dahua'
    'BC:32:5F' = 'Dahua'
    'F4:B1:C2' = 'Dahua'
}

function Normalize-Mac {
    param([string]$Mac)
    if ([string]::IsNullOrWhiteSpace($Mac)) { return '' }
    $m = $Mac.Trim().ToUpper() -replace '-', ':' -replace '\.', '' -replace '\s', ''
    if ($m -match '^[0-9A-F]{12}$') {
        return (($m -split '(.{2})' | Where-Object { $_ }) -join ':')
    }
    return $m
}

function Get-Oui {
    param([string]$Mac)
    $m = Normalize-Mac $Mac
    if ($m.Length -ge 8) { return $m.Substring(0,8) }
    return ''
}

function Import-ExternalOuiMap {
    # Loads a complete offline OUI database when available.
    # Keep any one of these files beside this script:
    #   oui.csv    -> OUI,Vendor or AA:BB:CC,Vendor
    #   oui.txt    -> IEEE public OUI text format
    #   manuf.txt  -> Wireshark manuf format
    try {
        $base = Split-Path -Parent $PSCommandPath
        if (-not $base) { $base = (Get-Location).Path }
        $files = @('oui.csv','oui.txt','manuf.txt','nmap-mac-prefixes') | ForEach-Object { Join-Path $base $_ } | Where-Object { Test-Path $_ }
        foreach ($file in $files) {
            $loadedFromThisFile = 0
            foreach ($line in [System.IO.File]::ReadLines($file)) {
                if ([string]::IsNullOrWhiteSpace($line)) { continue }
                $l = $line.Trim()
                if ($l.StartsWith('#')) { continue }
                $oui = ''
                $vendor = ''

                # Wireshark manuf: 00:11:22<TAB>Vendor
                if ($l -match '^([0-9A-Fa-f]{2})[:-]([0-9A-Fa-f]{2})[:-]([0-9A-Fa-f]{2})[,	 ]+(.+)$') {
                    $oui = ('{0}:{1}:{2}' -f $matches[1],$matches[2],$matches[3]).ToUpper()
                    $vendor = $matches[4].Trim(' ',',','`t','"')
                }
                # IEEE oui.txt: AABBCC     (base 16)        Vendor Name
                elseif ($l -match '^([0-9A-Fa-f]{6})\s+\(base 16\)\s+(.+)$') {
                    $x = $matches[1].ToUpper()
                    $oui = $x.Substring(0,2)+':'+$x.Substring(2,2)+':'+$x.Substring(4,2)
                    $vendor = $matches[2].Trim()
                }
                # CSV/simple: AABBCC,Vendor OR AABBCC Vendor
                elseif ($l -match '^([0-9A-Fa-f]{6})[,	 ]+(.+)$') {
                    $x = $matches[1].ToUpper()
                    $oui = $x.Substring(0,2)+':'+$x.Substring(2,2)+':'+$x.Substring(4,2)
                    $vendor = $matches[2].Trim(' ',',','`t','"')
                }

                if ($oui -and $vendor) {
                    # External complete database should be allowed to improve/replace built-in generic names.
                    $VendorMap[$oui] = $vendor
                    $loadedFromThisFile++
                    $script:OuiEntriesLoaded++
                }
            }
            if ($loadedFromThisFile -gt 0) { $script:OuiExternalFiles += (Split-Path -Leaf $file) }
        }
        if ($script:OuiEntriesLoaded -gt 0) {
            $script:OuiSourceMode = "External OUI database loaded: $($script:OuiEntriesLoaded) entries from $($script:OuiExternalFiles -join ', ')"
        }
    } catch {}
}

function New-OuiDatabaseHelpFile {
    try {
        $base = Split-Path -Parent $PSCommandPath
        if (-not $base) { $base = (Get-Location).Path }
        $helpFile = Join-Path $base 'README_OUI_DATABASE.txt'
        if (-not (Test-Path $helpFile)) {
@'
Complete OUI / Vendor Database Support
======================================

This scanner has a built-in common IT vendor map for PC, switch, firewall,
printer, IP phone, CCTV, biometric/Savior/ZKT and other common devices.

For complete vendor coverage, place ONE full OUI file beside this PS1:

1) oui.csv
   Format examples:
   00:11:22,Cisco Systems
   AABBCC,Dell Inc.

2) manuf.txt
   Wireshark manufacturer database format.

3) oui.txt
   IEEE public OUI text format.

After placing the file, reopen the scanner. It will load the OUI database offline.
No internet is needed during scanning.
'@ | Set-Content -Path $helpFile -Encoding UTF8
        }
    } catch {}
}

New-OuiDatabaseHelpFile
Import-ExternalOuiMap

function Get-VendorName {
    param([string]$Mac)
    $oui = Get-Oui $Mac
    if ($VendorMap.ContainsKey($oui)) { return $VendorMap[$oui] }
    if ($oui) { return 'Unknown Vendor' }
    return ''
}

function Get-DeviceType {
    param([string]$Vendor, [string]$HostName, [string]$Mac, [string]$Http, [string]$Https)
    $v = ($Vendor + ' ' + $HostName).ToLower()
    if (-not (Get-Oui $Mac)) { return 'Unknown' }

    # Biometric / attendance / access-control first because many show generic OEM names.
    if ($v -match 'zkteco|zksoftware|zk teco|savior|saviour|access control|biometric|fingerprint|attendance') { return 'Savior / ZKT / Biometric' }

    # Voice / PBX / IP phones.
    if ($v -match 'fanvil|yealink|grandstream|polycom|poly |avaya|snom|alcatel|mitel|cisco phone|ip phone') { return 'IP Phone' }

    # Network infra.
    if ($v -match 'cisco|aruba|procurve|hewlett packard enterprise|hpe|d-link|dlink|netgear|tp-link|tplink|mikrotik|ubiquiti|unifi|ruckus|extreme|brocade|h3c|huawei|juniper|arista|zyxel|meraki|allied telesis|edgecore|planet technology|switch|router|access point|wireless') { return 'Switch / Router / Network' }

    # Security appliances.
    if ($v -match 'fortinet|fortigate|palo alto|checkpoint|check point|sophos|sonicwall|watchguard|cyberoam|firewall') { return 'Firewall / Security' }

    # CCTV / NVR.
    if ($v -match 'hikvision|dahua|axis communications|hanwha|uniview|cp plus|camera|nvr|dvr|cctv') { return 'CCTV / Camera / NVR' }

    # Printers / label printers / scanners.
    if ($v -match 'brother|canon|epson|lexmark|ricoh|kyocera|xerox|zebra|honeywell|datamax|sato|printer|barcode') { return 'Printer / Scanner' }

    # NAS / storage.
    if ($v -match 'qnap|synology|netapp|nas|storage') { return 'NAS / Storage' }

    # Virtualization.
    if ($v -match 'vmware|virtualbox|hyper-v|microsoft virtual|kvm|xen|qemu') { return 'VM / Virtual Host' }

    # PCs, laptops, workstations and servers.
    if ($v -match 'dell|lenovo|hp|hewlett|acer|asus|msi|apple|microsoft|intel|gigabyte|asrock|toshiba|fujitsu|surface|thinkpad|elitebook|probook|latitude|optiplex|inspiron|precision|workstation|desktop|laptop|server|pc') { return 'PC / Laptop / Server' }

    # IoT and other devices with web UI.
    if (($Http -eq 'Open') -or ($Https -eq 'Open')) { return 'Web / Embedded IT Device' }
    return 'Unknown IT Device'
}

function Get-ArpMacForIp {
    param([string]$Ip)
    try {
        $arp = arp -a $Ip 2>$null | Out-String
        foreach ($line in ($arp -split "`r?`n")) {
            if ($line -match "\b$([regex]::Escape($Ip))\b\s+([0-9A-Fa-f:-]{17})\s+") {
                return Normalize-Mac $matches[1]
            }
        }
    } catch {}
    return ''
}

function Resolve-HostNameSafe {
    param([string]$Ip)

    # 1) DNS / reverse DNS
    try {
        $entry = [System.Net.Dns]::GetHostEntry($Ip)
        if ($entry.HostName) { return ($entry.HostName -replace '\.$','') }
    } catch {}

    # 2) NetBIOS fallback. Best for Windows PCs when DNS reverse entry is missing.
    try {
        $nbt = nbtstat -A $Ip 2>$null | Out-String
        foreach ($line in ($nbt -split "`r?`n")) {
            if ($line -match '^\s*([^\s<]+)\s+<00>\s+UNIQUE') {
                $name = $matches[1].Trim()
                if ($name -and $name -notmatch '^__MSBROWSE__$') { return $name }
            }
        }
    } catch {}

    # 3) nslookup fallback.
    try {
        $ns = nslookup $Ip 2>$null | Out-String
        foreach ($line in ($ns -split "`r?`n")) {
            if ($line -match 'Name:\s+(.+)$') {
                $name = $matches[1].Trim()
                if ($name) { return ($name -replace '\.$','') }
            }
        }
    } catch {}

    # 4) ping -a fallback.
    try {
        $pa = ping -a -n 1 -w 250 $Ip 2>$null | Out-String
        if ($pa -match 'Pinging\s+([^\s\[]+)\s+\[') {
            $name = $matches[1].Trim()
            if ($name -and $name -ne $Ip) { return $name }
        }
    } catch {}

    return ''
}

function Test-PortQuick {
    param([string]$Ip, [int]$Port, [int]$TimeoutMs = 250)
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($Ip, $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if ($ok -and $client.Connected) {
            $client.EndConnect($iar)
            $client.Close()
            return $true
        }
        $client.Close()
    } catch {}
    return $false
}

function Ping-Ip {
    param([string]$Ip)
    try {
        return Test-Connection -ComputerName $Ip -Count 1 -Quiet -ErrorAction SilentlyContinue
    } catch {
        try {
            $p = New-Object System.Net.NetworkInformation.Ping
            $r = $p.Send($Ip, 500)
            return ($r.Status -eq 'Success')
        } catch { return $false }
    }
}

function Add-Log {
    param([string]$Text)
    $time = Get-Date -Format 'HH:mm:ss'
    $txtLog.AppendText("[$time] $Text`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Add-ResultRow {
    param($Obj)
    $script:Results.Add($Obj) | Out-Null
    $rowIndex = $grid.Rows.Add()
    $row = $grid.Rows[$rowIndex]
    $row.Cells['SNo'].Value = $Obj.SNo
    $row.Cells['IPAddress'].Value = $Obj.IPAddress
    $row.Cells['MACAddress'].Value = $Obj.MACAddress
    $row.Cells['OUI'].Value = $Obj.OUI
    $row.Cells['Vendor'].Value = $Obj.Vendor
    $row.Cells['DeviceType'].Value = $Obj.DeviceType
    $row.Cells['HostName'].Value = $Obj.HostName
    $row.Cells['Status'].Value = $Obj.Status
    $row.Cells['HTTP'].Value = $Obj.HTTP
    $row.Cells['HTTPS'].Value = $Obj.HTTPS
    $row.Cells['URL'].Value = $Obj.URL

    if ($Obj.Status -eq 'Online') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(225,255,225) }
    elseif ($Obj.Status -eq 'ARP Only') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230,240,255) }
    else { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,235,235) }
}

# ---------------- GUI ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'LAN Device Explorer'
$form.Size = New-Object System.Drawing.Size(1220, 760)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

# Window icon setting
# Leave empty to hide the default WinForms icon from the title bar.
# Paste your Base64 icon value below if needed. Leave empty to hide the title-bar icon.
# Recommended: ICO Base64. PNG/JPG Base64 also supported.
$AppIconBase64 = @'

'@
try {
    if ([string]::IsNullOrWhiteSpace($AppIconBase64)) {
        $form.ShowIcon = $false
    } else {
        $iconBytes = [Convert]::FromBase64String($AppIconBase64)
        $msIcon = New-Object System.IO.MemoryStream(,$iconBytes)
        try {
            $form.Icon = New-Object System.Drawing.Icon($msIcon)
        } catch {
            $msIcon.Position = 0
            $bmpIcon = [System.Drawing.Bitmap]::FromStream($msIcon)
            $hIcon = $bmpIcon.GetHicon()
            $form.Icon = [System.Drawing.Icon]::FromHandle($hIcon)
        }
    }
} catch {
    $form.ShowIcon = $false
}

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'LAN Device Explorer'
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(20, 18)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = 'Discover IP Address, MAC Address, OUI Vendor, Device Type, Hostname and Web Services'
$lblSub.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(24, 58)
$form.Controls.Add($lblSub)

$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(20, 90)
$panel.Size = New-Object System.Drawing.Size(1160, 82)
$panel.BorderStyle = 'FixedSingle'
$panel.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($panel)

$lblSegment = New-Object System.Windows.Forms.Label
$lblSegment.Text = 'IP Segment'
$lblSegment.Location = New-Object System.Drawing.Point(15, 15)
$lblSegment.AutoSize = $true
$panel.Controls.Add($lblSegment)

$txtSegment = New-Object System.Windows.Forms.TextBox
$txtSegment.Text = '10.209.110'
$txtSegment.Location = New-Object System.Drawing.Point(15, 38)
$txtSegment.Size = New-Object System.Drawing.Size(180, 25)
$panel.Controls.Add($txtSegment)

$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Text = 'From'
$lblFrom.Location = New-Object System.Drawing.Point(225, 15)
$lblFrom.AutoSize = $true
$panel.Controls.Add($lblFrom)

$numFrom = New-Object System.Windows.Forms.NumericUpDown
$numFrom.Minimum = 1
$numFrom.Maximum = 254
$numFrom.Value = 1
$numFrom.Location = New-Object System.Drawing.Point(225, 38)
$numFrom.Size = New-Object System.Drawing.Size(80, 25)
$panel.Controls.Add($numFrom)

$lblTo = New-Object System.Windows.Forms.Label
$lblTo.Text = 'To'
$lblTo.Location = New-Object System.Drawing.Point(330, 15)
$lblTo.AutoSize = $true
$panel.Controls.Add($lblTo)

$numTo = New-Object System.Windows.Forms.NumericUpDown
$numTo.Minimum = 1
$numTo.Maximum = 254
$numTo.Value = 254
$numTo.Location = New-Object System.Drawing.Point(330, 38)
$numTo.Size = New-Object System.Drawing.Size(80, 25)
$panel.Controls.Add($numTo)

$chkHostname = New-Object System.Windows.Forms.CheckBox
$chkHostname.Text = 'Resolve Hostname'
$chkHostname.Checked = $true
$chkHostname.Location = New-Object System.Drawing.Point(445, 38)
$chkHostname.Size = New-Object System.Drawing.Size(145, 25)
$panel.Controls.Add($chkHostname)

$chkPorts = New-Object System.Windows.Forms.CheckBox
$chkPorts.Text = 'Check HTTP/HTTPS'
$chkPorts.Checked = $true
$chkPorts.Location = New-Object System.Drawing.Point(600, 38)
$chkPorts.Size = New-Object System.Drawing.Size(160, 25)
$panel.Controls.Add($chkPorts)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = 'Start Scan'
$btnStart.Location = New-Object System.Drawing.Point(795, 25)
$btnStart.Size = New-Object System.Drawing.Size(110, 38)
$btnStart.BackColor = [System.Drawing.Color]::FromArgb(42, 160, 80)
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.FlatStyle = 'Flat'
$panel.Controls.Add($btnStart)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = 'Stop'
$btnStop.Enabled = $false
$btnStop.Location = New-Object System.Drawing.Point(915, 25)
$btnStop.Size = New-Object System.Drawing.Size(90, 38)
$btnStop.BackColor = [System.Drawing.Color]::FromArgb(220, 70, 70)
$btnStop.ForeColor = [System.Drawing.Color]::White
$btnStop.FlatStyle = 'Flat'
$panel.Controls.Add($btnStop)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(1015, 25)
$btnExport.Size = New-Object System.Drawing.Size(120, 38)
$btnExport.BackColor = [System.Drawing.Color]::FromArgb(40, 105, 190)
$btnExport.ForeColor = [System.Drawing.Color]::White
$btnExport.FlatStyle = 'Flat'
$panel.Controls.Add($btnExport)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 185)
$progress.Size = New-Object System.Drawing.Size(1160, 20)
$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready'
$lblStatus.Location = New-Object System.Drawing.Point(20, 210)
$lblStatus.Size = New-Object System.Drawing.Size(1160, 25)
$form.Controls.Add($lblStatus)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 240)
$grid.Size = New-Object System.Drawing.Size(1160, 360)
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = 'FullRowSelect'
$grid.AutoSizeColumnsMode = 'Fill'
$grid.RowHeadersVisible = $false
$form.Controls.Add($grid)

$cols = @(
    @{Name='SNo'; Header='S.No'; Width=50},
    @{Name='IPAddress'; Header='IP Address'; Width=120},
    @{Name='MACAddress'; Header='MAC Address'; Width=140},
    @{Name='OUI'; Header='OUI'; Width=80},
    @{Name='Vendor'; Header='Vendor'; Width=155},
    @{Name='DeviceType'; Header='Device Type'; Width=165},
    @{Name='HostName'; Header='Host Name'; Width=170},
    @{Name='Status'; Header='Status'; Width=95},
    @{Name='HTTP'; Header='HTTP'; Width=70},
    @{Name='HTTPS'; Header='HTTPS'; Width=70},
    @{Name='URL'; Header='URL'; Width=190}
)
foreach ($c in $cols) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.Name
    $col.HeaderText = $c.Header
    $col.FillWeight = $c.Width
    [void]$grid.Columns.Add($col)
}

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 615)
$txtLog.Size = New-Object System.Drawing.Size(1160, 90)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($txtLog)

$btnStop.Add_Click({
    $script:StopScan = $true
    Add-Log 'Stop requested by user.'
})

$btnExport.Add_Click({
    if ($script:Results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No data to export.', 'Export', 'OK', 'Information') | Out-Null
        return
    }
    $outDir = Join-Path $env:USERPROFILE 'Desktop\LAN_MAC_Scanner_Output'
    if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $file = Join-Path $outDir ('LAN_MAC_Scan_' + (Get-Date -Format 'dd-MM-yyyy_HH-mm-ss') + '.csv')
    $script:Results | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Add-Log "Exported: $file"
    Start-Process explorer.exe "/select,`"$file`""
})

$btnStart.Add_Click({
    $script:StopScan = $false
    $script:Results.Clear()
    $grid.Rows.Clear()
    $txtLog.Clear()

    $segment = $txtSegment.Text.Trim().TrimEnd('.')
    $from = [int]$numFrom.Value
    $to = [int]$numTo.Value
    if ($from -gt $to) {
        [System.Windows.Forms.MessageBox]::Show('From value cannot be greater than To value.', 'Invalid Range', 'OK', 'Warning') | Out-Null
        return
    }
    if ($segment -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}$') {
        [System.Windows.Forms.MessageBox]::Show('Enter IP segment like 10.209.110', 'Invalid Segment', 'OK', 'Warning') | Out-Null
        return
    }

    $btnStart.Enabled = $false
    $btnStop.Enabled = $true
    $progress.Minimum = 0
    $progress.Maximum = ($to - $from + 1)
    $progress.Value = 0

    Add-Log "Scan started: $segment.$from to $segment.$to"
    $sno = 0
    $total = $to - $from + 1
    $done = 0

    for ($i = $from; $i -le $to; $i++) {
        if ($script:StopScan) { break }
        $ip = "$segment.$i"
        $done++
        $lblStatus.Text = "Scanning $ip   ($done / $total)"
        [System.Windows.Forms.Application]::DoEvents()

        $isOnline = Ping-Ip $ip
        $mac = Get-ArpMacForIp $ip

        # If ping failed, still try ARP after a short ping attempt. This catches devices blocking ping but present in ARP.
        if (-not $mac) {
            try { ping -n 1 -w 250 $ip | Out-Null } catch {}
            Start-Sleep -Milliseconds 30
            $mac = Get-ArpMacForIp $ip
        }

        $status = 'Offline'
        if ($isOnline) { $status = 'Online' }
        elseif ($mac) { $status = 'ARP Only' }

        $http = ''
        $https = ''
        $url = ''
        if ($chkPorts.Checked -and ($isOnline -or $mac)) {
            $httpOpen = Test-PortQuick $ip 80 250
            $httpsOpen = Test-PortQuick $ip 443 250
            $http = if ($httpOpen) { 'Open' } else { 'Closed' }
            $https = if ($httpsOpen) { 'Open' } else { 'Closed' }
            if ($httpOpen) { $url = "http://$ip/" }
            elseif ($httpsOpen) { $url = "https://$ip/" }
        }

        # Do not use variable name `$host` because `$Host` is a built-in read-only PowerShell variable.
        # This is important after PS1-to-EXE conversion also.
        $resolvedHostName = ''
        if ($chkHostname.Checked -and ($isOnline -or $mac)) { $resolvedHostName = Resolve-HostNameSafe $ip }

        if ($isOnline -or $mac) {
            $sno++
            $oui = Get-Oui $mac
            $vendor = Get-VendorName $mac
            $deviceType = Get-DeviceType -Vendor $vendor -HostName $resolvedHostName -Mac $mac -Http $http -Https $https
            $obj = [pscustomobject]@{
                SNo        = $sno
                IPAddress  = $ip
                MACAddress = $mac
                OUI        = $oui
                Vendor     = $vendor
                DeviceType = $deviceType
                HostName   = $resolvedHostName
                Status     = $status
                HTTP       = $http
                HTTPS      = $https
                URL        = $url
                ScanTime   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            }
            Add-ResultRow $obj
            Add-Log "Found $ip  MAC=$mac  OUI=$oui  Vendor=$vendor  Type=$deviceType  Host=$resolvedHostName  Status=$status"
        }

        if ($progress.Value -lt $progress.Maximum) { $progress.Value++ }
        [System.Windows.Forms.Application]::DoEvents()
    }

    $lblStatus.Text = "Completed. Found $($script:Results.Count) device(s)."
    Add-Log "Scan completed. Found $($script:Results.Count) device(s)."
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
})

[void]$form.ShowDialog()
