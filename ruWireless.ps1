## CSV File Location (If this doesn't exist, the script will attempt to create it. Users will need full control of the file.)
$csv = "$pwd\ruWireless.csv"

## Error log path (Optional but recommended. If this doesn't exist, the script will attempt to create it. Users will need full control of the file.)
$ErrorLogPath = "$pwd\PowerShell-PC-ruWireless-Error-Log.log"

Write-Host "Gathering inventory information..."

# Date
$Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss" #done

# Serial Number
$SN = Get-WmiObject -Class Win32_Bios | Select-Object -ExpandProperty SerialNumber #

# Username
$Username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

#Get Connection Type
$WirelessConnected = $null
$WiredConnected = $null



# Detecting PowerShell version, and call the best cmdlets
if ($PSVersionTable.PSVersion.Major -gt 2)
{
    # Using Get-CimInstance for PowerShell version 3.0 and higher
    $WirelessAdapters =  Get-CimInstance -Namespace "root\WMI" -Class MSNdis_PhysicalMediumType -Filter `
        'NdisPhysicalMediumType = 9'
    $WiredAdapters = Get-CimInstance -Namespace "root\WMI" -Class MSNdis_PhysicalMediumType -Filter `
        "NdisPhysicalMediumType = 0 and `
        (NOT InstanceName like '%pangp%') and `
        (NOT InstanceName like '%cisco%') and `
        (NOT InstanceName like '%juniper%') and `
        (NOT InstanceName like '%vpn%') and `
        (NOT InstanceName like 'Hyper-V%') and `
        (NOT InstanceName like 'VMware%') and `
        (NOT InstanceName like 'VirtualBox Host-Only%')"
    $ConnectedAdapters =  Get-CimInstance -Class win32_NetworkAdapter -Filter `
        'NetConnectionStatus = 2'
    $VPNAdapters =  Get-CimInstance -Class Win32_NetworkAdapterConfiguration -Filter `
        "Description like '%pangp%' `
        or Description like '%cisco%'  `
        or Description like '%juniper%' `
        or Description like '%vpn%'"
}
else
{
    # Needed this script to work on PowerShell 2.0 (don't ask)
    $WirelessAdapters = Get-WmiObject -Namespace "root\WMI" -Class MSNdis_PhysicalMediumType -Filter `
        'NdisPhysicalMediumType = 9'
    $WiredAdapters = Get-WmiObject -Namespace "root\WMI" -Class MSNdis_PhysicalMediumType -Filter `
        "NdisPhysicalMediumType = 0 and `
        (NOT InstanceName like '%pangp%') and `
        (NOT InstanceName like '%cisco%') and `
        (NOT InstanceName like '%juniper%') and `
        (NOT InstanceName like '%vpn%') and `
        (NOT InstanceName like 'Hyper-V%') and `
        (NOT InstanceName like 'VMware%') and `
        (NOT InstanceName like 'VirtualBox Host-Only%')"
    $ConnectedAdapters = Get-WmiObject -Class win32_NetworkAdapter -Filter `
        'NetConnectionStatus = 2'
    $VPNAdapters = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter `
        "Description like '%pangp%' `
        or Description like '%cisco%'  `
        or Description like '%juniper%' `
        or Description like '%vpn%'"
}

# Gather computer and user info here. 

# Date
$Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Username
$Username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

Foreach($Adapter in $ConnectedAdapters) {
    If($WirelessAdapters.InstanceName -contains $Adapter.Name)
    {
        $WirelessConnected = $true
    }
}

Foreach($Adapter in $ConnectedAdapters) {
    If($WiredAdapters.InstanceName -contains $Adapter.Name)
    {
        $WiredConnected = $true
    }
}

Foreach($Adapter in $ConnectedAdapters) {
    If($VPNAdapters.Index -contains $Adapter.DeviceID)
    {
        $VPNConnected = $true
    }
}

If(($WirelessConnected -ne $true) -and ($WiredConnected -eq $true)){ $ConnectionType="WIRED"}
If(($WirelessConnected -eq $true) -and ($WiredConnected -eq $true)){$ConnectionType="WIRED AND WIRELESS"}
If(($WirelessConnected -eq $true) -and ($WiredConnected -ne $true)){$ConnectionType="WIRELESS"}
If($VPNConnected -eq $true){$ConnectionType="VPN"}

# Function to write the inventory to the CSV file
function OutputToCSV {
  # CSV properties

  Write-Host "Adding inventory information to the CSV file..."
  $infoObject = New-Object PSObject
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Date Collected" -Value $Date
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Hostname" -Value $env:computername
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "User" -Value $Username
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Serial Number/Service Tag" -Value $SN
  Add-Member -InputObject $infoObject -MemberType NoteProperty -Name "Connection Type" -Value $ConnectionType

  $infoObject
  $infoColl += $infoObject

  # Output to CSV file
  try {
    $infoColl | Export-Csv -Path $csv -NoTypeInformation -Append
    Write-Host -ForegroundColor Green "Inventory was successfully updated!"
    # Clean up empty rows
    (Get-Content $csv) -notlike ",,,,,,,,,,,,,,,,,,,,*" | Set-Content $csv
    exit 0
  }
  catch {
    if (-not (Test-Path $ErrorLogPath))
    {
      New-Item -ItemType "file" -Path $ErrorLogPath
      icacls $ErrorLogPath /grant Everyone:F
    }
    Add-Content -Path $ErrorLogPath -Value "[$Date] $Username at $env:computername was unable to export to the inventory file at $csv."
   throw "Unable to export to the CSV file. Please check the permissions on the file."
    exit 1
  }
}

# Just in case the inventory CSV file doesn't exist, create the file and run the inventory.
if (-not (Test-Path $csv))
{
  Write-Host "Creating CSV file..."
  try {
    New-Item -ItemType "file" -Path $csv
    icacls $csv /grant Everyone:F
    OutputToCSV
  }
  catch {
    if (-not (Test-Path $ErrorLogPath))
    {
      New-Item -ItemType "file" -Path $ErrorLogPath
      icacls $ErrorLogPath /grant Everyone:F
    }
    Add-Content -Path $ErrorLogPath -Value "[$Date] $Username at $env:computername was unable to create the inventory file at $csv."
    throw "Unable to create the CSV file. Please check the permissions on the file."
    exit 1
  }
}

# Check to see if the CSV file exists then run the script.
function Check-IfCSVExists {
  Write-Host "Checking to see if the CSV file exists..."
  $import = Import-Csv $csv
  if ($import -match $env:computername)
  {
    try {
      (Get-Content $csv) -notmatch $env:computername | Set-Content $csv
      OutputToCSV
    }
    catch {
      if (-not (Test-Path $ErrorLogPath))
      {
        New-Item -ItemType "file" -Path $ErrorLogPath
        icacls $ErrorLogPath /grant Everyone:F
      }
      Add-Content -Path $ErrorLogPath -Value "[$Date] $Username at $env:computername was unable to import and/or modify the inventory file located at $csv."
      throw "Unable to import and/or modify the CSV file. Please check the permissions on the file."
      exit 1
    }
  }
  else
  {
    OutputToCSV
  }
}

Check-IfCSVExists
