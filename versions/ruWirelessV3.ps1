## Set up command line switches and what variables they map to.
[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True)]
    [Alias("List")]
    [ValidateScript({Test-Path -Path $_ -PathType Leaf})]
    [string]$ServerFile,
    [Parameter(Mandatory=$True)]
    [Alias("O")]
    [ValidateScript({Test-Path $_ -PathType 'Container'})]
    [string]$OutputPath,
    [Alias("DiskAlert")]
    [ValidateRange(0,100)]
    [int]$DiskAlertThreshold,
    [Alias("CpuAlert")]
    [ValidateRange(0,100)]
    [int]$CpuAlertThreshold,
    [Alias("MemAlert")]
    [ValidateRange(0,100)]
    [int]$MemAlertThreshold,
    [Alias("Refresh")]
    [ValidateRange(300,28800)]
    [int]$RefreshTime,
    [switch]$Light,
    [switch]$csv,
    [alias("Subject")]
    [string]$MailSubject,
    [Alias("SendTo")]
    [string]$MailTo,
    [Alias("From")]
    [string]$MailFrom,
    [Alias("Smtp")]
    [string]$SmtpServer,
    [Alias("User")]
    [string]$SmtpUser,
    [Alias("Pwd")]
    [string]$SmtpPwd,
    [switch]$UseSsl)

## Function to get the up time from a server.
Function Get-UpTime
{
    param([string] $LastBootTime)
    $Uptime = (Get-Date) - [System.Management.ManagementDateTimeconverter]::ToDateTime($LastBootTime)
    "$($Uptime.Days) days $($Uptime.Hours)h $($Uptime.Minutes)m"
}

# Date
$Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss" #done

# Serial Number
$SN = Get-WmiObject -Class Win32_Bios | Select-Object -ExpandProperty SerialNumber

# Username
$Username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
         
## Begining of the loop. At the bottom of the script the loop is broken if the refresh option is not configured.
Do
{
    ## If CSV is configured, setting the location and name of the report output. If CSV is not configured output a HTML file.
    If ($csv)
    {
        $OutputFile = "$OutputPath\WinServ-Status-Report.csv"
        
        ## If the CSV file already exists, clear it
        $csvT = Test-Path -Path $OutputFile

        If ($csvT)
        {
            Clear-Content -Path $OutputFile
        }
    }

    Else
    {
        $OutputFile = "$OutputPath\WinServ-Status-Report.htm"
    }

    $ServerList = Get-Content $ServerFile
    $Result = @()

    ## Using variables for HTML and CSS so we don't need to use escape characters below.
    $Green = "00e600"
    $Grey = "e6e6e6"
    $Red = "ff4d4d"
    $Black = "1a1a1a"
    $Yellow = "ffff4d"
    $CssError = "error"
    $CssFormat = "format"
    $CssSpinner = "spinner"
    $CssRect1 = "rect1"
    $CssRect2 = "rect2"
    $CssRect3 = "rect3"
    $CssRect4 = "rect4"
    $CssRect5 = "rect5"

    ## Sort Servers based on whether they are online or offline
    $ServerList = $ServerList | Sort-Object

    ForEach ($ServerName in $ServerList)
    {
        $PingStatus = Test-Connection -ComputerName $ServerName -Count 1 -Quiet

        If ($PingStatus -eq $False)
        {
            $ServersOffline += @($ServerName)
        }

        Else
        {
            $ServersOnline += @($ServerName)
        }
    }

    $ServerListFinal = $ServersOffline + $ServersOnline

    ## Look through the final servers list.
    ForEach ($ServerName in $ServerListFinal)
    {
        $PingStatus = Test-Connection -ComputerName $ServerName -Count 1 -Quiet

        ## If server responds, get the stats for the server.
        If ($PingStatus)
        {
            $CpuAlert = $false
            $MemAlert = $false
            $DiskAlert = $false
            $OperatingSystem = Get-WmiObject Win32_OperatingSystem -ComputerName $ServerName

            #$DateCollected =Get-Date -Format "MM/dd/yyyy HH:mm:ss"
            #$SN = Get-WmiObject -Class Win32_Bios | Select-Object -ExpandProperty SerialNumber
            #$Username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
            #$WirelessConnected = $null
            #$WiredConnected = $null

            $CpuUsage = Get-WmiObject Win32_Processor -Computername $ServerName | Measure-Object -Property LoadPercentage -Average | ForEach-Object {$_.Average; If($_.Average -ge $CpuAlertThreshold){$CpuAlert = $True};}
            $Uptime = Get-Uptime($OperatingSystem.LastBootUpTime)
            $MemUsage = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ServerName | ForEach-Object {“{0:N0}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) * 100)/ $_.TotalVisibleMemorySize); If((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory) * 100)/ $_.TotalVisibleMemorySize -ge $MemAlertThreshold){$MemAlert = $True};}
            $DiskUsage = Get-WmiObject Win32_LogicalDisk -ComputerName $ServerName | Where-Object {$_.DriveType -eq 3} | Foreach-Object {$_.DeviceID, [Math]::Round((($_.Size - $_.FreeSpace) * 100)/ $_.Size); If([Math]::Round((($_.Size - $_.FreeSpace) * 100)/ $_.Size) -ge $DiskAlertThreshold){$DiskAlert = $True};}
	    }
	
        ## Put the results together in an array.
        $Result += New-Object PSObject -Property @{
	        ServerName = $ServerName
		    Status = $PingStatus
            CpuUsage = $CpuUsage
            CpuAlert = $CpuAlert
		    Uptime = $Uptime
            MemUsage = $MemUsage
            MemAlert = $MemAlert
            DiskUsage = $DiskUsage
            DiskAlert = $DiskAlert

            DateCollected = $DateCollected
            SN = $SN
            Username = $Username
            ConnectionType = $ConnectionType

	    }

####

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

####

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
        
        ## Clear the variables after obtaining and storing the results, otherwise data is duplicated.
        If ($ServerListFinal)
        {
            Clear-Variable ServerListFinal
        }

        If ($ServersOffline)
        {
            Clear-Variable ServersOffline
        }

        If ($ServersOnline)
        {
            Clear-Variable ServersOnline
        }

        If ($PingStatus)
        {
            Clear-Variable PingStatus
        }

        If ($CpuUsage)
        {
            Clear-Variable CpuUsage
        }

        If ($Uptime)
        {
            Clear-Variable Uptime
        }

        If ($MemUsage)
        {
            Clear-Variable MemUsage
        }

        If ($DiskUsage)
        {
            Clear-Variable DiskUsage
        }

       if($DateCollected)
       {
           Clear-Variable DateCollected
       }

       if($SN)
       {
           Clear-Variable SN
       }

       if($Username)
       {
           Clear-Variable Username
       }

      if($ConnectionType)
      {
        Clear-Variable ConnectionType
      }

    }

    ## If there is a result put the report together.
    If ($Null -ne $Result)
    {
        ## If CSV report is specified, output a CSV file. If CSV is not configured output a HTML file.
        If ($csv)
        {
            ForEach($Entry in $Result)
            {
                If ($Entry.Status -eq $True)
                {
                    Add-Content -Path "$OutputFile" -Value "$($Entry.DateCollected),$($Entry.ServerName),$($Entry.SN),$($Entry.Username),$($Entry.ConnectionType),Online,CPU: $($Entry.CpuUsage),Mem: $($Entry.MemUsage),$($Entry.DiskUsage),$($Entry.Uptime)"
                }

                Else
                {
                    Add-Content -Path "$OutputFile" -Value "$($Entry.ServerName),Offline"
                }
            }
        }

        Else
        {
            ## If the light theme is specified, use a lighter css theme. If not, use the dark css theme.
            If ($Light)
            {
                $HTML = '<style type="text/css">
                    p {font-family:Gotham, "Helvetica Neue", Helvetica, Arial, sans-serif;font-size:14px}
                    p {color:#000000;}
                    #Header{font-family:Gotham, "Helvetica Neue", Helvetica, Arial, sans-serif;width:100%;border-collapse:collapse;}
                    #Header td, #Header th {font-size:14px;text-align:left;}
                    #Header tr.alt td {color:#ffffff;background-color:#404040;}
                    #Header tr:nth-child(even) {background-color:#404040;}
                    #Header tr:nth-child(odd) {background-color:#737373;}
                    body {background-color: #d9d9d9;}
                    .spinner {width: 40px;height: 20px;font-size: 14px;padding: 5px;}
                    .spinner > div {background-color: #00e600;height: 100%;width: 3px;display: inline-block;animation: sk-stretchdelay 3.2s infinite ease-in-out;}
                    .spinner .rect2 {animation-delay: -3.1s;}
                    .spinner .rect3 {animation-delay: -3.0s;}
                    .spinner .rect4 {animation-delay: -2.9s;}
                    .spinner .rect5 {animation-delay: -2.8s;}
                    @keyframes sk-stretchdelay {0%, 40%, 100% {transform: scaleY(0.4);} 20% {transform: scaleY(1.0);}}
                    .format {position: relative;overflow: hidden;padding: 5px;}
                    .error {-webkit-animation-name: alert;animation-duration: 4s;animation-iteration-count: infinite;animation-direction: alternate;padding: 5px;}
                    @keyframes alert {from {background-color:rgba(117,0,0,0);} to {background-color:rgba(117,0,0,1);}}
                    </style>
                    <head><meta http-equiv="refresh" content="300"></head>'

                $HTML += "<html><body>
                    <p><font color=#$Black>Last update: $(Get-Date -Format G)</font></p>
                    <table border=0 cellpadding=0 cellspacing=0 id=header>"
            }

            ## If the light theme is not specified, use a darker css theme.
            Else
            {
                $HTML = '<style type="text/css">
                    p {font-family:Gotham, "Helvetica Neue", Helvetica, Arial, sans-serif;font-size:14px}
                    p {color:#ffffff;}
                    #Header{font-family:Gotham, "Helvetica Neue", Helvetica, Arial, sans-serif;width:100%;border-collapse:collapse;}
                    #Header td, #Header th {font-size:14px;text-align:left;}
                    #Header tr:nth-child(even) {background-color:#0F0F0F;}
                    #Header tr:nth-child(odd) {background-color:#1B1B1B;}
                    body {background-color: #0F0F0F;}
                    .spinner {width: 40px;height: 20px;font-size: 14px;padding: 5px;}
                    .spinner > div {background-color: #00e600;height: 100%;width: 3px;display: inline-block;animation: sk-stretchdelay 3.2s infinite ease-in-out;}
                    .spinner .rect2 {animation-delay: -3.1s;}
                    .spinner .rect3 {animation-delay: -3.0s;}
                    .spinner .rect4 {animation-delay: -2.9s;}
                    .spinner .rect5 {animation-delay: -2.8s;}
                    @keyframes sk-stretchdelay {0%, 40%, 100% {transform: scaleY(0.4);} 20% {transform: scaleY(1.0);}}
                    .format {position: relative;overflow: hidden;padding: 5px;}
                    .error {animation-name: alert;animation-duration: 4s;animation-iteration-count: infinite;animation-direction: alternate;padding: 5px;}
                    @keyframes alert {from {background-color:rgba(117,0,0,0);} to {background-color:rgba(117,0,0,1);}}
                    </style>
                    <head><meta http-equiv="refresh" content="300"></head>'

                $HTML += "<html><body>
                    <p><font color=#$Grey>Last update: $(Get-Date -Format G)</font></p>
                    <table border=0 cellpadding=0 cellspacing=0 id=header>"
            }

            ## Highlight the alerts if the alerts are triggered.
            ForEach($Entry in $Result)
            {
                If ($RefreshTime -ne 0)
                {

                    If ($Entry.Status -eq $True)
                    {
                        $HTML += "<td><div class=$CssSpinner><div class=$CssRect1></div> <div class=$CssRect2></div> <div class=$CssRect3></div> <div class=$CssRect4></div> <div class=$CssRect5></div></div></td>"
                    }
                

                    Else
                    {
                        $HTML += "<td><div class=$CssError><font color=#$Red>OFFL</font></div></td>"
                    }
                }

                #
                If ($Entry.Status -eq $True)
                {
                    $HTML += "<td><div class=$CssFormat><font color=#$Green>$($Entry.DateCollected)</font></div></td>"
                }

                Else
                {
                    $HTML += "<td><div class=$CssError><font color=#$Red>$($Entry.DateCollected)</font></div></td>"
                }
                
                If ($Entry.Status -eq $True)
                {
                    $HTML += "<td><div class=$CssFormat><font color=#$Green>$($Entry.ServerName)</font></div></td>"
                }

                Else
                {
                    $HTML += "<td><div class=$CssError><font color=#$Red>$($Entry.ServerName)</font></div></td>"
                }

                If ($Entry.Status -eq $True)
                {
                    $HTML += "<td><div class=$CssFormat><font color=#$Green>$($Entry.SN)</font></div></td>"
                }

                Else
                {
                    $HTML += "<td><div class=$CssError><font color=#$Red>$($Entry.SN)</font></div></td>"
                }

                If ($Entry.Status -eq $True)
                {
                    $HTML += "<td><div class=$CssFormat><font color=#$Green>$($Entry.Username)</font></div></td>"
                }

                Else
                {
                   # $HTML += "<td><div class=$CssError><font color=#$Red>$($Entry.Username)</font></div></td>"
                }

                ###
                If ($Entry.Status -eq $true)
                {
                    $HTML += "<td><div class=$CssFormat><font color=#$Green>$($Entry.ConnectionType)</font></div></td>"
                }

                Else
                {
                    $HTML += "<td><div class=$CssError><font color=#$Red>$($Entry.ConnectionType)</font></div></td>"
                }
                ###

                If ($Null -ne $Entry.CpuUsage)
                {
                    If ($Entry.CpuAlert -eq $True)
                    {
                        $HTML += "<td><div class=$CssFormat><font color=#$Yellow>CPU: $($Entry.CpuUsage)%</font></div></td>"
                    }

                    Else
                    {
                        $HTML += "<td><div class=$CssFormat><font color=#$Green>CPU: $($Entry.CpuUsage)%</font></div></td>"
                    }
                }
            
                Else
                {
                    $HTML += "<td><div class=$CssError><font color=#$Red>OFFL</font></div></td>"
                }

                If ($Null -ne $Entry.MemUsage)
                {
                    If ($Entry.MemAlert -eq $True)
                    {
                        $HTML += "<td><div class=$CssFormat><font color=#$Yellow>Mem: $($Entry.MemUsage)%</font></div></td>"
                    }

                    Else
                    {
                        $HTML += "<td><div class=$CssFormat><font color=#$Green>Mem: $($Entry.MemUsage)%</font></div></td>"
                    }
                }

                Else
                {
                    $HTML += "<td><div class=$CssError><font color=#$Red>OFFL</font></div></td>"
                }

                If ($Null -ne $Entry.DiskUsage)
                {
                    If ($Entry.DiskAlert -eq $True)
                    {
                        $HTML += "<td><div class=$CssFormat><font color=#$Yellow>$($Entry.DiskUsage)%</font></div></td>"
                    }

                    Else
                    {
                        $HTML += "<td><div class=$CssFormat><font color=#$Green>$($Entry.DiskUsage)%</font></div></td>"
                    }
                }

                Else
                {
                    $HTML += "<td><div class=$CssError><font color=#$Red>OFFL</font></div></td>"
                }

                If ($Entry.Status -eq $True)
                {
                    $HTML += "<td><div class=$CssFormat><font color=#$Green>$($Entry.Uptime)</font></div></td>
                            </tr>"
                }

                Else
                {
                    $HTML += "<td><div class=$CssError><font color=#$Red>OFFL</font></div></td>
                            </tr>"
                }
            }

            ## Finish the HTML file.
            $HTML += "</table></body></html>"

            ## Output the HTML file
            $HTML | Out-File $OutputFile
        }

        ## If email was configured, set the variables for the email subject and body.
        If ($SmtpServer)
        {
            # If no subject is set, use the string below
            If ($Null -eq $MailSubject)
            {
                $MailSubject = "Server Status Report"
            }
        
            $MailBody = Get-Content -Path $OutputFile | Out-String

            ## If an email password was configured, create a variable with the username and password.
            If ($SmtpPwd)
            {
                $SmtpPwdEncrypt = Get-Content $SmtpPwd | ConvertTo-SecureString
                $SmtpCreds = New-Object System.Management.Automation.PSCredential -ArgumentList ($SmtpUser, $SmtpPwdEncrypt)

                ## If ssl was configured, send the email with ssl.
                If ($UseSsl)
                {
                    Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -BodyAsHtml -SmtpServer $SmtpServer -UseSsl -Credential $SmtpCreds
                }

                ## If ssl wasn't configured, send the email without ssl.
                Else
                {
                    Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -BodyAsHtml -SmtpServer $SmtpServer -Credential $SmtpCreds
                }
            }

            ## If an email username and password were not configured, send the email without authentication.
            Else
            {
                Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -BodyAsHtml -SmtpServer $SmtpServer
            }
        }

        ## If the refresh time option is configured, wait the specifed number of seconds then loop.
        If ($RefreshTime -ne 0)
        {
            Start-Sleep -Seconds $RefreshTime
        }
    }
}

## If the refresh time option is not configured, stop the loop.
Until ($RefreshTime -eq 0)

## End