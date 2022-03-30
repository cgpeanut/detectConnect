# README.md
PowerShell script that creates an inventory, IP and detects whether an agent is  connected via Wired, Wireless (WiFi) or both.
After it collects that information, it is outputted to a CSV file. It will first check the CSV file (if it exists) to see if the hostname already exists in the file. 
If hostname exists in the CSV file, it will overwrite it with the latest information so that the inventory is up to date and there is no duplicate information.
It is designed to be run as a login script and/or a scheduled/immediate task run by a domain user. Elevated privileges are not required.

The script detects the following VPN platforms
Palo Alto GlobalProtect
Cisco
Juniper
Dell VPN (Sonicwall)
F5 Networks VPN 

The script filters out the following virtual platform host switches
Hyper-V
VMware
Virtual Box

Update:  replaced the various where-object lines I had with -Filter

Note: If you want the script to output the result as a task sequence variable in ConfigMgr, in this case setting the results as the TSConnectionType variable, simply change the last line to these lines:

# To Do

## Features and Requirements

* The utility will display the Creation Date,  Hostname, User, Serial Number/Service Tag, and Connection Type. (done)
* The utility has been tested running on Windows 10 (done)
* The utility can display the results as either a CSV file. (done) 

### Generating A Password File

The password used for SMTP server authentication must be in an encrypted text file. To generate the password file, run the following command in PowerShell, on the computer that is going to run the script and logged in with the user that will be running the script. When you run the command you will be prompted for a username and password. Enter the username and password you want to use to authenticate to your SMTP server.

Please note: This is only required if you need to authenticate to the SMTP server when send the log via e-mail.

``` powershell

$creds = Get-Credential
$creds.Password | ConvertFrom-SecureString | Set-Content C:\Users\roxasrr\code\detectConnect\ps-script-pwd.txt
After running the commands, you will have a text file containing the encrypted password. When configuring the -Pwd switch enter the path and file name of this file.

# Configuration

Here’s a list of all the command line switches and example configurations.

The path to a TXT file containing the netbios names of the Agents computers you wish to check.
