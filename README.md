# README.md
PowerShell script that creates an inventory and detects whether an agent is  connected via Wired, Wireless (WiFi) or both.
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

# To Do
The script needs execute using the list of servers and output a ruWirelss.csv file. That contains the the following info:

"Date Collected","Hostname","User","Serial Number/Service Tag","Connection Type"
"2022-03-30 15:38:51","VAW5VMLFB3","VUMC\roxasrr","5VMLFB3","WIRELESS"

# if an agent computer is wireless perform a speed test using Speedtest by Ookla

     Server: United Communications [TN] - Chapel Hill, TN (id = 17094)
        ISP: Vanderbilt University Medical Center
    Latency:    31.85 ms   (2.13 ms jitter)
   Download:    40.06 Mbps (data used: 57.0 MB)
     Upload:    42.41 Mbps (data used: 63.1 MB)
Packet Loss:     3.4%
 Result URL: https://www.speedtest.net/result/c/fe69ef67-4bad-4259-b3f6-35cd71dca7e6