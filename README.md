# README.md
powershell script to detect whether the clients where connected via Wired, Wireless (WiFi) and/or VPN.

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

