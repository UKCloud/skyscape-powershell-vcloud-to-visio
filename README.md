# skyscape-powershell-vcloud-to-visio
Draws a vCloud Organisation in Microsoft Visio

This PowerShell script will survey your vCloud Organisation via the API using light weight Search-Cloud and direct API commands (This is VIEW data only, no changes are made), to create an in memory data object that depicts your vCloud Organisation. This is then used with Microsoft VISIO COM automation to draw the architecture for you.

Areas included in the structure are:

 - Virtual Datacenters
 - VAPPS
 - VM's - Including all Properties, stored in the shape data
 - VAPP Networks
 - OrgVDC Networks
 - Network Connections & IP Addresses Between VM NIC's and Networks
 - Network Connections between Networks and vShield Edges
 - vShield Edges
 - IPSEC VPN's between any vShield Edges within the Same Org

# Screenshots

 - Summary
![alt tag](https://github.com/skyscape-cloud-services/skyscape-powershell-vcloud-to-visio/blob/master/screenshots/Summary.jpg)
 - Server Properties
![alt tag](https://github.com/skyscape-cloud-services/skyscape-powershell-vcloud-to-visio/blob/master/screenshots/ServerData.PNG)
 - Network Properties
![alt tag](https://github.com/skyscape-cloud-services/skyscape-powershell-vcloud-to-visio/blob/master/screenshots/NetworkData.PNG)


# Requirements

To run the script, you will need PowerCLI installed with the vCloud extensions:

https://www.vmware.com/support/developer/PowerCLI/

You will also need Microsoft Visio installed - This was built using Visio 2013, so I am unsure how any different versions will respond.

You will also need to know:

 - The URL for your vCloud API endpoint - for example, Skyscape = api.vcd.portal.skyscapecloud.com
Please note, you do not require HTTP or HTTPS at the start of this.
 - Your vCloud Organisation Name/ID
 - Your vCloud Organisation UserName
 - Your vCloud Organisation Password

Note: For Skyscape customers, this is available from the Portal via the API dropdown option in the top right menu.

Update the variables at the top of the script:

 - $Global:API = "api.vcd.portal.skyscapecloud.com"
 - $Global:Username = "your_org_user_name"
 - $Global:Password = "your_org_password"
 - $Global:Org = "your_org"
 
Optionally, you can now also add filters for VDC's and VAPPS to do this, edit the following filter variables:

 - $Global:VAPPFilter = "*myvappname*"
 - $Global:VDCFilter = "*myvdcname*"

Save the file

Open PowerShell and run the PS1 file.

License and Authors
-------------------
Authors:
  * Peter Rossi (prossi@skyscapecloud.com)

Copyright 2016 Skyscape Cloud Services

Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with the License. You may obtain a copy of the License at

http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the specific language governing permissions and limitations under the License.

