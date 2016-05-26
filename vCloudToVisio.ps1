#**********************************************************
# VARIABLES
#**********************************************************


$Global:API = "api.vcd.portal.skyscapecloud.com"
$Global:Username = "your_org_user_name"
$Global:Password = "your_org_password"
$Global:Org = "your_org"
$Global:ShapeMaster = @()

#**********************************************************
# Imports and Settings
#**********************************************************


Import-Module VMware.VimAutomation.Cloud -force | Out-Null
Set-PowerCLIConfiguration -WebOperationTimeoutSeconds -1 -Scope Session -Confirm:$False

#**********************************************************
# Functions
#**********************************************************

Function Get-EdgeConfig ($EdgeGateway)  {  
    $Edgeview = $EdgeGateway | get-ciview
    $webclient = New-Object system.net.webclient
    $webclient.Headers.Add("x-vcloud-authorization",$Edgeview.Client.SessionKey)
    $webclient.Headers.Add("accept",$EdgeView.Type + ";version=5.1")
    [xml]$EGWConfXML = $webclient.DownloadString($EdgeView.href)
    $Holder = "" | Select Firewall,NAT,LoadBalancer,DHCP,VPN,Routing
    $Holder.Firewall = $EGWConfXML.EdgeGateway.Configuration.EdgegatewayServiceConfiguration.FirewallService.FirewallRule
    $Holder.NAT = $EGWConfXML.EdgeGateway.Configuration.EdgegatewayServiceConfiguration.NatService.NatRule
    $Holder.LoadBalancer = $EGWConfXML.EdgeGateway.Configuration.EdgegatewayServiceConfiguration.LoadBalancerService.VirtualServer
    $Holder.DHCP = $EGWConfXML.EdgeGateway.Configuration.EdgegatewayServiceConfiguration.GatewayDHCPService.Pool
	$Holder.VPN = $EGWConfXML.EdgeGateway.Configuration.EdgegatewayServiceConfiguration.GatewayIpsecVpnService
	$Holder.Routing = $EGWConfXML.EdgeGateway.Configuration.EdgeGatewayServiceConfiguration.StaticRoutingService
    Return $Holder
}


Function Add-ServerProperties($Shape,$VMData)
{
	#$PropertyList = @("Name","ID","PowerState","VAPPName","VAPPID","HardwareVersion","CPU","MemoryMB","StorageProfile","OS","Network1_Name","Network1_IP","Network1_MAC","Network1_Type","Network2_Name","Network2_IP","Network2_MAC","Network2_Type","Network3_Name","Network3_IP","Network3_MAC","Network3_Type","Network4_Name","Network4_IP","Network4_MAC","Network4_Type","Network5_Name","Network5_IP","Network5_MAC","Network5_Type")
	$Shape.Cells("Prop.Name").Formula = '"' + $VMData.Name + '"'
	$Shape.Cells("Prop.PowerState").Formula = '"' + $VMData.status + '"'
	$Shape.Cells("Prop.ID").Formula = '"' + $VMData.id + '"'
	$Shape.Cells("Prop.VAPPName").Formula = '"' + $VMData.ContainerName + '"'
	$Shape.Cells("Prop.VAPPID").Formula = '"' + $VMData.Container + '"'
	$Shape.Cells("Prop.HardwareVersion").Formula = '"' + $VMData.HardwareVersion + '"'
	$Shape.Cells("Prop.CPU").Formula = '"' + $VMData.NumberOfCpus + '"'
	$Shape.Cells("Prop.MemoryMB").Formula = '"' + $VMData.MemoryMB + '"'
	$Shape.Cells("Prop.StorageProfile").Formula = '"' + $VMData.StorageProfileName + '"'
	$Shape.Cells("Prop.OS").Formula = '"' + $VMData.GuestOS + '"'
	
	$I = 0
	$NICS = $VMData.virtualHardwareSection.item | ?{$_.ResourceType.Value -eq 10}
	ForEach($N in $NICS)
	{
		$I += 1
		$IP = ($N.Connection.AnyAttr | ?{$_.Name -eq 'vcloud:ipAddress'} | Select Value -first 1).value
		$Shape.Cells("Prop.Network$($I)_Name").Formula = '"' + $N.Connection.Value + '"'
		$Shape.Cells("Prop.Network$($I)_Type").Formula = '"' + $N.ResourceSubType.Value + '"'
		$Shape.Cells("Prop.Network$($I)_IP").Formula = '"' + $IP + '"'
		$Shape.Cells("Prop.Network$($I)_MAC").Formula = '"' + $N.Address + '"'
	}

}

Function Add-NetworkProperties($Shape,$NetworkData)
{
	$Shape.Cells("Prop.ID").Formula = '"' + $NetworkData.ID + '"'
	$Shape.Cells("Prop.Name").Formula = '"' + $NetworkData.Name + '"'
	$Shape.Cells("Prop.IpScopeId").Formula = '"' + $NetworkData.IpScopeId + '"'
	$Shape.Cells("Prop.IsIpScopeInherited").Formula = '"' + $NetworkData.IsIpScopeInherited + '"'
	$Shape.Cells("Prop.Gateway").Formula = '"' + $NetworkData.Gateway + '"'
	$Shape.Cells("Prop.Netmask").Formula = '"' + $NetworkData.Netmask + '"'
	$Shape.Cells("Prop.Dns1").Formula = '"' + $NetworkData.Dns1 + '"'
	$Shape.Cells("Prop.Dns2").Formula = '"' + $NetworkData.Dns2 + '"'
	$Shape.Cells("Prop.DnsSuffix").Formula = '"' + $NetworkData.DnsSuffix + '"'
	$LinkNetworkName = $NetworkData.AnyAttr | ?{$_.Name -eq "linknetworkname"} | Select -First 1
	$LinkStatus = $NetworkData.AnyAttr | ?{$_.Name -eq "isLinked"} | Select -First 1
	if(($LinkNetworkName | Measure-Object).count -gt 0)
	{
		$Shape.Cells("Prop.LinkNetworkName").Formula = '"' + $LinkNetworkName.Value + '"'
		$Shape.Cells("Prop.IsLinked").Formula = '"' + $LinkStatus.Value + '"'
	}
}

Function Map-Org()
{
	if(($global:DefaultCIServers | ?{$_.IsConnected -eq $True} | Measure-Object).count -gt 0)
	{
		write-host "Disconnecting from existing vCloud instances..."
		ForEach($CI in $global:DefaultCIServers | ?{$_.IsConnected -eq $True})
		{
			Write-Host "-Disconnect $($CI.Name)"
			Disconnect-CIServer -Server ($CI.Name) -Confirm:$false
		}
	}
	Write-Host "Connecting to $($Global:API)"
	Connect-CIServer -Server ($Global:API) -User ($Global:Username) -Password ($Global:Password) -Org ($Global:Org) | Out-Null
	Write-Host "Getting ORGS"
	$Orgs = Search-Cloud -QueryType Organization
	Write-Host "Getting VDC's"
	$VDCs = Search-Cloud -QueryType OrgVDC
	Write-Host "Getting VAPP's"
	$VAPPs = Search-Cloud -QueryType VAPP
	Write-Host "Getting Non Template VM's"
	$VMs = Search-Cloud -QueryType VM -filter 'IsVappTemplate==False'
	Write-Host "Getting Edge Gateways"
	$EdgeGWs = Search-Cloud -QueryType EdgeGateway
	Write-Host "Getting Org Networks"
	$OrgNetworks = Search-Cloud -QueryType OrgNetwork
	Write-Host "Getting Org VDC Networks"
	$OrgVDCNetworks = search-cloud -QueryType OrgVDCNetwork
	Write-Host "Getting Catalogs"
	$VappOrgVDCNetworkRelations = Search-Cloud -QueryType VappOrgVdcNetworkRelation
	$VappOrgNetworkRelations = Search-Cloud -QueryType VAppOrgNetworkRelation
	$Catalogs = search-cloud -QueryType Catalog
	Write-Host "Getting Catalog Items"
	$CatalogItems = search-cloud -querytype CatalogItem
	Write-Host "Splitting out vAppTemplates from Media Catalog Items"
	$CatalogItemsTemplates = $CatalogItems | ?{$_.EntityType -ne "media"}
	$CatalogItemsMedia = $CatalogItems | ?{$_.EntityType -eq "media"}
	Write-Host "Getting VApp Templates"
	$Templates = Search-cloud -querytype VAppTemplate
	Write-Host "Getting VM's that are Templates"
	$TemplateVMS = Search-Cloud -QueryType VM -filter 'IsVappTemplate==True'
	Write-Host "Getting additional data"
	$VAPPNetworks = Search-Cloud -QueryType VAPPNetwork
	#VMTree
	$VMSFinal = @()
	Write-Host "Getting VM Hardware,Network,OS details etc..."
	$VMTotal = ($VMs | Measure-Object).count
	$VMI = 0
	ForEach($VM in $VMs)
	{
		$VMI += 1
		$VMP = ($VMI/$VMTotal)*100
		Write-Progress -Status "Getting VM Data..." -Activity "$($VM.Name)" -PercentComplete $VMP -Id 0
		$T = Get-CIView -Id ($VM.Id)
		$Sections = $T.section
		$VM | add-member -membertype NoteProperty -name virtualHardwareSection -Value ($Sections | ?{$_.href -like "*virtualHardwareSection*"}) -force
		$VM | add-member -membertype NoteProperty -name operatingSystemSection -Value ($Sections | ?{$_.href -like "*operatingSystemSection*"}) -force
		$VM | add-member -membertype NoteProperty -name networkConnectionSection -Value ($Sections | ?{$_.href -like "*networkConnectionSection*"}) -force 
		$VM | add-member -membertype NoteProperty -name guestCustomizationSection -Value ($Sections | ?{$_.href -like "*guestCustomizationSection*"}) -force
		$VM | add-member -membertype NoteProperty -name runtimeInfoSection -Value ($Sections | ?{$_.href -like "*runtimeInfoSection*"}) -force
		$VMSFinal += $VM
	}
	
	Write-Host "Getting vShield Edge Configs"
	$EGWTotal = ($EdgeGWs | Measure-Object).count
	$EGWI = 0
	$EGWFinal = @()
	ForEach($E in $EdgeGWs)
	{
		$EGWI += 1
		$EGWP = ($EGWI/$EGWTotal)*100
		Write-Progress -Status "Getting Edge Data" -Activity "$($E.id)" -PercentComplete $EGWP -Id 1
		$EGW = Get-CIView -Id ($E.id)
		$DetailedConfiguration = Get-EdgeConfig $E
		$E | Add-Member -MemberType NoteProperty -Name Configuration -Value ($EGW.Configuration) -Force
		$E | Add-Member -MemberType NoteProperty -Name DetailedConfiguration -Value $DetailedConfiguration -Force
		
		$EGWFinal += $E
	}
	
	Write-Host "Processing data tree..."
	
	$FinalObj = "" | Select OrgObject,OrgJson,OrgJSONNoCatalog,ORGs,VDCs,Vapps,VMs,EdgeGws,OrgNetworks,OrgVDCNetworks,Catalogs,CatalogItems,CatalogItemTemplates,CatalogItemMedia,Templates,TemplateVMs,VappOrgVDCNetworkRelations,VappOrgNetworkRelations,VappNetworks
	$FinalObj.Orgs = $Orgs
	$FinalObj.VDCs = $VDCs
	$FinalObj.Vapps = $VAPPs
	$FinalObj.VMs = $VMsFinal
	$FinalObj.EdgeGws = $EGWFinal
	$FinalObj.OrgNetworks = $OrgNetworks
	$FinalObj.OrgVDCNetworks = $OrgVDCNetworks
	$FinalObj.Catalogs = $Catalogs
	$FinalObj.CatalogItems = $CatalogItems
	$FinalObj.CatalogItemTemplates = $CatalogItemsTemplates
	$FinalObj.CatalogItemMedia = $CatalogItemsMedia
	$FinalObj.Templates = $Templates
	$FinalObj.TemplateVMs = $TemplateVMS
	$FinalObj.VappOrgVDCNetworkRelations = $VappOrgVDCNetworkRelations
	$FinalObj.VappOrgNetworkRelations = $VappOrgNetworkRelations
	$FinalObj.VAPPNetworks = $VAPPNetworks
	
	Return $FinalObj
}

Function Save-ShapeLookup($Shape)
{
	$holder = "" | Select ID,Name,Text,PinX,PinY
	$holder.ID = $Shape.ID
	$holder.Name = $Shape.Name
	$holder.Text = $Shape.Text
	$holder.PinX = $Shape.PinX
	$holder.PinY = $Shape.PinY
	$Global:ShapeMaster += $holder
	Return $null	
}

Function Get-ShapeFromLookup($Page,$NameFilter,$TextFilter)
{
if($NameFilter)
{
	$Lookup = $Global:ShapeMaster | ?{$_.Name -like "*$NameFilter*"}
}
else
{
	$Lookup = $Global:ShapeMaster | ?{$_.Text -like "*$TextFilter*"}
}
$Final = @()
ForEach($Item in $Lookup)
{
	$Shape = $Page.Shapes.ItemFromID($Item.ID)
	$Final += $Shape

}
Return $Final
}

function Connect-VisioObject ($Page,$FirstObj, $SecondObj)
{
	$Connector = $Page.Drop($Page.Application.ConnectorToolDataObject, 0, 0)
	$connectBegin = $Connector.CellsU("BeginX").GlueTo($FirstObj.CellsU("PinX"))
	$connectEnd = $Connector.CellsU("EndX").GlueTo($SecondObj.CellsU("PinX"))
	Return $Connector
}


Function Draw-Vapp($page,$vms,$mastercontainer,$server,$data,$networkmaster,$DynamicConnectorMaster,$startx,$Starty)
{
	$networks = get-vappnetworks -data $data -vappid $vms[0].container
	write-host "Drawing VAPP $($vms[0].container) - $($vms[0].ContainerName)"
	$startxmm = $startx * 25.4
	$startymm = $Starty * 25.4
	write-host "Drawing Container at X: $startx Y: $starty"
	Write-Host "X - $startxmm"
	Write-Host "Y - $startymm"
	$container = $page.Drop($mastercontainer,$startx,$starty)
	#$GH = Read-Host "Shape Dropped"
	($container.cells("LocPinX")).formula = "=Width*0"
	($container.cells("LocPinY")).formula = "=Height*1"
	$XDrop = $container.cells("PinX").ResultIU * 25.4
	$YDrop = $container.cells("PinY").ResultIU * 25.4
	Write-Host "Container was dropped at X: $XDrop Y: $YDrop"
	$container.text = $vms[0].Containername
	$container.name = $vms[0].Container
	$SavedShape = Save-ShapeLookup -Shape $Container
	$X = $startx
	$MaxHeight = 0
	$firstserver = $null
	$firstnetwork = $null
	
	$VMCount = ($vms | Measure-Object).count
	$VMI = 0
		
	ForEach($VM in $VMS)
	{
		$VMI += 1
		$VMP = 	($VMI/$VMCount)*100
		Write-Progress -Status "Processing VAPP - $($Container.Text)" -Activity "$($VM.Name)-$($VM.id)" -PercentComplete $VMP -Id 2
		$holder = $page.Drop($Server,$X,$starty)
		Add-ServerProperties -Shape $holder -VMData $VM
		#$GH = Read-Host "Shape Dropped"
		if($firstserver -eq $null)
		{
			$firstserver = $holder
		}
		else
		{
			$holder.cells("PinY").ResultIU = $firstserver.cells("PinY").ResultIU
		}
		Write-Host "Drawing server at X: $X Y: $starty"
		$widthcell = ($holder.cells("TxtWidth")).formula = "=Width*1.6116"
		if($VM.Status -eq "POWERED_ON")
		{
			($Holder.cells("TextBkgnd")).formula = "=RGB(0,255,0)"
		}
		else
		{
			($Holder.cells("TextBkgnd")).formula = "=RGB(255,0,0)"
		}
		
		$holder.text = $VM.name
		$holder.name = $VM.id
		$SavedShape = Save-ShapeLookup -Shape $Holder
		$container.containerproperties.addmember($holder,1)
		$Temp = $holder.cells("TxtHeight").ResultIU + $holder.cells("Height").ResultIU
		if($Temp -gt $MaxHeight)
		{
			$MaxHeight = $Temp
		}
		$X += 1.5
	}
	$X = $startx-2
	$Y = $Starty-2
	$VMsOnPage = Get-ShapeFromLookup -Page $Page -NameFilter ":VM:"
	Write-Host "Processing VAPPNetworks"
	$VAPPNetworkCount = ($networks.VAPPNetworks | Measure-Object).count
	$VAPPNetworkI = 0
	ForEach($Network in $networks.VAPPNetworks)
	{
		$VAPPNetworkI +=1
		$VAPPNetworkP = ($VAPPNetworkI/$VAPPNetworkCount)*100
		Write-Progress -Status "Processing VAPP Networks" -Activity "$($Network.Id)" -PercentComplete $VAPPNetworkP -Id 3
		$holder = $page.Drop($networkmaster,$X,$Y)
		Add-NetworkProperties -Shape $Holder -NetworkData $Network
		#$GH = Read-Host "Shape Dropped"
		Write-Host "Drawing Network at X: $X Y: $Y"
		#$b = Read-Host "Dropping Network"
		$Y = $Y - 1.5
		$holder.name = $Network.Id + "-" + $vms[0].Container
		Write-Host "Setting VAPP NetworkID to $($holder.name)"
		$holder.text = $Network.Name
		$SavedShape = Save-ShapeLookup -Shape $holder
		$container.containerproperties.addmember($holder,1)
		
		$ConnectedVMs = @()
		ForEach($VM in $vms)
		{
			Write-Host "Looking for Network $($Network.OrgVdcNetworkName)"
			$NetworkName = $VM.NetworkConnectionSection.NetworkConnection | ?{$_.Network -eq $Network.Name} | Select -First 1
			if(($NetworkName | Measure-Object).count -gt 0)
			{
				$TheShape = $VMsOnPage | ?{$_.Name -eq $VM.id}
				#$theshape.autoconnect($holder,0,$DynamicConnectorMaster)
				#$con = $Page.Shapes[-1]
				$con = Connect-VisioObject -Page $Page -FirstObj $TheShape -SecondObj $Holder
				$con.Text = $NetworkName.IPAddress# + "`r`n" + $NetworkName.Macaddress + "`r`nMode: " + $NetworkName.IpAddressAllocationMode + "`r`nConnected: " + $NetworkName.IsConnected
				$con.Cells("TxtPinX").Formula = "= POINTALONGPATH( Geometry1.Path, 0 )"
				$con.Cells("TxtPinY").Formula = "= POINTALONGPATH( Geometry1.Path, 0.1 )-1"
				$SavedShape = Save-ShapeLookup -Shape $con
				$ConnectedVMs += $TheShape
			}
		}
		
		$lastnetworktextheight = $holder.cells("TxtHeight").ResultIU
	}
	
	$VAPPNetworkCount = ($networks.VDCNetworks | Measure-Object).count
	$VAPPNetworkI = 0
	ForEach($Network in $networks.VDCNetworks)
	{
		$VAPPNetworkI +=1
		$VAPPNetworkP = ($VAPPNetworkI/$VAPPNetworkCount)*100
		Write-Progress -Status "Processing VDC Networks" -Activity "$($Network.Id)" -PercentComplete $VAPPNetworkP -Id 4
		$holder = $page.Drop($networkmaster,$X,$Y)
		Write-Host "Drawing Network at X: $X Y: $Y"
		$Y = $Y - 1.5
		$holder.name = $Network.OrgVdcNetwork + "-" + $vms[0].Container
		Write-Host "Setting VAPP NetworkID to $($holder.name)"
		$holder.text = $Network.OrgVdcNetworkName
		$container.containerproperties.addmember($holder,1)
		
		#Check to see if it's a routed network
		write-host "Looking for $($Network.OrgVDCNetworkName) in VAPPNetworks"
		$Match = $networks.VAPPNetworks | ?{$_.Name -eq $Network.OrgVDCNetworkName} | Select -First 1
		Write-Host "Checking to see if this is a routed networking"
		$SavedShape = Save-ShapeLookup -Shape $holder
		if(($Match | Measure-Object).count -gt 0)
		{
			Write-Host "It Matches"
			$LinkType = ($Match.AnyAttr | ?{$_.Name -eq "linkType"} | Select Value).value
			if($LinkType -eq 0)
			{
				Write-Host "LinkType is 0"
				$InternalObj = Get-ShapeFromLookup -Page $Page -NameFilter "$($Match.Id)-$($vms[0].Container)"
				$con = Connect-VisioObject -Page $Page -FirstObj $InternalObj -SecondObj $Holder
				$SavedShape = Save-ShapeLookup -Shape $con
			}
			else
			{
				Write-Host "LinkType is 1"
			}
		}
		else
		{
			$ConnectedVMs = @()
			ForEach($VM in $vms)
			{
				Write-Host "Looking for Network $($Network.OrgVdcNetworkName)"
				$NetworkName = $VM.NetworkConnectionSection.NetworkConnection | ?{$_.Network -eq $Network.OrgVdcNetworkName} | Select -First 1
				if(($NetworkName | Measure-Object).count -gt 0)
				{
					$TheShape = $VMsOnPage | ?{$_.Name -eq $VM.id}
					$con = Connect-VisioObject -Page $Page -FirstObj $TheShape -SecondObj $Holder
					$con.Text = $NetworkName.IPAddress
					$con.Cells("TxtPinX").Formula = "= POINTALONGPATH( Geometry1.Path, 0 )"
					$con.Cells("TxtPinY").Formula = "= POINTALONGPATH( Geometry1.Path, 0 )-1"
					$SavedShape = Save-ShapeLookup -Shape $con
					$ConnectedVMs += $TheShape
				}
			}
		}
		$lastnetworktextheight = $holder.cells("TxtHeight").ResultIU
	}
	
	$MaxHeight += 1
	Write-Host "MaxHeight = $MaxHeight"
	$container.containerproperties.fittocontents()
	if(($container.cells("Width").ResultIU *25.4) -lt 50)
	{
		$container.cells("Width").ResultIU = (50/25.4)
	}
	
	return $container
}

Function Draw-GW($page,$data,$DynamicConnectorMaster,$FWMaster,$mastercontainer)
{
	$Y = 4
	$X = 1
	$Grouped = $data.EdgeGWs | Group VDC
	$AllVDCContainers = Get-ShapeFromLookup -Page $Page -NameFilter ":VDC:"
	$AllNetworks = Get-ShapeFromLookup -Page $Page -NameFilter ":NETWORK:"
	ForEach($Group in $Grouped)
	{
		$ThisVDCContainer = $AllVDCContainers | ?{$_.Name -eq $Group.Name}
		if(($ThisVDCContainer | Measure-Object).count -gt 0)
		{
			$VAPPContainerIDs = $ThisVDCContainer.ContainerProperties.GetMemberShapes(16)
			$VAPPContainers = @()
			ForEach($ID in $VAPPContainerIDs)
			{
				$VAPPContainers += $page.Shapes.ItemFromID($ID)
			}
			if(($VAPPContainers | Measure-Object).count -gt 0)
			{
				$MostLeft = $VAPPContainers | %{$_.Cells("PinX").ResultIU} | Sort -Descending | Select -Last 1
				$CurrentY = $VAPPContainers | %{$_.Cells("PinY").ResultIU} | Sort -Descending | Select -Last 1
				$Tallest = $VAPPContainers | %{$_.Cells("Height").ResultIU} | Sort -Descending | Select -First 1
			}
			else
			{
				$MostLeft = $ThisVDCContainer.Cells("PinX").ResultIU
				$CurrentY = $ThisVDCContainer.Cells("PinY").ResultIU
				$Tallest = $ThisVDCContainer.Cells("Height").ResultIU
			}
			
			$GWContainerX = $MostLeft
			$GWContainerY = $CurrentY - $Tallest - (50/25.4)
			Write-Host "GWContainerX = $GWContainerX GWContainerY = $GWContainerY"
			
			$GWContainer = $page.Drop($mastercontainer,$GWContainerX,$GWContainerY)
			$GWContainer.cells("LocPinX").formula = "=Width*0"
			$GWContainer.cells("LocPinY").formula = "=Height*1"
			$GWContainer.text = "vShield Edge Gateways"
			$SavedShape = Save-ShapeLookup -Shape $GWContainer
			$GWStartY = $GWContainer.cells("PinY").ResultIU + 1
			$GWStartX = $GWContainer.cells("PinX").ResultIU + 1
			
			ForEach($GW in $Group.Group)
			{
				$FW = $page.Drop($FWMaster,$GWStartX,$GWStartY)
				$GWContainer.containerproperties.addmember($FW,1)
				$FW.Text = $GW.Name
				$FW.Name = $GW.id
				($FW.cells("TxtWidth")).formula = "=Width*1.6116"
				$SavedShape = Save-ShapeLookup -Shape $FW
				$UplinkFW = $GW.Configuration.GatewayInterfaces.GatewayInterface | ?{$_.InterfaceType -eq "uplink"} | Select -First 1
				if(($UplinkFW | Measure-Object).count -gt 0)
				{
					$UFW = $page.drop($FWMaster,$GWStartX,$GWStartY-3)
					$GWContainer.containerproperties.addmember($UFW,1)
					$UFW.Text = $UplinkFW.Name
					$UFW.Name = $UplinkFW.id
					$IPS = $UplinkFW.SubnetParticipation | Select IPAddress
					$IPText = "" 
					$SavedShape = Save-ShapeLookup -Shape $UFW
					ForEach($IP in $IPS)
					{
						$IPText += "$($IP.IPAddress)`r`n"
					}
					$IPText = $IPText.TrimEnd("`r`n")
					$con = Connect-VisioObject -Page $Page -FirstObj $UFW -SecondObj $FW
					$con.text = $IPText
					$SavedShape = Save-ShapeLookup -Shape $Con
				}
				ForEach($Network in $GW.Configuration.GatewayInterfaces.GatewayInterface)
				{
					$ID = $Network.network.Href.Split("/")[-1]
					Write-Host "Looking for Shapes where name is like *network*$($ID)"
					$NetworksToConnect = $AllNetworks | ?{$_.Name -like "*$ID*"}
					ForEach($N in $NetworksToConnect)
					{
						Write-Host "Attempting to connect FW: $($FW.name) to Network $($N.name)"
						$con = Connect-VisioObject -Page $Page -FirstObj $N -SecondObj $FW
						$con.Name = $N.Name
						$SavedShape = Save-ShapeLookup -Shape $Con
					}
				}
				$GWStartX += 1.5
			
			}
			$ThisVDCContainer.ContainerProperties.AddMember($GWContainer,1)
		}
		else
		{
			Write-Host "This VDC doesn't exist..."
		}
	}
}

Function Get-VappNetworks($Data,$VAPPID)
{
	$VDCNetwork = $Data.VappOrgVDCNetworkRelations | ?{$_.id -eq $VAPPID}
	$OrgNetwork = $Data.VappOrgNetworkRelations | ?{$_.id -eq $VAPPID}
	$VAPPNetwork = $Data.VappNetworks | ?{$_.vapp -eq $VAPPID}
	$holder = "" | Select VDCNetworks,OrgNetworks,VAPPNetworks
	$holder.VDCNetworks = $VDCNetwork
	$holder.OrgNetworks = $OrgNetwork
	$holder.VappNetworks = $VAPPNetwork
	Return $Holder
}

Function Enable-DiagramServices($document)
{
$DServiceSet = [Microsoft.Office.Interop.Visio.VisDiagramServices]::visServiceStructureBasic
$document.DiagramServicesEnabled = $DServiceSet
}

Function Disable-DiagramServices($document)
{
$DServiceSet = [Microsoft.Office.Interop.Visio.VisDiagramServices]::visServiceNone
$document.DiagramServicesEnabled = $DServiceSet
}

#**********************************************************
# MAIN PROCESS
#**********************************************************

#Initialise VISIO
$Visio = New-Object -ComObject Visio.Application 
$documents = $Visio.Documents 
$stencilPath=$Visio.GetBuiltInStencilFile(2,0)
$stencil=$Visio.Documents.OpenEx($stencilPath,64)
$documents = $visio.Documents 
$document = $documents.Add("Basic Network Diagram.vst") 
$pages = $visio.ActiveDocument.Pages 
$page = $pages.Item(1) 

#Get Master Stencils
$NetworkStencil = $visio.Documents.Add("periph_m.vss") 
$ConnectorStencil = $visio.Documents.Add("CONNEC_U.VSSX")
$DynamicConnectorMaster = $ConnectorStencil.Masters.Item("Dynamic Connector")
$Server = $NetworkStencil.Masters.Item("Server")
$NetworkMaster = $NetworkStencil.Masters.Item("Ethernet")
$FWMaster = $NetworkStencil.Masters.Item("Firewall")
$MasterContainer = $stencil.Masters["Plain"]

#Edit Server Master Stencil Properties
$eServer = $server.open()
$EditShape = $eServer.Shapes[1]
$CurrentRowCount = $EditShape.RowCount(243)
For($I = 0;$I -le $CurrentRowCount;$I++)
{
	$EditShape.DeleteRow(243,0)
}
$PropertyList = @("Name","ID","PowerState","VAPPName","VAPPID","HardwareVersion","CPU","MemoryMB","StorageProfile","OS","Network1_Name","Network1_IP","Network1_MAC","Network1_Type","Network2_Name","Network2_IP","Network2_MAC","Network2_Type","Network3_Name","Network3_IP","Network3_MAC","Network3_Type","Network4_Name","Network4_IP","Network4_MAC","Network4_Type","Network5_Name","Network5_IP","Network5_MAC","Network5_Type")
ForEach($Prop in $PropertyList)
{
	$NewRow = $EditShape.addnamedrow(243,"$($Prop)",0)
}
$eServer.close()

#Edit Network Master Stencil Properties
$eNetwork = $NetworkMaster.open()
$EditShape = $eNetwork.Shapes[1]
$CurrentRowCount = $EditShape.RowCount(243)
For($I = 0;$I -le $CurrentRowCount;$I++)
{
	$EditShape.DeleteRow(243,0)
}
$PropertyList = @("Name","ID","IPScopeId","IsIpScopeInherited","Gateway","Netmask","Dns1","Dns2","DnsSuffix","LinkNetworkName","IsLinked")
ForEach($Prop in $PropertyList)
{
	$NewRow = $EditShape.addnamedrow(243,"$($Prop)",0)
}
$eNetwork.close()


#Get the data

$Data = Map-Org

#MAP The data

Enable-DiagramServices $Document

$X = 1
$Y = 1
$vdccount = 0
$firstvdccontainer = $null
$vdctotal =($data.vdcs|Measure-Object).count

#Draw VDC's and VAPPS
ForEach($VDC in $data.vdcs)
{
	$vdccount += 1
	$vdcP=($vdccount/$vdctotal)*100
	Write-Progress -Status "Processing VDC $($VDC.Name)"-Activity "Drawing VAPPS" -PercentComplete $vdcP -Id 0
	$VAPPS = $data.VMS | ?{$_.vdc -eq $vdc.id} | Group Container
	$thiscontainer = $null
	$firstcontainer = $null
	$Y = $Y + 10
	$X = $X + 20
	$vdccontainer = $page.Drop($mastercontainer,$X+1,$Y+1)
	($vdccontainer.cells("LocPinX")).formula = "=Width*0"
	($vdccontainer.cells("LocPinY")).formula = "=Height*1"
	$vdccontainer.text = $VDC.name
	$vdccontainer.name = $VDC.id
	$SavedShape = Save-ShapeLookup -Shape $vdccontainer
	$RightShift = 0
	$VAPPTotal =($VAPPS|Measure-Object).count
	$VAPPI = 0
	ForEach($Vapp in $VAPPS)
	{
		$VAPPI+= 1
		$VAPPP = ($VAPPI/$VAPPTotal)*100
		write-progress -Status "Processing VAPPS" -Activity "$($Vapp.name)" -PercentComplete $VAPPP -Id 1
		$thiscontainer = Draw-Vapp -page $page -vms ($vapp.group) -mastercontainer $mastercontainer -server $server -data $data -networkmaster $networkmaster -DynamicConnectorMaster $DynamicConnectorMaster -startx ($X+$RightShift) -starty $Y
		$RightShift += $thiscontainer.Cells("Width").ResultIU 
		$NextStart = ($RightShift + $X)*25.4 + ($thiscontainer.Cells("PinX").ResultIU * 25.4)
		Write-Host "Next VAPP Should be Drawn at X = $NextStart mm"
		Write-Host "Last VAPP Width = $($RightShift * 25.4)"
		$vdccontainer.containerproperties.addmember($thiscontainer,1)
	}
	if($thiscontainer -ne $null)
	{
		
		$vdccontainer.containerproperties.fittocontents()
		if($vdccontainer.cells("Width").ResultIU -lt (150/25.4))
		{
			$vdccontainer.cells("Width").formula = "150 mm"
		}
	}
	
}

Start-Sleep -Seconds 5

#Updated VAPP Container Shapes to look a bit different to VDC's
$VAPPContainers = $Global:ShapeMaster | ?{$_.Name -like "*VAPP*"} | ?{$_.Name -notlike "*NETWORK*"}
ForEach($VC in $VappContainers)
{
	$Page.shapes.ItemFromID($VC.id).CellsSRC(1, 2, 0).FormulaU = "2.25 pt"
	$Page.shapes.ItemFromID($VC.id).CellsSRC(1, 2, 2).FormulaU = "2"
	$Page.shapes.ItemFromID($VC.id).CellsSRC(1, 2, 3).FormulaU = "11.338582677165 pt"
}

Start-Sleep -Seconds 5

#Set VDC Heights to be Uniform
$VDCS = Get-ShapeFromLookup -Page $Page -NameFilter ":VDC:"
$Y = [double]$VDCS[0].Cells("PinY").ResultIU
$Tallest = $VDCS | %{$_.Cells("Height").ResultIU} | Sort | Select -last 1
$Top = $Y
ForEach($V in $VDCS)
{
		$PreviousTop = $V.Cells("PinY").ResultIU
		$V.Cells("PinY").ResultIU = "$($Top)"
		$NewTop = $V.Cells("PinY").ResultIU
		Write-Host "Current Top = $PreviousTop"
		Write-Host "New Top = $NewTop"
		$V.Cells("Height").ResultIU = $Tallest
}



Start-Sleep -Seconds 5

#Set VAPP Heights to be Uniform
$VAPPS = Get-ShapeFromLookup -Page $Page -NameFilter ":VAPP:" | ?{$_.name -notlike "*network*"}
$HighestVapp = $vapps | %{$_.Cells("PinY").ResultIU} | Sort | Select -first 1
$TallestVapp = $vapps | %{$_.Cells("Height").ResultIU} | Sort | Select -last 1
ForEach($V in $VAPPS)
{
	$V.Cells("PinY").ResultIU = $HighestVapp
	$V.Cells("Height").ResultIU = $TallestVapp

}

start-sleep -Seconds 5
#Draw vShield Edge Gateways
Draw-GW -page $page -data $data -DynamicConnectorMaster $DynamicConnectorMaster -FWMaster $FWMaster -mastercontainer $mastercontainer

Start-Sleep -Seconds 5
#Align VDC's against each other
$VDCS = Get-ShapeFromLookup -Page $Page -NameFilter ":VDC:"
#$page.Shapes | ?{$_.name -like "*VDC*"}

$counter = 0
$X = [double]$VDCS[0].Cells("PinX").ResultIU * 25.4
$GapMM = 20
$GapIU = $GapMM/25.4
$Width = 0
ForEach($V in $VDCS)
{
		$NewLocation = $X + $GapIU
		$CurrentLocation = $V.Cells("PinX").formula
		$NewLocation = "$($NewLocation) mm"
		$V.Cells("PinX").formula = "$($NewLocation)"
		$ActualNewLocation = $V.Cells("PinX").formula
		$WidthIU = $V.Cells("Width").ResultIU * 25.4
		$CurrentNew = $V.Cells("PinX").ResultIU * 25.4
		Write-Host "Moving VDC $($V.Text)"
		Write-Host "Current Position = $CurrentLocation"
		Write-Host "New Location = $NewLocation"
		Write-Host "VDC Width = $WidthIU"
		Write-Host "Actual New Location = $ActualNewLocation"
		
		$X = $CurrentNew + $WidthIU
}

$VPNS = @()
ForEach($EGW in $data.EdgeGws)
{
	$VPNData = $EGW.DetailedConfiguration.VPN.Tunnel | Select Name,PeerIpAddress,LocalIPAddress
	ForEach($Line in $VPNData)
	{
		$Line | Add-Member -MemberType NoteProperty -Name EdgeName -Value $EGW.Name -Force
		$Line | Add-Member -MemberType NoteProperty -Name ID -Value $EGW.Id -Force
		$VPNS += $Line
	}	
	
}

$P2PVPNS = @()
ForEach($Line in $VPNS)
{
	$Search = $VPNS | ?{($_.LocalIpAddress -eq $Line.PeerIpAddress) -and ($_.PeerIpAddress -eq $Line.LocalIpAddress)}
	
	if(($Search | Measure-Object).count -gt 0)
	{
		#This is a point to point VPN within this vCloud ORG
		$ExistingSearch = $P2PVPNS | ?{$_.LocalIP -eq $Search.LocalIpAddress} | ?{$_.PeerIP -eq $Search.PeerIpAddress}
		Write-Host "Looking for a LocalIP that matches $($Search.PeerIPAddress) and a PeerIP that matches $($Search.LocalIPAddress)"
		if(($ExistingSearch | Measure-Object).count -eq 0)
		{
			write-host "Didn't fine one, adding it"
			#Doesn't already exist in the list, adding it
			$Holder = "" | Select "Name","LocalIp","PeerIP","LocalID","PeerID"
			$Holder.Name = $Line.Name
			$Holder.LocalIP = $Line.LocalIPAddress
			$Holder.PeerIP = $Line.PeerIPAddress
			$Holder.LocalID = $Line.ID
			$Holder.PeerID = $Search.ID
			$P2PVPNS += $Holder
		}
		else
		{
			write-host "Found one, ignoring..."
		}
	}
}

#Draw the VPNS
ForEach($VPN in $P2PVPNS)
{
	$FirstFWID = $global:ShapeMaster | ?{$_.Name -eq $VPN.LocalID}
	$SecondFWID = $global:ShapeMaster | ?{$_.Name -eq $VPN.PeerID}
	$FirstFW = $Page.Shapes.ItemFromID($FirstFWID.ID)
	$SecondFW = $Page.Shapes.ItemFromID($SecondFWID.ID)
	$con = Connect-VisioObject -Page $Page -FirstObj $FirstFW -SecondObj $SecondFW
	$con.text = "$($VPN.Name): VPN $($VPN.LocalIP) to $($VPN.PeerIP)"
}
$Page.ResizeToFitContents()

Disable-DiagramServices $Document


