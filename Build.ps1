# This script takes input from the user and a CSV file and builds config files for Polycom phones
# Initially created 5/10/2018


# Company Name:
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$CompanyName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the company name.", "Company", "")

# SIP Server:
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$SIPServer = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the SIP Server IP.", "SIP Server", "")

# SIP Server:
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$PhonePass = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Phone Password.", "Phone Password", "")

# Number of Lines:
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$LineQuantity = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the number of lines for the phone display.", "LineQuantity", "6")

# Current Date
$Date = Get-Date

# Local users profile path
$UserPath = $env:USERPROFILE
$Path = "$UserPath\Desktop\$CompanyName"

# Make Directory
New-Item -ItemType directory -Path $Path

WHILE(1){
    # Wait for extensions.csv in the folder:
    [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
    $FileReady = [Microsoft.VisualBasic.Interaction]::MsgBox("Please place the extensions.csv into the $Path\ folder. Click OK when complete.", 1, "File Move")

	# Break the loop if Cancel is clicked
    if ($FileReady -eq "Cancel"){
        break;
    }

	# Break the loop if the file is there
    if (Test-Path -LiteralPath $test -PathType Leaf){
        break;
    }

}


# Site Config
	$site = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+"`r`n"
	$site += '<polycomConfig xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="polycomConfig.xsd">'+"`r`n"
	$site += "`t"+'<voIpProt>'+"`r`n"
	$site += "`t`t"+'<voIpProt.server voIpProt.server.1.address="'+$SIPServer+'" voIpProt.server.1.port="5060">'+"`r`n"
	$site += "`t`t"+'</voIpProt.server>'+"`r`n"
	$site += "`t"+'</voIpProt>'+"`r`n"
	$site += "`t"+'<msg>'+"`r`n"
	$site += "`t`t"+'<msg.mwi msg.mwi.1.callBackMode="contact" msg.mwi.1.callBack="*97">'+"`r`n"
	$site += "`t"+'</msg.mwi>'+"`r`n"
	$site += "`t"+'</msg>'+"`r`n"
	$site += "`t"+'<up up.oneTouchVoiceMail="1">'+"`r`n"
	$site += "`t"+'</up>'+"`r`n"
	$site += "`t"+'<call>'+"`r`n"
	$site += "`t`t"+'<call.internationalDialing call.internationalDialing.enabled="0">'+"`r`n"
	$site += "`t`t"+'</call.internationalDialing>'+"`r`n"
	$site += "`t"+'</call>'+"`r`n"
	$site += "`t"+'<reg reg.1.lineKeys="'+$LineQuantity+'">'+"`r`n"
	$site += "`t"+'</reg>'+"`r`n"
	$site += "`t"+'<tcpIpApp>'+"`r`n"
	$site += "`t`t"+'<tcpIpApp.sntp tcpIpApp.sntp.address="pool.ntp.org" tcpIpApp.sntp.gmtOffset="-28800" tcpIpApp.sntp.gmtOffsetcityID="4">'+"`r`n"
	$site += "`t`t`t"+'<tcpIpApp.sntp.address tcpIpApp.sntp.address.overrideDHCP="1">'+"`r`n"
	$site += "`t`t`t"+'</tcpIpApp.sntp.address>'+"`r`n"
	$site += "`t`t`t"+'<tcpIpApp.sntp.gmtOffset tcpIpApp.sntp.gmtOffset.overrideDHCP="1">'+"`r`n"
	$site += "`t`t`t"+'</tcpIpApp.sntp.gmtOffset>'+"`r`n"
	$site += "`t`t"+'</tcpIpApp.sntp>'+"`r`n"
	$site += "`t"+'</tcpIpApp>'+"`r`n"
	$site += "`t"+'<device device.set="1">'+"`r`n"
	$site += "`t`t"+'<device.auth device.auth.localAdminPassword="'+$PhonePass+'" device.auth.localUserPassword="'+$PhonePass+'">'+"`r`n"
	$site += "`t`t`t"+'<device.auth.localAdminPassword device.auth.localAdminPassword.set="1">'+"`r`n"
	$site += "`t`t`t"+'</device.auth.localAdminPassword>'+"`r`n"
	$site += "`t`t`t"+'<device.auth.localUserPassword device.auth.localUserPassword.set="1">'+"`r`n"
	$site += "`t`t`t"+'</device.auth.localUserPassword>'+"`r`n"
	$site += "`t`t"+'</device.auth>'+"`r`n"
	$site += "`t"+'</device>'+"`r`n"
	$site += '</polycomConfig>'+"`r`n"


	$PathSite = "$Path\site.cfg"
	[System.IO.File]::WriteAllLines($PathSite, $site)


# Extension Configs
	# Load CSV for the loop
	$csv = Import-Csv "$Path\extensions.csv" -Delimiter ","

	foreach($item in $csv){

		#"Extension = $($item.extension) and Name = $($item.name) and Secret = $($item.secret)"

		$extention = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`r`n"
		$extention += '<polycomConfig xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:noNamespaceSchemaLocation="polycomConfig.xsd">' + "`r`n"
		$extention += "`t" + '<reg reg.1.address="' + $($item.extension) + '" reg.1.auth.userId="' + $($item.extension) + '" reg.1.auth.password="' + $($item.secret) + '" reg.1.label="' + $($item.name) + '" reg.1.displayName="' + $($item.name) + '">' + "`r`n"
		$extention += "`t" + '</reg>' + "`r`n"
		$extention += '</polycomConfig>' + "`r`n"

		$PathExt = "$Path\$($item.extension).cfg"
		[System.IO.File]::WriteAllLines($PathExt, $extention)
	}