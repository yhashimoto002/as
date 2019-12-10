<#
.SYNOPSIS

Modify configuration of WS.
If you omit the MGMT hostname, setup proceed without MGMT (using XML).

.EXAMPLE

PS> .\Modify-ConfigWS.ps1 {WS IPaddress} [{MGMT Hostname}]
PS> .\Modify-ConfigWS.ps1 10.12.176.132 YUHASHIMOTO030 
PS> .\Modify-ConfigWS.ps1 10.12.176.132 

#>

param (
    [Parameter(Position=0, Mandatory=$true)]
    [string] $IpAddr,
    [Parameter(Position=1, Mandatory=$false)]
    [string] $Hostname
)


# check parameter
if (! ($IpAddr -match "\d+\.\d+\.\d+\.\d+"))
{
    Write-Host "Please set IP address of WS"
    exit 1
}


# do not change
$ws_install_dir = "\\$IpAddr\c$\Program Files\Votiro\SDS Web Service"
$fc_install_dir = "\\$IpAddr\c$\Program Files\Votiro\Votiro File Connector"
$ws_config_machine = "$ws_install_dir\config\machine.xml"
$ws_config_webapi = "$ws_install_dir\config\webapi.xml"
$ws_config_publish = "$ws_install_dir\Policy\publish.xml"
$ws_config_websrv = "$ws_install_dir\Policy\sanitize_websrv.xml"
$fc_config = "$fc_install_dir\FileConnector.config"


# load config
$conf = Get-Content (Join-Path $PSScriptRoot "settings.ini") | ? { $_ -match "=" } | ConvertFrom-StringData
$fe_config = $conf.FeConfig
$default_channel_in = $conf.DefaultChannelIn
$default_channel_out = $conf.DefaultChannelOut
$channel_name = $conf.ChannelName
$channel_in = $conf.ChannelIn
$channel_out = $conf.ChannelOut
$policy_config_name = $conf.PolicyConfigurationName
$policy_name = $conf.PolicyName

$simple_policy = @{}
$simple_policy.Add("CleanOffice", $conf.CleanOffice)
$simple_policy.Add("CleanPdf", $conf.CleanPdf)
$simple_policy.Add("CleanImages", $conf.CleanImages)
$simple_policy.Add("CleanCad", $conf.CleanCad)
$simple_policy.Add("ExtractEmls", $conf.ExtractEmls)
$simple_policy.Add("BlockPasswordProtectedArchives", $conf.BlockPasswordProtectedArchives)
$simple_policy.Add("BlockPasswordProtectedOffice", $conf.BlockPasswordProtectedOffice)
$simple_policy.Add("BlockPasswordProtectedPdfs", $conf.BlockPasswordProtectedPdfs)
$simple_policy.Add("BlockAllPasswordProtected", $conf.BlockAllPasswordProtected)
$simple_policy.Add("Blockunsupported", $conf.Blockunsupported)
$simple_policy.Add("ScanVirus", $conf.ScanVirus)
$simple_policy.Add("BlockUnknownFiles", $conf.BlockUnknownFiles)
$simple_policy.Add("ExtractArchiveFiles", $conf.ExtractArchiveFiles)
$simple_policy.Add("BlockEquationOleObject", $conf.BlockEquationOleObject)
$simple_policy.Add("BlockBinaryFiles", $conf.BlockBinaryFiles)
$simple_policy.Add("BlockScriptFiles", $conf.BlockScriptFiles)
$simple_policy.Add("BlockFakeFiles", $conf.BlockFakeFiles)


# backup
if(! (Test-Path $ws_config_machine".org"))
{
    cp $ws_config_machine -Destination $ws_config_machine".org"
}
if(! (Test-Path $ws_config_webapi".org"))
{
    cp $ws_config_webapi -Destination $ws_config_webapi".org"
}
if(! (Test-Path $ws_config_publish".org"))
{
    cp $ws_config_publish -Destination $ws_config_publish".org"
}
if(! (Test-Path $ws_config_websrv".org"))
{
    cp $ws_config_websrv -Destination $ws_config_websrv".org"
}
if(! (Test-Path $fc_config".org"))
{
    cp $fc_config -Destination $fc_config".org"
}


# parse XML
$ws_xml_machine = [xml](Get-Content $ws_config_machine)
$ws_xml_webapi = [xml](Get-Content $ws_config_webapi)
$ws_xml_publish = [xml](Get-Content $ws_config_publish)
$fc_xml = [xml](Get-Content $fc_config)


# modify WS and FC config
# if MGMT hostname is not given, XML is used instead
Write-Host "Modifying WS configuration ..."

# FileConnector.config
$fc_xml.FileConnector.MaxWebApiThreads = "64"
$fc_xml.FileConnector.WebApiSettings.WebApiTimeOutInMS = "3600000"
$fc_xml.FileConnector.WebApiSettings.WebApiStatusCheckIntervalInMS = "1000"
($fc_xml.FileConnector.Channels.Channel | ? { $_.Name -eq "DefaultChannel" }).In = $default_channel_in
($fc_xml.FileConnector.Channels.Channel | ? { $_.Name -eq "DefaultChannel" }).Out = $default_channel_out
($fc_xml.FileConnector.Channels.Channel | ? { $_.Name -eq "DefaultChannel" }).InProcessItemsLimit = "150"
$node = $fc_xml.FileConnector.Channels
$childnode = $node.Channel | ? { $_.Name -eq $channel_name}
if (! $childnode)
{
    $element_channel = $fc_xml.CreateElement("Channel")
    [void]$element_channel.SetAttribute("Name", $channel_name)
    [void]$element_channel.SetAttribute("In", $channel_in)
    [void]$element_channel.SetAttribute("Out", $channel_out)
    [void]$element_channel.SetAttribute("PolicyConfigurationName", $policy_config_name)
    [void]$element_channel.SetAttribute("IgnoreEmptyFile", "true")
    [void]$element_channel.SetAttribute("DeleteAfterSanitization", "true")
    [void]$element_channel.SetAttribute("InProcessItemsLimit", "150")
    [void]$element_channel.SetAttribute("FileNotChangedInSeconds", "2")
    [void]$element_channel.SetAttribute("ExcludedExtensions", ".partial,.part,.crdownload")
    [void]$element_channel.SetAttribute("IgnoreFilesWithoutExtension", "False")
    [void]$element_channel.SetAttribute("MaxFileSizeInBytes", "9223372036854775807")
    [void]$node.AppendChild($element_channel)
}

$node = $fc_xml.FileConnector.Policies
$childnode = $node.PolicyParams | ? { $_.Name -eq $policy_config_name}
if (! $childnode)
{
    $element_policyparams = $fc_xml.CreateElement("PolicyParams")
    [void]$element_policyparams.SetAttribute("Name", $policy_config_name)
    [void]$node.AppendChild($element_policyparams)
    $childnode = $node.PolicyParams | ? { $_.Name -eq $policy_config_name}
    $element_predefinedpolicy = $fc_xml.CreateElement("PredefinedPolicy")
    [void]$element_predefinedpolicy.SetAttribute("PolicyName",$policy_name)
    [void]$childnode.AppendChild($element_predefinedpolicy)
    $element_policyrules = $fc_xml.CreateElement("PolicyRules")
    [Xml.XmlNode]$childchildnode = $childnode.AppendChild($element_policyrules)
    foreach ($entry in $simple_policy.GetEnumerator())
    {
        $element_add = $fc_xml.CreateElement("add")
        [void]$element_add.SetAttribute("Name", $entry.Key)
        [void]$element_add.SetAttribute("Value", $entry.Value)
        [void]$childchildnode.AppendChild($element_add)
    }
}

if ($Hostname)
{
    # machine.xml
    $node = $ws_xml_machine.VotiroConfiguration.MachineSettings
    if (! $node.ManagementSettings)
    {
        $element_managementsettings = $ws_xml_machine.CreateElement("ManagementSettings")
        [void]$element_managementsettings.SetAttribute("SanitizationInfoServerUrl", "https://${hostname}:7070")
        if ($node.SandboxServiceSettings) # v8.3 or above
        {
            [void]$element_managementsettings.SetAttribute("BlobManagerServerUrl", "https://${hostname}:3030")
        }
        else
        {
            [void]$element_managementsettings.SetAttribute("BlobManagerServerUrl", "http://${hostname}:3030")
        }
        [void]$element_managementsettings.SetAttribute("AuthToken", "30acc6eb-16d9-4133-ae43-0f5b6d40a318")
        [void]$node.AppendChild($element_managementsettings)
    }
    else
    {
        $ws_xml_machine.VotiroConfiguration.MachineSettings.ManagementSettings.SanitizationInfoServerUrl = "https://${hostname}:7070"
        if ($node.SandboxServiceSettings) # v8.3 or above
        {
            $ws_xml_machine.VotiroConfiguration.MachineSettings.ManagementSettings.BlobManagerServerUrl = "https://${hostname}:3030"
        }
        else
        {
            $ws_xml_machine.VotiroConfiguration.MachineSettings.ManagementSettings.BlobManagerServerUrl = "http://${hostname}:3030"
        }
        $ws_xml_machine.VotiroConfiguration.MachineSettings.ManagementSettings.AuthToken = "30acc6eb-16d9-4133-ae43-0f5b6d40a318"
    }

    # webapi.xml
    $ws_xml_webapi.WebApiServerConfig.NamedPolicyFolderPath = ""
    $ws_xml_webapi.WebApiServerConfig.NamedPolicyServerUri = "https://${hostname}:7070"
    $ws_xml_webapi.WebApiServerConfig.AuthToken = "30acc6eb-16d9-4133-ae43-0f5b6d40a318"

    # publish.xml
    $node = $ws_xml_publish.ArrayOfPolicyRule.PolicyRule.Actions
    $childnode = $node.PolicyAction | ? { $_.type -eq "ReportToManagementAction"}
    if(! $childnode)
    {
        $element_policyaction = $ws_xml_publish.CreateElement("PolicyAction")
        [void]$element_policyaction.SetAttribute("type", $ws_xml_publish.ArrayOfPolicyRule.xsi, "ReportToManagementAction")
        $element_backupfolderpath = $ws_xml_publish.CreateElement("BackupFolderPath")
        $element_backupfolderpath.InnerText = "C:\backup"
        [void]$node.AppendChild($element_policyaction)
        $childnode = $node.PolicyAction | ? { $_.type -eq "ReportToManagementAction"}
        [void]$childnode.AppendChild($element_backupfolderpath)
    }

    # sanitize_websrv.xml
    cp $ws_config_websrv".org" -Destination $ws_config_websrv -Force

    # FileConnector.config
    $node = $fc_xml.FileConnector.Policies
    $childnode = $node.PolicyParams | ? { $_.Name -eq "DefaultPolicy"}
    $childnode.PredefinedPolicy.PolicyName = "Default Policy"
    $childnode = $node.PolicyParams | ? { $_.Name -eq $policy_config_name}
    $childnode.PredefinedPolicy.PolicyName = $policy_name
}
else
{
    # machine.xml
    $node = $ws_xml_machine.VotiroConfiguration.MachineSettings
    if ($node.ManagementSettings)
    {
        [void]$node.RemoveChild($node.ManagementSettings)
    }
    
    # webapi.xml
    $ws_xml_webapi.WebApiServerConfig.NamedPolicyFolderPath = "Policy"
    $ws_xml_webapi.WebApiServerConfig.NamedPolicyServerUri = ""
    $ws_xml_webapi.WebApiServerConfig.AuthToken = ""

    # publish.xml
    $node = $ws_xml_publish.ArrayOfPolicyRule.PolicyRule.Actions
    $childnode = $node.PolicyAction | ? { $_.type -eq "ReportToManagementAction"}
    if ($childnode)
    {
        [void]$node.RemoveChild($childnode)
    }

    # sanitize_websrv.xml
    cp $fe_config -Destination $ws_config_websrv -Force

    # FileConnector.config
    $node = $fc_xml.FileConnector.Policies
    $childnode = $node.PolicyParams | ? { $_.Name -eq "DefaultPolicy"}
    $childnode.PredefinedPolicy.PolicyName = ""
    $childnode = $node.PolicyParams | ? { $_.Name -eq $policy_config_name}
    $childnode.PredefinedPolicy.PolicyName = ""
}

$ws_xml_machine.Save($ws_config_machine)
$ws_xml_webapi.Save($ws_config_webapi)
$fc_xml.Save($fc_config)
$ws_xml_publish.Save($ws_config_publish)


# restart service
Write-Host "Restarting Service ..."
Get-Service -Name VotiroSAPI -ComputerName $IpAddr | Restart-Service
Get-Service -Name VotiroSNMC -ComputerName $IpAddr | Restart-Service
Get-Service -Name Votiro.FileConnector.WindowsService.exe -ComputerName $IpAddr | Restart-Service
Get-Service -ComputerName $IpAddr | Where-Object { $_.Name -eq 'VotiroSAPI' }
Get-Service -ComputerName $IpAddr | Where-Object { $_.Name -eq 'VotiroSNMC' }
Get-Service -ComputerName $IpAddr | Where-Object { $_.Name -eq 'Votiro.FileConnector.WindowsService.exe' }
