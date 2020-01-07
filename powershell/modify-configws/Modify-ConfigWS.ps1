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


# change if needed
$username = "Administrator"
$password = "Asgent123"

# do not change
$encryptedPassword = ConvertTo-SecureString $password -AsPlainText -Force
$mountDriveName = "V"
$credential = New-Object System.Management.Automation.PSCredential($username, $encryptedPassword)
try
{
    New-PSDrive -Name $mountDriveName -PSProvider FileSystem -Root "\\$IpAddr\c$" -Credential $credential -Persist -ErrorAction Stop | Out-Null
}
catch
{
    Remove-PSDrive $mountDriveName
    New-PSDrive -Name $mountDriveName -PSProvider FileSystem -Root "\\$IpAddr\c$" -Credential (Get-Credential) -Persist | Out-Null
}

$ws_install_dir = "${mountDriveName}:\Program Files\Votiro\SDS Web Service"
$fc_install_dir = "${mountDriveName}:\Program Files\Votiro\\Votiro File Connector"
$ws_config_machine = "$ws_install_dir\config\machine.xml"
$ws_config_webapi = "$ws_install_dir\config\webapi.xml"
$ws_config_publish = "$ws_install_dir\Policy\publish.xml"
$ws_config_websrv = "$ws_install_dir\Policy\sanitize_websrv.xml"
$fc_config = "$fc_install_dir\FileConnector.config"


# load config
$conf = Get-Content (Join-Path $PSScriptRoot "settings.ini") | Where-Object { $_ -match "=" } | ConvertFrom-StringData
$fe_config = $conf.FeConfig

$default_channel_in = $conf.DefaultChannelIn
$default_channel_out = $conf.DefaultChannelOut

$fe_channel_name = $conf.FeChannelName
$fe_channel_in = $conf.FeChannelIn
$fe_channel_out = $conf.FeChannelOut
$fe_policy_config_name = $conf.FePolicyConfigurationName
$fe_policy_name = $conf.FePolicyName

$fe_simple_policy = @{}
$fe_simple_policy.Add("CleanOffice", $conf.FeCleanOffice)
$fe_simple_policy.Add("CleanPdf", $conf.FeCleanPdf)
$fe_simple_policy.Add("CleanImages", $conf.FeCleanImages)
$fe_simple_policy.Add("ExtractEmls", $conf.FeExtractEmls)
$fe_simple_policy.Add("BlockPasswordProtectedArchives", $conf.FeBlockPasswordProtectedArchives)
$fe_simple_policy.Add("BlockPasswordProtectedOffice", $conf.FeBlockPasswordProtectedOffice)
$fe_simple_policy.Add("BlockPasswordProtectedPdfs", $conf.FeBlockPasswordProtectedPdfs)
$fe_simple_policy.Add("BlockAllPasswordProtected", $conf.FeBlockAllPasswordProtected)
$fe_simple_policy.Add("Blockunsupported", $conf.FeBlockunsupported)
$fe_simple_policy.Add("BlockFakeFiles", $conf.FeBlockFakeFiles)
$fe_simple_policy.Add("ScanVirus", $conf.FeScanVirus)
$fe_simple_policy.Add("BlockUnknownFiles", $conf.FeBlockUnknownFiles)
$fe_simple_policy.Add("ExtractArchiveFiles", $conf.FeExtractArchiveFiles)
$fe_simple_policy.Add("BlockEquationOleObject", $conf.FeBlockEquationOleObject)
$fe_simple_policy.Add("BlockBinaryFiles", $conf.FeBlockBinaryFiles)
$fe_simple_policy.Add("BlockScriptFiles", $conf.FeBlockScriptFiles)

$sample_channel_name = $conf.SampleChannelName
$sample_channel_in = $conf.SampleChannelIn
$sample_channel_out = $conf.SampleChannelOut
$sample_policy_config_name = $conf.SamplePolicyConfigurationName
$sample_policy_name = $conf.SamplePolicyName

$sample_simple_policy = @{}
$sample_simple_policy.Add("CleanOffice", $conf.SampleCleanOffice)
$sample_simple_policy.Add("CleanPdf", $conf.SampleCleanPdf)
$sample_simple_policy.Add("CleanImages", $conf.SampleCleanImages)
$sample_simple_policy.Add("ExtractEmls", $conf.SampleExtractEmls)
$sample_simple_policy.Add("BlockPasswordProtectedArchives", $conf.SampleBlockPasswordProtectedArchives)
$sample_simple_policy.Add("BlockPasswordProtectedOffice", $conf.SampleBlockPasswordProtectedOffice)
$sample_simple_policy.Add("BlockPasswordProtectedPdfs", $conf.SampleBlockPasswordProtectedPdfs)
$sample_simple_policy.Add("BlockAllPasswordProtected", $conf.SampleBlockAllPasswordProtected)
$sample_simple_policy.Add("Blockunsupported", $conf.SampleBlockunsupported)
$sample_simple_policy.Add("BlockFakeFiles", $conf.SampleBlockFakeFiles)
$sample_simple_policy.Add("ScanVirus", $conf.SampleScanVirus)
$sample_simple_policy.Add("BlockUnknownFiles", $conf.SampleBlockUnknownFiles)
$sample_simple_policy.Add("ExtractArchiveFiles", $conf.SampleExtractArchiveFiles)
$sample_simple_policy.Add("BlockEquationOleObject", $conf.SampleBlockEquationOleObject)
$sample_simple_policy.Add("BlockBinaryFiles", $conf.SampleBlockBinaryFiles)
$sample_simple_policy.Add("BlockScriptFiles", $conf.SampleBlockScriptFiles)


# mkdir
New-Item $default_channel_in.replace("c:", "${mountDriveName}:") -ItemType Directory -Force | Out-Null
New-Item $fe_channel_in.replace("c:", "${mountDriveName}:") -ItemType Directory -Force | Out-Null
New-Item $sample_channel_in.replace("c:", "${mountDriveName}:") -ItemType Directory -Force | Out-Null


# backup
if(! (Test-Path $ws_config_machine".org"))
{
    Copy-Item $ws_config_machine -Destination $ws_config_machine".org"
}
if(! (Test-Path $ws_config_webapi".org"))
{
    Copy-Item $ws_config_webapi -Destination $ws_config_webapi".org"
}
if(! (Test-Path $ws_config_publish".org"))
{
    Copy-Item $ws_config_publish -Destination $ws_config_publish".org"
}
if(! (Test-Path $ws_config_websrv".org"))
{
    Copy-Item $ws_config_websrv -Destination $ws_config_websrv".org"
}
if(! (Test-Path $fc_config".org"))
{
    Copy-Item $fc_config -Destination $fc_config".org"
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
# - default
($fc_xml.FileConnector.Channels.Channel | Where-Object { $_.Name -eq "DefaultChannel" }).In = $default_channel_in
($fc_xml.FileConnector.Channels.Channel | Where-Object { $_.Name -eq "DefaultChannel" }).Out = $default_channel_out
($fc_xml.FileConnector.Channels.Channel | Where-Object { $_.Name -eq "DefaultChannel" }).InProcessItemsLimit = "150"
# - fe
$node = $fc_xml.FileConnector.Channels
$childnode = $node.Channel | Where-Object { $_.Name -eq $fe_channel_name}
if (! $childnode)
{
    $element_channel = $fc_xml.CreateElement("Channel")
    [void]$element_channel.SetAttribute("Name", $fe_channel_name)
    [void]$element_channel.SetAttribute("In", $fe_channel_in)
    [void]$element_channel.SetAttribute("Out", $fe_channel_out)
    [void]$element_channel.SetAttribute("PolicyConfigurationName", $fe_policy_config_name)
    [void]$element_channel.SetAttribute("IgnoreEmptyFile", "true")
    [void]$element_channel.SetAttribute("DeleteAfterSanitization", "true")
    [void]$element_channel.SetAttribute("InProcessItemsLimit", "150")
    [void]$element_channel.SetAttribute("FileNotChangedInSeconds", "2")
    [void]$element_channel.SetAttribute("ExcludedExtensions", ".partial,.part,.crdownload")
    [void]$element_channel.SetAttribute("IgnoreFilesWithoutExtension", "False")
    [void]$element_channel.SetAttribute("MaxFileSizeInBytes", "9223372036854775807")
    [void]$node.AppendChild($element_channel)
}
# - sample
$childnode = $node.Channel | Where-Object { $_.Name -eq $sample_channel_name}
if (! $childnode)
{
    $element_channel = $fc_xml.CreateElement("Channel")
    [void]$element_channel.SetAttribute("Name", $sample_channel_name)
    [void]$element_channel.SetAttribute("In", $sample_channel_in)
    [void]$element_channel.SetAttribute("Out", $sample_channel_out)
    [void]$element_channel.SetAttribute("PolicyConfigurationName", $sample_policy_config_name)
    [void]$element_channel.SetAttribute("IgnoreEmptyFile", "true")
    [void]$element_channel.SetAttribute("DeleteAfterSanitization", "true")
    [void]$element_channel.SetAttribute("InProcessItemsLimit", "150")
    [void]$element_channel.SetAttribute("FileNotChangedInSeconds", "2")
    [void]$element_channel.SetAttribute("ExcludedExtensions", ".partial,.part,.crdownload")
    [void]$element_channel.SetAttribute("IgnoreFilesWithoutExtension", "False")
    [void]$element_channel.SetAttribute("MaxFileSizeInBytes", "9223372036854775807")
    [void]$node.AppendChild($element_channel)
}

# - fe
$node = $fc_xml.FileConnector.Policies
$childnode = $node.PolicyParams | Where-Object { $_.Name -eq $fe_policy_config_name}
if (! $childnode)
{
    $element_policyparams = $fc_xml.CreateElement("PolicyParams")
    [void]$element_policyparams.SetAttribute("Name", $fe_policy_config_name)
    [void]$node.AppendChild($element_policyparams)
    $childnode = $node.PolicyParams | Where-Object { $_.Name -eq $fe_policy_config_name}
    $element_predefinedpolicy = $fc_xml.CreateElement("PredefinedPolicy")
    [void]$element_predefinedpolicy.SetAttribute("PolicyName",$fe_policy_name)
    [void]$childnode.AppendChild($element_predefinedpolicy)
    $element_policyrules = $fc_xml.CreateElement("PolicyRules")
    [Xml.XmlNode]$childchildnode = $childnode.AppendChild($element_policyrules)
    foreach ($entry in $fe_simple_policy.GetEnumerator())
    {
        $element_add = $fc_xml.CreateElement("add")
        [void]$element_add.SetAttribute("Name", $entry.Key)
        [void]$element_add.SetAttribute("Value", $entry.Value)
        [void]$childchildnode.AppendChild($element_add)
    }
}
# - sample
$childnode = $node.PolicyParams | Where-Object { $_.Name -eq $sample_policy_config_name}
if (! $childnode)
{
    $element_policyparams = $fc_xml.CreateElement("PolicyParams")
    [void]$element_policyparams.SetAttribute("Name", $sample_policy_config_name)
    [void]$node.AppendChild($element_policyparams)
    $childnode = $node.PolicyParams | Where-Object { $_.Name -eq $sample_policy_config_name}
    $element_predefinedpolicy = $fc_xml.CreateElement("PredefinedPolicy")
    [void]$element_predefinedpolicy.SetAttribute("PolicyName",$sample_policy_name)
    [void]$childnode.AppendChild($element_predefinedpolicy)
    $element_policyrules = $fc_xml.CreateElement("PolicyRules")
    [Xml.XmlNode]$childchildnode = $childnode.AppendChild($element_policyrules)
    foreach ($entry in $sample_simple_policy.GetEnumerator())
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
    $childnode = $node.PolicyAction | Where-Object { $_.type -eq "ReportToManagementAction"}
    if(! $childnode)
    {
        $element_policyaction = $ws_xml_publish.CreateElement("PolicyAction")
        [void]$element_policyaction.SetAttribute("type", $ws_xml_publish.ArrayOfPolicyRule.xsi, "ReportToManagementAction")
        $element_backupfolderpath = $ws_xml_publish.CreateElement("BackupFolderPath")
        $element_backupfolderpath.InnerText = "C:\backup"
        [void]$node.AppendChild($element_policyaction)
        $childnode = $node.PolicyAction | Where-Object { $_.type -eq "ReportToManagementAction"}
        [void]$childnode.AppendChild($element_backupfolderpath)
    }

    # sanitize_websrv.xml
    Copy-Item $ws_config_websrv".org" -Destination $ws_config_websrv -Force

    # FileConnector.config
    $node = $fc_xml.FileConnector.Policies
    $childnode = $node.PolicyParams | Where-Object { $_.Name -eq "DefaultPolicy"}
    $childnode.PredefinedPolicy.PolicyName = "Default Policy"
    $childnode = $node.PolicyParams | Where-Object { $_.Name -eq $fe_policy_config_name}
    $childnode.PredefinedPolicy.PolicyName = $fe_policy_name
    $childnode = $node.PolicyParams | Where-Object { $_.Name -eq $sample_policy_config_name}
    $childnode.PredefinedPolicy.PolicyName = $sample_policy_name
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
    $childnode = $node.PolicyAction | Where-Object { $_.type -eq "ReportToManagementAction"}
    if ($childnode)
    {
        [void]$node.RemoveChild($childnode)
    }

    # sanitize_websrv.xml
    Copy-Item $fe_config -Destination $ws_config_websrv -Force

    # FileConnector.config
    $node = $fc_xml.FileConnector.Policies
    $childnode = $node.PolicyParams | Where-Object { $_.Name -eq "DefaultPolicy"}
    $childnode.PredefinedPolicy.PolicyName = ""
    $childnode = $node.PolicyParams | Where-Object { $_.Name -eq $fe_policy_config_name}
    $childnode.PredefinedPolicy.PolicyName = ""
    $childnode = $node.PolicyParams | Where-Object { $_.Name -eq $sample_policy_config_name}
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
$fcService = try {
    Get-Service -Name Votiro.FileConnector.WindowsService.exe -ComputerName $IpAddr -ErrorAction Stop
}
catch {
    Get-Service -Name "Votiro SDS File Connector Service" -ComputerName $IpAddr
}
$fcService | Restart-Service
Get-Service -ComputerName $IpAddr | Where-Object { $_.Name -eq 'VotiroSAPI' }
Get-Service -ComputerName $IpAddr | Where-Object { $_.Name -eq 'VotiroSNMC' }
Get-Service -ComputerName $IpAddr | Where-Object { $_.Name -eq $fcService.Name }

Remove-PSDrive -Name $mountDriveName