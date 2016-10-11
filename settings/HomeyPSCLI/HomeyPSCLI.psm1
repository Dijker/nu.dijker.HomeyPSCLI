### Homey PS CLI 

function Connect-Homey
<#
.Synopsis
   Connect-Homey (Homey by Athom http://www.athom.com)
   Beta 0.3.0
.DESCRIPTION
   Set IP Address and Bearer for your LOCAL Homey to store in a PowerShell variable Windows computer

.EXAMPLE
    Connect-Homey -IP 1.2.3.4 -Bearer abcdefg -ExportPath C:\HomeyBackup -WriteConfig

.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
   Beta beta beta.... 

.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://forum.athom.com/',
                  ConfirmImpact='Medium')]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$false, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   #Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        # [Alias("ip")] 
        [string]
        $IP,

        #Parameter 2 Bearer token for Homey
        [Parameter(ParameterSetName='Parameter Set 1')]
        [ValidatePattern("[0-9][a-f]*")]
        [ValidateLength(0,40)]
        [string]
        $Bearer, 

        # Param3 help description
        [Parameter(ParameterSetName='Parameter Set 1')]
        # [AllowNull()]
        # [AllowEmptyCollection()]
        # [AllowEmptyString()]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({$true})]
        #[ValidateRange(0,5)]
        [string]
        $ExportPath,

        [string]
        $CloudID, 
        [switch]
        $WriteConfig
    )

    $ScriptDirectory = Get-ScriptDirectory
    # Write-Host "ScriptDirectory : $ScriptDirectory"
    If (Test-Path "$ScriptDirectory\Config-HomeyPSCLI.ps1" ) {
            . $ScriptDirectory\Config-HomeyPSCLI.ps1
        } 
    If ($ExportPath -ne "") { $Global:_HomeysExportPath = $ExportPath }
    If ($IP -ne "") { $Global:_HomeysIP = $IP }
    If ($Bearer -ne "") { $Global:_HomeysBearer= $Bearer }
    If ($CloudID -ne "") { $Global:_HomeysCloudID = "$CloudID" }

    $Global:_HomeysProtocol = "https"

    $Global:_HomeysHeaders = @{"Authorization"="Bearer $_HomeysBearer"} 
    $Global:_HomeysContentType = "application/json"
    $Global:_HomeysCloudHostname = "$_HomeysCloudID.homey.athom.com"
    $Global:_HomeysCloudLocalHostname = "$_HomeysCloudID.homeylocal.com"
    If ($CloudID -ne "" ) {    
        $_HomeysResolvedLocalIP =  ([System.Net.Dns]::GetHostAddresses($_HomeysCloudLocalHostname)).IPAddressToString
    }

    $_SystemWR = $null
    $Global:_HomeysConnectHost = $null

    If (Test-NetConnection $_HomeysCloudLocalHostname -InformationLevel Quiet) { 
        "$_HomeysCloudLocalHostname ICMP Response OK" 
        $Global:_HomeysProtocol = "http"
        $Global:_HomeysConnectHost = $_HomeysCloudLocalHostname
        $_SystemWR = Invoke-WebRequest -Uri "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/system" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType

    } else { 
        If (Test-NetConnection $_HomeysIP -InformationLevel Quiet) {
            "$_HomeysIP ICMP Response OK" 
            $Global:_HomeysProtocol = "http"
            $Global:_HomeysConnectHost = $_HomeysIP
            $_SystemWR = Invoke-WebRequest -Uri "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/system" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
            $Global:_HomeysCloudID = $_SystemWR.Headers.'X-Homey-ID'
        } else {
            $_SystemWR = try {
                $Global:_HomeysCloudHostname = "$CloudID.homey.athom.com"
                $Global:_HomeyGetSystemApi = "$_HomeysProtocol`://$_HomeysCloudHostname/api/manager/system"
                Invoke-WebRequest -Uri "$_HomeyGetSystemApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -ErrorAction Ignore 
            } catch { 
                $_.Exception.Response
                Write-Host "Warning: Homey local and Cloud presentation not found! " -ForegroundColor Yellow
            }
            If ($_SystemWR -ne $null ) { 
                $Global:_HomeysConnectHost = $_HomeysCloudHostname
            } 
        }
    } 
    # Want to move Global URLs them to the correct functions
    $Global:_HomeyGetZonesApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/Zones/Zone" 
    $Global:_HomeyGetFoldersApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/Folder" 
    $Global:_HomeyGetFlowsApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/flow" 
    $Global:_HomeyGetSystemApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/system"
    $Global:_HomeyMyHomeyPSCLIApp = "$_HomeysProtocol`://$_HomeysConnectHost/app/nu.dijker.homeypscli"

    # $_SystemWR = Invoke-WebRequest -Uri "$_HomeyGetSystemApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    # $_SystemJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_SystemWR , [System.Collections.Hashtable])
    IF ( $_SystemWR.StatusCode -eq 200 ) { 
        "$_HomeysConnectHost HTTP Response OK" 
        $Global:_HomeyVersion = $_SystemWR.Headers.'X-Homey-Version'
        "Homey version $Global:_HomeyVersion "
        If ($_HomeysExportPath -ne $null ) { "Homeys Export configured to: $_HomeysExportPath"}
    }

    $return = Export-HomeySystemSettings -AppUri 'apps/app'
    if ('nu.dijker.homeypscli' -in $return.keys )  { 
        # HomeyPSCLI INstalled !!
        $_HomeysMyAppVersion = $return.'nu.dijker.homeypscli'.version.Split(".")
        $_HomeysMyAppVersionVal = 0+10*$_HomeysMyAppVersion[2] + 1000*$_HomeysMyAppVersion[1] +100000*$_HomeysMyAppVersion[0]
        $_CurrentMyAppVersion = (Get-Module HomeyPSCLI).Version
        $_CurrentMyAppVersionVal = $_CurrentMyAppVersion.Revision + 10*$_CurrentMyAppVersion.Build + 1000*$_CurrentMyAppVersion.Minor + 100000*$_CurrentMyAppVersion.Major
        $_CurrentMyAppVersionStr = "v{0}.{1}.{2}-{3}" -f $_CurrentMyAppVersion.Major,$_CurrentMyAppVersion.Minor,$_CurrentMyAppVersion.Build,$_CurrentMyAppVersion.Revision 
        If ($_HomeysMyAppVersionVal -gt $_CurrentMyAppVersionVal ) {
            Write-Host "New HomeyPSCLI version available on Homey!" -ForegroundColor Green

            Write-Host ".... Downloading upgrade ...." -ForegroundColor Green


            # "{0:yyyyMMddHHmmss}" -f (Get-ChildItem C:\Users\gdi.PQRNL\Documents\WindowsPowerShell\Modules\HomeyPSCLI\HomeyPSCLIn.psd1 ).LastWriteTime
            ('HomeyPSCLI.psd1','HomeyPSCLI.psm1' ) | ForEach-Object  {"$_"
                iF ( Test-Path "$ScriptDirectory\$_" ) {
                    $_LastWriteTime = "{0:yyyyMMddHHmmss}" -f (Get-ChildItem "$ScriptDirectory\$_" ).LastWriteTime
                    Rename-Item -Path "$ScriptDirectory\$_" "$_-$_LastWriteTime"
                }
                Invoke-WebRequest "$_HomeyMyHomeyPSCLIApp/settings/HomeyPSCLI/$_" -OutFile "$ScriptDirectory\$_"
            Write-Host "To activate, Reload module: Import-Module HomeyPSCLI  -Force" -ForegroundColor Yellow
            }
        }

    } Else { 
        Write-Host "Warning: Homey PSCLI not installed on Homey!" -ForegroundColor Yellow
        Write-Host "pls. install for automatic updates of your HomeyPSCLI Module" -ForegroundColor Yellow
    }
    If ($WriteConfig) {
        "`$Global:_HomeysBearer = ""$_HomeysBearer""" | Out-File -FilePath $ScriptDirectory\Config-HomeyPSCLI.ps1
        "`$Global:_HomeysCloudID = ""$_HomeysCloudID""" | Out-File -FilePath $ScriptDirectory\Config-HomeyPSCLI.ps1 -Append
        "`$Global:_HomeysExportPath = ""$_HomeysExportPath""" | Out-File -FilePath $ScriptDirectory\Config-HomeyPSCLI.ps1 -Append
        "`$Global:_HomeysIP = ""$_HomeysIP""" | Out-File -FilePath $ScriptDirectory\Config-HomeyPSCLI.ps1 -Append
    }
}


function Export-HomeyAppsVar
{
    param (
    [string] $AppUri )
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    $_HomeyUriGetAppVar = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/settings/app/$AppUri/variables"
    $_AppWR = Invoke-WebRequest -Uri "$_HomeyUriGetAppVar" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    $_AppJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_AppWR, [System.Collections.Hashtable])
    # $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
    return $_AppJSON.result 
}

function Export-HomeySystemSettings
{
    param (
    [string] $AppUri )
    $_HomeyUriGetAppVar = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/$AppUri"
    $_AppWR = Invoke-WebRequest -Uri "$_HomeyUriGetAppVar" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    $_AppJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_AppWR, [System.Collections.Hashtable])
    # $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
    return $_AppJSON.result 
}

function Import-HomeyAppsVar
{
    param (
    [string] $AppUri,
    [string] $JSONFile,
    [switch] $AddMissingVar
    )
    $_HomeyUriGetAppVar = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/settings/app/$AppUri/variables"
    If ($AppUri -eq 'nl.bevlogenheid.countdown' ) { 
        $_HomeyUriPutAppVar = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/settings/app/$AppUri/changedvariables" 
    }  Else { 
        $_HomeyUriPutAppVar = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/settings/app/$AppUri/variables" 
    }   

    If (Test-Path $JSONFile) {
        $NewAppsVar = Get-Content -Raw -Path $JSONFile | ConvertFrom-Json 
        $_AppWR = try {
            Invoke-WebRequest -Uri "$_HomeyUriGetAppVar" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -ErrorAction Ignore 
        } catch { 
            $_.Exception.Response
            Write-Host "Warning: App does not Exists! " -ForegroundColor Yellow
        }
        If ($_AppWR.StatusCode -eq 200) {
            $_AllCurrentVars = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_AppWR.Content, [System.Collections.Hashtable]).result
            If ($AddMissingVar) {
                ## $_AllVarsNew | ForEach-Object { if ($_.name -notin ($_AllVars.name) ) { $_ ;$_AllVars += $_ } }
                $NewAppsVar | ForEach-Object { if ($_.name -notin ($_AllCurrentVars.name) ) { $_AllCurrentVars += $_ } }
                $NewAppsVar = $_AllCurrentVars 
            } 
            $CompressedJSONVar = $NewAppsVar | ConvertTo-Json -Depth 99  -Compress
            $CompressedJSONVarValue = "{""value"":$CompressedJSONVar}"
            $_AppWR = Invoke-WebRequest -Uri "$_HomeyUriPutAppVar" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -Method Put -Body $CompressedJSONVarValue
            $_AppJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_AppWR, [System.Collections.Hashtable])
            # $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
            return $_AppJSON.result 
        }
    } Else {
        Write-Host "Error: File $JSONFile not found!" 
    }
}

function Get-HomeyPendingUpdate
{
    param (
    [switch] $Verbose,
    [switch] $InstallPendingUpdate
    )
    $_HomeyUriPendingUpdate = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/updates/update/"

    $_AppWR = Invoke-WebRequest -Uri "$_HomeyUriPendingUpdate" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    $_AppJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_AppWR, [System.Collections.Hashtable])
    $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
    If ($_AppJSON.result.size -gt 10) {
        If ($Verbose) {
        " Date    : {0}" -f $_AppJSON.result.date
        " Version : {0}" -f $_AppJSON.result.version
        " Size    : {0}" -f $_AppJSON.result.size
        $html = $_AppJSON.result.changelog.en
        @('br','/li' ) | % {$html = $html -replace "<$_[^>]*?>", "`n" }
        @('ul','li', '/ul' ) | % {$html = $html -replace "<$_[^>]*?>", "" }
        @('&amp;' ) | % {$html = $html -replace "$_", "&" }
        $html = $html -replace "`n`n", "`n"
        $html
        } Else { 
            return $_AppJSON.result 
        } 
        If ($InstallPendingUpdate) {
            $_HomeyUriPendingUpdate = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/updates/update/"
            $_AppWR = Invoke-WebRequest -Uri "$_HomeyUriPendingUpdate" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -Method Post
            Write-Host " Updating Homey Software, do not turn off the power! "
        } 
    } Else {
        If ($Verbose) {
        " No Updates available..."
        " Version : {0}" -f $Global:_HomeyVersion
        }
        return $_AppJSON.result 
        # return $false # No pending update 
    }
}

function New-HomeyFlow
{
    $Global:_HomeyUriFlowsApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/flow/" 
    $_FlowWR = Invoke-WebRequest -Uri "$_HomeyUriFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -Method Post
    $_FlowJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FlowWR, [System.Collections.Hashtable])
    $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
    return $_FlowJSON.result
}

function Remove-HomeyFlow
{
    param (
    [string] $ID
    )
    $Global:_HomeyUriFlowsApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/flow/$ID" 
    $_FlowWR = try {
        Invoke-WebRequest -Uri "$_HomeyUriFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -ErrorAction Ignore 
    } catch { 
        $_.Exception.Response
    }

    If ($_FlowWR.StatusCode -eq 200 ) {
        $_FlowWR = Invoke-WebRequest -Uri "$_HomeyUriFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -Method Delete
        $_FlowJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_AppWR, [System.Collections.Hashtable])
        # $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
    } Else {
        Write-Host "Error finding Flow ID: $ID" -ForegroundColor Yellow 
    }
    return $_FlowJSON.result
}

function Get-HomeyFolderPath
{
    Param( $_FolderJSON, $_Key)
    If( $_FolderJSON.$_Key.folder -eq $False ) {
        $_FolderName = Get-ValidFilename $_FolderJSON.$_Key.title
        # $_IllegalFileFolderChars.ToCharArray() | % {$_FolderName = $_FolderName -replace "$_", "-" }
        return  $_FolderName
    } else { # Next level Folders 
        $_TopFolderName = "$(Get-HomeyFolderPath $_FolderJSON "$($_FolderJSON.$_Key.folder)")"
        $_FolderName = Get-ValidFilename $_FolderJSON.$_Key.title

        # $_IllegalFileFolderChars.ToCharArray() | % {$_FolderName = $_FolderName -replace "$_", "-" }
        return "$_TopFolderName\$_FolderName"
    }
}

function Get-HomeyFoldersStructure
{
    $_ExportPathFolders =  "$_HomeysExportPath\Flows"
    $_FoldersWR = Invoke-WebRequest -Uri "$_HomeyGetFoldersApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    $_FoldersJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FoldersWR , [System.Collections.Hashtable])
    $Global:_HomeyVersion = $_FoldersWR.Headers.'X-Homey-Version'

    # Write-Host "-=-=- Folders "
    $_FoldersJSON.result.keys | ForEach-Object {

        $_FolderRelPath = Get-HomeyFolderPath $_FoldersJSON.result $_ 
        If (!(Test-Path $_ExportPathFolders\$_FolderRelPath)) { $return = New-Item $_ExportPathFolders\$_FolderRelPath -ItemType Directory } 
    } 
    return $_FoldersJSON
}

function Get-HomeyZonePath
{
    Param( $_ZonesJSON, $_Key)
    If( $_ZonesJSON.$_Key.parent -eq $False ) {
        $_ZoneName = Get-ValidFilename $_ZonesJSON.$_Key.name
        # $_IllegalFileFolderChars.ToCharArray() | % {$_ZoneName = $_ZoneName -replace "$_", "-" }
        return  $_ZoneName
    } else { # Next level Folders 
        $_TopZoneName = Get-HomeyZonePath $_ZonesJSON $_ZonesJSON.$_Key.parent
        $_ZoneName = Get-ValidFilename $_ZonesJSON.$_Key.name
        # $_IllegalFileFolderChars.ToCharArray() | % {$_ZoneName = $_ZoneName -replace "$_", "-" }
        return "$_TopZoneName\$_ZoneName" 
    }
}

function Export-HomeyZonesStructure
{
    $_ExportPathZones =  "$_HomeysExportPath\Zones"
    $_FoldersWR = Invoke-WebRequest -Uri "$_HomeyGetZonesApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    $_ZonesJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FoldersWR , [System.Collections.Hashtable]).result

    # Write-Host "-=-=- Folders "
    $_ZonesJSON.keys | ForEach-Object {
        $_ZonesRelPath = Get-HomeyZonePath $_ZonesJSON $_             
        If (!(Test-Path $_ExportPathZones\$_ZonesRelPath)) { $return = New-Item $_ExportPathZones\$_ZonesRelPath -ItemType Directory } 
    } 
    $_ZonesJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\Zones-v$_HomeyVersion-$_ExportDTSt.json"
    return $_ZonesJSON
}

function Get-HomeyZonesStructure
{
    $_FoldersWR = Invoke-WebRequest -Uri "$_HomeyGetZonesApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    $_ZonesJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FoldersWR , [System.Collections.Hashtable]).result

    return $_ZonesJSON
}

function Export-HomeyDevices
{
    # $_HomeyFlowFoldersArray[0]
    $Global:_ExportDTSt = "{0:yyyyMMddHHmmss}" -f (get-date)

    $_ExportPathDevices =  "$_HomeysExportPath\Zones"
    $_HomeyGetDeviceApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/devices/device"
    $_DevicesWR = Invoke-WebRequest -Uri "$_HomeyGetDeviceApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    $_DevicesJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_DevicesWR.Content, [System.Collections.Hashtable])
    $Global:_HomeyVersion = $_FlowsWR.Headers.'X-Homey-Version'
    $_HomeyDeviceZones = Get-HomeyZonesStructure

    $_DevicesJSON.result.keys | ForEach-Object {
        $_DeviceZoneID = $_DevicesJSON.result.$_.zone.id
        $_DeviceID = $_DevicesJSON.result.$_.id
        $_DeviceName = $_DevicesJSON.result.$_.name
        $_DeviceFileName = Get-ValidFilename $_DeviceName 
        # $_IllegalFileFolderChars.ToCharArray() | % {$_DeviceFileName = $_DeviceFileName -replace "$_", "-" }
        $_DeviceFolder = Get-HomeyZonePath $_HomeyDeviceZones $_DeviceZoneID
        $_DevicesJSON.result.$_ | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathDevices\$_DeviceFolder\$_DeviceFileName-$_DeviceID-v$_HomeyVersion-$_ExportDTSt.json" 
    } 
    return $_DevicesJSON.result.values
}

function Get-HomeyFlows
{
    $_FlowsWR = Invoke-WebRequest -Uri "$_HomeyGetFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    $_FlowsJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FlowsWR, [System.Collections.Hashtable])
    return $_FlowsJSON.result.Values
}

function Get-HomeyFlow
{
    param (
    [string] $ID
    )
    $Global:_HomeyUriFlowsApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/flow/$ID" 
    $_FlowWR = try {
        Invoke-WebRequest -Uri "$_HomeyUriFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -ErrorAction Ignore 
    } catch { 
        $_.Exception.Response
    }
    If ($_FlowWR.StatusCode -eq 200 ) {
        $_FlowWR = Invoke-WebRequest -Uri "$_HomeyUriFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
        $_FlowJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FlowWR, [System.Collections.Hashtable])
    } Else {
        Write-Host "Error finding Flow ID: $ID" -ForegroundColor Yellow 
    }
    return $_FlowJSON.result
}

function Set-HomeyFlow
{
    param (
    [string] $ID,
    [string] $CompressedJSONFlow
    )
    $Global:_HomeyUriFlowsApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/flow/$ID" 
    $_FlowWR = try {
        Invoke-WebRequest -Uri "$_HomeyUriFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -ErrorAction Ignore 
    } catch { 
        $_.Exception.Response
    }
    If ($_FlowWR.StatusCode -eq 200 ) {
        $_FlowWR = Invoke-WebRequest -Uri "$_HomeyUriFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -Method Put -Body $CompressedJSONFlow 
        $_FlowJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FlowWR, [System.Collections.Hashtable])
        # $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
    } Else {
        Write-Host "Error finding Flow ID: $ID" -ForegroundColor Yellow 
    }
    return $_FlowJSON.result
}

function Export-HomeyFlows
{
    $Global:_ExportDTSt = "{0:yyyyMMddHHmmss}" -f (get-date)

    $_ExportPathFolders =  "$_HomeysExportPath\Flows"
    $_FlowsWR = Invoke-WebRequest -Uri "$_HomeyGetFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    $_FlowsJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FlowsWR, [System.Collections.Hashtable])
    $Global:_HomeyVersion = $_FlowsWR.Headers.'X-Homey-Version'
    $_HomeyFlowFolders = Get-HomeyFoldersStructure 

    $_FlowsJSON.result.keys | ForEach-Object {
        $_FlowsFolderID = $_FlowsJSON.result.$_.folder
        $_FlowsID = $_FlowsJSON.result.$_.id
        $_FlowsTitle = Get-ValidFilename $_FlowsJSON.result.$_.title
        $_FlowFolder = Get-HomeyFolderPath $_HomeyFlowFolders.result $_FlowsFolderID 

        $_FlowsJSON.result.$_ | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathFolders\$_FlowFolder\$_FlowsTitle-$_FlowsID-v$_HomeyVersion-$_ExportDTSt.json" 
    } 
    return $_FlowsJSON.result.Values
}


function Export-HomeyConfig
<#
.Synopsis
   Get Export config from Homey (by Athom http://www.athom.com)
   Beta 0.3.0
.DESCRIPTION
   Get Export information from your LOCAL connected Homey to store on a Windows computer

.EXAMPLE
   Export-HomeyConfig  **** tbd -Scope [All,Flows,Devices]

.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
   Beta beta beta.... 

.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
{
    $Global:_HomeyGetFoldersApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/Folder/" 
    $Global:_HomeyGetFlowsApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/flow/" 
    $Global:_HomeysHeaders = @{"Authorization"="Bearer $_HomeysBearer"}
    $Global:_HomeysContentType = "application/json"
    $Global:_ExportDTSt = "{0:yyyyMMddHHmmss}" -f (get-date)
    $Global:_HomeyVersion =""
  
    $_FoldersJSON = Get-HomeyFoldersStructure 
    # Export to Folders.json
    $_FoldersJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\Folders-v$_HomeyVersion-$_ExportDTSt.json"

    $_FlowsJSON = Export-HomeyFlows 
    $_FlowsJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\Flows-v$_HomeyVersion-$_ExportDTSt.json" 

    $_ZonesJSON = Export-HomeyZonesStructure 

    $_DevicesJSON = Export-HomeyDevices
    $_DevicesJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\Devices-v$_HomeyVersion-$_ExportDTSt.json" 

    # Maybe Get these dynamic 
    # Test is App enabed ?
    @('net.i-dev.betterlogic','nl.bevlogenheid.countdown') | ForEach-Object {
        $_ExportPathAppsVar =  "$_HomeysExportPath\Apps\$_"
        If (!(Test-Path $_ExportPathAppsVar)) { $return = New-Item $_ExportPathAppsVar -ItemType Directory } 
        $return = Export-HomeyAppsVar $_ 
        $return  | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathAppsVar\Vars-$_-$_HomeyVersion-$_ExportDTSt.json" 
    
        }

    # Manager/ System default Homeys settings 
    @('apps/app', 'speech-input/settings', 'speech-output/settings', 'speech-output/voice', 'ledring/brightness', 'ledring/screensaver' ,
       'speaker/settings', 'geolocation' ,  'zwave/state' , 'users/user', 'updates/settings', 'system' ) | ForEach-Object {
        $_ExportPathAppsVar =  "$_HomeysExportPath\Settings\$_"
        If (!(Test-Path $_ExportPathAppsVar)) { $return = New-Item $_ExportPathAppsVar -ItemType Directory } 
        $return = Export-HomeySystemSettings -AppUri $_
        $_Filename = Get-ValidFilename $_
        # $_IllegalFileFolderChars.ToCharArray() | % {$_Filename = $_Filename -replace "$_", "-" }
        $return  | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathAppsVar\Vars-$_Filename-$_HomeyVersion-$_ExportDTSt.json" 
    }

}

function Import-HomeyFlow 
{
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   #Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        # [Alias("ip")] 
        [string] $JSONFile,
        [switch] $OverwriteFlow,
        [switch] $NewFlowID,
        [switch] $RestoreToRoot
    )
    If (Test-Path $JSONFile)     {
        $FlowObject = Get-Content -Raw -Path $JSONFile | ConvertFrom-Json 
        If ($RestoreToRoot) {$FlowObject.folder = $false}
        $FlowID = $FlowObject.id
        $Global:_HomeyUriFlowsApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/flow/$FlowID" 

        $_FlowWR = try {
            Invoke-WebRequest -Uri "$_HomeyUriFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -ErrorAction Ignore 
        } catch { 
            $_.Exception.Response
            If (!$NewFlowID) { Write-Host "Warning: Flow does not Exists! use -NewFlowID " -ForegroundColor Yellow}
        }
        If ($NewFlowID) {
            $NewFlow = New-HomeyFlow
            $FlowObject.id = $NewFlow.id
            $FlowID = $FlowObject.id
            If ($_FlowWR.StatusCode -eq 200) { 
                $FlowObject.title += " (2)" 
                $FlowObject.enabled = $false
            }
        }

        If (($_FlowWR.StatusCode -eq 200) -or $NewFlowID) {
            # Flow Exists
            If ($OverwriteFlow -or $NewFlowID) {
                $CompressedJSONFlow = $FlowObject | ConvertTo-Json -Depth 99  -Compress
                Set-HomeyFlow -ID $FlowID -CompressedJSONFlow $CompressedJSONFlow 
            } Else { Write-Host "Warning: Flow already Exists! use -OverwriteFlow " -ForegroundColor Yellow }
        }
    } Else {
        Write-Host "Error: File $JSONFile not found!" 
    }
 
}


### Internal Functions
function Get-ScriptDirectory {
    Split-Path -parent $PSCommandPath
}

function Get-ValidFilename {
    param( [string]$_RawFilename ) 
    # https://kb.acronis.com/content/39790
    $_IllegalFileFolderChars = '"/<>:'
    # Also fix for * ? \ ^ > |
    $_IllegalFileFolderChars.ToCharArray() | % {$_RawFilename  = $_RawFilename  -replace "$_", "-" }
    @('\\','\*','\?','\^','\|' ) | % {$_RawFilename = $_RawFilename -replace "$_", "-" }
    return $_RawFilename
}

# Base code when loading Module 
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
# Set environment 


Write-Host " HomeyPSCLI Loaded!..." -ForegroundColor Green 
Write-Host "" 
Write-Host " Use Connect-Homey to configure and Test connection:" 
Write-Host "     Connect-Homey [-IP <string>] [-Bearer <string>] [-ExportPath <string>] [-WriteConfig] " -ForegroundColor Green 
Write-Host "" 
Write-Host " use ""Get-Command -Module HomeyPSCLI"" to see the possible commands"
Write-Host " Have Phun with Homey from your PowerShell !" -ForegroundColor Green 

# Export using the manifest.
# Export-ModuleMember -Function Update-Something -Alias *