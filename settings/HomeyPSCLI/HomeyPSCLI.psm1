<#
.Synopsis
   HomeyPSCLI Module (Homey by Athom http://www.athom.com)
   Beta 0.0.10 20170701
   - Various updates
.DESCRIPTION
   Set IP Address and Bearer for your LOCAL Homey to store in a PowerShell variable Windows computer
   Import-Module HomeyPSCLI [-Force] [-Verbose]
.EXAMPLE
   Get-Command -Module HomeyPSCLI
.NOTES
   General notes
   Beta beta beta....
#>


function Connect-Homey
<#
.Synopsis
   Connect-Homey
   Beta 0.0.10
.DESCRIPTION
   Set IP Address or HostName and Bearer for your LOCAL Homey to store in a PowerShell variable on your Windows computer
   Optional set Export Path for exports of JSON Config files to your disk

.EXAMPLE
    Connect-Homey -IP 1.2.3.4 -Bearer abcdefg -ExportPath C:\HomeyBackup -WriteConfig

.NOTES
    -WriteConfig writes to a file in your module directory the IP/Bearer/CloudID/Exportpath ...
.FUNCTIONALITY
    Connect-Homey tests/connects to Homey, gets the version and sets Global Variables for direct use of other commandleds
#>
{
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1',
                  SupportsShouldProcess=$true,
                  PositionalBinding=$false,
                  HelpUri = 'https://github.com/Dijker/nu.dijker.HomeyPSCLI/wiki',
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

    Write-Verbose "function Connect-Homey"
    Write-Debug "Debug = $PSBoundParameters"
    Write-Verbose "Verbose = $PSBoundParameters"

    $ScriptDirectory = Get-ScriptDirectory
    # Write-Host "ScriptDirectory : $ScriptDirectory"
    If (Test-Path "$ScriptDirectory\Config-HomeyPSCLI.ps1" ) {
            . $ScriptDirectory\Config-HomeyPSCLI.ps1
        }
    If ($ExportPath -ne "") { $Global:_HomeysExportPath = $ExportPath }
    If ($IP -ne "") {
        $Global:_HomeysIP = $IP
        $Global:_HomeysCloudID = ""
    }
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

    If ( $_SystemWR.StatusCode -eq 200 ) {
        "$_HomeysConnectHost HTTP Response OK"
        $Global:_HomeyVersion = $_SystemWR.Headers.'X-Homey-Version'
        "Homey version $Global:_HomeyVersion "
        If ($_HomeysExportPath -ne $null ) { "Homeys Export configured to: $_HomeysExportPath"}
    }

    # !!
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
            }
            Write-Host "To activate, Reload module: Import-Module HomeyPSCLI  -Force" -ForegroundColor Yellow
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
        [Parameter(Mandatory=$true,
                ValueFromPipeline=$false,
                # ValueFromPipelineByPropertyName=$true,
                ValueFromRemainingArguments=$False,
                Position=0,
                ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('net.i-dev.betterlogic','nl.bevlogenheid.countdown')]
        [string] $AppUri
    )
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    Write-Verbose "function Export-HomeyAppsVar"
    try {
        $_HomeyUriGetAppVar = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/settings/app/$AppUri/variables"
        $_AppWR = Invoke-WebRequest -Uri "$_HomeyUriGetAppVar" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
        $_AppJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_AppWR, [System.Collections.Hashtable])
        # $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
        return $_AppJSON.result
    } Catch {return $null }
}

function Export-HomeySystemSettings
{
    param (
    [string] $AppUri )
    Write-Verbose "function Export-HomeySystemSettings"

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
    Write-Verbose "function Import-HomeyAppsVar"
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
            # $_.Exception.Response
            Write-Host "Warning: App not installed! " -ForegroundColor Yellow
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

function Get-HomeyReboot
{
    param (
    [switch] $Verbose
    )
    Write-Verbose "function Get-HomeyReboot"
    $_HomeyUriPendingUpdate = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/system/reboot"

    $_AppWR = Invoke-WebRequest -Uri "$_HomeyUriPendingUpdate" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -Method Post
    $_AppJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_AppWR, [System.Collections.Hashtable])
    $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
    if ( $Verbose ) {
        if (($_AppJSON.status -eq 200 ) -and ( $_AppJSON.result -eq $true )) {
            "Restarting"
        } else {
            "Failed, Result: {0} - Status: {1}" -f $_AppJSON.result ,$_AppJSON.status
        }
    } Else {
        return $_AppJSON
    }
}

function Get-HomeyPendingUpdate
{
    param (
    [switch] $Verbose,
    [switch] $InstallPendingUpdate
    )
    Write-Verbose "function Get-HomeyPendingUpdate"
    $_HomeyUriPendingUpdate = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/updates/update/?cache=0"

    $_AppWR = Invoke-WebRequest -Uri "$_HomeyUriPendingUpdate" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    $_AppJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_AppWR, [System.Collections.Hashtable])
    $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
    If ($_AppJSON.result.count -gt 0 ) {
        If ($Verbose) {
            $_AppJSON.result | ForEach-Object {
                        " Date    : {0}" -f $_.date
                        " Version : {0}" -f $_.version
                        " Size    : {0}" -f $_.size
                        $html = $_.changelog.en
                        if ( $html.length -eq 0 ) { $html = "<i>This update has no notes"}
                        @('br','/li' ) | % {$html = $html -replace "<$_[^>]*?>", "`n" }
                        @('ul','li', '/ul', 'p', '/p') | % {$html = $html -replace "<$_[^>]*?>", "" }
                        @('&amp;' ) | % {$html = $html -replace "$_", "&" }
                        $html += "`n"
                        $html = $html -replace "`n`n", "`n"
                        $html.Split("`n") | ForEach-Object { if ($_ -match '<i>') { $x=$_ ; @('i', '/i') | % { $x = $x -replace "<$_[^>]*?>", "" } ;  Write-Host $x -ForegroundColor Yellow  } else {$_ } }
            }
        } Else {
            return $_AppJSON.result
        }
        If ($InstallPendingUpdate) {
            Write-Host " Please wait fetching and starting update, do not turn off the power! " -ForegroundColor blue -BackgroundColor White

            $_HomeyUriPendingUpdate = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/updates/update/"
            $_AppWR = Invoke-WebRequest -Uri "$_HomeyUriPendingUpdate" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -Method Post
            # Add Testing Result !!!
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
    Write-Verbose "function New-HomeyFlow"
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
    Write-Verbose "function Remove-HomeyFlow"
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
    # Write-Verbose "function Get-HomeyFolderPath"

    If (!$_Key) {
        return ""
    } else {
        If( $_FolderJSON.$_Key.folder -eq $False ) {
            $_FolderName = Get-ValidFilename $_FolderJSON.$_Key.title
            return  $_FolderName
        } else { # Next level Folders
            $_TopFolderName = "$(Get-HomeyFolderPath $_FolderJSON "$($_FolderJSON.$_Key.folder)")"
            $_FolderName = Get-ValidFilename $_FolderJSON.$_Key.title

            return "$_TopFolderName\$_FolderName"
        }
    }
}

function Get-HomeyFoldersStructure
{
    Write-Verbose "function Get-HomeyFoldersStructure"
    $_ExportPathFolders =  "$_HomeysExportPath\Flows"
    $_FoldersWR = Invoke-WebRequest -Uri "$_HomeyGetFoldersApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    $_FoldersJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FoldersWR , [System.Collections.Hashtable]).result
    $Global:_HomeyVersion = $_FoldersWR.Headers.'X-Homey-Version'

    $VerboseMsg =  "Folders Count: {0}" -f $_FoldersJSON.Keys.Count
    Write-Verbose $VerboseMsg
    $_FoldersJSON.keys | ForEach-Object {

        $_FolderRelPath = Get-HomeyFolderPath $_FoldersJSON $_
        If (!(Test-Path $_ExportPathFolders\$_FolderRelPath)) { $return = New-Item $_ExportPathFolders\$_FolderRelPath -ItemType Directory ; Write-Verbose "Creating Folder: $_FolderRelPath" }
    }
    return $_FoldersJSON
}

function Get-HomeyZonePath
{
    Param( $_ZonesJSON, $_Key)
    # Write-Verbose "function Get-HomeyZonePath"
    If( $_ZonesJSON.$_Key.parent -eq $False ) {
        $_ZoneName = Get-ValidFilename $_ZonesJSON.$_Key.name
        return  $_ZoneName
    } else { # Next level Folders
        $_TopZoneName = Get-HomeyZonePath $_ZonesJSON $_ZonesJSON.$_Key.parent
        $_ZoneName = Get-ValidFilename $_ZonesJSON.$_Key.name
        return "$_TopZoneName\$_ZoneName"
    }
}

function Export-HomeyZonesStructure
{
    Write-Verbose "function Export-HomeyZonesStructure"
    $_ExportPathZones =  "$_HomeysExportPath\Zones"
    $_FoldersWR = Invoke-WebRequest -Uri "$_HomeyGetZonesApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    $_ZonesJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FoldersWR , [System.Collections.Hashtable]).result

    # Zones
    $VerboseMsg =  "Zones Count: {0}" -f $_ZonesJSON.keys.Count
    Write-Verbose $VerboseMsg
    $_ZonesJSON.keys | ForEach-Object {
        $_ZonesRelPath = Get-HomeyZonePath $_ZonesJSON $_
        If (!(Test-Path $_ExportPathZones\$_ZonesRelPath)) { $return = New-Item $_ExportPathZones\$_ZonesRelPath -ItemType Directory ;  Write-Verbose "Creating Zone: $_ZonesRelPath"}
    }

    # Write Only Incremental files
    $_LastFile = Get-ChildItem "$_HomeysExportPath\Zones-v*.json"  -Filter  *.json -File -ErrorAction SilentlyContinue | sort -Descending -property LastWriteTime
    If ($_LastFile -ne $null) {
        $_NewJSON = $_ZonesJSON | ConvertTo-Json -depth 99 | Out-String
        $_LastJSONFile = Get-Content -Raw $_LastFile[0]
        If ($_NewJSON -eq $_LastJSONFile) {$_WriteJSON = $false} else {$_WriteJSON = $true}
    } else { $_WriteJSON = $true }
    If ($_WriteJSON -eq $true ) {
        Write-Verbose "Writing Zones"
        $_ZonesJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\Zones-v$_HomeyVersion-$_ExportDTSt.json"
    }
    return $_ZonesJSON
}

function Get-HomeyZonesStructure
{
    Write-Verbose "function Get-HomeyZonesStructure"
    $_FoldersWR = Invoke-WebRequest -Uri "$_HomeyGetZonesApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Web.Extensions")
    $_ZonesJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FoldersWR , [System.Collections.Hashtable]).result

    return $_ZonesJSON
}

function Export-HomeyDevices
{
    Write-Verbose "function Export-HomeyDevices"
    $Global:_ExportDTSt = "{0:yyyyMMddHHmmss}" -f (get-date)

    $_ExportPathDevices =  "$_HomeysExportPath\Zones"
    $_HomeyGetDeviceApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/devices/device"
    $_DevicesWR = Invoke-WebRequest -Uri "$_HomeyGetDeviceApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    $_DevicesJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_DevicesWR.Content, [System.Collections.Hashtable])
    $Global:_HomeyVersion = $_DevicesWR.Headers.'X-Homey-Version'
    $_HomeyDeviceZones = Get-HomeyZonesStructure

    $VerboseMsg =  "Devices Count: {0}" -f $_DevicesJSON.result.keys.Count
    Write-Verbose $VerboseMsg

    $_DevicesJSON.result.keys | ForEach-Object {
        $_DeviceZoneID = $_DevicesJSON.result.$_.zone.id
        $_DeviceID = $_DevicesJSON.result.$_.id
        $_DeviceName = $_DevicesJSON.result.$_.name
        $_DeviceFileName = Get-ValidFilename $_DeviceName
        # $_IllegalFileFolderChars.ToCharArray() | % {$_DeviceFileName = $_DeviceFileName -replace "$_", "-" }
        $_DeviceFolder = Get-HomeyZonePath $_HomeyDeviceZones $_DeviceZoneID
        # Write-Verbose "Device: $_DeviceFolder\$_DeviceFileName"
        # $_DevicesJSON.result.$_ | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathDevices\$_DeviceFolder\$_DeviceFileName-$_DeviceID-v$_HomeyVersion-$_ExportDTSt.json"

        # Write Only Incremental files
        $_LastFile = Get-ChildItem "$_ExportPathDevices\$_DeviceFolder\$_DeviceFileName-$_DeviceID-v*.json"  -Filter  *.json -File -ErrorAction SilentlyContinue | sort -Descending -property LastWriteTime
        If ($_LastFile -ne $null) {
            $_NewJSON = $_DevicesJSON.result.$_ | ConvertTo-Json -depth 99 | Out-String
            $_LastJSONFile = Get-Content -Raw $_LastFile[0]
            If ($_NewJSON -eq $_LastJSONFile) {$_WriteJSON = $false} else {$_WriteJSON = $true}
        } else { $_WriteJSON = $true }
        If ($_WriteJSON -eq $true ) {
            Write-Verbose "Device: $_DeviceFolder\$_DeviceFileName"
            $_DevicesJSON.result.$_ | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathDevices\$_DeviceFolder\$_DeviceFileName-$_DeviceID-v$_HomeyVersion-$_ExportDTSt.json"
        }
    }
    return $_DevicesJSON.result.values
}

function Get-HomeyDevices
{
    Write-Verbose "function Get-HomeyDevices"
    $_HomeyGetDeviceApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/devices/device"
    $_DevicesWR = Invoke-WebRequest -Uri "$_HomeyGetDeviceApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    $_DevicesJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_DevicesWR.Content, [System.Collections.Hashtable])
    return $_DevicesJSON.result.values
}

function Debug-HomeyAppVariableUsage
{
    param (
        [Parameter(Mandatory=$true,
                   ValueFromPipeline=$false,
                   # ValueFromPipelineByPropertyName=$true,
                   ValueFromRemainingArguments=$False,
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('net.i-dev.betterlogic','nl.bevlogenheid.countdown')]
        [string]$ApplicationName
        #[switch]$FixMissingVars
    )
    # Set Vars
    # If ($ApplicationName -eq $null) {
    #    $ApplicationName = "net.i-dev.betterlogic"
    # }
    # $ApplicationName = "nl.bevlogenheid.countdown"
    $_HomeyUriGetAppVar = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/settings/app/$ApplicationName/variables"
    If ($ApplicationName -eq 'nl.bevlogenheid.countdown' ) {
        $_HomeyUriPutAppVar = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/settings/app/$ApplicationName/changedvariables"
    }  Else {
        $_HomeyUriPutAppVar = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/settings/app/$ApplicationName/variables"
    }
    $_AppWR = try {
        Invoke-WebRequest -Uri "$_HomeyUriGetAppVar" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -ErrorAction Ignore
    } catch {
        # $_.Exception.Response
        Write-Host "Warning: App not installed! " -ForegroundColor Yellow
        break
    }
    $AppVariables = Export-HomeyAppsVar -AppUri "$ApplicationName"

    If ( $AppVariables -eq $null) { $AppVariableNames = @() } else
    {
        $AppVariableNames = ($AppVariables).name
    }

    #get all devices
    $allDevices = Get-HomeyDevices
    $AppDevices = $allDevices | where {$_.driver.uri -eq "Homey:app:$ApplicationName"}
    # $AppDeviceVars = $AppDevices.data.id
    #get all flows
    $allFlows = Get-HomeyFlows

    #find flows with better logic variables in the
    #trigger, conditions or actions
    $cardst = $allFlows.trigger | where {$_.uri -contains "homey:app:$ApplicationName" }
    $cardsc = $allFlows.conditions | where {$_.uri -contains "homey:app:$ApplicationName" }
    $cardsa = $allFlows.actions | where {$_.uri -contains "homey:app:$ApplicationName" }

    $UsedVariables = @()
    $UsedVariables = ,$AppDevice.data.id + $cardst.args.variable.name + $cardsc.args.variable.name + $cardsa.args.variable.name | select -Unique

    #compare the list of better logic variables to
    #the list of flows containing better logic variables
    $comparison = Compare-Object -ReferenceObject $AppVariableNames -DifferenceObject $UsedVariables

    #write output on screen
    Write-Host -ForegroundColor Yellow "BetterLogic variables only existing in App:"
    $onlyInApp = ($comparison | where {$_.SideIndicator -eq "<="}).InputObject
    $onlyInApp | foreach {Write-Host $_}
    Write-Host -ForegroundColor Yellow "App variables only existing in flows or devices (missing in App Settings):"
    $onlyInFlows = ($comparison | where {$_.SideIndicator -eq "=>"}).InputObject
    $onlyInFlows | foreach {Write-Host $_}
    Write-Host -ForegroundColor Yellow "List of flows or devices containing orphaned variables:"
    foreach ($VarName in $onlyInFlows) {

        $flowsContainingOrphanedVars = @()
        $flowsContainingOrphanedVars += , ($allFlows | where {$_.trigger.args.variable.name -contains $VarName })
        $flowsContainingOrphanedVars += ($allFlows | where {$_.actions.args.variable.name -contains $VarName })
        $flowsContainingOrphanedVars += ($allFlows | where {$_.conditions.args.variable.name -contains $VarName })
        $flowsContainingOrphanedVars = $flowsContainingOrphanedVars | select -Unique


        $flowsNameContainingOrphanedVars = @()
        $flowsNameContainingOrphanedVars += ,($allFlows | where {$_.trigger.args.variable.name -contains $VarName }).title
        $flowsNameContainingOrphanedVars += ($allFlows | where {$_.actions.args.variable.name -contains $VarName }).title
        $flowsNameContainingOrphanedVars += ($allFlows | where {$_.conditions.args.variable.name -contains $VarName }).title
        $flowsNameContainingOrphanedVars = $flowsNameContainingOrphanedVars | select -Unique

        If ($flowsNameContainingOrphanedVars.Count -ne 0) {
            Write-Host -ForegroundColor Red $VarName
            Write-Host -ForegroundColor Yellow "Exists in these flows:"
            $flowsNameContainingOrphanedVars

        } Else {
            Write-Host -ForegroundColor Red $VarName
            Write-Host -ForegroundColor Yellow "Exists in this Device:"
            $DeviceContainingOrphanedVars = ($AppDevices |  where {$_.data.id -contains $VarName } ).name
            $DeviceContainingOrphanedVars
        }
        <# Still doesn't work, working on this.....
        If ( $FixMissingVars ) {
            # flowsContainingOrphanedVars
            If ( $flowsContainingOrphanedVars[0].trigger.args.variable.name -contains $VarName ) {
                $B =  $flowsContainingOrphanedVars[0].trigger.args.variable | ConvertTo-Json -Depth 99  -Compress
            } Else {
                If ( $flowsContainingOrphanedVars[0].actions.args.variable.name -contains $VarName ) {
                      $B =  $flowsContainingOrphanedVars[0].actions.args.variable | ConvertTo-Json -Depth 99  -Compress
                } Else {
                    If ( $flowsContainingOrphanedVars[0].conditions.args.variable.name -contains $VarName ) {
                         $B =  $flowsContainingOrphanedVars[0].conditions.args.variable | ConvertTo-Json -Depth 99  -Compress
                    }
                }

            }
            $CompressedJSONVarValue = "{""value"":{""variable"":$B}}"
            $_AppWR = Invoke-WebRequest -Uri "$_HomeyUriPutAppVar" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -Method Put -Body $CompressedJSONVarValue
            $_AppJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_AppWR, [System.Collections.Hashtable])
            # $Global:_HomeyVersion = $_AppWR.Headers.'X-Homey-Version'
            return $_AppJSON.result
        }
        #>
    }
}

function Get-HomeyFlows
{
    Write-Verbose "function Get-HomeyFlows"
    $_FlowsWR = Invoke-WebRequest -Uri "$_HomeyGetFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    $_FlowsJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FlowsWR, [System.Collections.Hashtable])
    return $_FlowsJSON.result.Values
}

function Get-HomeyFlow
{
    param (
    [string] $ID )
    Write-Verbose "function Get-HomeyFlow"

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
    [string] $CompressedJSONFlow )
    Write-Verbose "function Set-HomeyFlow"
    $Global:_HomeyUriFlowsApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/flow/$ID"
    $_FlowWR = try {
        Invoke-WebRequest -Uri "$_HomeyUriFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -ErrorAction Ignore
    } catch {
        $_.Exception.Response
    }
    If ($_FlowWR.StatusCode -eq 200 ) {
        $_FlowWR = Invoke-WebRequest -Uri "$_HomeyUriFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -Method Put -Body $CompressedJSONFlow
        $_FlowJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FlowWR, [System.Collections.Hashtable])
    } Else {
        Write-Host "Error finding Flow ID: $ID" -ForegroundColor Yellow
    }
    return $_FlowJSON.result
}

function Export-HomeyFlows
{
    Write-Verbose "Export-HomeyFlows"
    $Global:_ExportDTSt = "{0:yyyyMMddHHmmss}" -f (get-date)

    $_ExportPathFolders =  "$_HomeysExportPath\Flows"
    $_FlowsWR = Invoke-WebRequest -Uri "$_HomeyGetFlowsApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType
    $_FlowsJSON = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FlowsWR, [System.Collections.Hashtable])
    $Global:_HomeyVersion = $_FlowsWR.Headers.'X-Homey-Version'
    $_HomeyFlowFolders = Get-HomeyFoldersStructure

    $VerboseMsg =  "Flows Count: {0}" -f $_FlowsJSON.result.keys.Count
    Write-Verbose $VerboseMsg

    $_FlowsJSON.result.keys | ForEach-Object {
        $_FlowsFolderID = $_FlowsJSON.result.$_.folder
        $_FlowsID = $_FlowsJSON.result.$_.id
        $_FlowsTitle = Get-ValidFilename $_FlowsJSON.result.$_.title
        $_FlowFolder = Get-HomeyFolderPath $_HomeyFlowFolders $_FlowsFolderID

        # Write Only Incremental files
        $_LastFile = Get-ChildItem "$_ExportPathFolders\$_FlowFolder\$_FlowsTitle-$_FlowsID-v*.json"  -Filter  *.json -File -ErrorAction SilentlyContinue | sort -Descending -property LastWriteTime
        If ($_LastFile -ne $null) {
            $_NewJSON = $_FlowsJSON.result.$_ | ConvertTo-Json -depth 99 | Out-String
            $_LastJSONFile = Get-Content -Raw $_LastFile[0]
            If ($_NewJSON -eq $_LastJSONFile) {$_WriteJSON = $false} else {$_WriteJSON = $true}
        } else { $_WriteJSON = $true }
        If ($_WriteJSON -eq $true ) {
            Write-Verbose "Writing Flow: $_FlowFolder\$_FlowsTitle"
            $_FlowsJSON.result.$_ | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathFolders\$_FlowFolder\$_FlowsTitle-$_FlowsID-v$_HomeyVersion-$_ExportDTSt.json"
        }
    }
    #
    # return $_FlowsJSON.result.Values Removed Values 20161230 22u27
    return $_FlowsJSON.result
}


function Export-HomeyConfig
<#
.Synopsis
   Get Export config from Homey (by Athom http://www.athom.com)
   Beta 0.0.10
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
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1',
                  SupportsShouldProcess=$true,
                  PositionalBinding=$false,
                  HelpUri = 'https://github.com/Dijker/nu.dijker.HomeyPSCLI/wiki',
                  ConfirmImpact='Medium')]
    Param
    (
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

        [switch]
        $FlowBackupZip
        )
    Write-Verbose "function Export-HomeyConfig"
    If ($ExportPath -ne "") { $_HomeysExportPath = $ExportPath }
    $Global:_HomeyGetFoldersApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/Folder/"
    $Global:_HomeyGetFlowsApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/flow/flow/"
    $Global:_HomeysHeaders = @{"Authorization"="Bearer $_HomeysBearer"}
    $Global:_HomeysContentType = "application/json"
    $Global:_ExportDTSt = "{0:yyyyMMddHHmmss}" -f (get-date)
    $Global:_HomeyVersion = ""
    If (!(Test-Path $_HomeysExportPath)) { $return = New-Item $_HomeysExportPath -ItemType Directory }


    $_ExportPathStats = $_HomeysExportPath

    $return = Get-HomeyStatistics -JSON
    $_hostname = $return.hostname
    $_BackupDateZ = $return.date
    # $return  | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathAppsVar\Vars-$_Filename-v$_HomeyVersion-$_ExportDTSt.json"

    # Write Only Incremental files
    $_LastFile = Get-ChildItem "$_ExportPathStats\Stats-$_hostname-v*.json"  -Filter  *.json -File -ErrorAction SilentlyContinue | sort -Descending -property LastWriteTime
    If ($_LastFile -ne $null) {
        $_NewJSON = $return  | ConvertTo-Json -depth 99 | Out-String
        $_LastJSONFile = Get-Content -Raw $_LastFile[0]
        If ($_NewJSON -eq $_LastJSONFile) {$_WriteJSON = $false} else {$_WriteJSON = $true}
    } else { $_WriteJSON = $true }
    If ($_WriteJSON -eq $true ) {
        Write-Verbose "Writing Stats - $_Filename"
        $return  | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathStats\Stats-$_hostname-v$_HomeyVersion-$_ExportDTSt.json"
    }

    $_FoldersJSON = Get-HomeyFoldersStructure
    # Export to Folders.json
    # $_FoldersJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\FlowFolders-v$_HomeyVersion-$_ExportDTSt.json"

    # Write Only Incremental files
    $_LastFile = Get-ChildItem "$_HomeysExportPath\FlowFolders-$_hostname-v*.json"  -Filter  *.json -File -ErrorAction SilentlyContinue | sort -Descending -property LastWriteTime
    If ($_LastFile -ne $null) {
        $_NewJSON = $_FoldersJSON | ConvertTo-Json -depth 99 | Out-String
        $_LastJSONFile = Get-Content -Raw $_LastFile[0]
        If ($_NewJSON -eq $_LastJSONFile) {$_WriteJSON = $false} else {$_WriteJSON = $true}
    } else { $_WriteJSON = $true }
    If ($_WriteJSON -eq $true ) {
        Write-Verbose "Writing AllFolders"
        $_FoldersJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\FlowFolders-$_hostname-v$_HomeyVersion-$_ExportDTSt.json"
    }

    $_FlowsJSON = Export-HomeyFlows
    # $_FlowsJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\Flows-v$_HomeyVersion-$_ExportDTSt.json"

    # Write Only Incremental files
    $_LastFile = Get-ChildItem "$_HomeysExportPath\Flows-$_hostname-v*.json"  -Filter  *.json -File -ErrorAction SilentlyContinue | sort -Descending -property LastWriteTime
    If ($_LastFile -ne $null) {
        $_NewJSON = $_FlowsJSON | ConvertTo-Json -depth 99 | Out-String
        $_LastJSONFile = Get-Content -Raw $_LastFile[0]
        If ($_NewJSON -eq $_LastJSONFile) {$_WriteJSON = $false} else {$_WriteJSON = $true}
    } else { $_WriteJSON = $true }
    If ($_WriteJSON -eq $true ) {
        Write-Verbose "Writing All Flows"
        $_FlowsJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\Flows-$_hostname-v$_HomeyVersion-$_ExportDTSt.json"
    }

    If ($FlowBackupZip)  {
        $_FoldersJSON | ConvertTo-Json -depth 99 -Compress | Out-File -FilePath "$_HomeysExportPath\folders.json" -Encoding ascii
        $_FlowsJSON | ConvertTo-Json -depth 99 -Compress | Out-File -FilePath "$_HomeysExportPath\flows.json" -Encoding ascii
        # {"backUpDate":"2016-12-30T22:10:01.392Z","backUpVersion":2,"homeyName":"BerliozHomeyRD"}
        "{`"backUpDate`":`"$_BackupDateZ`",`"backUpVersion`":2,`"homeyName`":`"$_hostname`"}" | Out-File -FilePath "$_HomeysExportPath\backUpInfo.json" -Encoding ascii
        Compress-Archive "$_HomeysExportPath\flows.json","$_HomeysExportPath\folders.json","$_HomeysExportPath\backUpInfo.json" "$_HomeysExportPath\$_hostname-v$_HomeyVersion-$_ExportDTSt.zip"
        Remove-Item -Path "$_HomeysExportPath\flows.json","$_HomeysExportPath\folders.json","$_HomeysExportPath\backUpInfo.json"
    }
    $_ZonesJSON = Export-HomeyZonesStructure

    $_DevicesJSON = Export-HomeyDevices
    # $_DevicesJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\Devices-v$_HomeyVersion-$_ExportDTSt.json"

    # Write Only Incremental files
    $_LastFile = Get-ChildItem "$_HomeysExportPath\Devices-$_hostname-v*.json"  -Filter  *.json -File -ErrorAction SilentlyContinue | sort -Descending -property LastWriteTime
    If ($_LastFile -ne $null) {
        $_NewJSON = $_DevicesJSON | ConvertTo-Json -depth 99 | Out-String
        $_LastJSONFile = Get-Content -Raw $_LastFile[0]
        If ($_NewJSON -eq $_LastJSONFile) {$_WriteJSON = $false} else {$_WriteJSON = $true}
    } else { $_WriteJSON = $true }
    If ($_WriteJSON -eq $true ) {
        Write-Verbose "Writing All Devices"
        $_DevicesJSON | ConvertTo-Json -depth 99 | Out-File -FilePath "$_HomeysExportPath\Devices-$_hostname-v$_HomeyVersion-$_ExportDTSt.json"
    }

    # Maybe Get these dynamic
    # Test is App enabed ?
    @('net.i-dev.betterlogic','nl.bevlogenheid.countdown' ) | ForEach-Object {
        $_ExportPathAppsVar =  "$_HomeysExportPath\Apps\$_"
        If (!(Test-Path $_ExportPathAppsVar)) { $return = New-Item $_ExportPathAppsVar -ItemType Directory }
        $return = Export-HomeyAppsVar $_
        # $return  | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathAppsVar\Vars-$_-v$_HomeyVersion-$_ExportDTSt.json"

        # Write Only Incremental files
        $_LastFile = Get-ChildItem "$_ExportPathAppsVar\Vars-$_hostname-$_-v*.json"  -Filter  *.json -File -ErrorAction SilentlyContinue | sort -Descending -property LastWriteTime
        If ($_LastFile -ne $null) {
            $_NewJSON = $return  | ConvertTo-Json -depth 99 | Out-String
            $_LastJSONFile = Get-Content -Raw $_LastFile[0]
            If ($_NewJSON -eq $_LastJSONFile) {$_WriteJSON = $false} else {$_WriteJSON = $true}
        } else { $_WriteJSON = $true }
        If ($_WriteJSON -eq $true ) {
            Write-Verbose "Writing Vars - $_"
            $return  | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathAppsVar\Vars-$_hostname-$_-v$_HomeyVersion-$_ExportDTSt.json"
        }


        }

    # Manager/ System default Homeys settings
    @('apps/app', 'speech-input/settings', 'speech-output/settings', 'speech-output/voice', 'ledring/brightness', 'ledring/screensaver', 'flow/token',
       'speaker/settings', 'geolocation', 'notifications/notification', 'notifications/origin',  'zwave/state' , 'users/user', 'updates/settings', 'updates/update', 'system' , 'system/memory', 'system/storage') | ForEach-Object {
        $_ExportPathAppsVar =  "$_HomeysExportPath\Settings\$_"
        If (!(Test-Path $_ExportPathAppsVar)) { $return = New-Item $_ExportPathAppsVar -ItemType Directory }
        $return = Export-HomeySystemSettings -AppUri $_
        $_Filename = Get-ValidFilename $_
        # $return  | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathAppsVar\Vars-$_Filename-v$_HomeyVersion-$_ExportDTSt.json"

        # Write Only Incremental files
        $_LastFile = Get-ChildItem "$_ExportPathAppsVar\Vars-$_hostname-$_Filename-v*.json"  -Filter  *.json -File -ErrorAction SilentlyContinue | sort -Descending -property LastWriteTime
        If ($_LastFile -ne $null) {
            $_NewJSON = $return  | ConvertTo-Json -depth 99 | Out-String
            $_LastJSONFile = Get-Content -Raw $_LastFile[0]
            If ($_NewJSON -eq $_LastJSONFile) {$_WriteJSON = $false} else {$_WriteJSON = $true}
        } else { $_WriteJSON = $true }
        If ($_WriteJSON -eq $true ) {
            Write-Verbose "Writing Vars - $_Filename"
            $return  | ConvertTo-Json -depth 99 | Out-File -FilePath "$_ExportPathAppsVar\Vars-$_hostname-$_Filename-v$_HomeyVersion-$_ExportDTSt.json"
        }
    }

}

function Get-HomeyStatistics {
    [cmdletbinding(DefaultParameterSetName=’Object’)]
    Param (
        # Param1 help description
        [Parameter(ParameterSetName='JSON')]
        # [AllowNull()]
        # [AllowEmptyCollection()]
        # [AllowEmptyString()]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        # [ValidateScript({$true})]
        #[ValidateRange(0,5)]
        [switch]
        $JSON,

        # Param2 help description
        [Parameter(ParameterSetName='Object')]
        # [AllowNull()]
        # [AllowEmptyCollection()]
        # [AllowEmptyString()]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        # [ValidateScript({$true})]
        #[ValidateRange(0,5)]
        [switch]
        $Object
    )
    If (!$JSON -and !$Object) { $Object = $true }

    $return = Export-HomeySystemSettings -AppUri 'system'
    $apps = Export-HomeySystemSettings -AppUri 'apps/app'
    $return.Apps = $apps.Keys.count
    $return.AppsDisabled = ($apps.Values | Where-Object {$_.enabled -eq $false}).count
    $devices =  Get-HomeyDevices
    $return.Devices = $devices.count
    $return.DevicesOffline = ($devices | Where-Object {$_.online -eq $false }).count
    $return.Zones = (Get-HomeyZonesStructure).values.count
    $flows = Get-HomeyFlows
    $return.Flows = $flows.count
    $return.FlowsBroken = ( $flows | Where-Object { $_.broken -eq $True} ).count
    $return.Folders = (Get-HomeyFoldersStructure).values.count
    $tokens = Export-HomeySystemSettings -AppUri 'flow/token'
    $return.Tokens = $tokens.count
    $return.DeviceTokens = ($Tokens | Where-Object { $_.uriObj.type -eq 'device'}).count
    $return.AppTokens = ($Tokens | Where-Object { $_.uriObj.type -eq 'app'}).count

    If ($Object)  {
<#      $myObject = @(New-Object System.Object)
        $return.keys | ForEach-Object  { $V = $return.$_ | ConvertTo-Json -Depth 99 ; $myObject[0] | Add-Member -type NoteProperty -name "$_" -Value "$V" }
        return $myObject  #>
        return $return | ConvertTo-Json|  ConvertFrom-Json
    }
    If ($JSON)  { return $return }
}


Function Get-HomeyTokens {

    [cmdletbinding(DefaultParameterSetName=’Object’)]
    Param (
        # Param1 help description
        [Parameter(ParameterSetName='JSON')]
        # [AllowNull()]
        # [AllowEmptyCollection()]
        # [AllowEmptyString()]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        # [ValidateScript({$true})]
        #[ValidateRange(0,5)]
        [switch]
        $JSON,

        # Param2 help description
        [Parameter(ParameterSetName='Object')]
        # [AllowNull()]
        # [AllowEmptyCollection()]
        # [AllowEmptyString()]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        # [ValidateScript({$true})]
        #[ValidateRange(0,5)]
        [switch]
        $Object
    )
    If (!$JSON -and !$Object) { $Object = $true }

    $return = Export-HomeySystemSettings -AppUri 'flow/token'

    If ($Object)  {
<#      $myObject = @()
        $return | ForEach-Object {
            $Val = $_
            $myObject1 = @(New-Object System.Object)
            $Val.keys | ForEach-Object { $V = $Val.$_ | ConvertTo-Json -Depth 99 ; $myObject1 | Add-Member -type NoteProperty -name "$_" -Value "$V" }
            $myObject += $myObject1
        }
        # $return[0].keys | ForEach-Object  { $V = $return[0].$_ | ConvertTo-Json -Depth 99 ; $myObject | Add-Member -type NoteProperty -name "$_" -Value "$V" }
        return $myObject #>
        return $return | ConvertTo-Json|  ConvertFrom-Json
    }
    If ($JSON)  { return $return }
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
        [switch] $RestoreToRoot )
    Write-Verbose "function Import-HomeyFlow "

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


<#
.Synopsis
   Record RF waves from 433 or 868 Mhz radios to analize
.DESCRIPTION
   Record RF waves from 433 or 868 Mhz radios to analize.
   Homey produces a Jaged Array of timings of everythin Homey hears.

.EXAMPLE
   Example:
    Read-HomeyRFuC -F433 -Timeout 20 -FilePath .\OutputFile433.txt -Append
    Read-HomeyRFuC -F868 -Timeout 20 -FilePath .\OutputFile868.txt -Append

.INPUTS
.OUTPUTS
   Output from this cmdlet:
   - File with lines of timings if filepath specifed
   - Jagged Array of timings
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
#>
function Read-HomeyRFuC
{
    Param (
    [switch]$F433,
    [switch]$F868,
    [string]$FilePath,
    [Switch]$Append,
    [int16]$Timeout
    )

    If ($F433 -and $F868 ) { Throw "Error: Can't analyze 433 and 868 at once!" }
    $Frequency = "433"
    If ($F868 ) { $Frequency = "868"}
    If ($Timeout -eq $null ) {$Timeout = 5 }
    $CompressedJSONVarValue = "{""frequency"":""$Frequency"",""timeout"":""$Timeout""}"
    $_Homey_uC_RecApi = "$_HomeysProtocol`://$_HomeysConnectHost/api/manager/microcontroller/record"
    $_FlowWR = Invoke-WebRequest -Uri "$_Homey_uC_RecApi" -Headers $_HomeysHeaders  -ContentType $_HomeysContentType -Method Post -Body $CompressedJSONVarValue
    $Key = (New-Object System.Web.Script.Serialization.JavaScriptSerializer).Deserialize($_FlowWR.Content, [System.Collections.Hashtable])
    If ($FilePath -eq "") {
        return ($Key.result  | ForEach-Object { "$_"})
    } else {
        $Key.result  | ForEach-Object { "$_"} | Out-File -FilePath $FilePath -Append:$Append
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
    @('\[',']') | % {$_RawFilename = $_RawFilename -replace "$_", "-" }
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
