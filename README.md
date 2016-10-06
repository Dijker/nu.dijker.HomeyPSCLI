# PowerShell CLI for Homey

## Install
Download the Powershell Module files from your Homey to manage Homey from your Windows PowerShell Command Line.

Get your Bearer token from the Browser and your CloudID
Initialize the PowerShell CLI for Homey

Open your PowerShell or PowerShell ISE, and load the module: 
*C:\Homey>  Import-Module HomeyPSCLI*
Connect to Homey:
*PS> Connect-Homey -Bearer "Deaf001000bad1000effe000hace000215c001" -CloudID "Deaf010bad10effe0hace021" -ExportPath "C:\Homey\Backup" -WriteConfig*

To export as much info from Homey I could find to the Export path defined with Connect-Homey use the following:
*PS> Export-HomeyConfig*

Look at the files and folder structure after running this command.

*PS> Get-HomeyFlows  | Where-Object { $_.title -eq "StopMic" }*

*PS> Get-HomeyFlows  | Where-Object { $_.broken -eq $true } | ConvertTo-Json |  ConvertFrom-Json | FT -Property id, title, broken*

*PS> Remove-HomeyFlow  -ID  "767831a5-98b7-4d44-a389-e13f74a9de4a"*


For more information and examples view the Settings page after installing, go to the forum (*Link to be inserted*) and create Issues (bug reports, feature requests) on Github (https://github.com/Dijker/nu.dijker.HomeyPSCLI/issues)  

##Warning:

This scripts make it possible to import some of the exported information, fe flows and BetterLogic variables. Possibly you break something in Homey when you import across different Homeys, different firmware versions or after importing edited information. 
Using the import functions incorrectly can cause serious, system-wide problems that may require you to factory reset your Homey, restore firmware or even buy a new Homey to correct them. The Creator of the App and Athom cannot guarantee that any problems resulting from the use of these scripts can be solved. Use this tool at your own risk.
*The Creator of the App and Athom are not responsible!!*

##Version History:
0.0.1 First version created local (20161005)

## To Do 
* Make Connect-Homey more robust on imput parameters
* test and fix Import for Apps ( net.i-dev.betterlogic maybe works, nl.bevlogenheid.countdown isnt working atm)
* Create an Auto-Update function for the module.
* Create new functions: Export-Zones, Export-Devices, Export-InsightsTemplates, Export-InsightsLog, Clear-InsightsLog
* Redraw a icon.svg that looks good when displayed on Homey
* Connect multiple Homeys 

## Licensing
HomeyPSCLI is free for non-commercial use only. If you wish to use the module and functions/scripts in a company or commercially, you must purchase a site-license by contacting the Author.

