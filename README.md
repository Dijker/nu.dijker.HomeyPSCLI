# PowerShell CLI for Homey

## Install
Download the Powershell Module files from your Homey to manage Homey from your Windows PowerShell Command Line.

Appstore info in the APPSTORE.md View here: https://apps.athom.com/app/nu.dijker.homeypscli

For more information and examples view the Settings page after installing, go to the forum (*Link to be inserted*) and create Issues (bug reports, feature requests) on Github ( https://github.com/Dijker/nu.dijker.HomeyPSCLI/issues )  

##Warning:

This scripts make it possible to import some of the exported information, fe flows and BetterLogic variables. Possibly you break something in Homey when you import across different Homeys, different firmware versions or after importing edited information.
Using the import functions incorrectly can cause serious, system-wide problems that may require you to factory reset your Homey, restore firmware or even buy a new Homey to correct them. The Creator of the App and Athom cannot guarantee that any problems resulting from the use of these scripts can be solved. Use this tool at your own risk.
*The Creator of the App and Athom are not responsible!!*

##Version History:
* 0.0.8 (20170701)
  Various fixes
* Previous Updates
  Added new exports for Homey v1.1.x firmware
  Added new commandlets Get-HomeyStatistics, Get-HomeyTokens  
  Online Version History https://github.com/Dijker/nu.dijker.HomeyPSCLI/wiki/Release-Notes

## To Do
* Make Connect-Homey more robust on input parameters
* test and fix Import for Apps
	( net.i-dev.betterlogic maybe works, nl.bevlogenheid.countdown isnt working atm)
* Create new functions: Export-InsightsTemplates, Export-InsightsLog, Clear-InsightsLog
* Connect multiple Homeys
* Disable auto updating by switch
* Action and Condition reordering

## Licensing
HomeyPSCLI is free for non-commercial use only. If you wish to use the module and functions/scripts in a company or commercially, you must purchase a site-license by contacting the Author.
