# PowerShell CLI for Homey

## Install
Download the Powershell Module files from your Homey to manage Homey from your Windows PowerShell Command Line.

Get your Bearer token from the Browser and your CloudID
Initialize the PowerShell CLI for Homey

Open your PowerShell or PowerShell ISE, and load the module:

*C:\Homey>  Import-Module HomeyPSCLI*

Connect to Homey:

*PS> Connect-Homey -Bearer "Deaf001000bad1000effe000hace000215c001" -ExportPath "C:\Homey\Backup" -WriteConfig*

To export as much info from Homey I could find to the Export path defined with Connect-Homey use the following:
*PS> Export-HomeyConfig*

Look at the files and folder structure after running this command.

*PS> Get-HomeyFlows  | Where-Object { $_.title -eq "StopMic" }*

*PS> Get-HomeyFlows  | Where-Object { $_.broken -eq $true } | ConvertTo-Json |  ConvertFrom-Json | FT -Property id, title, broken*

*PS> Get-HomeyStatistics

*PS> Get-HomeyTokens

*PS> Remove-HomeyFlow  -ID  "767831a5-98b7-4d44-a389-e13f74a9de4a"*


For more information and examples view the Settings page after installing, go to the forum (*Link to be inserted*) and create Issues (bug reports, feature requests) on Github ( https://github.com/Dijker/nu.dijker.HomeyPSCLI/issues )  

You may leave me a donation if you love my work.
<form action="https://www.paypal.com/cgi-bin/webscr" method="post" target="_top">
<input type="hidden" name="cmd" value="_s-xclick">
<input type="hidden" name="encrypted" value="-----BEGIN PKCS7-----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-----END PKCS7-----
">
<input type="image" src="https://www.paypalobjects.com/nl_NL/NL/i/btn/btn_donate_LG.gif" border="0" name="submit" alt="PayPal, de veilige en complete manier van online betalen.">
<img alt="" border="0" src="https://www.paypalobjects.com/nl_NL/i/scr/pixel.gif" width="1" height="1">
</form>

##Warning:

This scripts make it possible to import some of the exported information, fe flows and BetterLogic variables. Possibly you break something in Homey when you import across different Homeys, different firmware versions or after importing edited information.
Using the import functions incorrectly can cause serious, system-wide problems that may require you to factory reset your Homey, restore firmware or even buy a new Homey to correct them. The Creator of the App and Athom cannot guarantee that any problems resulting from the use of these scripts can be solved. Use this tool at your own risk.
*The Creator of the App and Athom are not responsible!!*

##Version History:
* 0.0.10 (20170701)

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
