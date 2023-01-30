###############################################################################################################
#
# ABOUT THIS PROGRAM
#
#   Maconomy_Shortcut_Remediate.ps1
#   https://github.com/Headbolt/Maconomy_Shortcut_Remediate
#
#   This script was designed to create specific Maconomy shortcuts
#
#	Intended use is in Microsoft Endpoint Manager, as the "Remediate" half of a Proactive Remediation Script
#	The "Check" half is found here https://github.com/Headbolt/Maconomy_Shortcut_Check
#
###############################################################################################################
#
# HISTORY
#
#   Version: 1.0 - 27/01/2023
#
#   - 27/01/2023 - V1.0 - Created by Headbolt
#
###############################################################################################################
#
#   DEFINE VARIABLES & READ IN PARAMETERS
#
###############################################################################################################
#
$CreateShortcut=0 # Setting ExitCode Variable to an initial 0
$global:PathTestResult="0" # Setting Global PathTestResult Variable to an initial 0
#
$global:NewShortcut="c:\Users\Public\Desktop\Deltek Maconomy 2.5.1 - PROD.lnk" # Setting the new Shortcut Path
$global:TargetFile = "C:\Program Files (x86)\Deltek Maconomy 2.5.1\Maconomy.exe" # Set Path to the Application
$global:Arguments = "-a https://server.address.com -p 443 -s databasename -vmargs -Duser.home=%userprofile% -Dhttps.protocols=TLSv1.2" # Set Arguments required
#
###############################################################################################################
#
# SCRIPT CONTENTS - DO NOT MODIFY BELOW THIS LINE
#
###############################################################################################################
#
# Defining Functions
#
###############################################################################################################
#
# Path Test Function
#
Function PathTest ($Path){
	$PathTest=(Test-Path $global:Path)
	#
	Write-Host 'Checking For File' $global:Path	
	Write-Host
	If ($PathTest -eq $True)
		{
			Write-Host 'File' $global:Path 'Exists'
			$global:PathTestResult="1"
		}
	Else
		{
			Write-Host 'File' $global:Path 'Does Not Exist'
			$global:PathTestResult="0"
		}
	Write-Host
	Write-Host '###############################################################################################################'
	Write-Host
}
#
###############################################################################################################
#
# Creation Function
#
Function Creation {
	Write-Host 'Creating Shortcut' $global:NewShortcut
	$WScriptShell = New-Object -ComObject WScript.Shell
	$Shortcut = $WScriptShell.CreateShortcut($global:NewShortcut)
	$Shortcut.TargetPath = $global:TargetFile 
	$Shortcut.Arguments = $global:Arguments 
	$Shortcut.Save()
	Write-Host
	Write-Host '###############################################################################################################'
	Write-Host
}
#
###############################################################################################################
#
# Deletion Function
#
Function Deletion {
	Write-Host 'Attempting to Delete file' $global:Path
	Remove-Item -Path $global:Path -Force -erroraction 'silentlycontinue'
	Write-Host
	Write-Host '###############################################################################################################'
	Write-Host
}
#
###############################################################################################################
#
# End Of Function Definition
#
###############################################################################################################
#
# Begin Processing
#
###############################################################################################################
#
Write-Host
Write-Host '###############################################################################################################'
Write-Host
#
$global:Path="c:\Users\Public\Desktop\Deltek Maconomy 2.5.1.lnk" # Setting the Path variable for this test
PathTest # Testing the if the Path exists
If ($PathTestResult -eq "1")
	{
		$CreateShortcut++
		Deletion
	}
#
$global:Path="c:\Users\Public\Desktop\Deltek Maconomy 2.5.1 - PROD.lnk" # Setting the Path variable for this test
PathTest # Testing the if the Path exists
If ($PathTestResult -eq "0")
	{
		$CreateShortcut++
	}
#
If ($CreateShortcut -gt "0")
	{
		Creation
	}
#
