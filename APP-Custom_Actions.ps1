#==============================================================
# APP-Custom_Action
#==============================================================
#
# by Diagg/OSD-Couture.com 20/09/2016
#
# Version 1.0
# Release Date : 20/09/2016
# last modified :
#
# Changes :	20/09/2016 	v1.0	Initial release
#
#
# Application command line : 
#
#==============================================================
#
#==============================================================


#Requires -Version 4
#Requires -RunAsAdministrator 


##== Debug
$ErrorActionPreference = "stop"
#$ErrorActionPreference = "Continue"


##== Check MDT Module
If (!((Get-Module).name -eq "ZTIUtility")) { Import-Module "ZTIUtility" -ErrorAction SilentlyContinue}	
If (!((Get-Module).name -eq "ZTIUtility")) { Write-Host "ERROR : unable to import MDT Powershell Module !!!" ; [Environment]::Exit(99)} #unabel to load MDT env, exit with return code 99

##== Init Variables
$Global:CurrentScriptName = $MyInvocation.MyCommand.Name
$Global:CurrentScriptFullName = $MyInvocation.MyCommand.Path
$Global:CurrentScriptPath = split-path $MyInvocation.MyCommand.Path
$Global:MDTShare = $TSEnv:Deployroot



#== Init Aditionnal Function
."$Global:MDTShare\Scripts\DiaggFunctions.ps1"
Init-Function
If ($Global:tsenv -eq $false) { [Environment]::Exit(99)} #unabel to load MDT env, exit with return code 99
If ($args.count -eq 0) { [Environment]::Exit(98)} #no argument sent to App-Custom_Action, exit with return code 98
Log-ScriptEvent -value ("Received Arguments : " + $args[0])  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False

#== Checking Conditions before launching
IF ($args[0]="CheckConditions")
	{
		Log-ScriptEvent -value "Entering Check Condition Section !"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False
		#[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
		#[System.Windows.Forms.MessageBox]::Show("Press ok to continue. . .")
		#$TSEnv:AppConditions = "True"
		Log-ScriptEvent -value "Check Conditions Section completed, conditions evaluated to $TSEnv:AppConditions"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False
	}


#== Calling Stand Alone custom actions
IF ($args[0]="StandAloneCustomAction")
	{
		Log-ScriptEvent -value "Entering Stand Alone custom actions Section !"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False
	
		Log-ScriptEvent -value "Stand Alone custom actions Section completed !"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False	
	}



#== Calling custom actions before installation
IF ($args[0]="ActionsBeforeInstall")
	{
		Log-ScriptEvent -value "Entering Actions Before Install Section !"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False
		#If ($TSEnv:ISSERVEROS -eq "True")
		#	{
		#		Log-ScriptEvent -value "Server OS detected, installing LAPS with All sub modules"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False
		#		$TSEnv:AppParameters = ($TSEnv:AppParameters + " ADDLOCAL=CSE,Management.UI,Management.PS,Management.ADMX") 
		#	}
		#Else
		#	{
		#		Log-ScriptEvent -value "Client OS detected, installing LAPS client module only !"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False
		#	}
		Log-ScriptEvent -value "Actions Before Install Section completed !"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False
	}
	
#== Calling custom actions after successful installation
IF ($args[0]="ActionsAfterSuccessfullInstall")
	{
		Log-ScriptEvent -value "Entering Actions After Successful Install Section !"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False
	
		Log-ScriptEvent -value "Actions After Successful Install Section completed !"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False	
	}
	
#== Calling custom actions after failed installation
IF ($args[0]="ActionsAfterFailedfullInstall")
	{
		Log-ScriptEvent -value "Entering Actions After Failed Install Section !"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False
		
		Log-ScriptEvent -value "Actions After Failed Install Section completed !"  -Component $Global:CurrentScriptName -Severity  1 -OutToConsole $False	
	
	}	