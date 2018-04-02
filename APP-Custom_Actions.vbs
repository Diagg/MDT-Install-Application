'//############################
'//  This script by Diagg/OSD-Couture.com
'//   http://www.osd-couture.com/
'// 
'//  History : 
'//
'//  V1.0 	23/06/2015 - Initial release  
'//  V1.1	19/12/2015 - Added logging for each function
'// V2.0	30/06/2016 -  Added Function  CheckConditions
'//
'//
'//############################


Function CheckConditions

	oLogging.CreateEntry "Begin Processing Pre-Check conditions ! ", LogTypeInfo
	
	'If oEnvironment.item("model") = "HP Spectre Pro x360 G2" or oEnvironment.item("model") = "HP EliteBook 840 G3" or oEnvironment.item("model") = "HP ProBook 430 G2" Then
	'	oEnvironment.Item("AppConditions") = "True"
	'	oLogging.CreateEntry "Device is Compliant, Pre-Check validation is successfull ! ", LogTypeInfo
	'Else	
	'	oEnvironment.Item("AppConditions") = "False"
	'	oLogging.CreateEntry "Device is not Compliant, Pre-Check validation has failed ! ", LogTypeInfo
	'End If
		
	oLogging.CreateEntry "Finished Processing Pre-Check conditions ! ", LogTypeInfo


End Function


Function StandAloneCustomAction

	oLogging.CreateEntry "begin Processing Stand alone custom Actions !", LogTypeInfo
	
	'oLogging.CreateEntry "Disabling Remote registry service ! ", LogTypeInfo
	'Manage_Services "RemoteRegistry", "Disabled", "Stopped"

	oLogging.CreateEntry "Finished Processing Custom Actions ! ", LogTypeInfo
	
End Function





Function ActionsBeforeInstall

	oLogging.CreateEntry "begin Processing Custom Actions before launching Application ! ", LogTypeInfo

	'Dim sSource, sDest, bOverwrite, sFile
	
	'sFile = "C:\Folder1\test.txt"	
	'sSource = "C:\Folder1"
	'sDest = "C:\Folder2\"
	'bOverwrite = True

	'Copy File
	'oFileHandling.CopyFile sFile, sDest, bOverwrite
	
	'Move File
	'oFileHandling.MoveFile sFile, sDest

	'Copy Folder
	'oFileHandling.CopyFolder sSource, sDest, bOverwrite
	
	'Move Folder
	'oFileHandling.MoveFolder sSource, sDest


	'Delete Files
	'oFileHandling.DeleteFile sTargetFile "C:\temp\myfile.txt"

	'Delete Folders
	'oFileHandling.RemoveFolder "C:\temp"

	'Add Reg
	'oShell.RegWrite "HKCU\KeyName\","", "REG_SZ" 'create default value
	'oShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\Hidden",1,"REG_DWORD"

	'Delete Reg
	'oShell.RegDelete ("HKCU\Control Panel\Desktop\MyKey\MySubKey\")

	'Create Shortcut
	'Dim strDesktop, olnk
	
	'strDesktop = WshShell.SpecialFolders("Desktop")
	'Set olnk = oCreateShortcut(strDesktop & "\Shortcut Script.lnk")
   
	'olnk.TargetPath = "C:\Program Files\MyApp\MyProgram.EXE"
	'olnk.Arguments = ""
	'olnk.Description = "MyProgram"
	'olnk.HotKey = "ALT+CTRL+F"
	'olnk.IconLocation = "C:\Program Files\MyApp\MyProgram.EXE, 2"
	'olnk.WindowStyle = "1"
	'olnk.WorkingDirectory = "C:\Program Files\MyApp"
	'olnk.Save

	'Create Folder
	'If Not oFSO.FolderExists(sSource) Then
	'	oFSO.CreateFolder sSource
	'End If

	'Rename File
	'If  oFSO.FileExists(sFile) Then	
	'	oFSO.GetFile(sFile).Name = "Hello.txt"
	'End If

	'Rename Folder
	'If  oFSO.FolderExists(sSource) Then
	'	oFSO.GetFolder(sSource).Name = "YourFolder"
	'End If
	
	oLogging.CreateEntry "Finished Processing Custom Actions before launching Application ! ", LogTypeInfo
	
End Function


Function ActionsAfterSuccessfullInstall
		
	oLogging.CreateEntry "Begin Processing Custom Actions After successful install ! ", LogTypeInfo

	'Copying  Active Setup  file
	'oFileHandling.CopyFile oEnvironment.Item("AppFolderPath") & "PinMDT.vbs", oEnv("AllUsersProfile") &"\Active-Setup\", 1
	
	
	'oLogging.CreateEntry "Disabling Firewall service", LogTypeInfo
	'Disable_Services (aTargetSvcs)
	
	
	oLogging.CreateEntry "Finished Processing Custom Actions After successful install ! ", LogTypeInfo

End Function


Function ActionsAfterFailedfullInstall

	oLogging.CreateEntry "Begin Processing Custom Actions After failed install ! ", LogTypeInfo
	
	oLogging.CreateEntry "Finished Processing Custom Actions After failed install ! ", LogTypeInfo

End Function

