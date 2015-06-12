Option Explicit

' DECLARE VARIABLES
Dim objWMIService, objOperatingSystem, objShell
Dim colOperatingSystems
Dim strOSVersion, strOSCaption,strKMSKey, strComputerName

'SET CONSTANTS
CONST strKMSServer = "IP_OR_DNS"
CONST strKMSPort = "PORT"

' GET OS VERSION
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colOperatingSystems = objWMIService.ExecQuery("Select version from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystems
	strOSVersion = left(objOperatingSystem.version,3)
Next

' CHECK OS VERSION AND QUIT IF XP OR EARLIER SINCE NO SLMGR.VBS ON PRE VISTA MACHINES
if strOSVersion < 6 Then
	wscript.echo "A KMS Key does not exist for this version of Windows."
	wscript.quit
End If

' GET COMPUTER NAME FOR UNIQUE IDENTIFIER
Set objShell = wscript.createobject("WScript.Shell")
strComputerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")

' CHECK IF WINDOWS IS ACTIVATED
If isWindowsActivated() = True Then
	wscript.echo "Windows is Already Activated"
	wscript.quit
Else
	wscript.echo "Not Activated - Activating Now"
End If

' GET OS CAPTION INFORMATION
Set colOperatingSystems = objWMIService.ExecQuery("Select caption from Win32_OperatingSystem")

For Each objOperatingSystem in colOperatingSystems
	strOSCaption = objOperatingSystem.caption
Next

Set colOperatingSystems = Nothing

' SELECT APPROPRIATE KMS KEY USING CAPTION INFORMATION
Select Case 1
	Case instr(strOSCaption,"Microsoft Windows 7 Professional"): strKMSKey = "FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4"
	Case instr(strOSCaption,"Microsoft Windows 8 Pro"): strKMSKey = "NG4HW-VH26C-733KW-K6F98-J8CK4"
	Case instr(strOSCaption,"Microsoft Windows 8.1 Pro"): strKMSKey = "GCRJD-8NW9H-F2CDX-CCM8D-9D6T9"
	Case instr(strOSCaption,"Microsoft Windows Server 2008 R2 Standard"): strKMSKey = "YC6KT-GKW9T-YTKYR-T4X34-R7VHC"
	Case instr(strOSCaption,"Microsoft Windows Server 2012 Standard"): strKMSKey = "XC9B7-NBPP2-83J2H-RHMBY-92BT4"
	Case instr(strOSCaption,"Microsoft Windows Server 2012 R2 Standard"): strKMSKey = "D2N9P-3P6X9-2R39C-7RTCD-MDVJX"
	Case Else
		wscript.echo "Could not Identify Operating System."
		wscript.quit
End Select

' RUN SLMGR.VBS TO ASSIGN KMS KEY, KMS SERVER, AND ACTIVATE
objShell.Run "wscript //B C:\windows\system32\slmgr.vbs /ipk " & strKMSKey,1,True
objShell.Run "wscript //B C:\windows\system32\slmgr.vbs /skms " & strKMSServer & ":" & strKMSPort,1,True
objShell.Run "wscript //B C:\windows\system32\slmgr.vbs /ato",1,True

' RUN FINAL CHECK TO MAKE SURE WINDOWS ACTIVATED SUCCESSFULLY
If isWindowsActivated() = True Then
	wscript.echo "Successfully Activated Windows"
Else
	wscript.echo "Windows Failed to Activate"
End If

Set objShell = Nothing
Set objWMIService = Nothing

Function isWindowsActivated()
	' DECLARE VARIABLES
	Dim colProducts, strQuery, objProduct
	
	' SET THE DEFAULT TO BE NOT ACTIVATED
	isWindowsActivated = False

	' CHECK FOR LICENCE INFORMATION USING WMI QUERY
	strQuery = "SELECT LicenseStatus FROM SoftwareLicensingProduct WHERE PartialProductKey <> null"
	Set colProducts = objWMIService.ExecQuery(strQuery)
	For Each objProduct in colProducts
		If objProduct.LicenseStatus = 1 Then
			isWindowsActivated = True
		End If 
	Next
End Function