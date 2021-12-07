; Title: TrackerBoards
; Version: 1.1
; Original Author: scram
; Modifying Author: Devin Dulude

; Description: Uses an .ini to determine Paragon install location, Paragon Module, and the Active Directory username. Reads the file, runs that paragon program, keys in username and password,
; and then if module is ORM, search for a teal colored pixel (which indicates Patient Tracking icon), move cursor to that location and click on it.

Sleep(1000)
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>

Local $TBdirectory = "C:\Windows\TB\"
Local $settingsFileName = "tb_settings.ini"
Local $ED = "emergency_department.exe"
Local $ORM = "operating_room_management.exe"

; Read the ini file for path, module, and username.
Local $PGFilePath = IniRead($TBdirectory & $settingsFileName, "PATH", "PGFilePath", "-1")
Local $PGModule = IniRead($TBdirectory & $settingsFileName, "PATH", "PGModule", "-1")
Local $ADUserName = IniRead($TBdirectory & $settingsFileName, "AD", "ADUserName", "-1")

; Verify path exists
If FileExists($PGFilePath) Then
   ; If module is ORM or ED, continue. Else, report error and quit.
   If (StringCompare($PGModule, $ORM) == 0) Or (StringCompare($PGModule, $ED) == 0) Then

	  Run($PGFilePath&$PGModule)
	  WinWaitActive("Application Logon")
	  Send($ADUserName)
	  Send("{TAB}")
	  Send("*REDACTEDFORGITHUB")
	  Send("{ENTER}")

	  ; This script has been modified to click on Patient Tracking within Operating Room Management. If the computer wants to use windows scaling, only 125% scaling is compatible.
	  ; IMPORTANT - Don't use windows preset scaling options as it interefes with autoit's ability to move the cursor accurately. You need to enter it in the advanced scaling options.
	  ; Devin Dulude 4/13/2020

	  ; if we are using ORM, have script click on Patient Tracking, Else we must be using emergency_department and therefore nothing left to do.
	  ; emergency_department automatically opens to the tracker. ORM does not.

	  If StringCompare($PGModule, $ORM) == 0 Then

		 Local $paragonWindow = "Paragon Operating"
		 ; wait until paragon is the active window
		 WinWaitActive($paragonWindow)

		 ; get coordinates of paragon window
		 Local $winCoords = WinGetPos($paragonWindow)
		 Sleep(100) ; just sleep 100ms to give the program time to make sure the window is active


		 ; If the WinGetPos doesnt error out, then grab coordinates of the pixel - else report window not found
		 If Not @error Then
			; debug information to print coordinates
			;MsgBox($MB_SYSTEMMODAL, "TrackerBoard Script", "Coordinates " & $winCoords[0] & ", " & $winCoords[1] & ", " & $winCoords[2] & ", " & $winCoords[3])

			; this searches entire desktop - dont need to scan entire desktop, but keeping it just in case
			;Local $ptCord = PixelSearch(0,0,@DeskTopWidth,@DeskTopHeight,57472)

			; this one searches paragon window specifically
			Local $ptCord = PixelSearch($winCoords[0],$winCoords[1],$winCoords[2],$winCoords[3],57472)

			; if PixelSearch doesnt error out, then move cursor to pixel location then click - else report the error
			If Not @error Then
			 ;  MsgBox ($MB_SYSTEMMODAL, "TrackerBoard Script", "Found pixel at " & $ptCord[0] & ", " & $ptCord[1])
			   MouseClick($MOUSE_CLICK_LEFT, $ptCord[0], $ptCord[1], 1, 5)
			Else
			   MsgBox ($MB_SYSTEMMODAL, "TrackerBoard Script", "Script failed to find Patient Tracking. Please click on it manually.")
			EndIf

		 Else
			MsgBox ($MB_SYSTEMMODAL, "TrackerBoard Script", "Window not found. Please click on Patient Tracking manually.")
		 EndIf

	  EndIf

	  Else
		 MsgBox ($MB_SYSTEMMODAL, "TrackerBoard Script", $settingsFileName & " configured incorrectly." & @LF & @LF & "Found: " & $PGModule & @LF & @LF & "Expected: " & $ED & " OR " & $ORM)
	  EndIf

   Else
	  MsgBox ($MB_SYSTEMMODAL, "TrackerBoard Script", $settingsFileName & " configured incorrectly." & @LF & @LF & $PGFilePath & " does not exist!")
   EndIf

Exit