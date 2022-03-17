#include <File.au3>
#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIComboBox.au3>
#include <MsgBoxConstants.au3>
#include <Date.au3>
#include <WinAPISysWin.au3>

; Press the ESC key to terminate the script
HotKeySet("{ESC}", "Terminate")

; Student Elements
; aWords[student][0] - first name
; aWords[student][1] - last name
; aWords[student][2] - stu num
; aWords[student][3] - expected start
; aWords[student][4] - address
; aWords[student][5] - city
; aWords[student][6] - zip
; aWords[student][7] - email
; aWords[student][8] - phone

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;; Globals ;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

$delay = 125 ; delay used to control the speed of clicking, copy/paste, etc.

; Might need these for shipping portion (future)
$color1 = 0x000000 ; color of the text in Contact Manager (black)
$color2 = 0xFFFFFF ; color of the text in Contact Manager (white)

$rowStart = 1
$rowNumber = 1

Local $aWords[20][9]

; Ask user if they would like to run the script.
Local $iAnswer = MsgBox(BitOR($MB_YESNO, $MB_SYSTEMMODAL), "Student Info Grabber", "Please read the associated user guide before using this script" & @LF & "To terminate the script during execution press the ESC key" & @LF & "Would you like to proceed?")

; If no (7), then terminate execution
If $iAnswer = 7 Then
	Exit
 EndIf

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;; Continuing Script ;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Get the handle of the CNS window via window title
Local $hWndCNS = WinGetHandle("CampusNexus Student - Charter College - \\Remote")

; Display an error message if there was a problem in getting the window handle (make sure CNS is open)
If @error Then
        MsgBox($MB_SYSTEMMODAL, "", "An error occurred when trying to retrieve the window handle of CNS.")
        Exit
	 EndIf

start()

Func start()

   ; Bring CNS to the foreground
   WinActivate($hWndCNS)

   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;;;;; Opening Contact Manager ;;;;;
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

	;Click Daily
   ;MouseClick("left", 165, 35, 1)

   ;Click Contact Manager
   ;MouseClick("left", 210, 55, 1, 30)

   ;Click Sub Contact Manager
   ;MouseClick("left", 410, 55, 1, 30)

   ; Click State Column (sorts by state)
   ;MouseClick("left", 1025, 390, 1)

   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;;;;; Copy Student Info Into Array ;;;;;
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

   ; MsgBox($MB_SYSTEMMODAL, "", "Make sure the list is sorted by state" & @LF & "Click OK to continue")

   ; Loop Variables
   Local $iRowNum = 0
   Local $i = 0
   Local $j = 0
   Local $curRow = $rowStart

   While $iRowNum < $rowNumber
	  ; Click Start Row
	  MouseClick("left", 470, 390 + $curRow * 15, 1)

	  ; Click Student Folder
	  MouseClick("left", 15, 100, 1)

	  Sleep($delay)

	  ; Click Edit
	  MouseClick("left", 1080, 785, 1)

	  Sleep($delay * 10)

	  ; Copy First Name
	  Call("sInfo", 1050, 380, $j, $i)
	  $i = $i + 1

	  ; Copy Last Name
	  Call("sInfo", 875, 380, $j, $i)
	  $i = $i + 1

	  ; Copy Student Number
	  Call("sInfo", 1190, 360, $j, $i)
	  $i = $i + 1

	  ; Copy Expected Start
	  Call("sInfo", 1180, 605, $j, $i)
	  $i = $i + 1

	  ; Copy Address
	  Call("sInfo", 960, 440, $j, $i)
	  $i = $i + 1

	  ; Copy City
	  Call("sInfo", 875, 470, $j, $i)
	  $i = $i + 1

	  ; Copy Zip Code - double click
	  MouseClick("left", 1000, 470, 1)
	  Call("sInfo", 1000, 470, $j, $i)
	  $i = $i + 1

	  ; Copy Email - double click
	  MouseClick("left", 1000, 620, 1)
	  Call("sInfo", 1000, 620, $j, $i)
	  $i = $i + 1

	  ; Copy Phone - Click and drag
	  PixelSearch(780, 575, 850, 580, $color1)
	  If Not @error Then
		 MouseClickDrag($MOUSE_CLICK_LEFT, 870, 580, 770, 580)

	  ; Copy selected text
	  Send("^c")

	  Sleep($delay * 5)

	  $aWords[$j][$i] = ClipGet()

		 Else
		 MouseClickDrag($MOUSE_CLICK_LEFT, 870, 515, 770, 515)

	  ; Copy selected text
	  Send("^c")

	  Sleep($delay * 5)

	  $aWords[$j][$i] = ClipGet()
	  EndIf

	  ; Display the data returned by ClipGet.
		  ;MsgBox($MB_SYSTEMMODAL, "", "The following data is stored in the clipboard: " & @CRLF & $aWords[$j][4] & " " & $aWords[$j][5] & " " & $aWords[$j][6] & " " & $aWords[$j][7]  & " " & $aWords[$j][8])

	  $i = 0
	  $j = $j + 1
	  $iRowNum = $iRowNum + 1
	  $curRow = $curRow + 1

	  ; Click Cancel Button
	  MouseClick("left", 1200, 785, 1)

	  ; Click Close Button
	  MouseClick("left", 1265, 785, 1)
   WEnd

   spreadsheet()
EndFunc

Func spreadsheet()
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;;;;; Switch to Browser w/Spreadsheet ;;;;;
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

   ; Minimize CNS
	  ;MouseClick("left", 1765, 10, 1)

   Sleep($delay * 10)

   Local $aList = _WinAPI_EnumWindowsTop ( )

   For $i = 1 To $aList[0][0]
	   If $aList[$i][1] = "MozillaWindowClass" Then
		   WinActivate($aList[$i][0])
	   EndIf
   Next

   ; Get the handle of the Browser window via Class
   Local $hWndBrowser = WinGetHandle("[CLASS:MozillaWindowClass]")

   ; Display an error message if there was a problem in getting the window handle (make sure browser is open)
   If @error Then
		   MsgBox($MB_SYSTEMMODAL, "", "An error occurred when trying to retrieve the window handle of Browser")
		   Exit
		EndIf

   ; Bring Browser to the foreground
   WinActivate($hWndBrowser, "")

   Sleep($delay * 10)

   ;;; Check to see if browser is open on spreadsheet
   ; Get the handle of the Browser window via Class
   $hWndWebpage = WinGetHandle("Student Computer Tracking.xlsx — Mozilla Firefox")

   ; Display an error message if there was a problem in getting the window handle (make sure webpage is open)
   If @error Then
		   Send("^t")
		   Sleep($delay * 5)
		   ClipPut("https://prospectedu-my.sharepoint.com/:x:/r/personal/william_christensen_prospecteducation_com/_layouts/15/Doc.aspx?sourcedoc=%7BAE32C238-0CF9-4E31-A2AF-75570AF50491%7D&file=Student%20Computer%20Tracking.xlsx")
		   Sleep($delay)
		   Send("^v")
		   Sleep($delay * 5)
		   Send("{ENTER}")
		EndIf

   WinActivate($hWndWebpage, "")

   Sleep($delay * 20)

   ; Click Student Name Column
   MouseClick("left", 620, 440, 1)

   Sleep($delay * 10)

   ; Go to the bottom of the column
   Send("^{DOWN}")
   Sleep($delay * 10)
   Send("{DOWN}")

   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;;;;; Paste Student Info to Spreadsheet ;;;;;
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

   ; Loop Variables
   $iRowNum = 0
   $i = 0
   $j = 0

   While $iRowNum < $rowNumber

	  ; Enter student name
	  Send($aWords[$j][$i] & " " & $aWords[$j][$i + 1], 1)
	  $i = $i + 2

	  Send("{RIGHT}")

	  ; Enter student Number
	  Send($aWords[$j][$i], 1)
	  $i = $i + 1

	  Send("{RIGHT}")

	  ; Enter Campus
	  Send("Van")

	  Send("{RIGHT}")

	  ; Enter expected start date
	  Send($aWords[$j][$i], 1)
	  $i = $i + 1

	  Send("{RIGHT}")

	  ; Enter in the shipping date
	  Send(_NowDate(), 1)

	  ; Go to next empty row
	  Send("{DOWN}")
	  Send("{LEFT}")
	  Send("{LEFT}")
	  Send("{LEFT}")
	  Send("{LEFT}")

	  $i = 0
	  $j = $j + 1
	  $iRowNum = $iRowNum + 1
   WEnd

   shipping()
EndFunc

Func shipping()
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;;;;; Paste Student Info into Shipping Form ;;;;;
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

   ; Ask user to navigate to the UPS shipping page
   $iAnswer = MsgBox(BitOR($MB_YESNO, $MB_SYSTEMMODAL), "Student Info Grabber", "Navigate to the UPS shipping page and Log in" & @LF & "Click Yes to continue the script")

   ; If no (7), then terminate execution
   If $iAnswer = 7 Then
	   Exit
	EndIf


   ; Loop Variables
   $iRowNum = 0
   $i = 0
   $j = 0

   ; Iterating through each student
   While $iRowNum < $rowNumber

	  ; Click Shipping
	  MouseClick("left", 465, 165, 1)

	  ; Click Create a Shipment
	  MouseClick("left", 285, 311, 1)

	  Sleep($delay * 20)

	  ; Click Enter New Address
	  MouseClick("left", 295, 595, 1)

	  ; Click Company or Name
	  MouseClick("left", 370, 695, 1)

	  ; Enter student name
	  Send($aWords[$j][0] & " " & $aWords[$j][1], 1)

	  Call("sendTabs", 3)

	  ; Enter Address
	  Send($aWords[$j][4], 1)
	  Call("sendTabs", 3)

	  ; Enter City
	  Send($aWords[$j][5], 1)
	  Call("sendTabs", 2)

	  ; Enter Zip
	  Send($aWords[$j][6], 1)
	  Send("{TAB}")

	  ; Enter phone
	  Send($aWords[$j][8], 1)
	  Call("sendTabs", 2)

	  ; Enter Email
	  Send($aWords[$j][7], 1)
	  Call("sendTabs", 2)
	  Send("{SPACE}")

	  ; Navigate to Weight
	  Call("sendTabs", 10)

	  ; Weight
	  Send("3")
	  Call("sendTabs", 1)

	  ; Dimensions
	  Send("16")
	  Send("{TAB}")

	  Send("9")
	  Send("{TAB}")

	  Send("3")
	  Call("sendTabs", 3)

	  ; Value
	  Send("250")

	  ; Navigate to Other Reference
	  Call("sendTabs", 16)

	  ; Value
	  Send("Student Tablet")

	  $j = $j + 1
	  $iRowNum = $iRowNum + 1

	  ; Ask user to navigate to the UPS shipping page
   $iAnswer = MsgBox(BitOR($MB_YESNO, $MB_SYSTEMMODAL), "Student Info Grabber", "Ensure to put in the state and verify info" & @LF & "Continue the script after the current student is finished" & @LF & "Click Yes to continue the script")

   ; If no (7), then terminate execution
   If $iAnswer = 7 Then
	   Exit
	EndIf

   WEnd

EndFunc

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;; Functions ;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

; Click TAB $num of times
Func sendTabs($num)
   Local $itr = 0;
   While $itr < $num
	  Send("{TAB}")
	  $itr = $itr + 1
	  Sleep($delay)
   WEnd
EndFunc

; Function that copies elements in student master and copies
; elements from clipboard to an array
Func sInfo($x, $y, $student, $sElement)
   ; Click Expected Start
   MouseClick("left", $x, $y, 1)

   ; Copy selected text
   Send("^c")

   Sleep($delay * 5)

   $aWords[$student][$sElement] = ClipGet()
EndFunc

; Function to terminate script with ESC being the hotkey
Func Terminate()
   $hGUI = GUICreate("Menu") ; will create a dialog box that when displayed is centered

   GUISetBkColor(0x00E0FFFF)
   GUISetFont(9, 300)

   GUICtrlCreateTab(10, 10, 480, 480)
		; Create tabitems

		Local $test = "blah"

        GUICtrlCreateTabItem("Control")
        Local $id1 = GUICtrlCreateButton("Beginning", 10, 40, 120, 20)
        Local $id2 = GUICtrlCreateButton("Spreadsheet", 10, 70, 120, 20)
		Local $id3 = GUICtrlCreateButton("Shipping", 10, 100, 120, 20)

		GUICtrlCreateTabItem("Student Info")
		;Local $idComboBox = GUICtrlCreateCombo("", 20, 50, 160, 150)
        ;GUICtrlSetData($idComboBox, $aWords[0][0]&"|"&$aWords[1][0]&"|"&$aWords[2][0]&"|"&$aWords[3][0]&"|"&$aWords[4][0]&"|"&$aWords[5][0]&"|"&$aWords[6][0]&"|"&$aWords[7][0]&"|"&$aWords[8][0]&"|"&$aWords[9][0]&"|"&$aWords[10][0]&"|"&$aWords[11][0]&"|"&$aWords[12][0]&"|"&$aWords[13][0]&"|"&$aWords[14][0]&"|"&$aWords[15][0]&"|"&$aWords[16][0]&"|"&$aWords[17][0]&"|"&$aWords[18][0]&"|"&$aWords[19][0], $aWords[0][0]); default Jon

		; Cuurent GenericCombo text
		 $sCurrGenericData = ""

		 ; Create GenericCombo
		 $hGenericCombo = GUICtrlCreateCombo("", 10, 40, 100, 20)
		 GUICtrlSetData(-1, $aWords[0][0]&"|Option B")
		 _GUICtrlComboBox_SetEditText($hGenericCombo, "Select")  ; Sets the text, but does not add to the list for selection

		 ; Create subordinate combos and hide them
		 $hGuiBoxForA1 = GUICtrlCreateCombo("",  10, 100, 100, 20)
		 GUICtrlSetData(-1, "Blah A 1.1|Blah A 1.2")
		 GUICtrlSetState(-1, $GUI_HIDE)
		 $hGuiBoxForA2 = GUICtrlCreateCombo("", 260, 100, 100, 20)
		 GUICtrlSetData(-1, "Blah A 2.1|Blah A 2.2")
		 GUICtrlSetState(-1, $GUI_HIDE)

		 $hGuiBoxForB1 = GUICtrlCreateCombo("",  10, 100, 100, 20)
		 GUICtrlSetData(-1, "Blah B 1.1|Blah B 1.2")
		 GUICtrlSetState(-1, $GUI_HIDE)
		 $hGuiBoxForB2 = GUICtrlCreateCombo("", 260, 100, 100, 20)
		 GUICtrlSetData(-1, "Blah B 2.1|Blah B 2.2")
		 GUICtrlSetState(-1, $GUI_HIDE)

		 ; Close Tab definiton
		 GUICtrlCreateTabItem("")

        GUICtrlSetState($id1, $GUI_CHECKED)

        GUISetState(@SW_SHOW)

        Local $idMsg
        ; Loop until the user exits.
        While 1
                $idMsg = GUIGetMsg()
                Select
                        Case $idMsg = $GUI_EVENT_CLOSE
                                ExitLoop
                        Case $idMsg = $id1
                                start()
								ExitLoop
                        Case $idMsg = $id2
                                spreadsheet()
								ExitLoop
						Case $idMsg = $id3
                                shipping()
								ExitLoop
							 ;Case $idMsg = $idComboBox
								;GUICtrlCreateLabel("label1", 30, 80, 50, 20)
                                ;MsgBox($MB_SYSTEMMODAL, "", "The combobox is currently displaying: " &$idComboBox)
							 EndSelect

			   ; Read GenericCombo selection
			 $sGenericData = GUICtrlRead($hGenericCombo)
			 ; If it has changed AND the combo is closed - if we do not check for closure, the code actions while we scan the dropdown list
			 If $sGenericData <> $sCurrGenericData And _GUICtrlComboBox_GetDroppedState($hGenericCombo) = False Then
				 ; Reset the current text so prevent flicker from constant redrawing
				 $sCurrGenericData = $sGenericData
				 ; Hide/Show the combos depending on the selection
				 Switch $sGenericData
					 Case $aWords[0][0]
						 GUICtrlSetState($hGuiBoxForB1, $GUI_HIDE)
						 GUICtrlSetState($hGuiBoxForB2, $GUI_HIDE)
						 GUICtrlSetState($hGuiBoxForA1, $GUI_SHOW)
						 GUICtrlSetState($hGuiBoxForA2, $GUI_SHOW)
					 Case "Option B"
						 GUICtrlSetState($hGuiBoxForA1, $GUI_HIDE)
						 GUICtrlSetState($hGuiBoxForA2, $GUI_HIDE)
						 GUICtrlSetState($hGuiBoxForB1, $GUI_SHOW)
						 GUICtrlSetState($hGuiBoxForB2, $GUI_SHOW)
				 EndSwitch
			 EndIf
        WEnd
    Exit
EndFunc   ;==>Terminate
