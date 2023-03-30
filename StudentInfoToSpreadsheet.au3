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
; aWords[student][10] - program

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;;; Globals ;;;;;
;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

$delay = 100 ; delay used to control the speed of clicking, copy/paste, etc.

Local $bNumkey = False

; Might need these for shipping portion (future)
$color1 = 0x000000 ; color of the editable text in Student Master (black)
$color2 = 0xFFFFFF ; color of the background text in Student Master (white)
$color3 = 0x6D6D6D ; color of the greyed out text in Student Master (grey)
$color4 = 0xBFFFCF ; color of the light green text in Student Master (green)
$colorBlue = 0x409EFE ; color of the light blue text in Contact Method (Blue)

$rowStart = 4
$rowNumber = 1

; $aWords[NumStudents][NumFields]
Local $aWords[20][11]

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
	  ; Click Contact Manager
	  MouseClick("left", 58, 76, 1)
	  Sleep($delay * 5)

	  ; Click Start Row
	  MouseClick("left", 470, 390 + $curRow * 15, 1)
	  Sleep($delay * 5)

	  ; Click Student Folder
	  MouseClick("left", 15, 100, 1)
	  Sleep($delay * 5)

	  Sleep($delay)

	  ; Wait for Student Master
	  PixelSearch(951, 637, 985, 646, $color4)
	  While @error
		 Sleep(20)
		 PixelSearch(951, 637, 985, 646, $color4)
	  WEnd

	  ; Click Edit
	  MouseClick("left", 1080, 785, 1)
	  Sleep($delay * 5)

	  ;Sleep($delay * 10)

	  ; Wait for Student Master
	  PixelSearch(777, 597, 900, 604, $color1)
	  While @error
		 Sleep(20)
		 PixelSearch(777, 597, 900, 604, $color1)
	  WEnd

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
	  $i = $i + 1

	  ; Display the data returned by ClipGet.
		  ;MsgBox($MB_SYSTEMMODAL, "", "The following data is stored in the clipboard: " & @CRLF & $aWords[$j][4] & " " & $aWords[$j][5] & " " & $aWords[$j][6] & " " & $aWords[$j][7]  & " " & $aWords[$j][8])


	  ; Click Cancel Button
	  MouseClick("left", 1200, 785, 1)

	  ; Click Close Button
	  MouseClick("left", 1265, 785, 1)

	  Sleep($delay * 5)

	  ; Click Contact Method
	  MouseClick("left", 15, 285, 1)

	  Sleep($delay * 5)

	  ; Wait for Contact Method to load - Window
	  ;PixelSearch(1147, 778, 1193, 786, $color1)
	  ;Sleep( $delay )
	  ;While @error
		 ;Sleep(20)
		; PixelSearch(1147, 778, 1193, 786, $color1)
	  ;WEnd

	  ; Wait for Contact Method to load - Window
	  PixelSearch(840, 284, 993, 298, $colorBlue)
	  ;Sleep( $delay )
	  While @error
		 Sleep(20)
		 PixelSearch(840, 284, 993, 298, $colorBlue)
	  WEnd

	  ; Wait for Contact Method to load - Cell
	  PixelSearch(811, 501, 892, 604, $color1)
	  ;Sleep( $delay )
	  While @error
		 Sleep(20)
		 PixelSearch(811, 501, 892, 604, $color1)
	  WEnd

	  ; Copy State - double click
	  MouseClick("left", 995, 395, 1)
	  Call("sInfo", 995, 395, $j, $i)
	  $i = $i + 1

	  Sleep($delay * 5)

	  ; Click Close
	  MouseClick("left", 1170, 780, 1)

	  Sleep($delay * 5)

	  ; Click Academic Records
	  MouseClick("left", 62, 992, 1)
	  Sleep($delay * 5)

	  ; Click Enrollment
	  MouseClick("left", 18, 192, 1)
	  Sleep($delay * 5)

	  ; Wait for Enrollment to enable - Window Loads
	  PixelSearch(1050, 825, 1060, 835, $color1)
	  While Not @error
		 Sleep(20)
		 PixelSearch(1050, 825, 1060, 835, $color1)
	  WEnd

	  ; Wait for Enrollment to enable - Cell Loads
	  PixelSearch(795, 513, 843, 519, $color3)
	  While Not @error
		 Sleep(20)
		 PixelSearch(1050, 825, 1060, 835, $color3)
	  WEnd

	  ; Click Enrollment Information
	  MouseClick("left", 865, 482, 1)
	  Sleep($delay * 5)

	  ; Wait for Enrollment Info tab to enable - Academic Advisor bottom left
	  PixelSearch(912, 704, 979, 712, $color4)
	  While @error
		 Sleep(20)
		 PixelSearch(912, 704, 979, 712, $color4)
	  WEnd

	  ; Click Edit
	  MouseClick("left", 943, 449, 1)

	  Sleep($delay * 10)

	  PixelSearch(804, 535, 847, 545, $color3)
	  While Not @error
		 Sleep(20)
		 PixelSearch(804, 535, 847, 545, $color3)
	  WEnd

	  ; Double Click Program and save
	  MouseClick("left", 896, 542, 1)
	  Call("sInfo", 896, 542, $j, $i)
	  Sleep($delay * 5)

	  ; Click Cancel
	  MouseClick("left", 1188, 834, 1)
	  Sleep($delay * 5)

	  ; Click Close
	  MouseClick("left", 1261, 835, 1)
	  Sleep($delay * 5)

	  ; Click Edit Activity
	  ;MouseClick("left", 520, 850, 1)

	  ;Sleep($delay * 10)

	  ;PixelSearch(780, 710, 840, 711, $color1)

	 ; If Not @error Then
		 ; Copy Program - double click
		; MouseClick("left", 835, 710, 1)
		; Call("sInfo", 835, 710, $j, $i)

		; Else

		 ; Copy Program - double click
		; MouseClick("left", 835, 725, 1)
		; Call("sInfo", 835, 725, $j, $i)

	;  EndIf

	  ;Sleep($delay * 5)

	  ; Click Close
	  ;MouseClick("left", 1170, 870, 1)

	  ;Sleep($delay * 5)

	  ; Click No
	 ; MouseClick("left", 1020, 620, 1)

	  ;Sleep($delay * 5)

	  $i = 0
	  $j = $j + 1
	  $iRowNum = $iRowNum + 1
	  $curRow = $curRow + 1
   WEnd

   spreadsheet()
EndFunc

Func spreadsheet()
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;;;;; Switch to Browser w/Spreadsheet ;;;;;
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

   ; Minimize CNS
	  ;MouseClick("left", 1765, 10, 1)

   Sleep($delay * 5)

   ;Local $aList = _WinAPI_EnumWindowsTop ( )

   ;For $i = 1 To $aList[0][0]
	  ;If $aList[$i][1] = "MozillaWindowClass" Then
	 ; ; If $aList[$i][1] = "Chrome_WidgetWin_1" Then
	;	   WinActivate($aList[$i][0])
	;   EndIf
  ; Next

   ; Get the handle of the Browser window via Class
   ; Local $hWndBrowser = WinGetHandle("[CLASS:MozillaWindowClass]")
   Local $hWndBrowser = WinGetHandle("[CLASS:XLMAIN]")
   ;Local $hWndBrowser = WinGetHandle("[CLASS:Chrome_WidgetWin_1]")

   ; Display an error message if there was a problem in getting the window handle (make sure browser is open)
   If @error Then
		   MsgBox($MB_SYSTEMMODAL, "", "An error occurred when trying to retrieve the window handle of Browser")
		   Exit
		EndIf

   ; Bring Browser to the foreground
   WinActivate($hWndBrowser, "")

   Sleep($delay * 5)

   ;;; Check to see if browser is open on spreadsheet
   ; Get the handle of the Browser window via Class
   ;
   ;$hWndWebpage = WinGetHandle("Student Computer Tracking.xlsx — Mozilla Firefox")
   ;
   ;$hWndWebpage = WinGetHandle("2023 Student Computer Tracking.xlsx — Mozilla Firefox")
   ;
   $hWndWebpage = WinGetHandle("2023 Student Computer Tracking.xlsx - Excel")
   ;
   ;$hWndWebpage = WinGetHandle("Student Computer Tracking.xlsx — Google Chrome")

   ; Display an error message if there was a problem in getting the window handle (make sure webpage is open)
   If @error Then
		   ;Send("^t")
		   ;Sleep($delay * 5)
		   ;ClipPut("https://prospectedu-my.sharepoint.com/:x:/r/personal/william_christensen_prospecteducation_com/_layouts/15/Doc.aspx?sourcedoc=%7BAE32C238-0CF9-4E31-A2AF-75570AF50491%7D&file=Student%20Computer%20Tracking.xlsx")
		   ;ClipPut("https://prospectedu-my.sharepoint.com/:x:/r/personal/william_christensen_prospecteducation_com/_layouts/15/doc2.aspx?sourcedoc=%7B45f2c78e-f62f-4d26-8d97-6b4615afb586%7D")
		   ;Sleep($delay)
		   ;Send("^v")
		   ;Sleep($delay * 5)
		   ;Send("{ENTER}")
		   MsgBox($MB_SYSTEMMODAL, "", "Please open the 2023 Student Computer Tracking spreadsheet before continuing")
		EndIf

   WinActivate($hWndWebpage, "")

   Sleep($delay * 10)

   Local $count = 0

   While $count < 15
	  Send("^{PGUP}")
	  $count = $count + 1
	  WEnd

   Send("^{HOME}")
   Send("^{UP}")
   Send("^{UP}")

   ; Click Student Name Column
   ;MouseClick("left", 570, 330, 1)

   Sleep($delay * 5)

   ; Go to the bottom of the column
   Send("^{DOWN}")
   Sleep($delay * 10)
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
	  ;Send("Van")

	  ;Send("{RIGHT}")

	  ; Enter expected start date
	  ;Send($aWords[$j][$i], 1)
	  $i = $i + 1

	  ;Send("{RIGHT}")

	  ; Enter Program
	  Send($aWords[$j][$i + 6], 1)
	  If (($aWords[$j][$i + 6] == "CMOAA") Or ($aWords[$j][$i + 6] == "AASMBC") Or ($aWords[$j][$i + 6] == "AASMBCDC")) Then
		 $bNumkey = True
	  EndIf

	  Send("{RIGHT}")
	  Send("{RIGHT}")

	  ; Enter in the shipping date
	  Send(_NowDate(), 1)

	  Send("{RIGHT}")

	  ; Enter Notes
	  Send("Shipped from PSC")


	  ; Go to next empty row
	  Send("{DOWN}")
	  Send("{LEFT}")
	  Send("{LEFT}")
	  Send("{LEFT}")
	  Send("{LEFT}")
	  Send("{LEFT}")

	  $i = 0
	  $j = $j + 1
	  $iRowNum = $iRowNum + 1
   WEnd

   If $bNumkey Then
	  MsgBox($MB_SYSTEMMODAL, "", "ALERT! At least one program needs a numkey pad!")
   EndIf

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
	 ; MouseClick("left", 465, 165, 1)

	  ; Click Create a Shipment
	  ;MouseClick("left", 285, 311, 1)

	  ; Skip CAD Students
	  While $aWords[$j][$i + 6] == "CAD"
		 $j = $j + 1
		 $iRowNum = $iRowNum + 1
		 If $iRowNum < $rowNumber Then
			Exit
		 EndIf
	  WEnd

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
	  Call("sendTabs", 1)

	  ; Enter State
	  Call("state", $aWords[$j][9])
	  Call("sendTabs", 1)

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
	  Call("sendTabs", 2)

	  ; Dimensions
	  If (($aWords[$j][10] == "CMOAA") Or ($aWords[$j][10] == "AASMBC") Or ($aWords[$j][10] == "AASMBCDC")) Then
		 Send("18")
		 Send("{TAB}")

		 Send("12")
		 Send("{TAB}")

		 Send("4")
		 Call("sendTabs", 3)
	  Else
		 Send("18")
		 Send("{TAB}")

		 Send("12")
		 Send("{TAB}")

		 Send("4")
		 Call("sendTabs", 3)

	 EndIf


	  ; Value
	  Send("250")

	  ; Navigate to Other Reference
	  Call("sendTabs", 14)

	  ; Value
	  Send("Student Tablet")

	  ; Schedule UPS Pickup
	  Call("sendTabs", 4)
	  Send("{SPACE}")

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

; Click TAB $num of times
Func sendDown($num)
   Local $itr = 0;
   While $itr < $num
	  Send("{DOWN}")
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
        Local $id1 = GUICtrlCreateButton("Beginning", 10, 40, 100, 20)
        Local $id2 = GUICtrlCreateButton("Spreadsheet", 110, 40, 100, 20)
		Local $id3 = GUICtrlCreateButton("Shipping", 210, 40, 100, 20)

		GUICtrlCreateTabItem("Student Info")
		;Local $idComboBox = GUICtrlCreateCombo("", 20, 50, 160, 150)
        ;GUICtrlSetData($idComboBox, $aWords[0][0]&"|"&$aWords[1][0]&"|"&$aWords[2][0]&"|"&$aWords[3][0]&"|"&$aWords[4][0]&"|"&$aWords[5][0]&"|"&$aWords[6][0]&"|"&$aWords[7][0]&"|"&$aWords[8][0]&"|"&$aWords[9][0]&"|"&$aWords[10][0]&"|"&$aWords[11][0]&"|"&$aWords[12][0]&"|"&$aWords[13][0]&"|"&$aWords[14][0]&"|"&$aWords[15][0]&"|"&$aWords[16][0]&"|"&$aWords[17][0]&"|"&$aWords[18][0]&"|"&$aWords[19][0], $aWords[0][0]); default Jon

		; Cuurent GenericCombo text
		 $sCurrGenericData = ""

		 ; Create GenericCombo
		 $hGenericCombo = GUICtrlCreateCombo("", 10, 40, 100, 20)
		 GUICtrlSetData(-1, $aWords[0][0]&"|"&$aWords[1][0]&"|"&$aWords[2][0]&"|"&$aWords[3][0]&"|"&$aWords[4][0]&"|"&$aWords[5][0]&"|"&$aWords[6][0]&"|"&$aWords[7][0]&"|"&$aWords[8][0]&"|"&$aWords[9][0]&"|"&$aWords[10][0]&"|"&$aWords[11][0]&"|"&$aWords[12][0]&"|"&$aWords[13][0]&"|"&$aWords[14][0]&"|"&$aWords[15][0]&"|"&$aWords[16][0]&"|"&$aWords[17][0]&"|"&$aWords[18][0]&"|"&$aWords[19][0])
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
					Local $it = 0
					Local $ySpace = 100
					While $it < 9
					   GUICtrlCreateLabel($aWords[0][$it], 10, $ySpace, 200, 50)
					   $it = $it + 1
					   $ySpace = $ySpace + 20
					 WEnd
						 ;GUICtrlSetState($hGuiBoxForB1, $GUI_HIDE)
						 ;GUICtrlSetState($hGuiBoxForB2, $GUI_HIDE)
						 ;GUICtrlSetState($hGuiBoxForA1, $GUI_SHOW)
						 ;GUICtrlSetState($hGuiBoxForA2, $GUI_SHOW)
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

 Func state($state)
	Switch $state
	  Case "AL"
		 Send("A")
	  Case "AK"
		 Send("A")
		 Call("sendDown", 1)
	  Case "AZ"
		 Send("A")
		 Call("sendDown", 2)
	  Case "AR"
		 Send("A")
		 Call("sendDown", 3)

	  Case "CA"
		 Send("C")
	  Case "CO"
		 Send("C")
		 Call("sendDown", 1)
	  Case "CT"
		 Send("C")
		 Call("sendDown", 2)

	  Case "DE"
		 Send("D")
	  Case "DC"
		 Send("D")
		 Call("sendDown", 1)

	  Case "GA"
		 Send("G")

	  Case "HI"
		 Send("H")

	  Case "ID"
		 Send("I")
	  Case "IL"
		 Send("I")
		 Call("sendDown", 1)
	  Case "IN"
		 Send("I")
		 Call("sendDown", 2)
	  Case "IA"
		 Send("I")
		 Call("sendDown", 3)

	  Case "KS"
		 Send("K")
	  Case "KY"
		 Send("K")
		 Call("sendDown", 1)

	  Case "LA"
		 Send("L")

	  Case "ME"
		 Send("M")
	  Case "MD"
		 Send("M")
		 Call("sendDown", 1)
	  Case "MA"
		 Send("M")
		 Call("sendDown", 2)
	  Case "MI"
		 Send("M")
		 Call("sendDown", 3)
	  Case "MN"
		 Send("M")
		 Call("sendDown", 4)
	  Case "MS"
		 Send("M")
		 Call("sendDown", 5)
	  Case "MO"
		 Send("M")
		 Call("sendDown", 6)
	  Case "MT"
		 Send("M")
		 Call("sendDown", 7)

	  Case "NE"
		 Send("N")
	  Case "NV"
		 Send("N")
		 Call("sendDown", 1)
	  Case "NH"
		 Send("N")
		 Call("sendDown", 2)
	  Case "NJ"
		 Send("N")
		 Call("sendDown", 3)
	  Case "NM"
		 Send("N")
		 Call("sendDown", 4)
	  Case "NY"
		 Send("N")
		 Call("sendDown", 5)
	  Case "NC"
		 Send("N")
		 Call("sendDown", 6)
	  Case "ND"
		 Send("N")
		 Call("sendDown", 7)

	  Case "OH"
		 Send("O")
	  Case "OK"
		 Send("O")
		 Call("sendDown", 1)
	  Case "OR"
		 Send("O")
		 Call("sendDown", 2)

	  Case "PA"
		 Send("P")

	  Case "RI"
		 Send("R")

	  Case "SC"
		 Send("S")
	  Case "SD"
		 Send("S")
		 Call("sendDown", 1)

	  Case "TN"
		 Send("T")
	  Case "TX"
		 Send("T")
		 Call("sendDown", 1)

	  Case "UT"
		 Send("U")

	  Case "VT"
		 Send("V")
	  Case "VA"
		 Send("V")
		 Call("sendDown", 1)

	  Case "WA"
		 Send("W")
	  Case "WV"
		 Send("W")
		 Call("sendDown", 1)
	  Case "WI"
		 Send("W")
		 Call("sendDown", 2)
	  Case "WY"
		 Send("W")
		 Call("sendDown", 3)
	  EndSwitch

 EndFunc
