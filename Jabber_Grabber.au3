#include <File.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>

;
; AutoIt Version: 3.3.14.5
; Platform:       Win 10
; Author:         Joshua Smith, Specialized Support Intern
; Date:			  July 19, 2021
;
; Script Function:
;   Reads usernames from .txt file. Iterates through users and looks up various information on Jabber. Copies information to notepad. Save as a .csv file.
;

$delay = 150
$delay2 = 500

HotKeySet("{ESC}", "Terminate")

; As user if they would like to run the script.
Local $iAnswer = MsgBox(BitOR($MB_YESNO, $MB_SYSTEMMODAL), "Jabber Grabber", "Please read the associated user guide before using this script" & @LF & "To terminate Jabber Grabber during execution press the ESC key" & @LF & "Would you like to proceed?")

; If no (7), then terminate execution
If $iAnswer = 7 Then
	Exit
 EndIf

; Import .txt file that contains usernames, each separated by a line
$file = FileOpenDialog("Select a file", @ScriptDir & "\", "Text files (*.txt)")

;
If @error Then
   MsgBox(4096, "", "Error reading file")
   Exit
EndIf

; Open file with usernames separated by a line
FileOpen($file, 0)

; Get handles for Notepad and Jabber
; Used to activate window
$hWnd_notepad = WinGetHandle("[CLASS:Notepad]")

; Notepad not open
If @error Then
   Run("notepad.exe") ; opens Notepad
   Sleep($delay2 * 4)
   ; Used to activate window
   $hWnd_notepad = WinGetHandle("[CLASS:Notepad]")
EndIf

; Notepad can't open
If @error Then
   MsgBox(4096, "", "Can't open Notepad")
   Exit
EndIf

$hWnd_jabber = WinGetHandle("[TITLE:Cisco Jabber]")

; Color used for PixelSearch
$jabber_phone = 0x24AB31; Green phone icon
$jabber_text = 0x000000 ; black text - zip code

; Used to know if Jabber profile has an extra line or not
$v2 = False

; Iterate through usernames
For $i = 1 to _FileCountLines($file)
    $line = FileReadLine($file, $i)

; Open Notepad
WinActivate($hWnd_notepad)

; Opening quote in Notepad
Send('"')

; Click search box in Jabber
WinActivate($hWnd_jabber)
MouseClick("left", 282, 55, 4)

; Enter username into Jabber search box
Send($line)
Sleep($delay2 * 2)

; Hover over user pic
MouseMove(98, 127, 0)
Sleep($delay2 * 5)

; Looks to see if Jabber phone icon is present. If not, then loop is continued.
$search = PixelSearch(231,108,332,131,$jabber_phone) ; left, top, right, bottom, color
If @error Then
   WinActivate($hWnd_notepad)
   Sleep($delay2)
   Send($line)
   Send(" not found")
   Send('"')
   Send("{ENTER}")
   ContinueLoop
EndIf

; Hover - click View Profile
MouseClick("left", 572, 368, 1)
Sleep($delay2 * 3)

; Look for extra line
$search = PixelSearch(600,610,641,620,$jabber_text) ; left, top, right, bottom, color

; If extra line not present, then deviate clicks
If @error Then
   $v2 = True
EndIf

; Left-click display name
MouseClick("left", 854, 216, 1)
Sleep($delay)

CopyPasteNote()

; Closing quote
Send('"')

; Comma delimiter
Send(",")
Sleep($delay)

; Opening quote city, state
Send('"')

; Open Jabber
OpenJabber()

; Click City
If not $v2 Then
   MouseClick("left", 721, 564, 1)
Else
   MouseClick("left", 676, 529, 1)
EndIf

Sleep($delay)

CopyPasteNote()

; Comma delimiter
Send(",")

; Open Jabber
OpenJabber()

; Click State
If not $v2 Then
   MouseClick("left", 653, 586, 1)
Else
   MouseClick("left", 653, 555, 1)
EndIf

Sleep($delay2)

CopyPasteNote()

; Closing quote
Send('"')

; Comma delimiter
Send(",")

; Open Jabber
OpenJabber()

; Click Job Title
If not $v2 Then
   MouseClick("left", 806, 472, 1)
Else
   MouseClick("left", 778, 439, 1)
EndIf

Sleep($delay)

CopyPasteNote()

; Comma delimiter
Send(",")
Sleep($delay)

Send($line)

Send("{ENTER}")

$v2 = False

Next

FileClose($file)

Func CopyPasteNote()
   ; Copy
   Send("{CTRLDOWN}c{CTRLUP}")

   ; Open Notepad
   WinActivate($hWnd_notepad)
   Sleep($delay)

   ; Paste
   Send("{CTRLDOWN}v{CTRLUP}")
EndFunc

Func OpenJabber()
   Send("{ALT DOWN}")
   Send("{TAB}")
   Send("{ALT UP}")
   Sleep($delay2 * 3)
EndFunc

Func Terminate()
    Exit
EndFunc   ;==>Terminate
