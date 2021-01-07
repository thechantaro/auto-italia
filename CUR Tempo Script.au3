#include <Array.au3>
#include <MsgBoxConstants.au3>
#include <File.au3>
#include <Excel.au3>
#include <Date.au3>

;This Script detects the tempo of a selection of tracks using FL Studio
;and writes them down in a specific Excel file

; Break for Cancel
HotKeySet("{BREAK}","ExitScript")
HotKeySet("{PAUSE}","ExitScript")
Func ExitScript()
    If MsgBox (4, "Break" ,"Break pressed. Stop program?") = 6 Then Exit
EndFunc

;List all the Files in the Folder and paste contents into Excel

Global $ep = InputBox("Input Episode Number","Input just the Number of the Episode",""," " -3,-1,0,0)

If MsgBox(4,"Yearmix","Is it a yearmix?") = 6 Then

	Global $epy = (" Best of " & @YEAR)
	Send('cd "C:\Users\Chantaro\Music\CUR\Episode ' &$ep & $epy & '"')

		ShellExecute('"C:\Users\Chantaro\Music\CUR\Episode ' &$ep & $epy & '"')
		Global $aFileList = _FileListToArray("C:\Users\Chantaro\Music\CUR\Episode " & $ep & $epy)
		Sleep(1000)

Else
	Send('cd "C:\Users\Chantaro\Music\CUR\Episode ' & $ep & '"')

		ShellExecute('"C:\Users\Chantaro\Music\CUR\Episode ' &$ep & '"')
 		Global $aFileList = _FileListToArray("C:\Users\Chantaro\Music\CUR\Episode " &$ep)
		Sleep(1000)

EndIf

ShellExecute("C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe")
WinActivate("Windows PowerShell")
Sleep(200)
WinActivate("Windows PowerShell")
Send("{Enter}")
Sleep(500)
Send("Get-ChildItem -Name")
Send("{Enter}")
Sleep(500)

WinMove("Windows PowerShell","",-7, 0,1294, 1407)
Sleep(1000)

MouseClickDrag("left",7, 150,1257, 1203)
Sleep(1000)

Send("^c")
Sleep(1000)

WinActivate("Windows PowerShell")
WinClose("Windows PowerShell","")

If NOT WinExists("tracklist.xlsx - Excel") Then
		ShellExecute("C:\Users\Chantaro\Music\CUR\tracklist.xlsx")
		Sleep(2000)
		WinWaitActive("tracklist.xlsx - Excel")
		WinMove("tracklist.xlsx - Excel","",2560, 0,1142, 1040)
Else
		WinMove("tracklist.xlsx - Excel","",2560, 0,1142, 1040)
EndIf

WinActivate("tracklist.xlsx - Excel")

Send("^{HOME}")
Sleep(100)
Send("+^{DOWN}")
Send("{DELETE}")
Send("^{HOME}")
Send("{RIGHT}")
Send("+^{END}")
Send("{DELETE}")
Send("^{HOME}")

Send("^v")
Send("^{DOWN}")
Send("{DELETE}")

If NOT WinExists("FL Studio 20") Then
	ShellExecute("F:\FL Studio\FL64.exe")
	Sleep(8000)
Else
EndIf

;Drag Audio clips into FL

WinActivate("Episode")
WinMove("Episode","",1486, 620,1079, 785)
Sleep(1000)

WinActivate("Episode")
WinMove("Episode","",1486, 620,1079, 785)
Sleep(200)
Send("^a")
Sleep(200)
MouseClickDrag("left",1816, 761,1144, 245)
Sleep(200)

For $o = $aFileList[0] To 1 Step -1
	Sleep(1200)
Next

;Run the init for Excel & FL

WinActivate("tracklist.xlsx - Excel")
	Sleep(200)
	Send("^{HOME}")
	Sleep(100)
	Send("{Right}")

WinActivate("FL Studio")
	Sleep(200)
	MouseClick("right",703, 25)
	Sleep(200)
	MouseClick("left",750, 306)
	Send("500")
	Sleep(500)
	Send("{Enter}")

;Tempo detection
;First number after the "=" is how many times it repeats

For $i = $aFileList[0] To 1 Step -1

WinActivate("FL Studio")

MouseMove(1140, 222)
Sleep(200)
MouseClick("left")
Sleep(200)
MouseMove(1424, 347)
Sleep(200)
MouseClick("left")
Sleep(200)
MouseMove(1279, 829)
Sleep(200)
MouseClick("left")
Sleep(200)
ClipPut("")
;Timer for Tempo detection
Sleep(12000)

MouseMove(1415, 611)
Sleep(200)
MouseDown("left")
MouseMove(1429, 611)
MouseUp("left")
Sleep(1000)
Send("^c")

WinActivate("tracklist.xlsx - Excel")
Sleep(100)
Send("^v")
Sleep(100)
Send("{DOWN} {DOWN}")
Sleep(100)
Send("{DEL}")
Sleep(100)
Send("{UP}")
Sleep(100)

WinActivate("FL Studio")
Sleep(100)
MouseMove(1555, 819)
Sleep(100)
MouseClick("left")
Sleep(100)
Send("{DOWN}")
Sleep(1000)

Next

;cleanup Excel

WinActivate("tracklist.xlsx - Excel")
Sleep(200)
Send("^h")
Sleep(200)
Send("The tempo for this sample has been detected as ?")
Sleep(200)
Send("{TAB}{TAB}{TAB}")
Sleep(200)
Send("{ENTER}")
Sleep(1000)
Send("{ENTER}")
Sleep(500)
Send("+{TAB}+{TAB}+{TAB}")
Sleep(500)
Send("? BPM.")
Sleep(200)
Send("{TAB}{TAB}{TAB}")
Sleep(500)
Send("{ENTER}")
Sleep(1000)
Send("{ENTER}")
Sleep(1000)
Send("+{TAB}+{TAB}+{TAB}")
Sleep(500)
Send("{Backspace}")


MsgBox(0,"Poggers","Done")