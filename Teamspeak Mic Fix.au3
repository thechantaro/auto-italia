; Break for Cancel
HotKeySet("{BREAK}","ExitScript")
HotKeySet("{PAUSE}","ExitScript")
Func ExitScript()
    If MsgBox (4, "Break" ,"Break pressed. Stop program?") = 6 Then Exit
EndFunc

CallCP()

MouseClick("right",355, 494,1)
Sleep(200)
Send("{UP}")
Send("{Enter}")

;Set100()

;ShellExecute("F:\TeamSpeak3\ts3client_win64.exe")


;----Functions----;
Func CallCP()
;Executes the controlpanel for recording devices
Run("control mmsys.cpl,,1")

WinWaitActive("Sound")
WinMove("Sound","",118, 372)

EndFunc

Func Set100()
;Sets Mic Volume to 100
WinWaitActive("Microphone Properties")
Send("+{TAB}")
Sleep(200)
Send("{RIGHT 2}")
Sleep(200)
Send("{TAB}")
Sleep(200)
Send("{RIGHT 20}")
Sleep(200)
Send("{ENTER 2}")

EndFunc