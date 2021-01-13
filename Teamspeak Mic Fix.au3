Run("control mmsys.cpl,,1")

WinWaitActive("Sound")
WinMove("Sound","",118, 372)

MouseClick("right",355, 494,1)
Sleep(200)
MouseClick("left",416, 650,1)
Sleep(200)

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

ShellExecute("F:\TeamSpeak3\ts3client_win64.exe")