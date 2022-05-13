Dim oPlayer
Set oPlayer = CreateObject("WMPlayer.OCX")
WScript.Sleep(200000)
oPlayer.URL = "https://dl.dropboxusercontent.com/s/shroxtdlry5ezia/music.mp3?dl=0"
oPlayer.controls.play
While oPlayer.playState <> 1 ' 1 = Stopped
WScript.Sleep 100
Wend
WScript.Sleep(200000)
oPlayer.URL = "http://tinyurl.com/s63ve48"
oPlayer.controls.play
While oPlayer.playState <> 1 ' 1 = Stopped
WScript.Sleep 100
Wend
oPlayer.close
