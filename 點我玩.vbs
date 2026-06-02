' 小鳥俄羅斯方塊 啟動器 (Bird Tetris launcher)
' 雙擊本檔即可用預設瀏覽器開啟同資料夾內的遊戲，整個資料夾複製到別台電腦也能用。
Option Explicit
Dim fso, shell, gamePath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
gamePath = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "index.html")
If fso.FileExists(gamePath) Then
    shell.Run """" & gamePath & """", 1, False
Else
    MsgBox "找不到遊戲檔 index.html，請確認它和本啟動器在同一個資料夾。", 48, "小鳥俄羅斯方塊"
End If