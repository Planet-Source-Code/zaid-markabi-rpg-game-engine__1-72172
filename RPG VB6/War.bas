Attribute VB_Name = "War"

Global EnemyNum As Integer
Global EnemyName(9) As String
Global EnemyHelath(9) As Integer
Global EnemyAttack(9) As Integer
Global EnemyMoney(9) As Integer
Global EnemyPoints(9) As Integer

Global DamageSelected As Integer

Sub Create_War(WarNum As Integer)
Dim i As Integer
Dim RoomPicture As String
Dim TempXX() As String

Open App.Path + "\Data\War\War" + Format(WarNum, "000") + ".txt" For Input As #4
Input #4, RoomPicture
frmMain.Earth.Picture = LoadPicture(App.Path + "\Data\War\War" + RoomPicture + ".jpg")
Input #4, EnemyNum

For i = 0 To EnemyNum - 1
Input #4, EnemyName(i)
Open App.Path + "\Data\Enemy\" + EnemyName(i) + ".txt" For Input As #5
Input #5, EnemyHelath(i)
Input #5, EnemyAttack(i)
Input #5, EnemyMoney(i)
Input #5, EnemyPoints(i)
Close #5

frmMain.EnemyBlock(i).Picture = LoadPicture(App.Path + "\Data\Enemy\" + EnemyName(i) + ".emf")

Next

Close #4

GameMenuWarPos = 0
GameMenuWarPosEnm = 0
frmMain.MenuPosWar.Top = frmMain.MenuWar(GameMenuWarPos).Top
frmMain.MenuPosWar2.Top = frmMain.EnemyBlock(GameMenuWarPosEnm).Top
frmMain.MenuPosWar2.Left = frmMain.EnemyBlock(GameMenuWarPosEnm).Left

For i = 0 To frmMain.PlayerFacePct.Count - 1
TempXX() = Split(PlayerName, " ")
If Left(PlayerListName(i), Len(TempXX(0))) = TempXX(0) Then
PlayerListSelected = i
Exit For
End If
Next
End Sub
