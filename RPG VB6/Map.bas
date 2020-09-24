Attribute VB_Name = "Map"

Sub Load_Map(MapNumberID As Integer)
On Error GoTo Err

frmMain.MapNameLbl.Caption = ""
frmMain.IntroBlock.Visible = True

MapNum = MapNumberID
Dim Command32 As String
Dim Command32splt() As String

Open App.Path + "\Data\Maps\Map" + Format(MapNumberID, "000") + ".txt" For Input As #1
Input #1, MapName
Input #1, x
MapSizeX = Int(x)
Input #1, x
MapSizeY = Int(x)
Dim MapSzX, MapSzY As Integer
For MapSzY = 1 To MapSizeY
For MapSzX = 1 To MapSizeX
Input #1, BlockTag(MapSzX, MapSzY)

Open App.Path + "\Data\Gfx\Block" + BlockTag(MapSzX, MapSzY) + ".txt" For Input As #2
Input #2, BlockMode(MapSzX, MapSzY)
Input #2, BlockHave(MapSzX, MapSzY)
If BlockHave(MapSzX, MapSzY) = "Door" Or BlockHave(MapSzX, MapSzY) = "DoorLock" Then
Input #2, x
BlockDoorGotoMap(MapSzX, MapSzY) = Mid(x, 4, 3)
Input #2, BlockDoorGotoPosX(MapSzX, MapSzY)
Input #2, BlockDoorGotoPosY(MapSzX, MapSzY)
End If
If BlockHave(MapSzX, MapSzY) = "Commands" Then
Input #2, BlockCommandsFile(MapSzX, MapSzY)
End If
If BlockHave(MapSzX, MapSzY) = "Seller" Then
Input #2, BlockCommandsFile(MapSzX, MapSzY)
End If
If BlockHave(MapSzX, MapSzY) = "War" Then
Input #2, BlockCommandsFile(MapSzX, MapSzY)
End If
Close #2

Next
Next

Do
Input #1, Command32
Command32splt() = Split(Command32, " ")
Select Case Command32splt(0)
Case Is = "POS="
If Not PlayerPosX = -1 Then PlayerPosX = Int(Command32splt(1))
If Not PlayerPosY = -1 Then PlayerPosY = Int(Command32splt(2))
End Select
Loop

Err:
Close #1

frmMain.MapNameEffect.Enabled = True
End Sub

Sub Refresh_All()
Dim i As Integer
MapStartX = PlayerPosX - ((BlockSizeX - 2) \ 2)
MapStartY = PlayerPosY - ((BlockSizeY - 2) \ 2)
If MapStartX < 0 Then MapStartX = -1
If MapStartY < 0 Then MapStartY = -1

i = 0
Dim IX, IY As Integer
For IY = 1 To BlockSizeY
For IX = 1 To BlockSizeX

If BlockTag(MapStartX + IX, MapStartY + IY) = "" Then
frmMain.Block(i).Picture = frmMain.Picture
Else

If CheckPicSource("Map:" + Format(MapNum) + ",Source:" + Format(MapStartX + IX) + ":" + Format(MapStartY + IY)) = BlockTag(MapStartX + IX, MapStartY + IY) Or CheckPicSource("Map:" + Format(MapNum) + ",Source:" + Format(MapStartX + IX) + ":" + Format(MapStartY + IY)) = "" Then
frmMain.Block(i).Picture = LoadPicture(App.Path + "\Data\Gfx\Block" + BlockTag(MapStartX + IX, MapStartY + IY) + ".bmp")
Else
frmMain.Block(i).Picture = LoadPicture(App.Path + "\Data\Gfx\Block" + CheckPicSource("Map:" + Format(MapNum) + ",Source:" + Format(MapStartX + IX) + ":" + Format(MapStartY + IY)) + ".bmp")
End If

End If

i = i + 1
Next
Next

frmMain.MapBlock.Left = -frmMain.Block(0).Width
frmMain.MapBlock.Top = -frmMain.Block(0).Height
End Sub


Sub Refresh_Map_Pos()

Refresh_All

PlayerPosBlock = (PlayerPosX - MapStartX - 1) + BlockSizeX * (PlayerPosY - MapStartY - 1)
If MapStartX > -1 Then frmMain.PlayerBlock.Left = frmMain.Block(PlayerPosBlock).Left
If MapStartY > -1 Then frmMain.PlayerBlock.Top = frmMain.Block(PlayerPosBlock).Top

Refresh_Player

End Sub
