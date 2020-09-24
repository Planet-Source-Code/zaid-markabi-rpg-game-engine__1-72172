Attribute VB_Name = "SetPosPreview"

Global PrevMapNum As String

Global PrevBlockTag(99, 99) As String

Global PrevMapSizeX As Integer
Global PrevMapSizeY As Integer

Global PrevMapStartX As Integer
Global PrevMapStartY As Integer

Global PrevSelectedPos As String
Global PrevSelectedPosX As Integer
Global PrevSelectedPosY As Integer

Global PrevSelectedPosTag As String

Global PrevSelectedPosOld As String
Global PrevSelectedPosXOld As Integer
Global PrevSelectedPosYOld As Integer

Sub Refresh_Map_Prev()
Dim i As Integer
i = 0

Dim IX, IY As Integer

For IY = 1 To 12
For IX = 1 To 14

If PrevBlockTag(PrevMapStartX + IX, PrevMapStartY + IY) = "" Then
SetPos.Block(i).Picture = SetPos.Picture
Else
SetPos.Block(i).Picture = LoadPicture(App.Path + "\Data\Gfx\Block" + PrevBlockTag(PrevMapStartX + IX, PrevMapStartY + IY) + ".bmp")
End If

i = i + 1
Next
Next

End Sub

Sub Load_Prev()
On Error GoTo 1
Dim x As String
Dim i, i2 As Integer

Open App.Path + "\Data\Maps\" + PrevMapNum For Input As #1
Input #1, x
SetPos.Caption = "Room : " + x
Input #1, x
PrevMapSizeX = Int(x)
Input #1, x
PrevMapSizeY = Int(x)
For i2 = 0 To PrevMapSizeY - 1
For i = 0 To PrevMapSizeX - 1
Input #1, PrevBlockTag(i + 1, i2 + 1)
Next
Next
Input #1, x
PrevSelectedPos = x
GoTo 2
1:
PrevSelectedPos = "POS= 1 1"
2:
Close #1

SetPos.Label1.Caption = PrevSelectedPos
PrevSelectedPosOld = PrevSelectedPos

Dim XX() As String
XX() = Split(PrevSelectedPos, " ")
PrevSelectedPosX = Int(XX(1))
PrevSelectedPosY = Int(XX(2))
SetPos.Block(PrevSelectedPosX - 1 + (14 * (PrevSelectedPosY - 1))).BorderStyle = 1
SetPos.Block(PrevSelectedPosX - 1 + (14 * (PrevSelectedPosY - 1))).Appearance = 1
PrevSelectedPosXOld = PrevSelectedPosX
PrevSelectedPosYOld = PrevSelectedPosY
SetPos.HScroll1.Value = 0
SetPos.VScroll1.Value = 0
SetPos.HScroll1.Max = PrevMapSizeX - 12
SetPos.VScroll1.Max = PrevMapSizeY - 10
PrevMapStartX = 0
PrevMapStartY = 0

Refresh_Map_Prev

End Sub
