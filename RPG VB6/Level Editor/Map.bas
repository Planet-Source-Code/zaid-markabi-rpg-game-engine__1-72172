Attribute VB_Name = "Map"

Global BlockTag(99, 99) As String

Global MapSizeX As Integer
Global MapSizeY As Integer

Global MapStartX As Integer
Global MapStartY As Integer

Sub Refresh_Map()
On Error Resume Next
Dim i As Integer
i = 0

Dim IX, IY As Integer

For IY = 1 To 12
For IX = 1 To 14

If BlockTag(MapStartX + IX, MapStartY + IY) = "" Then
frmMain.Block(i).Picture = frmMain.Picture
Else
frmMain.Block(i).Picture = LoadPicture(App.Path + "\Data\Gfx\Block" + BlockTag(MapStartX + IX, MapStartY + IY) + ".bmp")
End If

i = i + 1
Next
Next

End Sub
