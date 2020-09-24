Attribute VB_Name = "SaveGameChanges"

Global CheckPointID(999) As String
Global CheckPointValue(999) As String
Global CheckPointNum As Integer

Global CheckSourceID(999) As String
Global CheckSource(999) As String
Global CheckSourceNum As Integer


Sub AddNewChekhPoint(ID As String, Value As String)

Dim i As Integer
For i = 0 To CheckPointNum - 1
If CheckPointID(i) = ID Then
CheckPointValue(i) = Value
Exit Sub
End If
Next

CheckPointID(CheckPointNum) = ID
CheckPointValue(CheckPointNum) = Value
CheckPointNum = CheckPointNum + 1
End Sub

Function CheckPoint(ID As String) As String
Dim i As Integer
For i = 0 To CheckPointNum - 1
If CheckPointID(i) = ID Then
CheckPoint = CheckPointValue(i)
Exit Function
End If
Next
CheckPoint = ""
End Function

Sub AddNewPicSource(ID As String, Source As String)

Dim i As Integer
For i = 0 To CheckSourceNum - 1
If CheckSourceID(i) = ID Then
CheckSource(i) = Source
Exit Sub
End If
Next

CheckSourceID(CheckSourceNum) = ID
CheckSource(CheckSourceNum) = Source
CheckSourceNum = CheckSourceNum + 1
End Sub

Function CheckPicSource(ID As String) As String
Dim i As Integer
For i = 0 To CheckSourceNum - 1
If CheckSourceID(i) = ID Then
CheckPicSource = CheckSource(i)
Exit Function
End If
Next
CheckPicSource = ""
End Function
