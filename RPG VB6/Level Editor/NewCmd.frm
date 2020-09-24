VERSION 5.00
Begin VB.Form NewCmd 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Command File"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   255
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete All"
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move Down"
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move Up"
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      Height          =   255
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New File"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "#code#"
      Top             =   1320
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "NewCmd.frx":0000
      Left            =   5040
      List            =   "NewCmd.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add"
      Height          =   255
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   840
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "NewCmd.frx":00EF
      Left            =   120
      List            =   "NewCmd.frx":00F1
      TabIndex        =   3
      Top             =   720
      Width           =   4695
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1560
      Pattern         =   "*.txt*"
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Combo5 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "NewCmd.frx":00F3
      Left            =   120
      List            =   "NewCmd.frx":00F5
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   4920
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   4920
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Commands :"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   870
   End
End
Attribute VB_Name = "NewCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo5_Change()
Combo5_Click
End Sub

Private Sub Combo5_Click()
On Error GoTo 1
Dim x As String
List1.Clear
Open App.Path + "\Data\Gfx\" + Combo5.Text For Input As #1
Do While Not x = "End Commands"
Input #1, x
List1.AddItem x
Loop
1:
Close #1
End Sub

Private Sub Command1_Click()
If List1.ListCount = 0 Then Command2_Click

Select Case Combo1.ListIndex
Case Is = 0
PrevMapNum = frmMain.Combo1.Text
SetPos.Show 1
Text1.Text = "Change " + Format(PrevSelectedPosX) + " " + Format(PrevSelectedPosY)
Case Is = 1
PrevMapNum = frmMain.Combo1.Text
SetPos.Show 1
SetPicturePrev = PrevSelectedPosTag
SetPic.Show 1
Text1.Text = "Picture " + Format(PrevSelectedPosX) + " " + Format(PrevSelectedPosY) + " " + SetPicturePrev
Case Is = 2
SetMovePrev = "Right"
SetMove.Show 1
Text1.Text = "Move " + SetMovePrev
Case Is = 3
SetPlayerPrev = "Zaid"
SetPlayer.Show 1
Text1.Text = "Player " + SetPlayerPrev
Case Is = 4
SetPlayerPrev = "Zaid"
SetPlayer.Show 1
Text1.Text = "Unlock " + SetPlayerPrev
Case Is = 5
SetPlayerPrev = "Zaid"
SetPlayer.Show 1
Text1.Text = "Lock " + SetPlayerPrev
Case Is = 6
Text1.Text = "Start"
Case Is = 7
SetPlayerPrev = "Zaid"
SetPlayer.Show 1
SetText.Show 1
Text1.Text = "Talk " + SetPlayerPrev + " " + SetTextPrev
Case Is = 8
Text1.Text = "End"
Case Is = 9
SetPlayerPrev = "Zaid"
SetPlayer.Show 1
SetMovePrev = "Right"
SetMove.Show 1
Text1.Text = "Show> " + SetPlayerPrev + " " + SetMovePrev
Case Is = 10
SetPlayerPrev = "Zaid"
SetPlayer.Show 1
SetMovePrev = "Right"
SetMove.Show 1
Text1.Text = "Move> " + SetPlayerPrev + " " + SetMovePrev

End Select

List1.AddItem Text1.Text, List1.ListCount - 1

Open App.Path + "\Data\Gfx\" + Combo5.Text For Output As #1
For i = 0 To List1.ListCount - 1
Write #1, List1.List(i)
Next
Close #1
End Sub

Private Sub Command2_Click()
Dim x As String
1:
x = InputBox("Command File :", "Add", "001")

If Not x Like ("###") Then GoTo 1

Open App.Path + "\Data\Gfx\Commands" + x + ".txt" For Output As #2
Write #2, "End Commands"
Close #2

Form_Load

Combo5.Text = "Commands" + x + ".txt"
End Sub

Private Sub Command3_Click()
On Error Resume Next
If List1.ListCount > 0 And Not List1.ListIndex = List1.ListCount - 1 Then List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command4_Click()
On Error Resume Next
If List1.ListIndex = 0 Then Exit Sub
Dim x As String
x = List1.List(List1.ListIndex - 1)
List1.List(List1.ListIndex - 1) = List1.List(List1.ListIndex)
List1.List(List1.ListIndex) = x
List1.ListIndex = List1.ListIndex - 1
End Sub

Private Sub Command5_Click()
On Error Resume Next
If List1.ListIndex = List1.ListCount - 2 Then Exit Sub
Dim x As String
x = List1.List(List1.ListIndex + 1)
List1.List(List1.ListIndex + 1) = List1.List(List1.ListIndex)
List1.List(List1.ListIndex) = x
List1.ListIndex = List1.ListIndex + 1
End Sub

Private Sub Command6_Click()
If MsgBox("Are you sure ?", vbYesNo + vbInformation, "Delete Commands") = vbYes Then
List1.Clear
List1.AddItem "End Commands"

Open App.Path + "\Data\Gfx\" + Combo5.Text For Output As #1
For i = 0 To List1.ListCount - 1
Write #1, List1.List(i)
Next
Close #1
End If
End Sub

Private Sub Command7_Click()
Command8_Click
frmMain.File1.Path = App.Path + "\Data\Gfx\"
frmMain.File1.Refresh
frmMain.Combo5.Clear
For i = 0 To frmMain.File1.ListCount - 1
frmMain.Combo5.AddItem File1.List(i)
Next
frmMain.Combo5.Text = Combo5.Text
Unload Me
End Sub

Private Sub Command8_Click()
Open App.Path + "\Data\Gfx\" + Combo5.Text For Output As #1
For i = 0 To List1.ListCount - 1
Write #1, List1.List(i)
Next
Close #1
End Sub

Private Sub Form_Load()
On Error Resume Next
File1.Path = App.Path + "\Data\Gfx\"
File1.Refresh
Combo5.Clear
For i = 0 To File1.ListCount - 1
If Left(File1.List(i), 8) = "Commands" Then Combo5.AddItem File1.List(i)
Next
Combo5.Text = Combo5.List(0)
End Sub

Private Sub List1_Click()
Dim x As String
x = InputBox("Enter new code :", "Edit", List1.List(List1.ListIndex))
If Not x = "" Then List1.List(List1.ListIndex) = x
End Sub
