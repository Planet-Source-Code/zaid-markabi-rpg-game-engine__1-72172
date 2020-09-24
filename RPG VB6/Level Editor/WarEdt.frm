VERSION 5.00
Begin VB.Form WarEdt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "War Editor"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Preview"
      Height          =   255
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   255
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create New Enemry"
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "WarEdt.frx":0000
      Left            =   2880
      List            =   "WarEdt.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add to List"
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "001"
      Top             =   840
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1080
      Pattern         =   "*.txt*"
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Combo5 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "WarEdt.frx":0004
      Left            =   120
      List            =   "WarEdt.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "WarEdt.frx":0008
      Left            =   120
      List            =   "WarEdt.frx":000A
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New File"
      Height          =   255
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   255
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enemies :"
      Height          =   195
      Index           =   0
      Left            =   2880
      TabIndex        =   11
      Top             =   840
      Width           =   690
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   2760
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enemies :"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Room Number :"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1110
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Wars :"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "WarEdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
On Error Resume Next
Image1.Picture = LoadPicture(App.Path + "\Data\Enemy\" + Left(Combo1.Text, Len(Combo1.Text) - 4) + ".emf")
End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Combo5_Change()
Dim EnemNum, i As Integer
Open App.Path + "\Data\War\" + Combo5.Text For Input As #1
Input #1, x
Text1.Text = x
Input #1, EnemNum
List1.Clear
For i = 1 To EnemNum
Input #1, x
List1.AddItem x
Next
Close #1
End Sub

Private Sub Combo5_Click()
Combo5_Change
End Sub

Private Sub Command1_Click()
List1.AddItem Left(Combo1.Text, Len(Combo1.Text) - 4)
Open App.Path + "\Data\War\" + Combo5.Text For Output As #1
Write #1, Text1.Text
Write #1, List1.ListCount
For i = 0 To List1.ListCount - 1
Write #1, List1.List(i)
Next
Close #1
End Sub

Private Sub Command2_Click()
Dim x As String
1:
x = InputBox("War File :", "Add", "001")

If Not x Like ("###") Then GoTo 1

Open App.Path + "\Data\War\War" + x + ".txt" For Output As #2
Write #2, "001"
Write #2, 0
Close #2

Form_Load

Combo5.Text = "War" + x + ".txt"
End Sub

Private Sub Command3_Click()
EnemyEdt.Show
End Sub

Private Sub Command4_Click()
On Error Resume Next
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command5_Click()
Command8_Click
frmMain.Combo8.Text = Combo5.Text
Unload Me
End Sub

Private Sub Command6_Click()
On Error Resume Next
WarPrev.Show
For i = 0 To List1.ListCount - 1
WarPrev.EnemyBlock(i).Picture = LoadPicture(App.Path + "\Data\Enemy\" + List1.List(i) + ".emf")
Next
WarPrev.Earth.Picture = LoadPicture(App.Path + "\Data\War\" + Text1.Text + ".jpg")
End Sub

Private Sub Command8_Click()
Open App.Path + "\Data\War\" + Combo5.Text For Output As #1
Write #1, Text1.Text
Write #1, List1.ListCount
For i = 0 To List1.ListCount - 1
Write #1, List1.List(i)
Next
Close #1
End Sub

Private Sub Form_Load()
File1.Path = App.Path + "\Data\War\"
File1.Refresh
Combo5.Clear
For i = 0 To File1.ListCount - 1
If Left(File1.List(i), 3) = "War" Then Combo5.AddItem File1.List(i)
Next
Combo5.Text = Combo5.List(0)

File1.Path = App.Path + "\Data\Enemy\"
File1.Refresh
Combo1.Clear
For i = 0 To File1.ListCount - 1
Combo1.AddItem File1.List(i)
Next
Combo1.Text = Combo1.List(0)
End Sub
