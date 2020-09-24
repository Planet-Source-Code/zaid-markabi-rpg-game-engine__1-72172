VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form EnemyEdt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enemy Editor"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2880
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1080
      Pattern         =   "*.txt*"
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Text            =   "0"
      ToolTipText     =   "each 1000 points  +1 Lvl"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Text            =   "0"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "0"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Text            =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New Enemy"
      Height          =   255
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox Combo5 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "EnemyEdt.frx":0000
      Left            =   120
      List            =   "EnemyEdt.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2760
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Add new EMF Enemy file"
      Filter          =   "EMF only"
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Give Points :"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Give Money :"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enemy Attack :"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Shape Shape1 
      Height          =   2295
      Left            =   120
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enemy Health :"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enemies :"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "EnemyEdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo5_Change()
On Error Resume Next
Open App.Path + "\Data\Enemy\" + Combo5.Text For Input As #1
Input #1, x
Text1.Text = x
Input #1, x
Text2.Text = x
Input #1, x
Text3.Text = x
Input #1, x
Text4.Text = x
Close #1

Image1.Picture = LoadPicture(App.Path + "\Data\Enemy\" + Left(Combo5.Text, Len(Combo5.Text) - 4) + ".emf")
End Sub

Private Sub Combo5_Click()
Combo5_Change
End Sub

Private Sub Command1_Click()
Command8_Click
Unload Me
End Sub

Private Sub Command2_Click()
Dim x As String
x = InputBox("Enemy Name :", "Add", "Dog")

Open App.Path + "\Data\Enemy\" + x + ".txt" For Output As #2
Write #2, 0
Write #2, 0
Write #2, 0
Write #2, 0
Close #2

CommonDialog1.ShowOpen
FileCopy CommonDialog1.FileName, App.Path + "\Data\Enemy\" + x + ".emf"

Form_Load

Combo5.Text = x + ".txt"
End Sub

Private Sub Command8_Click()
On Error Resume Next
Open App.Path + "\Data\Enemy\" + Combo5.Text For Output As #1
Write #1, Text1.Text
Write #1, Text2.Text
Write #1, Text3.Text
Write #1, Text4.Text
Close #1
End Sub

Private Sub Form_Load()
On Error Resume Next
File1.Pattern = "*.txt*"
File1.Path = App.Path + "\Data\Enemy\"
File1.Refresh
Combo5.Clear
For i = 0 To File1.ListCount - 1
Combo5.AddItem File1.List(i)
Next
Combo5.Text = Combo5.List(0)
End Sub
