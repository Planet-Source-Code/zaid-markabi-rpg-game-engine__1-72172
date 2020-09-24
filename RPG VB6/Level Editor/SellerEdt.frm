VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form SellerEdt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sellers"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add New Magic"
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   1320
      Pattern         =   "*.txt*"
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ComboBox Text4 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SellerEdt.frx":0000
      Left            =   1440
      List            =   "SellerEdt.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1440
      TabIndex        =   11
      Text            =   "0"
      Top             =   2880
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
   Begin VB.ComboBox Combo5 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "SellerEdt.frx":0004
      Left            =   120
      List            =   "SellerEdt.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New File"
      Height          =   255
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Add new EMF Magic file"
      Filter          =   "EMF only"
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   2415
      Left            =   120
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magic Vaule :"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magic Name :"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magic Price :"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Energy Price :"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Potion Price :"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sellers :"
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "SellerEdt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo5_Change()
On Error Resume Next
Open App.Path + "\Data\Weapons\" + Combo5.Text For Input As #1
Input #1, x
Text1.Text = x
Input #1, x
Text2.Text = x
Input #1, x
Text3.Text = x
Input #1, x
Text4.Text = x + ".emf"
Input #1, x
Text5.Text = x
Close #1
End Sub

Private Sub Combo5_Click()
Combo5_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
Open App.Path + "\Data\Weapons\" + Combo5.Text For Output As #1
Write #1, Text1.Text
Write #1, Text2.Text
Write #1, Text3.Text
Write #1, Left(Text4.Text, Len(Text4.Text) - 4)
Write #1, Text5.Text
Close #1
End Sub

Private Sub Command2_Click()
Dim x As String
1:
x = InputBox("Seller File :", "Add", "001")

If Not x Like ("###") Then GoTo 1

Open App.Path + "\Data\Weapons\Seller" + x + ".txt" For Output As #2
Write #2, "0"
Write #2, "0"
Write #2, "0"
Write #2, "Magic.emf"
Write #2, "0"
Close #2

Form_Load

Combo5.Text = "Seller" + x + ".txt"
End Sub

Private Sub Command3_Click()
On Error GoTo 1
CommonDialog1.ShowOpen
FileCopy CommonDialog1.FileName, App.Path + "\Data\Weapons\" + CommonDialog1.FileTitle
On Error Resume Next
File1.Pattern = "*.emf*"
File1.Path = App.Path + "\Data\Weapons\"
File1.Refresh
Text4.Clear
For i = 0 To File1.ListCount - 1
Text4.AddItem File1.List(i)
Next
Text4.Text = CommonDialog1.FileTitle
1:
End Sub

Private Sub Command4_Click()
Command1_Click
frmMain.File1.Path = App.Path + "\Data\Weapons\"
frmMain.File1.Refresh
frmMain.Combo6.Clear
For i = 0 To File1.ListCount - 1
frmMain.Combo6.AddItem File1.List(i)
Next
frmMain.Combo6.Text = Combo5.Text
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
File1.Pattern = "*.emf*"
File1.Path = App.Path + "\Data\Weapons\"
File1.Refresh
Text4.Clear
For i = 0 To File1.ListCount - 1
Text4.AddItem File1.List(i)
Next
Text4.Text = Text4.List(0)

File1.Pattern = "*.txt*"
File1.Path = App.Path + "\Data\Weapons\"
File1.Refresh
Combo5.Clear
For i = 0 To File1.ListCount - 1
If Left(File1.List(i), 6) = "Seller" Then Combo5.AddItem File1.List(i)
Next
Combo5.Text = Combo5.List(0)
End Sub

Private Sub Text4_Change()
On Error Resume Next
Image1.Picture = LoadPicture(App.Path + "\Data\Weapons\" + Text4.Text)
End Sub

Private Sub Text4_Click()
Text4_Change
End Sub
