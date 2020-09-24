VERSION 5.00
Begin VB.Form frmAddBlock 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add new Block"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Object"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   240
      Pattern         =   "*.bmp*"
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Block Image :"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmAddBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command17_Click()
frmAddObj.Show 1
End Sub

Private Sub Command2_Click()
On Error Resume Next
SavePicture Image1.Picture, App.Path + "\Data\Temp\Block.bmp"
Unload Me
End Sub

Private Sub File1_Click()
On Error Resume Next
Image1.Picture = LoadPicture(File1.Path + "\" + File1.List(File1.ListIndex))
End Sub

Private Sub Form_Load()
File1.Path = App.Path + "\Data\Gfx\"
End Sub
