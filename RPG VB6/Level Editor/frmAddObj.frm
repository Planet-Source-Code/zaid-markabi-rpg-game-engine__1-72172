VERSION 5.00
Begin VB.Form frmAddObj 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Add Object"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   217
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+"
      Height          =   255
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   255
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   255
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Left            =   240
      Pattern         =   "*.emf*"
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   1920
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   0
      Top             =   720
      Width           =   1095
      Begin VB.Image Image2 
         Height          =   1095
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00404040&
      Caption         =   " Add Object"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Objects :"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   630
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   5
      Height          =   2895
      Left            =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "frmAddObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo 1
Image2.Width = Image2.Width + 20
Image2.Height = Image2.Height + 20
Image2.Left = Image2.Left - 10
Image2.TOp = Image2.TOp - 10
1:
End Sub

Private Sub Command17_Click()
On Error GoTo 1
Image2.Width = Image2.Width - 20
Image2.Height = Image2.Height - 20
Image2.Left = Image2.Left + 10
Image2.TOp = Image2.TOp + 10
1:
End Sub

Private Sub Command2_Click()
frmAddBlock.Image1.Picture = CaptureScreen(Picture1.hwnd)
DoEvents
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub File1_Click()
On Error Resume Next
Image2.Picture = LoadPicture(App.Path + "\Data\Tools\" + File1.List(File1.ListIndex))
End Sub

Private Sub Form_Load()
File1.Path = App.Path + "\Data\Tools\"
Image1.Picture = frmAddBlock.Image1.Picture
End Sub
