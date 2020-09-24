VERSION 5.00
Begin VB.Form SetPic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set Picture"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ok"
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   120
      Pattern         =   "*.bmp*"
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   2040
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "SetPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub File1_Click()
On Error Resume Next
Image1.Picture = LoadPicture(App.Path + "\Data\Gfx\" + File1.List(File1.ListIndex))
SetPicturePrev = Mid(File1.List(File1.ListIndex), 6, 2)
End Sub

Private Sub Form_Load()
File1.Path = App.Path + "\Data\Gfx\"
File1.Refresh

For i = 0 To File1.ListCount - 1
If File1.List(i) = "Block" + SetPicturePrev + ".bmp" Then File1.ListIndex = i
Next
End Sub
