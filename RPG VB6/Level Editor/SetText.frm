VERSION 5.00
Begin VB.Form SetText 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Text"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text here"
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OK"
      Height          =   255
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label TextLn2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5535
   End
End
Attribute VB_Name = "SetText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SetTextPrev = Text1.Text
Unload Me
End Sub

Private Sub Text1_Change()
If InStr(1, Text1.Text, Chr(34)) > 0 Or InStr(1, Text1.Text, ",") > 0 Then
MsgBox "You can't write [ " + Chr(34) + "  <or>  ,  ] !", vbExclamation, "Error"
End If
TextLn2.Caption = Text1.Text
End Sub

Private Sub Text1_Click()
Text1_Change
End Sub
