VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RPG Editor"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   14055
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   2265
      ScaleWidth      =   2625
      TabIndex        =   58
      Top             =   4440
      Width           =   2655
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enemies"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   1800
         Width           =   2295
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Wars"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sellers"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Commands Files"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   " Game's Resources "
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
         Index           =   4
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   11400
      ScaleHeight     =   1905
      ScaleWidth      =   2625
      TabIndex        =   44
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Remove"
         Height          =   255
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add"
         Height          =   255
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1560
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1440
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         DialogTitle     =   "Select Icon"
         Filter          =   "any file"
      End
      Begin VB.ComboBox Combo7 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   " Game Blocks"
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
         Index           =   2
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         Height          =   255
         Left            =   840
         TabIndex        =   48
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "---"
         Height          =   255
         Left            =   840
         TabIndex        =   47
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Block :"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   45
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1065
      ScaleWidth      =   2625
      TabIndex        =   15
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Load"
         Height          =   255
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Save"
         Height          =   255
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refresh"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   1560
         Pattern         =   "*.txt*"
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   " Rooms"
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
         TabIndex        =   53
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   11400
      ScaleHeight     =   4545
      ScaleWidth      =   2625
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   240
         ScaleHeight     =   1905
         ScaleWidth      =   2145
         TabIndex        =   62
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
         Begin VB.ComboBox Combo8 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmMain.frx":0442
            Left            =   120
            List            =   "frmMain.frx":0444
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   360
            Width           =   1935
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Edit"
            Height          =   255
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "War File :"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   65
            Top             =   120
            Width           =   675
         End
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create"
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3960
         Width           =   975
      End
      Begin VB.PictureBox Picture6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   240
         ScaleHeight     =   1905
         ScaleWidth      =   2145
         TabIndex        =   38
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton Command11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Edit"
            Height          =   255
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox Combo6 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmMain.frx":0446
            Left            =   120
            List            =   "frmMain.frx":0448
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Seller File :"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Width           =   765
         End
      End
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   240
         ScaleHeight     =   1905
         ScaleWidth      =   2145
         TabIndex        =   35
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton Command5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Edit"
            Height          =   255
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   840
            Width           =   615
         End
         Begin VB.ComboBox Combo5 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmMain.frx":044A
            Left            =   120
            List            =   "frmMain.frx":044C
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Command File :"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   36
            Top             =   120
            Width           =   1080
         End
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   240
         ScaleHeight     =   1905
         ScaleWidth      =   2145
         TabIndex        =   26
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton Command7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Edit"
            Height          =   255
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   600
            TabIndex        =   31
            Text            =   "?"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   600
            TabIndex        =   30
            Text            =   "?"
            Top             =   1440
            Width           =   735
         End
         Begin VB.ComboBox Combo4 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "? = Default"
            Height          =   195
            Left            =   1200
            TabIndex        =   34
            Top             =   840
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X ="
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   33
            Top             =   1080
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y ="
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   32
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Position :"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Goto Map Number :"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   1395
         End
      End
      Begin VB.ComboBox Combo3 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMain.frx":044E
         Left            =   240
         List            =   "frmMain.frx":046A
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1320
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMain.frx":04AE
         Left            =   120
         List            =   "frmMain.frx":04B8
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   " Add new Block"
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
         Index           =   3
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Events :"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post new Block :"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7215
      Left            =   11040
      Max             =   1
      TabIndex        =   11
      Top             =   0
      Width           =   375
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   2640
      Max             =   1
      TabIndex        =   10
      Top             =   7200
      Width           =   8415
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3345
      ScaleWidth      =   2625
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Apply"
         Height          =   255
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         Height          =   255
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         TabIndex        =   19
         Text            =   "POS= 1 1"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show gird"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         TabIndex        =   8
         Text            =   "10"
         Top             =   1680
         Width           =   735
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   600
         TabIndex        =   5
         Text            =   "12"
         Top             =   1320
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Syria - S.A.R"
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   " Map Setting"
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
         Index           =   1
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID="
         Height          =   195
         Left            =   60
         TabIndex        =   52
         Top             =   3120
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player Pos ( Default ) :"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Y ="
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X ="
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map Size :"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map Name :"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   870
      End
   End
   Begin VB.PictureBox MapBlock 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   2640
      RightToLeft     =   -1  'True
      ScaleHeight     =   7215
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.Image PlayerBlock 
         Height          =   615
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   0
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   1
         Left            =   600
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   2
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   3
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   4
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   5
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   6
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   7
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   8
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   9
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   10
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   11
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   12
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   13
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   0
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   14
         Left            =   0
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   15
         Left            =   600
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   16
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   17
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   18
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   19
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   20
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   21
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   22
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   23
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   24
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   25
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   26
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   27
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   28
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   29
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   30
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   31
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   32
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   33
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   34
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   35
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   36
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   37
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   38
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   39
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   40
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   41
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   42
         Left            =   0
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   43
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   44
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   45
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   46
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   47
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   48
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   49
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   50
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   51
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   52
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   53
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   54
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   55
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   56
         Left            =   0
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   57
         Left            =   600
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   58
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   59
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   60
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   61
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   62
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   63
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   64
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   65
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   66
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   67
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   68
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   69
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   70
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   71
         Left            =   600
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   72
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   73
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   74
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   75
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   76
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   77
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   78
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   79
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   80
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   81
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   82
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   83
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   84
         Left            =   0
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   85
         Left            =   600
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   86
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   87
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   88
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   89
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   90
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   91
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   92
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   93
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   94
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   95
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   96
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   97
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   3600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   98
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   99
         Left            =   600
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   100
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   101
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   102
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   103
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   104
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   105
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   106
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   107
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   108
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   109
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   110
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   111
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   4200
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   112
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   113
         Left            =   600
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   114
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   115
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   116
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   117
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   118
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   119
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   120
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   121
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   122
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   123
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   124
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   125
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   4800
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   126
         Left            =   0
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   127
         Left            =   600
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   128
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   129
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   130
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   131
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   132
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   133
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   134
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   135
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   136
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   137
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   138
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   139
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   5400
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   140
         Left            =   0
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   141
         Left            =   600
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   142
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   143
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   144
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   145
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   146
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   147
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   148
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   149
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   150
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   151
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   152
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   153
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   6000
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   154
         Left            =   0
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   155
         Left            =   600
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   156
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   157
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   158
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   159
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   160
         Left            =   3600
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   161
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   162
         Left            =   4800
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   163
         Left            =   5400
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   164
         Left            =   6000
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   165
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   166
         Left            =   7200
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
      Begin VB.Image Block 
         Height          =   615
         Index           =   167
         Left            =   7800
         Stretch         =   -1  'True
         Top             =   6600
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, i2 As Integer
Dim x As String

Private Sub Block_Click(Index As Integer)
On Error Resume Next
i = 0
Dim IX, IY As Integer
Dim IXx, IYy As Integer
For IY = 1 To 10
For IX = 1 To 14
If i = Index Then
IXx = IX + MapStartX
IYy = IY + MapStartY
BlockTag(IXx, IYy) = BlockPost
End If
i = i + 1
Next
Next
Refresh_Map
End Sub

Private Sub Block_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
i = 0
Dim IX, IY As Integer
Dim IXx, IYy As Integer
For IY = 1 To 10
For IX = 1 To 14
If i = Index Then
IXx = IX + MapStartX
IYy = IY + MapStartY
Label6.Caption = "ID=" + BlockTag(IXx, IYy)
End If
i = i + 1
Next
Next

End Sub

Private Sub Check1_Click()
On Error Resume Next
For i = 0 To Block.Count - 1
Block(i).BorderStyle = Check1.Value
Block(i).Appearance = 0
Next
End Sub

Private Sub Combo1_Click()
On Error GoTo 1
If Left(Combo1.Text, 3) = "Map" Then
Open App.Path + "\Data\Maps\" + Combo1.Text For Input As #1
Input #1, x
Text1.Text = x
Input #1, x
Text2.Text = x
Input #1, x
Text3.Text = x
For i2 = 0 To Int(Text3.Text) - 1
For i = 0 To Int(Text2.Text) - 1
Input #1, BlockTag(i + 1, i2 + 1)
Next
Next
Input #1, x
Text5.Text = x
GoTo 2
1:
Text5.Text = "POS= 1 1"
2:
Close #1
End If
Command1_Click
Refresh_Map
End Sub

Private Sub Combo3_Click()
On Error Resume Next
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture9.Visible = False

Select Case Combo3.ListIndex
Case Is = 3, 4
Picture4.Visible = True
Case Is = 5
Picture5.Visible = True
Case Is = 6
Picture6.Visible = True
Case Is = 7
Picture9.Visible = True

End Select
End Sub

Private Sub Combo7_Change()
On Error Resume Next
Image1.Picture = LoadPicture(App.Path + "\Data\Gfx\" + Left(Combo7.Text, Len(Combo7.Text) - 4) + ".bmp")

Open App.Path + "\Data\Gfx\" + Combo7.Text For Input As #1
Input #1, x
Label4.Caption = x
Input #1, x
Label5.Caption = x
Close #1

BlockPost = Mid(Combo7.Text, 6, 2)
End Sub

Private Sub Combo7_Click()
Combo7_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Int(Text2.Text) < 12 Then Text2.Text = "12"
If Int(Text3.Text) < 10 Then Text3.Text = "10"
MapSizeX = Int(Text2.Text)
MapSizeY = Int(Text3.Text)
HScroll1.Value = 0
VScroll1.Value = 0
HScroll1.Max = MapSizeX - 12
VScroll1.Max = MapSizeY - 10
MapStartX = 0
MapStartY = 0
End Sub

Private Sub Command10_Click()
On Error Resume Next
If MsgBox("Are you sure ?", vbYesNo + vbInformation, "Delete block") = vbYes Then
Kill App.Path + "\Data\Gfx\" + Left(Combo7.Text, Len(Combo7.Text) - 4) + ".bmp"
Kill App.Path + "\Data\Gfx\" + Combo7.Text
Command2_Click
End If
End Sub

Private Sub Command11_Click()
SellerEdt.Show
End Sub

Private Sub Command12_Click()
NewCmd.Show
End Sub

Private Sub Command13_Click()
SellerEdt.Show
End Sub

Private Sub Command14_Click()
WarEdt.Show
End Sub

Private Sub Command15_Click()
WarEdt.Show
End Sub

Private Sub Command16_Click()
EnemyEdt.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next

Command3_Click

File1.Path = App.Path + "\Data\Maps\"
File1.Refresh
Combo1.Clear
Combo4.Clear
For i = 0 To File1.ListCount - 1
If Left(File1.List(i), 3) = "Map" Then Combo1.AddItem File1.List(i)
Combo4.AddItem File1.List(i)
Next

Combo1.Text = Combo1.List(0)
Combo4.Text = Combo4.List(0)

File1.Path = App.Path + "\Data\Gfx\"
File1.Refresh
Combo5.Clear
Combo7.Clear
For i = 0 To File1.ListCount - 1
If Left(File1.List(i), 8) = "Commands" Then Combo5.AddItem File1.List(i)
If Left(File1.List(i), 5) = "Block" Then Combo7.AddItem File1.List(i)
Next

Combo5.Text = Combo5.List(0)
Combo7.Text = Combo7.List(0)

File1.Path = App.Path + "\Data\Weapons\"
File1.Refresh
Combo6.Clear
For i = 0 To File1.ListCount - 1
Combo6.AddItem File1.List(i)
Next
Combo6.Text = Combo6.List(0)

File1.Path = App.Path + "\Data\War\"
File1.Refresh
Combo8.Clear
For i = 0 To File1.ListCount - 1
Combo8.AddItem File1.List(i)
Next
Combo8.Text = Combo8.List(0)
End Sub

Private Sub Command3_Click()
On Error Resume Next
Open App.Path + "\Data\Maps\" + Combo1.Text For Output As #1
Write #1, Text1.Text
Write #1, Text2.Text
Write #1, Text3.Text
For i2 = 0 To Int(Text3.Text) - 1
For i = 0 To Int(Text2.Text) - 1
Write #1, BlockTag(i + 1, i2 + 1)
Next
Next
Write #1, Text5.Text
Close #1
End Sub

Private Sub Command4_Click()
Combo1_Click
End Sub

Private Sub Command5_Click()
On Error Resume Next
NewCmd.Show
NewCmd.Combo5.Text = Combo5.Text
End Sub


Private Sub Command6_Click()
On Error Resume Next
PrevMapNum = Combo1.Text
SetPos.Show 1
Text5.Text = PrevSelectedPos
End Sub

Private Sub Command7_Click()
PrevMapNum = Combo4.Text
SetPos.Show 1
Text6.Text = Format(PrevSelectedPosX)
Text4.Text = Format(PrevSelectedPosY)
Combo4.Text = PrevMapNum
End Sub

Private Sub Command8_Click()
On Error GoTo 2
Dim x As String
1:
x = InputBox("Block Name :", "Add new Block", "FF")
If x = "" Then Exit Sub
If Len(x) > 2 Then GoTo 1

If MsgBox("Do you like to upload new image for this block ? ( or you will select old picture )", vbYesNo + vbQuestion, "Picture") = vbYes Then
CommonDialog1.ShowOpen
FileCopy CommonDialog1.FileName, App.Path + "\Data\Gfx\Block" + x + ".bmp"
Else
frmAddBlock.Show 1
FileCopy App.Path + "\Data\Temp\Block.bmp", App.Path + "\Data\Gfx\Block" + x + ".bmp"
End If

Open App.Path + "\Data\Gfx\Block" + x + ".txt" For Output As #1
Write #1, Combo2.Text
Write #1, Combo3.Text
Select Case Combo3.ListIndex
Case Is = 3, 4
Write #1, Combo4.Text
Write #1, Text6.Text
Write #1, Text4.Text
Case Is = 5
Write #1, Combo5.Text
Case Is = 6
Write #1, Combo6.Text
End Select
Close #1

Command2_Click

Combo7.Text = "Block" + x + ".txt"
2:
Picture2.Visible = False
End Sub

Private Sub Command9_Click()
Picture2.Visible = True
End Sub

Private Sub Form_Load()
On Error Resume Next

MapSizeX = 20
MapSizeY = 14
HScroll1.Max = MapSizeX - 12
VScroll1.Max = MapSizeY - 10
MapStartX = 0
MapStartY = 0

Combo2.Text = Combo2.List(0)
Combo3.Text = Combo3.List(0)

Refresh_Map
Command2_Click
End Sub

Private Sub HScroll1_Scroll()
On Error Resume Next
MapStartX = HScroll1.Value
Refresh_Map
End Sub

Private Sub VScroll1_Scroll()
On Error Resume Next
MapStartY = VScroll1.Value
Refresh_Map
End Sub
