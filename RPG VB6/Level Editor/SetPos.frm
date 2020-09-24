VERSION 5.00
Begin VB.Form SetPos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3120
      Pattern         =   "*.txt*"
      TabIndex        =   7
      Top             =   7560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ok"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7560
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   0
      Max             =   1
      TabIndex        =   2
      Top             =   7200
      Width           =   8415
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7215
      Left            =   8400
      Max             =   1
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
   Begin VB.PictureBox MapBlock 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   7215
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.Image Block 
         Height          =   615
         Index           =   167
         Left            =   7800
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
         Index           =   165
         Left            =   6600
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
         Index           =   163
         Left            =   5400
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
         Index           =   161
         Left            =   4200
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
         Index           =   159
         Left            =   3000
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
         Index           =   157
         Left            =   1800
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
         Index           =   155
         Left            =   600
         Stretch         =   -1  'True
         Top             =   6600
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
         Index           =   153
         Left            =   7800
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
         Index           =   151
         Left            =   6600
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
         Index           =   149
         Left            =   5400
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
         Index           =   147
         Left            =   4200
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
         Index           =   145
         Left            =   3000
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
         Index           =   143
         Left            =   1800
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
         Index           =   141
         Left            =   600
         Stretch         =   -1  'True
         Top             =   6000
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
         Index           =   139
         Left            =   7800
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
         Index           =   137
         Left            =   6600
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
         Index           =   135
         Left            =   5400
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
         Index           =   133
         Left            =   4200
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
         Index           =   131
         Left            =   3000
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
         Index           =   129
         Left            =   1800
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
         Index           =   127
         Left            =   600
         Stretch         =   -1  'True
         Top             =   5400
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
         Index           =   125
         Left            =   7800
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
         Index           =   123
         Left            =   6600
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
         Index           =   121
         Left            =   5400
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
         Index           =   119
         Left            =   4200
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
         Index           =   117
         Left            =   3000
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
         Index           =   115
         Left            =   1800
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
         Index           =   113
         Left            =   600
         Stretch         =   -1  'True
         Top             =   4800
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
         Index           =   111
         Left            =   7800
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
         Index           =   109
         Left            =   6600
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
         Index           =   107
         Left            =   5400
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
         Index           =   105
         Left            =   4200
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
         Index           =   103
         Left            =   3000
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
         Index           =   101
         Left            =   1800
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
         Index           =   99
         Left            =   600
         Stretch         =   -1  'True
         Top             =   4200
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
         Index           =   97
         Left            =   7800
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
         Index           =   95
         Left            =   6600
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
         Index           =   93
         Left            =   5400
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
         Index           =   91
         Left            =   4200
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
         Index           =   89
         Left            =   3000
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
         Index           =   87
         Left            =   1800
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
         Index           =   85
         Left            =   600
         Stretch         =   -1  'True
         Top             =   3600
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
         Index           =   83
         Left            =   7800
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
         Index           =   81
         Left            =   6600
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
         Index           =   79
         Left            =   5400
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
         Index           =   77
         Left            =   4200
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
         Index           =   75
         Left            =   3000
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
         Index           =   73
         Left            =   1800
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
         Index           =   71
         Left            =   600
         Stretch         =   -1  'True
         Top             =   3000
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
         Index           =   69
         Left            =   7800
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
         Index           =   67
         Left            =   6600
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
         Index           =   65
         Left            =   5400
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
         Index           =   63
         Left            =   4200
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
         Index           =   61
         Left            =   3000
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
         Index           =   59
         Left            =   1800
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
         Index           =   57
         Left            =   600
         Stretch         =   -1  'True
         Top             =   2400
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
         Index           =   55
         Left            =   7800
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
         Index           =   53
         Left            =   6600
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
         Index           =   51
         Left            =   5400
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
         Index           =   49
         Left            =   4200
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
         Index           =   47
         Left            =   3000
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
         Index           =   45
         Left            =   1800
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
         Index           =   43
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1800
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
         Index           =   41
         Left            =   7800
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
         Index           =   39
         Left            =   6600
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
         Index           =   37
         Left            =   5400
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
         Index           =   35
         Left            =   4200
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
         Index           =   33
         Left            =   3000
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
         Index           =   31
         Left            =   1800
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
         Index           =   29
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1200
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
         Index           =   27
         Left            =   7800
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
         Index           =   25
         Left            =   6600
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
         Index           =   23
         Left            =   5400
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
         Index           =   21
         Left            =   4200
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
         Index           =   19
         Left            =   3000
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
         Index           =   17
         Left            =   1800
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
         Index           =   15
         Left            =   600
         Stretch         =   -1  'True
         Top             =   600
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
         Index           =   13
         Left            =   7800
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
         Index           =   11
         Left            =   6600
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
         Index           =   9
         Left            =   5400
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
         Index           =   7
         Left            =   4200
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
         Index           =   5
         Left            =   3000
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
         Index           =   3
         Left            =   1800
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
         Index           =   1
         Left            =   600
         Stretch         =   -1  'True
         Top             =   0
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
      Begin VB.Image PlayerBlock 
         Height          =   615
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POS= 1 1"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   7680
      Width           =   690
   End
End
Attribute VB_Name = "SetPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Block_Click(Index As Integer)
For i = 0 To Block.Count - 1
Block(i).BorderStyle = 0
Block(i).Appearance = 0
Next
Block(Index).BorderStyle = 1
Block(Index).Appearance = 1

i = 0
Dim IX, IY As Integer
For IY = 1 To 10
For IX = 1 To 14
If i = Index Then
PrevSelectedPosX = IX + PrevMapStartX
PrevSelectedPosY = IY + PrevMapStartY
PrevSelectedPos = "POS= " + Format(PrevSelectedPosX) + " " + Format(PrevSelectedPosY)
Label1.Caption = PrevSelectedPos
PrevSelectedPosTag = PrevBlockTag(PrevSelectedPosX, PrevSelectedPosY)
End If
i = i + 1
Next
Next
End Sub

Private Sub Combo1_Click()
PrevMapNum = Combo1.Text
Load_Prev
End Sub

Private Sub Command1_Click()
PrevSelectedPos = PrevSelectedPosOld
PrevSelectedPosX = PrevSelectedPosXOld
PrevSelectedPosY = PrevSelectedPosYOld
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Load_Prev

File1.Path = App.Path + "\Data\Maps\"
File1.Refresh
Combo1.Clear
For i = 0 To File1.ListCount - 1
If Left(File1.List(i), 3) = "Map" Then Combo1.AddItem File1.List(i)
Next
End Sub

Private Sub HScroll1_Scroll()
PrevMapStartX = HScroll1.Value
Refresh_Map_Prev
End Sub

Private Sub VScroll1_Scroll()
PrevMapStartY = VScroll1.Value
Refresh_Map_Prev
End Sub
