VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RPG"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   13755
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   13755
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox ItemBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   6000
      ScaleHeight     =   6015
      ScaleWidth      =   7215
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   4320
         Top             =   1800
      End
      Begin VB.Label MenuItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check Points NO. = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   480
         TabIndex        =   71
         Top             =   3360
         Width           =   2040
      End
      Begin VB.Label MenuItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Map Name = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   480
         TabIndex        =   70
         Top             =   2880
         Width           =   1365
      End
      Begin VB.Label MenuItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player Name = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   480
         TabIndex        =   69
         Top             =   2400
         Width           =   1590
      End
      Begin VB.Label MenuItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Points = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   480
         TabIndex        =   68
         Top             =   1920
         Width           =   885
      End
      Begin VB.Label MenuItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Money = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   480
         TabIndex        =   67
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label MenuItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Energy NO. = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   480
         TabIndex        =   66
         Top             =   960
         Width           =   1425
      End
      Begin VB.Label MenuItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Potion NO. = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   480
         TabIndex        =   65
         Top             =   480
         Width           =   1350
      End
   End
   Begin VB.PictureBox SellerBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   9120
      ScaleHeight     =   6015
      ScaleWidth      =   7215
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Timer SellerUpdate 
         Interval        =   100
         Left            =   6600
         Top             =   4800
      End
      Begin VB.Label MenuLblBuy5 
         BackStyle       =   0  'Transparent
         Caption         =   "Energy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   62
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label MenuLblBuy5 
         BackStyle       =   0  'Transparent
         Caption         =   "Potion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   63
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Image BuyFinger5 
         Height          =   555
         Left            =   240
         Picture         =   "frmMain.frx":13AF6
         Top             =   3000
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image MenuBuy5 
         Height          =   375
         Index           =   0
         Left            =   240
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Image MenuBuy5 
         Height          =   375
         Index           =   1
         Left            =   240
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label MenuLblBuy4 
         BackStyle       =   0  'Transparent
         Caption         =   "Buy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4200
         TabIndex        =   61
         Top             =   3840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Image BuyFinger4 
         Height          =   555
         Left            =   3720
         Picture         =   "frmMain.frx":14396
         Top             =   3840
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image MenuBuy4 
         Height          =   375
         Left            =   3720
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "  Magic"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   60
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label PrcMagic 
         BackStyle       =   0  'Transparent
         Caption         =   "Price : 0 $"
         Height          =   255
         Left            =   3600
         TabIndex        =   59
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Label AtkMagic 
         BackStyle       =   0  'Transparent
         Caption         =   "Attack : 0"
         Height          =   255
         Left            =   3600
         TabIndex        =   58
         Top             =   4320
         Width           =   2295
      End
      Begin VB.Image MagicEffect 
         Height          =   2295
         Left            =   3600
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Shape Shape2 
         Height          =   3050
         Left            =   3480
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label MoneyPlyrHav 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "000 $"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4800
         TabIndex        =   57
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Image ImgOnly 
         Height          =   375
         Index           =   0
         Left            =   4800
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Image BuyFinger2 
         Height          =   555
         Left            =   240
         Picture         =   "frmMain.frx":14C36
         Top             =   840
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label MenuLblBuy2 
         BackStyle       =   0  'Transparent
         Caption         =   "Magic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   54
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label MenuLblBuy2 
         BackStyle       =   0  'Transparent
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   53
         Top             =   840
         Width           =   1695
      End
      Begin VB.Image BuyFinger3 
         Height          =   555
         Left            =   240
         Picture         =   "frmMain.frx":154D6
         Top             =   1800
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label MenuLblBuy3 
         BackStyle       =   0  'Transparent
         Caption         =   "Energy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   56
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label MenuLblBuy3 
         BackStyle       =   0  'Transparent
         Caption         =   "Potion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   55
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Image MenuBuy3 
         Height          =   375
         Index           =   1
         Left            =   240
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Image MenuBuy3 
         Height          =   375
         Index           =   0
         Left            =   240
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Image MenuBuy2 
         Height          =   375
         Index           =   1
         Left            =   2520
         Top             =   840
         Width           =   2175
      End
      Begin VB.Image MenuBuy2 
         Height          =   375
         Index           =   0
         Left            =   240
         Top             =   840
         Width           =   2175
      End
      Begin VB.Image BuyFinger 
         Height          =   555
         Left            =   240
         Picture         =   "frmMain.frx":15D76
         Top             =   240
         Width           =   450
      End
      Begin VB.Label MenuLblBuy 
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   5280
         TabIndex        =   52
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label MenuLblBuy 
         BackStyle       =   0  'Transparent
         Caption         =   "Sell"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   3000
         TabIndex        =   51
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label MenuLblBuy 
         BackStyle       =   0  'Transparent
         Caption         =   "Buy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   50
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image MenuBuy 
         Height          =   375
         Index           =   2
         Left            =   4800
         Top             =   240
         Width           =   2175
      End
      Begin VB.Image MenuBuy 
         Height          =   375
         Index           =   1
         Left            =   2520
         Top             =   240
         Width           =   2175
      End
      Begin VB.Image MenuBuy 
         Height          =   375
         Index           =   0
         Left            =   240
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.PictureBox WarBox 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   6360
      ScaleHeight     =   6015
      ScaleWidth      =   7215
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Timer WarUpdate 
         Interval        =   100
         Left            =   4800
         Top             =   5400
      End
      Begin VB.Timer RemoveEffect 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   360
         Top             =   600
      End
      Begin VB.Timer EnemyWar 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1440
         Top             =   240
      End
      Begin VB.Image EnemyBlock 
         Height          =   1095
         Index           =   3
         Left            =   3840
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Image EnemyBlock 
         Height          =   1095
         Index           =   9
         Left            =   3360
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Image EnemyBlock 
         Height          =   1095
         Index           =   8
         Left            =   2400
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Image EnemyBlock 
         Height          =   1095
         Index           =   7
         Left            =   2280
         Top             =   240
         Width           =   1095
      End
      Begin VB.Image EnemyBlock 
         Height          =   1095
         Index           =   6
         Left            =   3720
         Top             =   720
         Width           =   1095
      End
      Begin VB.Image EnemyBlock 
         Height          =   1095
         Index           =   5
         Left            =   5040
         Top             =   960
         Width           =   1095
      End
      Begin VB.Image EnemyBlock 
         Height          =   1095
         Index           =   4
         Left            =   5400
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Image EnemyBlock 
         Height          =   1095
         Index           =   2
         Left            =   2040
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Image EnemyBlock 
         Height          =   1095
         Index           =   1
         Left            =   360
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label PlyrNam 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Player Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   48
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label PlyrEnrg 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " MP : 000 \ 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   47
         Top             =   5640
         Width           =   1695
      End
      Begin VB.Label PlyrHlth 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   " HP : 000 \ 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5400
         TabIndex        =   46
         Top             =   5280
         Width           =   1695
      End
      Begin VB.Label ItemLblWar 
         BackStyle       =   0  'Transparent
         Caption         =   "Energy"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   45
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label ItemLblWar 
         BackStyle       =   0  'Transparent
         Caption         =   "Potion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   44
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Image MenuItemWar 
         Height          =   555
         Left            =   2880
         Picture         =   "frmMain.frx":16616
         Top             =   4560
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image MenuItemsWar 
         Height          =   375
         Index           =   1
         Left            =   2880
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Image MenuItemsWar 
         Height          =   375
         Index           =   0
         Left            =   2880
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Image WeaponEffect 
         Height          =   495
         Left            =   600
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image MenuPosWar2 
         Height          =   555
         Left            =   0
         Picture         =   "frmMain.frx":16EB6
         Top             =   0
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label MenuLblWar 
         BackStyle       =   0  'Transparent
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   43
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label MenuLblWar 
         BackStyle       =   0  'Transparent
         Caption         =   "Magic"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   42
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label MenuLblWar 
         BackStyle       =   0  'Transparent
         Caption         =   "Attack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   41
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Image MenuPosWar 
         Height          =   555
         Left            =   240
         Picture         =   "frmMain.frx":17756
         Top             =   4560
         Width           =   450
      End
      Begin VB.Image MenuWar 
         Height          =   375
         Index           =   2
         Left            =   240
         Top             =   5520
         Width           =   2175
      End
      Begin VB.Image MenuWar 
         Height          =   375
         Index           =   1
         Left            =   240
         Top             =   5040
         Width           =   2175
      End
      Begin VB.Image MenuWar 
         Height          =   375
         Index           =   0
         Left            =   240
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Image EnemyBlock 
         Height          =   1095
         Index           =   0
         Left            =   1080
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Earth 
         Height          =   4455
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7215
      End
      Begin VB.Image BackImgWar 
         Height          =   1575
         Left            =   0
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   7215
      End
   End
   Begin VB.PictureBox StatusBox 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   8760
      ScaleHeight     =   6015
      ScaleWidth      =   7335
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   7335
      Begin VB.Label PlayerNameLblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "PlayerName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   5760
         Width           =   3135
      End
      Begin VB.Image MenuPos2 
         Height          =   555
         Left            =   0
         Picture         =   "frmMain.frx":17FF6
         Top             =   960
         Width           =   450
      End
      Begin VB.Label PowerLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl : 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   38
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label EnergyLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "MP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   37
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label HealthLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "HP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5160
         TabIndex        =   36
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label PowerLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl : 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   5160
         TabIndex        =   35
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label EnergyLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "MP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   5160
         TabIndex        =   34
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label HealthLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "HP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   5160
         TabIndex        =   33
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label PowerLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl : 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   32
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label EnergyLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "MP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   31
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label HealthLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "HP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   30
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label PowerLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl : 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   29
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label EnergyLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "MP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   28
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label HealthLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "HP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5160
         TabIndex        =   27
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label PowerLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl : 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   26
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label EnergyLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "MP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   25
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label HealthLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "HP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   24
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label PowerLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl : 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   23
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label EnergyLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "MP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   22
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label HealthLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "HP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   21
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label PowerLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl : 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   20
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label EnergyLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "MP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   19
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label HealthLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "HP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   18
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label PowerLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lvl : 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   17
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label EnergyLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "MP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   16
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label HealthLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "HP : 000 / 000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         Height          =   1275
         Index           =   7
         Left            =   5040
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         Height          =   1275
         Index           =   6
         Left            =   5040
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         Height          =   1275
         Index           =   5
         Left            =   5040
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         Height          =   1275
         Index           =   4
         Left            =   5040
         Top             =   120
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         Height          =   1275
         Index           =   3
         Left            =   1440
         Top             =   4440
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         Height          =   1275
         Index           =   2
         Left            =   1440
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Shape Shape1 
         Height          =   1275
         Index           =   1
         Left            =   1440
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Image PlayerFacePct 
         Height          =   1275
         Index           =   3
         Left            =   120
         Picture         =   "frmMain.frx":18896
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   1275
      End
      Begin VB.Image PlayerFacePct 
         Height          =   1275
         Index           =   2
         Left            =   120
         Picture         =   "frmMain.frx":1EAF4
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1275
      End
      Begin VB.Image PlayerFacePct 
         Height          =   1275
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":24D52
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Image PlayerFacePct 
         Height          =   1275
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":2AFB0
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1275
      End
      Begin VB.Image PlayerFacePct 
         Height          =   1275
         Index           =   4
         Left            =   3720
         Picture         =   "frmMain.frx":3120E
         Stretch         =   -1  'True
         Top             =   120
         Width           =   1275
      End
      Begin VB.Image PlayerFacePct 
         Height          =   1275
         Index           =   5
         Left            =   3720
         Picture         =   "frmMain.frx":3746C
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Image PlayerFacePct 
         Height          =   1275
         Index           =   6
         Left            =   3720
         Picture         =   "frmMain.frx":3D6CA
         Stretch         =   -1  'True
         Top             =   3000
         Width           =   1275
      End
      Begin VB.Image PlayerFacePct 
         Height          =   1275
         Index           =   7
         Left            =   3720
         Picture         =   "frmMain.frx":43928
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         Height          =   1275
         Index           =   0
         Left            =   1440
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.PictureBox GameMenu 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   6840
      ScaleHeight     =   6015
      ScaleWidth      =   7335
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   7335
      Begin VB.Timer MenuSlide2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3840
         Top             =   4200
      End
      Begin VB.Timer MenuSlide 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3120
         Top             =   4200
      End
      Begin VB.Image MenuPos 
         Height          =   555
         Left            =   240
         Picture         =   "frmMain.frx":49B86
         Top             =   240
         Width           =   450
      End
      Begin VB.Label MenuLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "End"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   14
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Image MenuImage 
         Height          =   375
         Index           =   6
         Left            =   240
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label MenuLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   13
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Image MenuImage 
         Height          =   375
         Index           =   5
         Left            =   240
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label MenuLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   720
         TabIndex        =   12
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label MenuLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Weapons"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   11
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label MenuLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   10
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label MenuLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Players"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label MenuLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Continue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
      Begin VB.Image MenuImage 
         Height          =   375
         Index           =   4
         Left            =   240
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Image MenuImage 
         Height          =   375
         Index           =   3
         Left            =   240
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Image MenuImage 
         Height          =   375
         Index           =   2
         Left            =   240
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Image MenuImage 
         Height          =   375
         Index           =   1
         Left            =   240
         Top             =   960
         Width           =   2175
      End
      Begin VB.Image MenuImage 
         Height          =   375
         Index           =   0
         Left            =   240
         Picture         =   "frmMain.frx":4A426
         Top             =   240
         Width           =   2190
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H005B3633&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   7215
      TabIndex        =   1
      Top             =   6000
      Width           =   7215
      Begin VB.Timer ReadPlyrPosition 
         Interval        =   250
         Left            =   6720
         Top             =   120
      End
      Begin VB.Timer ClearTextLines 
         Enabled         =   0   'False
         Interval        =   3000
         Left            =   240
         Top             =   120
      End
      Begin VB.Image PlayerTalking 
         Height          =   1335
         Left            =   5760
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label TextLn2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label TextLn1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.PictureBox MapBlock 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7215
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   7215
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      Begin VB.Timer FriendsMovement 
         Enabled         =   0   'False
         Interval        =   60
         Left            =   1320
         Top             =   120
      End
      Begin VB.Timer PlayerMovement 
         Enabled         =   0   'False
         Interval        =   60
         Left            =   720
         Top             =   120
      End
      Begin VB.PictureBox IntroBlock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   600
         Picture         =   "frmMain.frx":4CF8A
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   615
         Begin VB.Timer MapNameEffect 
            Enabled         =   0   'False
            Interval        =   25
            Left            =   0
            Top             =   0
         End
         Begin VB.Label MapNameLbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Map Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   5520
            Width           =   6495
         End
      End
      Begin VB.Image FriendBlock 
         Height          =   615
         Index           =   6
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image FriendBlock 
         Height          =   615
         Index           =   5
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image FriendBlock 
         Height          =   615
         Index           =   4
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image FriendBlock 
         Height          =   615
         Index           =   3
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image FriendBlock 
         Height          =   615
         Index           =   2
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image FriendBlock 
         Height          =   615
         Index           =   1
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image FriendBlock 
         Height          =   615
         Index           =   0
         Left            =   1800
         Stretch         =   -1  'True
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image PlayerBlock 
         Height          =   615
         Left            =   1200
         Stretch         =   -1  'True
         Top             =   600
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
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ClearTextLines_Timer()
TextLn1.Caption = ""
TextLn2.Caption = ""
PlayerTalking.Picture = Me.Picture
GameSleep = False
ClearTextLines.Enabled = False
End Sub

Private Sub EnemyWar_Timer()

PlayerListHealth(PlayerListSelected) = PlayerListHealth(PlayerListSelected) - EnemyAttack(GameMenuWarPosEnm)

MenuPosWar.Visible = True
EnemyWar.Enabled = False
End Sub

Private Sub Form_Activate()
Picture2.SetFocus
End Sub

Private Sub Form_Load()
IntroBlock.Width = MapBlock.Width
IntroBlock.Height = MapBlock.Height
Me.Width = MapBlock.Width
GameMenu.Left = MapBlock.Width
GameMenu.Picture = IntroBlock.Picture
StatusBox.Left = MapBlock.Width
StatusBox.Picture = IntroBlock.Picture
BackImgWar.Picture = IntroBlock.Picture
SellerBox.Picture = IntroBlock.Picture
ItemBox.Picture = IntroBlock.Picture
BlockSizeX = 14
BlockSizeY = 12
frmMain.MapBlock.Left = -frmMain.Block(0).Width
frmMain.MapBlock.Top = -frmMain.Block(0).Height
Me.Width = Picture2.Width + 80
WarBox.Left = 0
ItemBox.Left = 0
Dim i As Integer
For i = 1 To MenuImage.Count - 1
MenuImage(i).Picture = MenuImage(0).Picture
Next
For i = 0 To MenuWar.Count - 1
MenuWar(i).Picture = MenuImage(0).Picture
Next
For i = 0 To MenuItemsWar.Count - 1
MenuItemsWar(i).Picture = MenuImage(0).Picture
Next
For i = 0 To MenuBuy.Count - 1
MenuBuy(i).Picture = MenuImage(0).Picture
Next
For i = 0 To MenuBuy2.Count - 1
MenuBuy2(i).Picture = MenuImage(0).Picture
Next
For i = 0 To MenuBuy3.Count - 1
MenuBuy3(i).Picture = MenuImage(0).Picture
Next
For i = 0 To MenuBuy5.Count - 1
MenuBuy5(i).Picture = MenuImage(0).Picture
Next
For i = 0 To ImgOnly.Count - 1
ImgOnly(i).Picture = MenuImage(0).Picture
Next
MenuBuy4.Picture = MenuBuy3(0).Picture

For i = 0 To frmMain.Block.Count - 1
frmMain.Block(i).Picture = frmMain.Picture
frmMain.Block(i).BorderStyle = 0
Next

Create_Player

Load_Map 0
Refresh_All

Refresh_Player

End Sub

Private Sub FriendsMovement_Timer()
Dim i As Integer
Dim SpeedMovement As Integer
SpeedMovement = 60

Dim NumOfFinished As Integer

For i = 0 To 6
If PlayerGMWalkEnbl(i) = True Then

NumOfFinished = NumOfFinished + 1

If FriendBlock(i).Left < Block(PlayerGMBlock(i)).Left Then
FriendBlock(i).Left = FriendBlock(i).Left + SpeedMovement
frmMain.FriendBlock(i).Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerGMName(i) + PlayerGMDirection(i) + Format(PlayerGMWalk(i)) + ".emf")
Else
PlayerGMWalkEnbl(i) = False
End If

If FriendBlock(i).Left > Block(PlayerGMBlock(i)).Left Then
FriendBlock(i).Left = FriendBlock(i).Left - SpeedMovement
frmMain.FriendBlock(i).Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerGMName(i) + PlayerGMDirection(i) + Format(PlayerGMWalk(i)) + ".emf")
Else
PlayerGMWalkEnbl(i) = False
End If

If FriendBlock(i).Top > Block(PlayerGMBlock(i)).Top Then
FriendBlock(i).Top = FriendBlock(i).Top - SpeedMovement
frmMain.FriendBlock(i).Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerGMName(i) + PlayerGMDirection(i) + Format(PlayerGMWalk(i)) + ".emf")
Else
PlayerGMWalkEnbl(i) = False
End If
If FriendBlock(i).Top < Block(PlayerGMBlock(i)).Top Then
FriendBlock(i).Top = FriendBlock(i).Top + SpeedMovement
frmMain.FriendBlock(i).Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerGMName(i) + PlayerGMDirection(i) + Format(PlayerGMWalk(i)) + ".emf")
Else
PlayerGMWalkEnbl(i) = False
End If

PlayerGMWalk(i) = PlayerGMWalk(i) + 1
If PlayerGMWalk(i) = 4 Then PlayerGMWalk(i) = 1

End If
Next

If NumOfFinished = 0 Then
For i = 0 To 6
If PlayerGMDirection(i) = "U" Then PlayerGMDirection(i) = "D": GoTo 1
If PlayerGMDirection(i) = "D" Then PlayerGMDirection(i) = "U": GoTo 1
If PlayerGMDirection(i) = "L" Then PlayerGMDirection(i) = "R": GoTo 1
If PlayerGMDirection(i) = "R" Then PlayerGMDirection(i) = "L": GoTo 1
1:
Next
Refresh_Friends
GameSleep = False
FriendsMovement.Enabled = False
End If
End Sub

Private Sub MapNameEffect_Timer()
If Len(frmMain.MapNameLbl.Caption) < Len(MapName) Then
frmMain.MapNameLbl.Caption = Left(MapName, Len(frmMain.MapNameLbl.Caption) + 1)
If frmMain.MapNameLbl.Caption = MapName Then MapNameEffect.Interval = 50 * Len(frmMain.MapNameLbl.Caption)
Else
MapNameEffect.Interval = 25
MapNameEffect.Enabled = False
IntroBlock.Visible = False
End If
End Sub

Private Sub MenuSlide_Timer()
If GameMenu.Visible = True Then
If GameMenu.Left > 0 Then
GameMenu.Left = GameMenu.Left - 250
Else
MenuSlide.Enabled = False
End If
Else
GameMenu.Left = MapBlock.Width
MenuSlide.Enabled = False
End If
End Sub

Private Sub MenuSlide2_Timer()
If StatusBox.Visible = True Then
If StatusBox.Left > 0 Then
StatusBox.Left = StatusBox.Left - 250
Else
MenuSlide2.Enabled = False
End If
Else
StatusBox.Left = MapBlock.Width
MenuSlide2.Enabled = False
End If
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
'MsgBox KeyCode

Dim XX() As String

If KeyCode = 13 Then KeyCode = 32 ' Enter = Space

If KeyCode = 32 And Right(TextLn1.Caption, 1) = ":" Then
ClearTextLines_Timer
End If

If GameSeller = True Then

If BuyFinger5.Visible = True Then
If KeyCode = 40 Or KeyCode = 39 Then ' Down Right
GameMenuPosBuy4 = GameMenuPosBuy4 + 1
If GameMenuPosBuy4 > MenuBuy5.Count - 1 Then GameMenuPosBuy4 = 0
BuyFinger5.Top = MenuBuy5(GameMenuPosBuy4).Top
TextLn1.Caption = MenuLblBuy5(GameMenuPosBuy4).Caption
If GameMenuPosBuy4 = 0 Then
TextLn2.Caption = "NO. = " + Format(PotionNum)
Else
TextLn2.Caption = "NO. = " + Format(EnergyNum)
End If
End If
If KeyCode = 38 Or KeyCode = 37 Then ' Up Left
GameMenuPosBuy4 = GameMenuPosBuy4 - 1
If GameMenuPosBuy4 < 0 Then GameMenuPosBuy4 = MenuBuy5.Count - 1
BuyFinger5.Top = MenuBuy5(GameMenuPosBuy4).Top
TextLn1.Caption = MenuLblBuy5(GameMenuPosBuy4).Caption
If GameMenuPosBuy4 = 0 Then
TextLn2.Caption = "NO. = " + Format(PotionNum)
Else
TextLn2.Caption = "NO. = " + Format(EnergyNum)
End If
End If
If KeyCode = 32 Then ' Space
Select Case GameMenuPosBuy4
Case Is = 0 ' sell Potion
If PotionNum > 0 Then
PotionNum = PotionNum - 1
PlayerMoney = PlayerMoney + (PotionPrice \ 2)
TextLn2.Caption = "NO. = " + Format(PotionNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
Else
TextLn2.Caption = "Don't have potions !" + vbCrLf + "NO. = " + Format(PotionNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
End If
Case Is = 1
If EnergyNum > 0 Then
EnergyNum = EnergyNum - 1
PlayerMoney = PlayerMoney + (EnergyPrice \ 2)
TextLn2.Caption = "NO. = " + Format(EnergyNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
Else
TextLn2.Caption = "Don't have more !" + vbCrLf + "NO. = " + Format(EnergyNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
End If
End Select
End If
If KeyCode = 27 Then ' Esc
BuyFinger5.Visible = False
BuyFinger.Visible = True
Exit Sub
End If
End If

If BuyFinger3.Visible = True Then
If KeyCode = 40 Or KeyCode = 39 Then ' Down Right
GameMenuPosBuy3 = GameMenuPosBuy3 + 1
If GameMenuPosBuy3 > MenuBuy3.Count - 1 Then GameMenuPosBuy3 = 0
BuyFinger3.Top = MenuBuy3(GameMenuPosBuy3).Top
TextLn1.Caption = MenuLblBuy3(GameMenuPosBuy3).Caption
If GameMenuPosBuy3 = 0 Then
TextLn2.Caption = "NO. = " + Format(PotionNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
Else
TextLn2.Caption = "NO. = " + Format(EnergyNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
End If
End If
If KeyCode = 38 Or KeyCode = 37 Then ' Up Left
GameMenuPosBuy3 = GameMenuPosBuy3 - 1
If GameMenuPosBuy3 < 0 Then GameMenuPosBuy3 = MenuBuy3.Count - 1
BuyFinger3.Top = MenuBuy3(GameMenuPosBuy3).Top
TextLn1.Caption = MenuLblBuy3(GameMenuPosBuy3).Caption
If GameMenuPosBuy3 = 0 Then
TextLn2.Caption = "NO. = " + Format(PotionNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
Else
TextLn2.Caption = "NO. = " + Format(EnergyNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
End If
End If
If KeyCode = 32 Then ' Space
Select Case GameMenuPosBuy3
Case Is = 0 ' buy Potion
If PlayerMoney >= PotionPrice Then
PlayerMoney = PlayerMoney - PotionPrice
PotionNum = PotionNum + 1
TextLn2.Caption = "NO. = " + Format(PotionNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
Else
TextLn2.Caption = "No enough cash !" + vbCrLf + "NO. = " + Format(PotionNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
End If
Case Is = 1
If PlayerMoney >= EnergyPrice Then
PlayerMoney = PlayerMoney - EnergyPrice
EnergyNum = EnergyNum + 1
TextLn2.Caption = "NO. = " + Format(EnergyNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
Else
TextLn2.Caption = "No enough cash !" + vbCrLf + "NO. = " + Format(EnergyNum)
TextLn2.Caption = TextLn2.Caption + vbCrLf + "Money = " + Format(PlayerMoney) + " $"
End If
End Select
End If
If KeyCode = 27 Then ' Esc
BuyFinger3.Visible = False
BuyFinger2.Visible = True
Exit Sub
End If
End If

If BuyFinger2.Visible = True Then
If KeyCode = 40 Or KeyCode = 39 Then ' Down Right
GameMenuPosBuy2 = GameMenuPosBuy2 + 1
If GameMenuPosBuy2 > MenuBuy2.Count - 1 Then GameMenuPosBuy2 = 0
BuyFinger2.Left = MenuBuy2(GameMenuPosBuy2).Left
End If
If KeyCode = 38 Or KeyCode = 37 Then ' Up Left
GameMenuPosBuy2 = GameMenuPosBuy2 - 1
If GameMenuPosBuy2 < 0 Then GameMenuPosBuy2 = MenuBuy2.Count - 1
BuyFinger2.Left = MenuBuy2(GameMenuPosBuy2).Left
End If
If KeyCode = 32 Then ' Space
Select Case GameMenuPosBuy2
Case Is = 0
BuyFinger2.Visible = False
BuyFinger3.Visible = True
GameMenuPosBuy3 = 0
TextLn1.Caption = MenuLblBuy3(GameMenuPosBuy3).Caption
TextLn2.Caption = "NO. = " + Format(PotionNum)
BuyFinger3.Top = MenuBuy3(GameMenuPosBuy3).Top
Case Is = 1
If BuyFinger4.Visible = False Then
BuyFinger4.Visible = True
MenuLblBuy4.Visible = True
MenuBuy4.Visible = True
TextLn1.Caption = "Magic Improvement"
TextLn2.Caption = "Attack : " + Format(PlayerListMagic(PlayerListSelected)) + " >> " + Format(MagicVaule)
Else
BuyFinger4.Visible = False
MenuLblBuy4.Visible = False
MenuBuy4.Visible = False
If PlayerMoney >= MagicPrice Then
PlayerMoney = PlayerMoney - MagicPrice
PlayerListMagic(PlayerListSelected) = MagicVaule
PlayerListMagicPicture(PlayerListSelected) = MagicName
TextLn1.Caption = "Magic Improved !"
TextLn2.Caption = "Attack : " + Format(MagicVaule)
Else
TextLn1.Caption = "No enough cash !"
TextLn2.Caption = "Attack : " + Format(PlayerListMagic(PlayerListSelected))
End If
End If
End Select
End If
If KeyCode = 27 Then ' Esc
BuyFinger2.Visible = False
BuyFinger.Visible = True
BuyFinger4.Visible = False
MenuLblBuy4.Visible = False
MenuBuy4.Visible = False
Exit Sub
End If
End If

If BuyFinger.Visible = True Then
If KeyCode = 40 Or KeyCode = 39 Then ' Down Right
GameMenuPosBuy = GameMenuPosBuy + 1
If GameMenuPosBuy > MenuBuy.Count - 1 Then GameMenuPosBuy = 0
BuyFinger.Left = MenuBuy(GameMenuPosBuy).Left
BuyFinger4.Visible = False
MenuLblBuy4.Visible = False
MenuBuy4.Visible = False
End If
If KeyCode = 38 Or KeyCode = 37 Then ' Up Left
GameMenuPosBuy = GameMenuPosBuy - 1
If GameMenuPosBuy < 0 Then GameMenuPosBuy = MenuBuy.Count - 1
BuyFinger.Left = MenuBuy(GameMenuPosBuy).Left
BuyFinger4.Visible = False
MenuLblBuy4.Visible = False
MenuBuy4.Visible = False
End If
If KeyCode = 32 Then ' Space
Select Case GameMenuPosBuy
Case Is = 0
BuyFinger.Visible = False
BuyFinger2.Visible = True
GameMenuPosBuy2 = 0
BuyFinger2.Left = MenuBuy2(GameMenuPosBuy2).Left
Case Is = 1
BuyFinger.Visible = False
BuyFinger5.Visible = True
GameMenuPosBuy4 = 0
BuyFinger5.Top = MenuBuy5(GameMenuPosBuy4).Top
Case Is = 2
If PlayerDirection = "R" Then PlayerPosX = PlayerPosX - 1
If PlayerDirection = "L" Then PlayerPosX = PlayerPosX + 1
If PlayerDirection = "U" Then PlayerPosY = PlayerPosY + 1
If PlayerDirection = "D" Then PlayerPosY = PlayerPosY - 1
Refresh_Player
GameSeller = False
GameSleep = False
SellerBox.Visible = False
End Select
End If
If KeyCode = 27 Then ' Esc
If PlayerDirection = "R" Then PlayerPosX = PlayerPosX - 1
If PlayerDirection = "L" Then PlayerPosX = PlayerPosX + 1
If PlayerDirection = "U" Then PlayerPosY = PlayerPosY + 1
If PlayerDirection = "D" Then PlayerPosY = PlayerPosY - 1
Refresh_Player
GameSeller = False
GameSleep = False
SellerBox.Visible = False
Exit Sub
End If

End If

End If

If GameWar = True And EnemyWar.Enabled = False Then
If MenuPosWar.Visible = True Then
If KeyCode = 40 Or KeyCode = 39 Then ' Down
GameMenuWarPos = GameMenuWarPos + 1
If GameMenuWarPos > MenuWar.Count - 1 Then GameMenuWarPos = 0
MenuPosWar.Top = MenuWar(GameMenuWarPos).Top
End If
If KeyCode = 38 Or KeyCode = 37 Then ' Up
GameMenuWarPos = GameMenuWarPos - 1
If GameMenuWarPos < 0 Then GameMenuWarPos = MenuWar.Count - 1
MenuPosWar.Top = MenuWar(GameMenuWarPos).Top
End If
End If
If MenuPosWar2.Visible = True Then
If KeyCode = 40 Or KeyCode = 39 Then ' Down
GameMenuWarPosEnm = GameMenuWarPosEnm + 1
If GameMenuWarPosEnm > EnemyNum - 1 Then GameMenuWarPosEnm = 0
Do While EnemyBlock(GameMenuWarPosEnm).Visible = False
GameMenuWarPosEnm = GameMenuWarPosEnm + 1
If GameMenuWarPosEnm > EnemyNum - 1 Then GameMenuWarPosEnm = 0
Loop
MenuPosWar2.Top = EnemyBlock(GameMenuWarPosEnm).Top
MenuPosWar2.Left = EnemyBlock(GameMenuWarPosEnm).Left
TextLn1.Caption = EnemyName(GameMenuWarPosEnm)
TextLn2.Caption = "HP : " + Format(EnemyHelath(GameMenuWarPosEnm)) + vbCrLf + "Atk : " + Format(EnemyAttack(GameMenuWarPosEnm))
End If
If KeyCode = 38 Or KeyCode = 37 Then ' Up
GameMenuWarPosEnm = GameMenuWarPosEnm - 1
If GameMenuWarPosEnm < 0 Then GameMenuWarPosEnm = EnemyNum - 1
Do While EnemyBlock(GameMenuWarPosEnm).Visible = False
GameMenuWarPosEnm = GameMenuWarPosEnm - 1
If GameMenuWarPosEnm < 0 Then GameMenuWarPosEnm = EnemyNum - 1
Loop
MenuPosWar2.Top = EnemyBlock(GameMenuWarPosEnm).Top
MenuPosWar2.Left = EnemyBlock(GameMenuWarPosEnm).Left
TextLn1.Caption = EnemyName(GameMenuWarPosEnm)
TextLn2.Caption = "HP : " + Format(EnemyHelath(GameMenuWarPosEnm)) + vbCrLf + "Atk : " + Format(EnemyAttack(GameMenuWarPosEnm))
End If
End If

If MenuItemWar.Visible = True Then
If KeyCode = 40 Or KeyCode = 39 Then ' Down
GameMenuPos2 = GameMenuPos2 + 1
If GameMenuPos2 > MenuItemsWar.Count - 1 Then GameMenuPos2 = 0
MenuItemWar.Top = MenuItemsWar(GameMenuPos2).Top
MenuItemWar.Left = MenuItemsWar(GameMenuPos2).Left
TextLn1.Caption = ItemLblWar(GameMenuPos2).Caption
If GameMenuPos2 = 0 Then
TextLn2.Caption = "NO. = " + Format(PotionNum, "000")
Else
TextLn2.Caption = "NO. = " + Format(EnergyNum, "000")
End If
End If
If KeyCode = 38 Or KeyCode = 37 Then ' Up
GameMenuPos2 = GameMenuPos2 - 1
If GameMenuPos2 < 0 Then GameMenuPos2 = MenuItemsWar.Count - 1
MenuItemWar.Top = MenuItemsWar(GameMenuPos2).Top
MenuItemWar.Left = MenuItemsWar(GameMenuPos2).Left
TextLn1.Caption = ItemLblWar(GameMenuPos2).Caption
If GameMenuPos2 = 0 Then
TextLn2.Caption = "NO. = " + Format(PotionNum, "000")
Else
TextLn2.Caption = "NO. = " + Format(EnergyNum, "000")
End If
End If
If KeyCode = 32 Then ' Space
Select Case GameMenuPos2
Case Is = 0
If PotionNum > 0 Then
PotionNum = PotionNum - 1
PlayerListHealth(PlayerListSelected) = PlayerListHealthMax(PlayerListSelected)
TextLn2.Caption = "Potion had used !"
Else
TextLn2.Caption = "You don't have enough potions !"
End If
Case Is = 1
If EnergyNum > 0 Then
EnergyNum = EnergyNum - 1
PlayerListEnergy(PlayerListSelected) = PlayerListEnergyMax(PlayerListSelected)
TextLn2.Caption = "Energy had used !"
Else
TextLn2.Caption = "You don't have enough !"
End If
End Select
End If
End If

If KeyCode = 27 Then ' Esc
MenuPosWar2.Visible = False
MenuItemWar.Visible = False
MenuPosWar.Visible = True
End If
If KeyCode = 32 Then ' Space
If MenuPosWar.Visible = True Then
Select Case GameMenuWarPos
Case Is = 0
MenuPosWar2.Visible = True
MenuPosWar.Visible = False
Do While EnemyBlock(GameMenuWarPosEnm).Visible = False
GameMenuWarPosEnm = GameMenuWarPosEnm + 1
If GameMenuWarPosEnm > EnemyNum - 1 Then GameMenuWarPosEnm = 0
Loop
DamageSelected = PlayerListPower(PlayerListSelected)
TextLn1.Caption = EnemyName(GameMenuWarPosEnm)
TextLn2.Caption = "HP : " + Format(EnemyHelath(GameMenuWarPosEnm)) + vbCrLf + "Atk : " + Format(EnemyAttack(GameMenuWarPosEnm))
WeaponEffect.Picture = LoadPicture(App.Path + "\Data\Weapons\Attack.emf")
MenuPosWar2.Top = EnemyBlock(GameMenuWarPosEnm).Top
MenuPosWar2.Left = EnemyBlock(GameMenuWarPosEnm).Left
Exit Sub
Case Is = 1
MenuPosWar2.Visible = True
MenuPosWar.Visible = False
Do While EnemyBlock(GameMenuWarPosEnm).Visible = False
GameMenuWarPosEnm = GameMenuWarPosEnm + 1
If GameMenuWarPosEnm > EnemyNum - 1 Then GameMenuWarPosEnm = 0
Loop
DamageSelected = PlayerListMagic(PlayerListSelected)
TextLn1.Caption = EnemyName(GameMenuWarPosEnm)
TextLn2.Caption = "HP : " + Format(EnemyHelath(GameMenuWarPosEnm)) + vbCrLf + "Atk : " + Format(EnemyAttack(GameMenuWarPosEnm))
WeaponEffect.Picture = LoadPicture(App.Path + "\Data\Weapons\" + PlayerListMagicPicture(GameMenuWarPosEnm) + ".emf")
MenuPosWar2.Top = EnemyBlock(GameMenuWarPosEnm).Top
MenuPosWar2.Left = EnemyBlock(GameMenuWarPosEnm).Left
Exit Sub
Case Is = 2
MenuItemWar.Visible = True
MenuPosWar.Visible = False
GameMenuPos2 = 0
MenuItemWar.Top = MenuItemsWar(GameMenuPos2).Top
MenuItemWar.Left = MenuItemsWar(GameMenuPos2).Left
TextLn1.Caption = ItemLblWar(GameMenuPos2).Caption
TextLn2.Caption = "NO. = " + Format(PotionNum, "000")
End Select
End If
If MenuPosWar2.Visible = True Then
WeaponEffect.Top = EnemyBlock(GameMenuWarPosEnm).Top + (EnemyBlock(GameMenuWarPosEnm).Height \ 2) - (WeaponEffect.Height \ 2)
WeaponEffect.Left = EnemyBlock(GameMenuWarPosEnm).Left + (EnemyBlock(GameMenuWarPosEnm).Width \ 2) - (WeaponEffect.Width \ 2)
WeaponEffect.Visible = True
RemoveEffect.Enabled = True
EnemyWar.Enabled = True
MenuPosWar.Visible = False
MenuPosWar2.Visible = False
End If
End If
End If

If GameSleep = True Then
Exit Sub
End If

Dim i, ii As Integer
Dim TempPosX, TempPosY As Integer
Dim TempX As String
Dim TempXX() As String

Dim ItemColct As String
Dim PosAddX, PosAddY As Integer

' ## Menu ##

If GameMenu.Visible = True And StatusBox.Visible = False Then
If KeyCode = 40 Or KeyCode = 39 Then ' Down
GameMenuPos = GameMenuPos + 1
If GameMenuPos > MenuImage.Count - 1 Then GameMenuPos = 0
End If
If KeyCode = 38 Or KeyCode = 37 Then ' Up
GameMenuPos = GameMenuPos - 1
If GameMenuPos < 0 Then GameMenuPos = MenuImage.Count - 1

End If
If KeyCode = 32 Then ' Space
Select Case GameMenuPos
Case Is = 0
GameMenu.Visible = Not GameMenu.Visible
MenuSlide.Enabled = True
Case Is = 1
For i = 0 To HealthLbl.Count - 1
HealthLbl(i).Caption = "HP : " + Format(PlayerListHealth(i), "000") + " \ " + Format(PlayerListHealthMax(i), "000")
EnergyLbl(i).Caption = "MP : " + Format(PlayerListEnergy(i), "000") + " \ " + Format(PlayerListEnergyMax(i), "000")
PowerLbl(i).Caption = "Lvl : " + Format(PlayerListPower(i), "000")
If PlayerListLocked(i) = True Then
HealthLbl(i).Visible = False
EnergyLbl(i).Visible = False
PowerLbl(i).Visible = False
Shape1(i).Visible = False
End If
Next
StatusBox.Visible = True
PlayerNameLblStatus.Caption = PlayerListName(GameMenuPos2)
MenuSlide2.Enabled = True
Case Is = 2
ItemBox.Visible = True
Case Is = 4
Load_Game
Refresh_Player
Refresh_All
GameMenu.Visible = Not GameMenu.Visible
StatusBox.Visible = False
MenuSlide.Enabled = True
Case Is = 5
Save_Game
GameMenu.Visible = Not GameMenu.Visible
StatusBox.Visible = False
MenuSlide.Enabled = True
Case Is = 6
End
End Select
End If
MenuPos.Top = MenuImage(GameMenuPos).Top
End If

If ItemBox.Visible = True Then
If KeyCode = 27 Then ' Esc
ItemBox.Visible = False
End If
End If

If StatusBox.Visible = True Then
If KeyCode = 32 Then ' Space
If PlayerListLocked(GameMenuPos2) = False Then
PlayerListSelected = GameMenuPos2
XX() = Split(PlayerListName(GameMenuPos2), " ")
PlayerName = XX(0)
For i = 0 To Shape1.Count - 1
Shape1(i).BorderStyle = 0
Next
Shape1(GameMenuPos2).BorderStyle = 1
Refresh_Player
TextLn1.Caption = PlayerName + " Selected !"
ClearTextLines.Enabled = True
Else
TextLn1.Caption = "Locked !"
ClearTextLines.Enabled = True
End If
End If
If KeyCode = 40 Then  ' Down
GameMenuPos2 = GameMenuPos2 + 1
If GameMenuPos2 > PlayerFacePct.Count - 1 Then GameMenuPos2 = 0
PlayerNameLblStatus.Caption = PlayerListName(GameMenuPos2)
MenuPos2.Top = PlayerFacePct(GameMenuPos2).Top + PlayerFacePct(GameMenuPos2).Height - MenuPos2.Height + 120
MenuPos2.Left = PlayerFacePct(GameMenuPos2).Left - 120
End If
If KeyCode = 39 Then  ' Right
GameMenuPos2 = GameMenuPos2 + 4
If GameMenuPos2 > PlayerFacePct.Count - 1 Then GameMenuPos2 = GameMenuPos2 - 4
PlayerNameLblStatus.Caption = PlayerListName(GameMenuPos2)
MenuPos2.Top = PlayerFacePct(GameMenuPos2).Top + PlayerFacePct(GameMenuPos2).Height - MenuPos2.Height + 120
MenuPos2.Left = PlayerFacePct(GameMenuPos2).Left - 120
End If
If KeyCode = 38 Then  ' Up
GameMenuPos2 = GameMenuPos2 - 1
If GameMenuPos2 < 0 Then GameMenuPos2 = PlayerFacePct.Count - 1
PlayerNameLblStatus.Caption = PlayerListName(GameMenuPos2)
MenuPos2.Top = PlayerFacePct(GameMenuPos2).Top + PlayerFacePct(GameMenuPos2).Height - MenuPos2.Height + 120
MenuPos2.Left = PlayerFacePct(GameMenuPos2).Left - 120
End If
If KeyCode = 37 Then  ' Left
GameMenuPos2 = GameMenuPos2 - 4
If GameMenuPos2 < 0 Then GameMenuPos2 = GameMenuPos2 + 4
PlayerNameLblStatus.Caption = PlayerListName(GameMenuPos2)
MenuPos2.Top = PlayerFacePct(GameMenuPos2).Top + PlayerFacePct(GameMenuPos2).Height - MenuPos2.Height + 120
MenuPos2.Left = PlayerFacePct(GameMenuPos2).Left - 120
End If
End If
If KeyCode = 27 Then ' Esc
GameMenu.Visible = Not GameMenu.Visible
StatusBox.Visible = False
MenuSlide.Enabled = True
End If

' ## Movements ##

If PlayerMovement.Enabled = False And GameMenu.Visible = False And StatusBox.Visible = False And IntroBlock.Visible = False Then

PosAddX = 0
PosAddY = 0
Select Case BlockHave(PlayerPosX, PlayerPosY)
Case Is = "Door"
ItemColct = "Door"
If BlockHave(PlayerPosX + PosAddX, PlayerPosY + PosAddY) = ItemColct And (CheckPoint("Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY)) = "" Or CheckPoint("Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY)) = ItemColct) Then
If Not BlockDoorGotoPosX(PlayerPosX + PosAddX, PlayerPosY + PosAddY) = "?" Then
TempPosX = BlockDoorGotoPosX(PlayerPosX + PosAddX, PlayerPosY + PosAddY)
Else
TempPosX = -1
End If
If Not BlockDoorGotoPosY(PlayerPosX + PosAddX, PlayerPosY + PosAddY) = "?" Then
TempPosY = BlockDoorGotoPosY(PlayerPosX + PosAddX, PlayerPosY + PosAddY)
Else
TempPosY = -1
End If
Load_Map Int(BlockDoorGotoMap(PlayerPosX + PosAddX, PlayerPosY + PosAddY))
If Not TempPosX = -1 Then PlayerPosX = TempPosX
If Not TempPosY = -1 Then PlayerPosY = TempPosY
Refresh_All
Refresh_Player
Else
TextLn1.Caption = "Closed,.. I can't !"
TextLn2.Caption = ""
ClearTextLines.Enabled = True
End If
Case Is = "DoorLock"
ItemColct = "DoorLock"
If BlockHave(PlayerPosX + PosAddX, PlayerPosY + PosAddY) = ItemColct And (CheckPoint("Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY)) = "" Or CheckPoint("Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY)) = ItemColct) Then
AddNewChekhPoint "Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY), "Nothing"
If Not BlockDoorGotoPosX(PlayerPosX + PosAddX, PlayerPosY + PosAddY) = "?" Then
TempPosX = BlockDoorGotoPosX(PlayerPosX + PosAddX, PlayerPosY + PosAddY)
Else
TempPosX = -1
End If
If Not BlockDoorGotoPosY(PlayerPosX + PosAddX, PlayerPosY + PosAddY) = "?" Then
TempPosY = BlockDoorGotoPosY(PlayerPosX + PosAddX, PlayerPosY + PosAddY)
Else
TempPosY = -1
End If
Load_Map Int(BlockDoorGotoMap(PlayerPosX + PosAddX, PlayerPosY + PosAddY))
If Not TempPosX = -1 Then PlayerPosX = TempPosX
If Not TempPosY = -1 Then PlayerPosY = TempPosY
Refresh_All
Refresh_Player
Else
TextLn1.Caption = "Closed,.. I can't !"
TextLn2.Caption = ""
ClearTextLines.Enabled = True
End If
Case Is = "Commands"
ItemColct = "Commands"
If BlockHave(PlayerPosX + PosAddX, PlayerPosY + PosAddY) = ItemColct And (CheckPoint("Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY)) = "" Or CheckPoint("Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY)) = ItemColct) Then
AddNewChekhPoint "Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY), "Nothing"
If CheckPoint("Command:" + BlockCommandsFile(PlayerPosX + PosAddX, PlayerPosY + PosAddY)) = "Done" Then TempX = "End Commands"
AddNewChekhPoint "Command:" + BlockCommandsFile(PlayerPosX + PosAddX, PlayerPosY + PosAddY), "Done"
Open App.Path + "\Data\Gfx\" + BlockCommandsFile(PlayerPosX + PosAddX, PlayerPosY + PosAddY) For Input As #3
Do While TempX <> "End Commands"
Input #3, TempX
TempXX() = Split(TempX, " ")
Select Case TempXX(0)
Case Is = "Change"
If CheckPoint("Map:" + Format(MapNum) + ",FloorWall:" + Format(TempXX(1)) + ":" + Format(TempXX(2))) = "Floor" Or BlockMode(Int(TempXX(1)), Int(TempXX(2))) = "Floor" Then
AddNewChekhPoint "Map:" + Format(MapNum) + ",FloorWall:" + Format(TempXX(1)) + ":" + Format(TempXX(2)), "Wall"
Else
AddNewChekhPoint "Map:" + Format(MapNum) + ",FloorWall:" + Format(TempXX(1)) + ":" + Format(TempXX(2)), "Floor"
End If
Case Is = "Talk"
TextLn1.Caption = TempXX(1) + " :"
TextLn2.Caption = Right(TempX, Len(TempX) - Len(TempXX(1)) - Len(TempXX(0)) - 2)
For i = 0 To PlayerFacePct.Count - 1
If Left(PlayerListName(i), Len(TempXX(1))) = TempXX(1) Then
PlayerTalking.Picture = PlayerFacePct(i).Picture
Exit For
End If
Next
GameSleep = True
Case Is = "Move"
If TempXX(1) = "Right" Then
Picture2_KeyDown 39, -1
End If
If TempXX(1) = "Left" Then
Picture2_KeyDown 37, -1
End If
If TempXX(1) = "Up" Then
Picture2_KeyDown 38, -1
End If
If TempXX(1) = "Down" Then
Picture2_KeyDown 40, -1
End If
GameSleep = True
Case Is = "Player"
PlayerName = TempXX(1)
Refresh_Player
Case Is = "Start"
GameTalking = True
ii = 0
For i = 0 To 7
If Not PlayerName = Left(PlayerListName(i), Len(PlayerName)) Then
TempXX() = Split(PlayerListName(i), " ")
PlayerGMName(ii) = TempXX(0)
If PlayerListLocked(ii) = False Then FriendBlock(ii).Visible = True
PlayerGMWalk(ii) = 1
frmMain.FriendBlock(ii).Left = frmMain.Block(PlayerPosBlock).Left
frmMain.FriendBlock(ii).Top = frmMain.Block(PlayerPosBlock).Top
Select Case ii
Case Is = 0
PlayerGMDirection(ii) = "U"
PlayerGMBlock(ii) = PlayerPosBlock - BlockSizeX
Case Is = 1
PlayerGMDirection(ii) = "U"
PlayerGMBlock(ii) = PlayerPosBlock - BlockSizeX - 1
Case Is = 2
PlayerGMDirection(ii) = "U"
PlayerGMBlock(ii) = PlayerPosBlock - BlockSizeX + 1
Case Is = 3
PlayerGMDirection(ii) = "L"
PlayerGMBlock(ii) = PlayerPosBlock - 1
Case Is = 4
PlayerGMDirection(ii) = "R"
PlayerGMBlock(ii) = PlayerPosBlock + 1
Case Is = 5
PlayerGMDirection(ii) = "D"
PlayerGMBlock(ii) = PlayerPosBlock + BlockSizeX - 1
Case Is = 6
PlayerGMDirection(ii) = "D"
PlayerGMBlock(ii) = PlayerPosBlock + BlockSizeX + 1
End Select
PlayerGMWalkEnbl(ii) = True
ii = ii + 1
End If
Next
Refresh_Friends
GameSleep = True
FriendsMovement.Enabled = True
Case Is = "End"
GameTalking = False
For i = 0 To 6
PlayerGMWalkEnbl(i) = False
FriendBlock(i).Visible = False
Next
Case Is = "Move>"
For i = 0 To 6
If PlayerGMName(i) = TempXX(1) Then
If TempXX(2) = "Right" Then
PlayerGMDirection(i) = "R"
PlayerGMBlock(i) = PlayerGMBlock(i) + 1
FriendBlock(i).Visible = True
PlayerGMWalkEnbl(i) = True
PlayerGMWalk(i) = 1
GameSleep = True
GameTalking = True
FriendsMovement.Enabled = True
End If
If TempXX(2) = "Left" Then
PlayerGMDirection(i) = "L"
PlayerGMBlock(i) = PlayerGMBlock(i) - 1
FriendBlock(i).Visible = True
PlayerGMWalkEnbl(i) = True
PlayerGMWalk(i) = 1
GameSleep = True
GameTalking = True
FriendsMovement.Enabled = True
End If
If TempXX(2) = "Up" Then
PlayerGMDirection(i) = "U"
PlayerGMBlock(i) = PlayerGMBlock(i) - BlockSizeX
FriendBlock(i).Visible = True
PlayerGMWalkEnbl(i) = True
PlayerGMWalk(i) = 1
GameSleep = True
GameTalking = True
FriendsMovement.Enabled = True
End If
If TempXX(2) = "Down" Then
PlayerGMDirection(i) = "D"
PlayerGMBlock(i) = PlayerGMBlock(i) + BlockSizeX
FriendBlock(i).Visible = True
PlayerGMWalkEnbl(i) = True
PlayerGMWalk(i) = 1
GameSleep = True
GameTalking = True
FriendsMovement.Enabled = True
End If
End If
Next
Case Is = "Show>"
For i = 0 To 6
If Not TempXX(1) = PlayerName Then
PlayerGMName(i) = TempXX(1)
PlayerGMWalk(i) = 1
FriendBlock(i).Visible = True
PlayerGMWalkEnbl(i) = True
If TempXX(2) = "Right" Then
PlayerGMBlock(i) = PlayerPosBlock + 1
PlayerGMDirection(i) = "R"
End If
If TempXX(2) = "Left" Then
PlayerGMBlock(i) = PlayerPosBlock - 1
PlayerGMDirection(i) = "L"
End If
If TempXX(2) = "Up" Then
PlayerGMBlock(i) = PlayerPosBlock - BlockSizeX
PlayerGMDirection(i) = "U"
End If
If TempXX(2) = "Down" Then
PlayerGMBlock(i) = PlayerPosBlock + BlockSizeX
PlayerGMDirection(i) = "D"
End If
End If
Next
GameTalking = True
Refresh_Friends
GameSleep = True
FriendsMovement.Enabled = True
Case Is = "Picture"
AddNewPicSource "Map:" + Format(MapNum) + ",Source:" + Format(TempXX(1)) + ":" + Format(TempXX(2)), Format(TempXX(3))
Refresh_All
Case Is = "Unlock"
For i = 0 To 7
If PlayerListName(i) = TempXX(1) Then
PlayerListLocked(i) = False
End If
Next
Case Is = "Lock"
For i = 0 To 7
If PlayerListName(i) = TempXX(1) Then
PlayerListLocked(i) = True
End If
Next
Case Is = "War"
GameWar = True
GameSleep = True
WarBox.Visible = True
Create_War Int(Int(BlockCommandsFile(PlayerPosX, PlayerPosY)))

End Select
Do While GameSleep = True
DoEvents
Loop
Loop
Close #3
End If

Case Is = "Seller"
ItemColct = "Seller"
GameSeller = True
GameSleep = True
Load_Seller BlockCommandsFile(PlayerPosX + PosAddX, PlayerPosY + PosAddY)

End Select

If KeyCode = 32 Then ' Space
Select Case PlayerDirection
Case Is = "R"
PosAddX = 1
PosAddY = 0
Case Is = "L"
PosAddX = -1
PosAddY = 0
Case Is = "U"
PosAddX = 0
PosAddY = -1
Case Is = "D"
PosAddX = 0
PosAddY = 1
End Select
Select Case BlockHave(PlayerPosX + PosAddX, PlayerPosY + PosAddY)
Case Is = "Potion"
ItemColct = "Potion"
If BlockHave(PlayerPosX + PosAddX, PlayerPosY + PosAddY) = ItemColct And (CheckPoint("Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY)) = "" Or CheckPoint("Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY)) = ItemColct) Then
If Not TextLn1.Caption = "" Then
PotionNum = PotionNum + 1
AddNewChekhPoint "Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY), "Nothing"
TextLn1.Caption = ItemColct + " collected ."
TextLn2.Caption = ""
End If
If Not TextLn1.Caption = ItemColct + " collected ." Then
TextLn1.Caption = "There is " + ItemColct + " , Do you like to Collect it ?"
TextLn2.Caption = "Press on SPACE bar to collect ."
ClearTextLines.Enabled = True
End If
Else
TextLn1.Caption = "There is Nothing ."
TextLn2.Caption = ""
ClearTextLines.Enabled = True
End If
Case Is = "Energy"
ItemColct = "Energy"
If BlockHave(PlayerPosX + PosAddX, PlayerPosY + PosAddY) = ItemColct And (CheckPoint("Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY)) = "" Or CheckPoint("Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY)) = ItemColct) Then
If Not TextLn1.Caption = "" Then
EnergyNum = EnergyNum + 1
AddNewChekhPoint "Map:" + Format(MapNum) + ",Block:" + Format(PlayerPosX + PosAddX) + ":" + Format(PlayerPosY + PosAddY), "Nothing"
TextLn1.Caption = ItemColct + " collected ."
TextLn2.Caption = ""
End If
If Not TextLn1.Caption = ItemColct + " collected ." Then
TextLn1.Caption = "There is " + ItemColct + " , Do you like to Collect it ?"
TextLn2.Caption = "Press on SPACE bar to collect ."
ClearTextLines.Enabled = True
End If
Else
TextLn1.Caption = "There is Nothing ."
TextLn2.Caption = ""
ClearTextLines.Enabled = True
End If

End Select
End If

If GameTalking = False Then

If KeyCode = 39 Then ' Right
PlayerDirection = "R"
Refresh_Player
If BlockMode(PlayerPosX + 1, PlayerPosY) = "Floor" And (CheckPoint("Map:" + Format(MapNum) + ",FloorWall:" + Format(PlayerPosX + 1) + ":" + Format(PlayerPosY)) = "" Or CheckPoint("Map:" + Format(MapNum) + ",FloorWall:" + Format(PlayerPosX + 1) + ":" + Format(PlayerPosY)) = "Floor") Then
PlayerMovement.Enabled = True
End If
End If

If KeyCode = 37 And PlayerPosX > 1 Then ' Left
PlayerDirection = "L"
Refresh_Player
If BlockMode(PlayerPosX - 1, PlayerPosY) = "Floor" And (CheckPoint("Map:" + Format(MapNum) + ",FloorWall:" + Format(PlayerPosX - 1) + ":" + Format(PlayerPosY)) = "" Or CheckPoint("Map:" + Format(MapNum) + ",FloorWall:" + Format(PlayerPosX - 1) + ":" + Format(PlayerPosY)) = "Floor") Then
PlayerMovement.Enabled = True
End If
End If

If KeyCode = 38 And PlayerPosY > 1 Then ' Up
PlayerDirection = "U"
Refresh_Player
If BlockMode(PlayerPosX, PlayerPosY - 1) = "Floor" And (CheckPoint("Map:" + Format(MapNum) + ",FloorWall:" + Format(PlayerPosX) + ":" + Format(PlayerPosY - 1)) = "" Or CheckPoint("Map:" + Format(MapNum) + ",FloorWall:" + Format(PlayerPosX) + ":" + Format(PlayerPosY - 1)) = "Floor") Then
PlayerMovement.Enabled = True
End If
End If

If KeyCode = 40 Then ' Down
PlayerDirection = "D"
Refresh_Player
If BlockMode(PlayerPosX, PlayerPosY + 1) = "Floor" And (CheckPoint("Map:" + Format(MapNum) + ",FloorWall:" + Format(PlayerPosX) + ":" + Format(PlayerPosY + 1)) = "" Or CheckPoint("Map:" + Format(MapNum) + ",FloorWall:" + Format(PlayerPosX) + ":" + Format(PlayerPosY + 1)) = "Floor") Then
PlayerMovement.Enabled = True
End If
End If

End If

End If
End Sub

Private Sub PlayerMovement_Timer()
Dim SpeedMovement As Integer
SpeedMovement = 120

If PlayerDirection = "R" Then
MapStartX = PlayerPosX - ((BlockSizeX - 2) \ 2) + 1
MapStartY = PlayerPosY - ((BlockSizeY - 2) \ 2)
If MapStartX < 0 Then MapStartX = -1
If MapStartY < 0 Then MapStartY = -1
If PlayerBlock.Left < Block(PlayerPosBlock + 1).Left Then
If Not MapStartX = -1 Then
MapBlock.Left = MapBlock.Left - SpeedMovement
End If
PlayerBlock.Left = PlayerBlock.Left + SpeedMovement
frmMain.PlayerBlock.Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerName + PlayerDirection + Format(PlayerPosWalk) + ".emf")
PlayerPosWalk = PlayerPosWalk + 1
If PlayerPosWalk = 4 Then PlayerPosWalk = 1
Else
PlayerPosX = PlayerPosX + 1
RePos_Player
Refresh_Map_Pos
PlayerMovement.Enabled = False
GameSleep = False
End If
End If

If PlayerDirection = "L" Then
MapStartX = PlayerPosX - ((BlockSizeX - 2) \ 2)
MapStartY = PlayerPosY - ((BlockSizeY - 2) \ 2)
If MapStartX < 0 Then MapStartX = -1
If MapStartY < 0 Then MapStartY = -1
If PlayerBlock.Left > Block(PlayerPosBlock - 1).Left Then
If Not MapStartX = -1 Then
MapBlock.Left = MapBlock.Left + SpeedMovement
End If
PlayerBlock.Left = PlayerBlock.Left - SpeedMovement
frmMain.PlayerBlock.Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerName + PlayerDirection + Format(PlayerPosWalk) + ".emf")
PlayerPosWalk = PlayerPosWalk + 1
If PlayerPosWalk = 4 Then PlayerPosWalk = 1
Else
PlayerPosX = PlayerPosX - 1
RePos_Player
Refresh_Map_Pos
PlayerMovement.Enabled = False
GameSleep = False
End If
End If

If PlayerDirection = "U" Then
MapStartX = PlayerPosX - ((BlockSizeX - 2) \ 2)
MapStartY = PlayerPosY - ((BlockSizeY - 2) \ 2)
If MapStartX < 0 Then MapStartX = -1
If MapStartY < 0 Then MapStartY = -1
If PlayerBlock.Top > Block(PlayerPosBlock - BlockSizeX).Top Then
If Not MapStartY = -1 Then
MapBlock.Top = MapBlock.Top + SpeedMovement
End If
PlayerBlock.Top = PlayerBlock.Top - SpeedMovement
frmMain.PlayerBlock.Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerName + PlayerDirection + Format(PlayerPosWalk) + ".emf")
PlayerPosWalk = PlayerPosWalk + 1
If PlayerPosWalk = 4 Then PlayerPosWalk = 1
Else
PlayerPosY = PlayerPosY - 1
RePos_Player
Refresh_Map_Pos
PlayerMovement.Enabled = False
GameSleep = False
End If
End If

If PlayerDirection = "D" Then
MapStartX = PlayerPosX - ((BlockSizeX - 2) \ 2)
MapStartY = PlayerPosY - ((BlockSizeY - 2) \ 2) + 1
If MapStartX < 0 Then MapStartX = -1
If MapStartY < 0 Then MapStartY = -1
If PlayerBlock.Top < Block(PlayerPosBlock + BlockSizeX).Top Then
If Not MapStartY = -1 Then
MapBlock.Top = MapBlock.Top - SpeedMovement
End If
PlayerBlock.Top = PlayerBlock.Top + SpeedMovement
frmMain.PlayerBlock.Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerName + PlayerDirection + Format(PlayerPosWalk) + ".emf")
PlayerPosWalk = PlayerPosWalk + 1
If PlayerPosWalk = 4 Then PlayerPosWalk = 1
Else
PlayerPosY = PlayerPosY + 1
RePos_Player
Refresh_Map_Pos
PlayerMovement.Enabled = False
GameSleep = False
End If
End If
End Sub

Private Sub ReadPlyrPosition_Timer()
Call Picture2_KeyDown(-1, -1)
End Sub

Private Sub RemoveEffect_Timer()

EnemyHelath(GameMenuWarPosEnm) = EnemyHelath(GameMenuWarPosEnm) - DamageSelected
If EnemyHelath(GameMenuWarPosEnm) < 1 Then EnemyBlock(GameMenuWarPosEnm).Visible = False

WeaponEffect.Visible = False
RemoveEffect.Enabled = False
End Sub

Private Sub SellerUpdate_Timer()
If GameSeller = False Then Exit Sub

MoneyPlyrHav.Caption = Format(PlayerMoney) + " $"

Dim i As Integer
For i = 0 To MenuBuy2.Count - 1
MenuBuy2(i).Visible = BuyFinger2.Visible
MenuLblBuy2(i).Visible = BuyFinger2.Visible
Next
For i = 0 To MenuBuy3.Count - 1
MenuBuy3(i).Visible = BuyFinger3.Visible
MenuLblBuy3(i).Visible = BuyFinger3.Visible
Next
For i = 0 To MenuBuy5.Count - 1
MenuBuy5(i).Visible = BuyFinger5.Visible
MenuLblBuy5(i).Visible = BuyFinger5.Visible
Next
End Sub

Private Sub Timer2_Timer()
If ItemBox.Visible = False Then Exit Sub

MenuItems(0).Caption = "Potion NO. = " + Format(PotionNum)
MenuItems(1).Caption = "Energy NO. = " + Format(EnergyNum)
MenuItems(2).Caption = "Money = " + Format(PlayerMoney)
MenuItems(3).Caption = "Points = " + Format(PlayerPoints)
MenuItems(4).Caption = "Player Name = " + Format(PlayerName)
MenuItems(5).Caption = "Map Name = " + Format(MapName)
MenuItems(6).Caption = "Check Points NO. = " + Format(CheckPointNum)

End Sub

Private Sub WarUpdate_Timer()
If GameWar = True Then

Dim i, ii As Integer
Dim TempXX() As String

For i = 0 To ItemLblWar.Count - 1
ItemLblWar(i).Visible = MenuItemWar.Visible
MenuItemsWar(i).Visible = MenuItemWar.Visible
Next

PlayerTalking.Picture = PlayerFacePct(PlayerListSelected).Picture

PlyrHlth.Caption = " HP : " + Format(PlayerListHealth(PlayerListSelected), "000") + " \ " + Format(PlayerListHealthMax(PlayerListSelected), "000")
PlyrEnrg.Caption = " MP : " + Format(PlayerListEnergy(PlayerListSelected), "000") + " \ " + Format(PlayerListEnergyMax(PlayerListSelected), "000")
PlyrNam.Caption = PlayerListName(PlayerListSelected)

If EnemyHelath(GameMenuWarPosEnm) < 1 Then EnemyBlock(GameMenuWarPosEnm).Visible = False

For i = 0 To EnemyBlock.Count - 1
If EnemyBlock(i).Visible = False Then ii = ii + 1
Next

If ii = EnemyNum Then
PlayerMoney = PlayerMoney + EnemyMoney(GameMenuWarPosEnm)
PlayerPoints = PlayerPoints + EnemyPoints(GameMenuWarPosEnm)
TextLn1.Caption = "You won the war !"
TextLn2.Caption = "++ " + Format(EnemyMoney(GameMenuWarPosEnm)) + " $"
TextLn2.Caption = TextLn2.Caption + vbCrLf
TextLn2.Caption = TextLn2.Caption + "++ " + Format(EnemyPoints(GameMenuWarPosEnm)) + " Points"
Do While PlayerPoints > 999
PlayerPoints = PlayerPoints - 1000
PlayerListPower(PlayerListSelected) = PlayerListPower(PlayerListSelected) + 1
TextLn1.Caption = "Level Up !"
Loop
ClearTextLines.Enabled = True
GameWar = False
GameSleep = False
WarBox.Visible = False
End If

End If
End Sub
