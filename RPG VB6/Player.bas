Attribute VB_Name = "Player"

Global PlayerName As String
Global PlayerPosX As Integer
Global PlayerPosY As Integer
Global PlayerDirection As String
Global PlayerPosBlock As Integer
Global PlayerPosWalk As Integer

Global PlayerListName(7) As String
Global PlayerListHealth(7) As Integer
Global PlayerListEnergy(7) As Integer
Global PlayerListHealthMax(7) As Integer
Global PlayerListEnergyMax(7) As Integer
Global PlayerListPower(7) As Integer
Global PlayerListMagic(7) As Integer
Global PlayerListMagicPicture(7) As String
Global PlayerListLocked(7) As Boolean

Global PlayerListSelected As Integer

Global PlayerGMName(6) As String
Global PlayerGMDirection(6) As String
Global PlayerGMBlock(6) As String
Global PlayerGMWalk(6) As Integer
Global PlayerGMWalkEnbl(6) As Boolean

Sub Create_Player()
PlayerName = "Zaid"
PlayerDirection = "R"
PlayerPosWalk = 1
Dim i As Integer
For i = 0 To 7
PlayerListHealth(i) = 0
PlayerListHealthMax(i) = 10
PlayerListEnergy(i) = 30
PlayerListEnergyMax(i) = 30
PlayerListPower(i) = 1
PlayerListMagic(i) = 1
PlayerListMagicPicture(i) = "Magic"
Next

PlayerListName(0) = "Zaid Markabi"
PlayerListLocked(0) = False
PlayerListHealth(0) = 25
PlayerListHealthMax(0) = 25

PlayerListName(1) = "Yazan Markabi"
PlayerListLocked(1) = True
PlayerListHealth(1) = 25
PlayerListHealthMax(1) = 25

PlayerListName(2) = "Mohammad Sourity"
PlayerListLocked(2) = True
PlayerListHealth(2) = 20
PlayerListHealthMax(2) = 20

PlayerListName(3) = "Sana Sourity"
PlayerListLocked(3) = True
PlayerListHealth(3) = 15
PlayerListHealthMax(3) = 15

PlayerListName(4) = "Akiad abo Kora"
PlayerListLocked(4) = True
PlayerListHealth(4) = 25
PlayerListHealthMax(4) = 25

PlayerListName(5) = "Muddar Tinawy"
PlayerListLocked(5) = True
PlayerListHealth(5) = 40
PlayerListHealthMax(5) = 40

PlayerListName(6) = "Omar Tinawy"
PlayerListLocked(6) = True
PlayerListHealth(6) = 30
PlayerListHealthMax(6) = 30

PlayerListName(7) = "Lynda Tinawy"
PlayerListLocked(7) = True
PlayerListHealth(7) = 25
PlayerListHealthMax(7) = 25

PotionNum = 3
EnergyNum = 2
PlayerMoney = 1200
End Sub

Sub Refresh_Player()
frmMain.PlayerBlock.Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerName + PlayerDirection + Format(PlayerPosWalk) + ".emf")

PlayerPosBlock = (PlayerPosX - MapStartX - 1) + BlockSizeX * (PlayerPosY - MapStartY - 1)

frmMain.PlayerBlock.Left = frmMain.Block(PlayerPosBlock).Left
frmMain.PlayerBlock.Top = frmMain.Block(PlayerPosBlock).Top
End Sub

Sub RePos_Player()
frmMain.PlayerBlock.Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerName + PlayerDirection + Format(PlayerPosWalk) + ".emf")
End Sub

Sub Refresh_Friends()
Dim i As Integer
For i = 0 To 6
frmMain.FriendBlock(i).Picture = LoadPicture(App.Path + "\Data\Players\" + PlayerGMName(i) + PlayerGMDirection(i) + Format(PlayerGMWalk(i)) + ".emf")

frmMain.FriendBlock(i).Left = frmMain.Block(PlayerGMBlock(i)).Left
frmMain.FriendBlock(i).Top = frmMain.Block(PlayerGMBlock(i)).Top
Next
End Sub

