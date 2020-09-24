Attribute VB_Name = "SaveLoad"

Sub Load_Game()
On Error GoTo 1
Dim LdDt As String
Dim i, i2 As Integer

Open App.Path + "\Data\Save\Data1.txt" For Input As #1
Input #1, LdDt
Input #1, LdDt
If Not Len(LdDt) = 23 Then End
Input #1, PlayerName
Input #1, PlayerPosX
Input #1, PlayerPosY
Input #1, PlayerDirection
Input #1, PlayerPosBlock
Input #1, PlayerPosWalk
For i = 0 To 7
Input #1, PlayerListName(i)
Input #1, PlayerListHealth(i)
Input #1, PlayerListEnergy(i)
Input #1, PlayerListHealthMax(i)
Input #1, PlayerListEnergyMax(i)
Input #1, PlayerListPower(i)
Input #1, PlayerListMagic(i)
Input #1, PlayerListMagicPicture(i)
Input #1, PlayerListLocked(i)
Next
Input #1, PlayerListSelected
For i = 0 To 6
Input #1, PlayerGMName(i)
Input #1, PlayerGMDirection(i)
Input #1, PlayerGMBlock(i)
Input #1, PlayerGMWalk(i)
Input #1, PlayerGMWalkEnbl(i)
Next
Input #1, PotionNum
Input #1, EnergyNum
Input #1, PlayerMoney
Input #1, PlayerPoints
Input #1, LdDt
If Not Len(LdDt) = 12 Then End
Input #1, MapSizeX
Input #1, MapSizeY
For i = 0 To MapSizeX + 1
For i2 = 0 To MapSizeY + 1
Input #1, BlockTag(i, i2)
Input #1, BlockMode(i, i2)
Input #1, BlockHave(i, i2)
Input #1, BlockDoorGotoMap(i, i2)
Input #1, BlockDoorGotoPosX(i, i2)
Input #1, BlockDoorGotoPosY(i, i2)
Input #1, BlockCommandsFile(i, i2)
Next
Next
Input #1, BlockSizeX
Input #1, BlockSizeY
Input #1, MapNum
Input #1, MapName
Input #1, MapStartX
Input #1, MapStartY
Input #1, GameMenuPos
Input #1, GameSleep
Input #1, GameTalking
Input #1, GameWar
Input #1, GameSeller
Input #1, GameMenuWarPos
Input #1, GameMenuWarPosEnm
Input #1, GameMenuPos2
Input #1, GameMenuPosBuy
Input #1, GameMenuPosBuy2
Input #1, GameMenuPosBuy3
Input #1, GameMenuPosBuy4
Input #1, CheckPointNum
For i = 0 To CheckPointNum
Input #1, CheckPointID(i)
Input #1, CheckPointValue(i)
Next
Input #1, CheckSourceNum
For i = 0 To CheckSourceNum
Input #1, CheckSourceID(i)
Input #1, CheckSource(i)
Next
Input #1, PotionPrice
Input #1, EnergyPrice
Input #1, MagicPrice
Input #1, MagicName
Input #1, MagicVaule
Input #1, LdDt
If Not Len(LdDt) = 12 Then End
Input #1, EnemyNum
For i = 0 To EnemyNum
Input #1, EnemyName(i)
Input #1, EnemyHelath(i)
Input #1, EnemyAttack(i)
Input #1, EnemyMoney(i)
Input #1, EnemyPoints(i)
Next
Input #1, DamageSelected
Close #1
1:
End Sub

Sub Save_Game()
On Error GoTo 1
Dim i, i2 As Integer

Open App.Path + "\Data\Save\Data1.txt" For Output As #1
Write #1, "Game_Save_Data_32"
Write #1, "Powered by Zaid Markabi"
Write #1, PlayerName
Write #1, PlayerPosX
Write #1, PlayerPosY
Write #1, PlayerDirection
Write #1, PlayerPosBlock
Write #1, PlayerPosWalk
For i = 0 To 7
Write #1, PlayerListName(i)
Write #1, PlayerListHealth(i)
Write #1, PlayerListEnergy(i)
Write #1, PlayerListHealthMax(i)
Write #1, PlayerListEnergyMax(i)
Write #1, PlayerListPower(i)
Write #1, PlayerListMagic(i)
Write #1, PlayerListMagicPicture(i)
Write #1, PlayerListLocked(i)
Next
Write #1, PlayerListSelected
For i = 0 To 6
Write #1, PlayerGMName(i)
Write #1, PlayerGMDirection(i)
Write #1, PlayerGMBlock(i)
Write #1, PlayerGMWalk(i)
Write #1, PlayerGMWalkEnbl(i)
Next
Write #1, PotionNum
Write #1, EnergyNum
Write #1, PlayerMoney
Write #1, PlayerPoints
Write #1, "Zaid Markabi"
Write #1, MapSizeX
Write #1, MapSizeY
For i = 0 To MapSizeX + 1
For i2 = 0 To MapSizeY + 1
Write #1, BlockTag(i, i2)
Write #1, BlockMode(i, i2)
Write #1, BlockHave(i, i2)
Write #1, BlockDoorGotoMap(i, i2)
Write #1, BlockDoorGotoPosX(i, i2)
Write #1, BlockDoorGotoPosY(i, i2)
Write #1, BlockCommandsFile(i, i2)
Next
Next
Write #1, BlockSizeX
Write #1, BlockSizeY
Write #1, MapNum
Write #1, MapName
Write #1, MapStartX
Write #1, MapStartY
Write #1, GameMenuPos
Write #1, GameSleep
Write #1, GameTalking
Write #1, GameWar
Write #1, GameSeller
Write #1, GameMenuWarPos
Write #1, GameMenuWarPosEnm
Write #1, GameMenuPos2
Write #1, GameMenuPosBuy
Write #1, GameMenuPosBuy2
Write #1, GameMenuPosBuy3
Write #1, GameMenuPosBuy4
Write #1, CheckPointNum
For i = 0 To CheckPointNum
Write #1, CheckPointID(i)
Write #1, CheckPointValue(i)
Next
Write #1, CheckSourceNum
For i = 0 To CheckSourceNum
Write #1, CheckSourceID(i)
Write #1, CheckSource(i)
Next
Write #1, PotionPrice
Write #1, EnergyPrice
Write #1, MagicPrice
Write #1, MagicName
Write #1, MagicVaule
Write #1, "Zaid Markabi"
Write #1, EnemyNum
For i = 0 To EnemyNum
Write #1, EnemyName(i)
Write #1, EnemyHelath(i)
Write #1, EnemyAttack(i)
Write #1, EnemyMoney(i)
Write #1, EnemyPoints(i)
Next
Write #1, DamageSelected
Write #1, "#End Data#"
Close #1
1:
End Sub
