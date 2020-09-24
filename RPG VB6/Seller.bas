Attribute VB_Name = "Seller"

Global PotionPrice As Integer
Global EnergyPrice As Integer

Global MagicPrice As Integer
Global MagicName As String
Global MagicVaule As Integer

Sub Load_Seller(ID As String)
frmMain.SellerBox.Visible = True

Open App.Path + "\Data\Weapons\Seller" + ID + ".txt" For Input As 6
Input #6, PotionPrice
Input #6, EnergyPrice
Input #6, MagicPrice
Input #6, MagicName
Input #6, MagicVaule

Close #6

frmMain.MagicEffect.Picture = LoadPicture(App.Path + "\Data\Weapons\" + MagicName + ".emf")
frmMain.AtkMagic.Caption = "Attack : " + Format(MagicVaule)
frmMain.PrcMagic.Caption = "Price : " + Format(MagicPrice) + " $"
End Sub
