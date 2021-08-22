VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EvilMonster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MonsterName As String, MonsterType As Monster2_Types
Public Attack As Single, Speed As Single, MaxSpeed As Single, HP As Single, MaxHP As Single '魔兽各项能力值
Public info As String
Public X As Single, Y As Single, W As Single, H As Single
Public Eating As Boolean
Public MonsterIndex As Integer
Public LastCallTime As Long
Public IceTime As Long
Public ChemicalMode As Boolean
Public DarkMode As Boolean
Public FireTime As Long
Public BookPic As String
Public ThunderMode As Boolean
Public OranSpeed As Single
Public Sub UpdateMove()
Randomize
If MonsterType = Fly2 And Int(Rnd * 100) < 10 Then X = X - 76
If Eating = False And GetTickCount - IceTime >= 5000 Then X = X - Speed
If FightWindow.SpicalBack = "Back4-day" Then X = X - Speed
End Sub
Public Sub UpdateEffect()
Select Case MonsterType
    Case Monster2_Types.Normal  'Normal
        Call NormalMonster
    Case Monster2_Types.ghost
    
    Case Monster2_Types.NewANew
        Call CallMonster
        Call NormalMonster
    Case Monster2_Types.CopyHP
        Call CopyHPMonster
    Case Monster2_Types.MoveMon
        Call MoveMonster
    Case Monster2_Types.ThrowAround
        Call ThrowAroundMonster
    Case Else
        Call NormalMonster
End Select
Randomize
If ChemicalMode = True Then HP = HP - 8
If GetTickCount - FireTime <= 5000 Then HP = HP - 10
If MonsterType = ReSelf Then HP = HP + 5

If ThunderMode = True Then
ThunderMode = False
HP = HP - 8
    If Int(Rnd * 100) < 3 Then
        Ring "Thunder"
        Dim EMIndex As Integer
        EMIndex = MonsterIndex
        If Int(Rnd * 100) < 10 Then NewPresent "g 10", EMonster(EMIndex).X, EMonster(EMIndex).Y
        If Int(Rnd * 100) < 7 Then NewPresent "g 20", EMonster(EMIndex).X, EMonster(EMIndex).Y
        If Int(Rnd * 100) < 10 Then NewPresent "f 1", EMonster(EMIndex).X, EMonster(EMIndex).Y
        Set EMonster(EMIndex) = EMonster(UBound(EMonster))
        EMonster(EMIndex).MonsterIndex = EMIndex
        ReDim Preserve EMonster(UBound(EMonster) - 1)
    End If
End If
End Sub
Public Sub LoadMonster(MonsterInformation As String, Index As Integer)  '初始化魔兽
If MonsterInformation = "深空之莲" Then MonsterInformation = "深空莲花"
Attack = Val(ReadINI("Monster", "Attack", App.Path & "\monster\emonster\" & MonsterInformation & ".ini"))
Speed = Val(ReadINI("Monster", "Speed", App.Path & "\monster\emonster\" & MonsterInformation & ".ini"))
HP = Val(ReadINI("Monster", "HP", App.Path & "\monster\emonster\" & MonsterInformation & ".ini"))
MonsterType = Val(ReadINI("Monster", "Type", App.Path & "\monster\emonster\" & MonsterInformation & ".ini"))
BookPic = ReadINI("Monster", "BookPic", App.Path & "\monster\emonster\" & MonsterInformation & ".ini")
info = ReadINI("Monster", "Info", App.Path & "\monster\emonster\" & MonsterInformation & ".ini")
H = ReadINI("Monster", "Height", App.Path & "\monster\emonster\" & MonsterInformation & ".ini")
MonsterName = MonsterInformation
MaxSpeed = Speed
OranSpeed = Speed
MaxHP = HP:  MonsterIndex = Index
End Sub
Private Sub CallMonster()
If GetTickCount - LastCallTime >= 10000 Then
LastCallTime = GetTickCount
NewEMonster "魔蝠", X + 73, Y
End If
End Sub
Private Sub NormalMonster()
temp = Split(FMonsterOnLine((Y - FightWindow.Friends(0).top) / 76), ";")
Eating = False
For i = 0 To UBound(temp)
If temp(i) <> "" Then
    If Val(temp(i)) <= UBound(FMonster) Then
    If Abs(FMonster(temp(i)).X - X) <= 73 Then
    Eating = True
    FMonster(temp(i)).HP = FMonster(temp(i)).HP - Attack
    If MonsterType = Becomess Then
    NewEMonster MonsterName, FMonster(temp(i)).X, FMonster(temp(i)).Y
    MakeFMonsterDead temp(i)
    End If
    If MonsterType = CopyHP And HP + Attack <= MaxHP Then HP = HP + Attack
    If FMonster(temp(i)).HP <= 0.001 Then MakeFMonsterDead temp(i)  'Dead
    Exit Sub
    End If
    End If
End If
Next
End Sub
Private Sub ThrowAroundMonster()
temp = Split(FMonsterOnLine((Y - FightWindow.Friends(0).top) / 76), ";")
Eating = False
For i = 0 To UBound(temp)
If temp(i) <> "" Then
    If Val(temp(i)) <= UBound(FMonster) Then
    If Abs(FMonster(temp(i)).X - X) <= 73 Then
    Eating = True
    FMonster(temp(i)).HP = FMonster(temp(i)).HP - Attack
    Randomize
    FMonster(temp(i)).X = FMonster(temp(i)).X + Int(Rnd * (Dad.ScaleWidth - FMonster(temp(i)).X))
    If FMonster(temp(i)).HP <= 0.001 Then MakeFMonsterDead temp(i)  'Dead
    Exit Sub
    End If
    End If
End If
Next
End Sub
Private Sub MoveMonster()
temp = Split(FMonsterOnLine((Y - FightWindow.Friends(0).top) / 76), ";")
Eating = False
For i = 0 To UBound(temp)
If temp(i) <> "" Then
    If Val(temp(i)) <= UBound(FMonster) Then
    If Abs(FMonster(temp(i)).X - X) <= 73 Then
    Eating = True
    FMonster(temp(i)).HP = FMonster(temp(i)).HP - Attack
    FMonster(temp(i)).X = FMonster(temp(i)).X + 76
    If FMonster(temp(i)).HP <= 0.001 Then MakeFMonsterDead temp(i)  'Dead
    Exit Sub
    End If
    End If
End If
Next
End Sub
Private Sub CopyHPMonster()
temp = Split(FMonsterOnLine((Y - FightWindow.Friends(0).top) / 76), ";")
Eating = False
For i = 0 To UBound(temp)
If temp(i) <> "" Then
    If Val(temp(i)) <= UBound(FMonster) Then
    If Abs(FMonster(temp(i)).X - X) <= 73 Then
    Eating = True
    FMonster(temp(i)).HP = FMonster(temp(i)).HP - Attack
    HP = HP + Attack
    If MonsterType = CopyHP And HP + Attack <= MaxHP Then HP = HP + Attack
    If FMonster(temp(i)).HP <= 0.001 Then MakeFMonsterDead temp(i)  'Dead
    Exit Sub
    End If
    End If
End If
Next
End Sub

