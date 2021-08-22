VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FriendlyMonster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public MonsterName As String, MonsterType As Monster2_Types
Public Attack As Single, Speed As Single, During As Single, HP As Single, MaxHP As Single  '魔兽各项能力值
Public CDTime As Long, Spend As Long, info As String, AttackEx As Boolean
Public W As Single, H As Single, LastFireTime As Long, LastSuperTimer As Long
Public X As Single, Y As Single
Public HasFunny As Boolean 'For Burn Monsters (1 items)
Public SpicalPic As String '有时候需要显示一些特别Spical的图片
Public MonsterIndex As Integer
Public NowEMonster As Integer
Public BookPic As String
Public SuperMode As Boolean, SuperTimes As Long
Public NeedWater As Boolean
Public OranSpeed As Single, OranDuring As Single
Public Level As Long
Public Sub UpdateClick()
Select Case MonsterType
    
End Select
End Sub
Public Sub UpdateMove()
Select Case MonsterType
    Case Monster2_Types.Burn '生产类
    Call BurnEffect
    Case Monster2_Types.BurnSuper
    Call BurnSuperEffect
End Select
End Sub
Public Sub UpdateSuper()
Select Case MonsterType
    Case Monster2_Types.Burn
    ReDim Preserve Funnys(UBound(Funnys) + 1)
    Funnys(UBound(Funnys)).X = X
    Funnys(UBound(Funnys)).Y = Y + 36
    Case Monster2_Types.Define
    Ring "get"
    HP = MaxHP
    Case Else
    Call PutFire
End Select
End Sub
Public Sub Update()
If SuperTimes > 10 Then SuperTimes = 0: SuperMode = False
If SuperMode = True And SuperTimes <= 10 Then
Call UpdateSuper
SuperTimes = SuperTimes + 1
Exit Sub
End If

Select Case MonsterType
    Case Monster2_Types.Normal   'Normal
        Call NormalMonster
    Case Monster2_Types.HateFlower
        Call MicaoMonster
    Case Monster2_Types.OhAngle
    
    Case Monster2_Types.BeatOut
        Call NormalMonster
    Case Monster2_Types.Burn
        Call BurnMonster
    Case Monster2_Types.BOOM
        Call BOOMMonster
    Case Monster2_Types.boom2
        Call BOOMMonster
    Case Monster2_Types.GoBack
        Call GoBackMonster
    Case Monster2_Types.Define
        Call DefineMonster
    Case Monster2_Types.boom3
        Call BOOMMonster3
    Case Monster2_Types.CopyHP
        Call CopyHPChooseEMonster
        Call CopyHPMonster
    Case Monster2_Types.BurnSuper
        Call BurnSuperMonster
    Case Monster2_Types.ThreeLine
        Call ThreeLineMonster
    Case Monster2_Types.IceAll
        Call IceMonster2
    Case Monster2_Types.PurpleFunny
        Call PurpleMonster
    Case Monster2_Types.NoFar
        Call NoFarMonster
    Case Monster2_Types.DarkNess
        Call NightMonster
    Case Monster2_Types.FireAll
        Call FireAllMonster
    Case Monster2_Types.CanMove
        Call MoveMonsteryee
    Case Monster2_Types.cloud
    Case Monster2_Types.FunnyAttack
        Call FunnyAttackMonster
    Case Monster2_Types.WaterAll
        Call WaterAllMonster
    Case Monster2_Types.Mirror
        Call MirrorMonster
    Case Else
        Call NormalMonster
End Select
End Sub
Public Sub LoadMonster(MonsterInformation As String, Index As Integer)  '初始化魔兽
Attack = Val(ReadINI("Monster", "Attack", App.Path & "\monster\fmonster\" & MonsterInformation & ".ini")) * (1 + Level / 10)
Speed = Val(ReadINI("Monster", "Speed", App.Path & "\monster\fmonster\" & MonsterInformation & ".ini")) * (1 + Level / 10)
During = Val(ReadINI("Monster", "During", App.Path & "\monster\fmonster\" & MonsterInformation & ".ini")) / (1 + Level / 10)
HP = Val(ReadINI("Monster", "HP", App.Path & "\monster\fmonster\" & MonsterInformation & ".ini")) * (1 + Level / 10)
Spend = Val(ReadINI("Monster", "Spend", App.Path & "\monster\fmonster\" & MonsterInformation & ".ini"))
CDTime = Val(ReadINI("Monster", "CD", App.Path & "\monster\fmonster\" & MonsterInformation & ".ini"))
info = ReadINI("Monster", "Info", App.Path & "\monster\fmonster\" & MonsterInformation & ".ini")
BookPic = ReadINI("Monster", "BookPic", App.Path & "\monster\fmonster\" & MonsterInformation & ".ini")
MonsterType = Val(ReadINI("Monster", "Type", App.Path & "\monster\fmonster\" & MonsterInformation & ".ini"))
MonsterName = MonsterInformation
OranSpeed = Speed
OranDuring = During
MaxHP = HP:  MonsterIndex = Index
LastFireTime = GetTickCount - 1500
End Sub
Private Sub MoveMonsteryee()

temp = Split(EMonsterOnLine((Y - FightWindow.Friends(0).top) / 76), ";")
Dim Shortest As Single, ShortIndex As Integer
Shortest = 1E+17
For i = 0 To UBound(temp)
    If Val(temp(i)) <> 0 And Val(temp(i)) <= UBound(EMonster) Then
        If Abs(X - EMonster(temp(i)).X) < Shortest Then
        Shortest = Abs(X - EMonster(temp(i)).X)
        ShortIndex = temp(i)
        End If
    End If
Next

    If ShortIndex <= UBound(EMonster) Then
        If Abs(X - EMonster(ShortIndex).X) <= 150 Then
        X = X - 150
        Exit Sub '逃跑要紧 兄弟
        End If
        If EMonster(ShortIndex).X > X And EMonster(ShortIndex).X - X > 304 Then
        X = X + 76
        End If
        If EMonster(ShortIndex).X < X And X - EMonster(ShortIndex).X > 304 Then
        X = X - 76
        End If
        
        If (EMonster(ShortIndex).X >= X And GetTickCount - LastFireTime >= During) Then
        Call PutFire
        End If
        
    End If

End Sub
Private Sub CopyHPChooseEMonster()
If UBound(EMonster) > 0 Then
Randomize
If NowEMonster > UBound(EMonster) Or NowEMonster = 0 Then NowEMonster = Int(Rnd * (UBound(EMonster) - 1) + 1)
End If
End Sub
Private Sub CopyHPMonster()
If GetTickCount - LastFireTime >= During Then
    LastFireTime = GetTickCount
    If UBound(EMonster) > 0 Then
    If NowEMonster <= UBound(EMonster) Then
    EMonster(NowEMonster).DarkMode = True
    Ring MonsterName
    EMonster(NowEMonster).HP = EMonster(NowEMonster).HP - Attack
    
    If EMonster(NowEMonster).HP <= 0.001 Then
    Randomize
    If Int(Rnd * 100) < 10 Then NewPresent "g 10", EMonster(NowEMonster).X, EMonster(NowEMonster).Y
    If Int(Rnd * 100) < 7 Then NewPresent "g 20", EMonster(NowEMonster).X, EMonster(NowEMonster).Y
    If Int(Rnd * 100) < 10 Then NewPresent "f 1", EMonster(NowEMonster).X, EMonster(NowEMonster).Y
    Set EMonster(NowEMonster) = EMonster(UBound(EMonster))
    EMonster(NowEMonster).MonsterIndex = NowEMonster
    ReDim Preserve EMonster(UBound(EMonster) - 1)
    Randomize
    NowEMonster = Int(Rnd * (UBound(EMonster) - 1) + 1)
    End If
        
        If UBound(FMonster) > 0 Then
        For i = 1 To UBound(FMonster)
        If i <= UBound(FMonster) Then
        If i <> MonsterIndex And FMonster(i).HP < FMonster(i).MaxHP And FMonster(i).MonsterType <> CopyHP Then
        FMonster(i).HP = FMonster(i).HP + Attack
        If FMonster(i).HP > FMonster(i).MaxHP Then FMonster(i).HP = FMonster(i).MaxHP
        End If
        End If
        Next
        End If
    End If
    End If
End If
End Sub
Private Sub FunnyAttackMonster()
If GetTickCount - LastFireTime >= 20000 Then
ReDim Preserve Funnys(UBound(Funnys) + 1)
With Funnys(UBound(Funnys))
   .X = X
   .Y = Y
End With
LastFireTime = GetTickCount
End If
Call NormalMonster
End Sub
Private Sub BurnEffect()
If HasFunny = True Then
HasFunny = False
FunnyCounts = FunnyCounts + 1
LastFireTime = GetTickCount - 1500
Ring "get"
End If
End Sub
Private Sub NightMonster()
If GetTickCount - LastFireTime >= During Then
Ring MonsterName
FightWindow.SpicalBack = ""
MakeFMonsterDead MonsterIndex
End If
End Sub
Private Sub BurnMonster()
If GetTickCount - LastFireTime >= During Then
HasFunny = True
End If
If HasFunny = True Then LastFireTime = GetTickCount
End Sub
Private Sub BurnSuperEffect()
If HasFunny = True Then
HasFunny = False
FightWindow.SuperCounts = FightWindow.SuperCounts + 1
LastFireTime = GetTickCount - 1500
Ring "get"
End If
End Sub
Private Sub BurnSuperMonster()
If GetTickCount - LastFireTime >= During Then
HasFunny = True
End If
If HasFunny = True Then LastFireTime = GetTickCount
End Sub
Private Sub MicaoMonster()
If GetTickCount - LastFireTime >= During Then
temp = Split(EMonsterOnLine((Y - FightWindow.Friends(0).top) / 76), ";")
For i = 0 To UBound(temp)
    If (Val(temp(i)) <> 0 And Val(temp(i)) <= UBound(EMonster)) Then
        If EMonster(temp(i)).X > X Then
        LastFireTime2 = LastFireTime
        Call PutFire
            If EMonster(temp(i)).MonsterName <> "魔花" And EMonster(temp(i)).MonsterName <> "龙花" And EMonster(temp(i)).MonsterName <> "幽灵之花" Then
            SpicalPic = ""
            LastFireTime = GetTickCount
            Else
            If SpicalPic = "" Then Ring "super"
            SpicalPic = "谜草Crazy"
            LastFireTime = LastFireTime2
            End If
        Exit For
        End If
    End If
Next

End If
End Sub
Private Sub GoBackMonster()
If GetTickCount - LastFireTime >= During Then
        During = 99999999
        Ring MonsterName
        SpicalPic = MonsterName & "Attack"
        If UBound(EMonster) > 0 Then
        For i = 1 To UBound(EMonster)
        Ring "冰雪女王"
        If i > UBound(EMonster) Then Exit For
        EMonster(i).IceTime = GetTickCount
        'Wait 10
        Ring "法师"
        If i > UBound(EMonster) Then Exit For
        EMonster(i).X = FightWindow.ScaleWidth - 36
        'Wait 10
        Next
        End If
        MakeFMonsterDead MonsterIndex
End If
End Sub
Private Sub PurpleMonster()
If GetTickCount - LastFireTime >= During Then
        During = 99999999
        Ring MonsterName
        SpicalPic = MonsterName & "0"
        If UBound(FMonster) > 0 Then
        For i = 1 To UBound(FMonster)
        If i > UBound(FMonster) Then Exit For
        FMonster(i).Attack = FMonster(i).Attack * 1.5
        FMonster(i).AttackEx = True
        Next
        End If
        MakeFMonsterDead MonsterIndex
End If
End Sub
Private Sub IceMonster2()
If GetTickCount - LastFireTime >= During Then
        During = 99999999
        Ring MonsterName
        SpicalPic = MonsterName & "0"
        If UBound(EMonster) > 0 Then
        For i = 1 To UBound(EMonster)
        If i > UBound(EMonster) Then Exit For
        EMonster(i).IceTime = GetTickCount
        Next
        End If
        MakeFMonsterDead MonsterIndex
End If
End Sub
Private Sub WaterAllMonster()
If GetTickCount - LastFireTime >= During Then
        During = 99999999
        Ring MonsterName
        SpicalPic = MonsterName & "0"
        If UBound(FMonster) > 0 Then
        For i = 1 To UBound(FMonster)
        If i > UBound(FMonster) Then Exit For
        If FMonster(i).NeedWater = True Then
        FMonster(i).NeedWater = False
        FMonster(i).Speed = FMonster(i).Speed * 2
        FMonster(i).During = FMonster(i).During / 2
        End If
        Next
        End If
        MakeFMonsterDead MonsterIndex
End If
End Sub
Private Sub BOOMMonster3()
If GetTickCount - LastFireTime >= During Then
        During = 99999999
        Ring MonsterName
        SpicalPic = MonsterName & "0"
        If UBound(EMonster) > 0 Then
        For i = 1 To UBound(EMonster)
        If i > UBound(EMonster) Then Exit For
        FightWindow.NewFire MonsterIndex, EMonster(i).X - 144, EMonster(i).Y
        Next
        End If
        MakeFMonsterDead MonsterIndex
End If
End Sub
Private Sub FireAllMonster()
If GetTickCount - LastFireTime >= During Then
        During = 99999999
        Ring MonsterName
        SpicalPic = MonsterName & "0"
        If UBound(EMonster) > 0 Then
        For i = 1 To UBound(EMonster)
        If i > UBound(EMonster) Then Exit For
        EMonster(i).FireTime = GetTickCount
        Next
        End If
        MakeFMonsterDead MonsterIndex
End If
End Sub
Private Sub DefineMonster()
Select Case HP
    Case Is > MaxHP * 0.7
    SpicalPic = MonsterName & "0"
    Case Is > MaxHP * 0.4
    SpicalPic = MonsterName & "1"
    Case Else
    SpicalPic = MonsterName & "2"
End Select
End Sub
Private Sub BOOMMonster()
If GetTickCount - LastFireTime >= During Then
        Ring MonsterName
        Call FightWindow.NewFire(MonsterIndex, -1, -1)
        SpicalPic = MonsterName & "Attack"
        During = 99999999
        MakeFMonsterDead MonsterIndex
End If
End Sub
Private Sub NoFarMonster()
If GetTickCount - LastFireTime >= During Then
temp = Split(EMonsterOnLine((Y - FightWindow.Friends(0).top) / 76), ";")
For i = 0 To UBound(temp)
    If Val(temp(i)) <> 0 And Val(temp(i)) <= UBound(EMonster) Then
        If (EMonster(temp(i)).X > X And EMonster(temp(i)).X - X <= 180) Then
        Call PutFire
        Exit For
        End If
    End If
Next

End If
End Sub
Private Sub MirrorMonster()
If GetTickCount - LastFireTime >= During Then
temp = Split(EMonsterOnLine((Y - FightWindow.Friends(0).top) / 76), ";")
For i = 0 To UBound(temp)
    If Val(temp(i)) <> 0 And Val(temp(i)) <= UBound(EMonster) Then
        If EMonster(temp(i)).X < X Then
        Call PutFire
        Exit For
        End If
    End If
Next

End If
End Sub
Private Sub NormalMonster()
If GetTickCount - LastFireTime >= During Then
temp = Split(EMonsterOnLine((Y - FightWindow.Friends(0).top) / 76), ";")
For i = 0 To UBound(temp)
    If Val(temp(i)) <> 0 And Val(temp(i)) <= UBound(EMonster) Then
        If EMonster(temp(i)).X > X Then
        Call PutFire
        Exit For
        End If
    End If
Next

End If
End Sub
Sub PutFire()
Ring MonsterName
Call FightWindow.NewFire(MonsterIndex, -1, -1)
LastFireTime = GetTickCount
End Sub
Private Sub ThreeLineMonster()
If GetTickCount - LastFireTime >= During Then
temp = Split(EMonsterOnLine((Y - FightWindow.Friends(0).top) / 76), ";")

For i = 0 To UBound(temp)
    If Val(temp(i)) <> 0 And Val(temp(i)) <= UBound(EMonster) Then
        If EMonster(temp(i)).X > X Then
        Call PutFire
        Exit For
        End If
    End If
Next

If (Y - 76 - FightWindow.Friends(0).top) / 76 >= 0 Then
temp = Split(EMonsterOnLine((Y - 76 - FightWindow.Friends(0).top) / 76), ";")
For i = 0 To UBound(temp)
    If Val(temp(i)) <> 0 And Val(temp(i)) <= UBound(EMonster) Then
        If EMonster(temp(i)).X > X Then
        Call PutFire
        Exit For
        End If
    End If
Next
End If

If (Y + 76 - FightWindow.Friends(0).top) / 76 <= 4 Then
temp = Split(EMonsterOnLine((Y + 76 - FightWindow.Friends(0).top) / 76), ";")
For i = 0 To UBound(temp)
    If Val(temp(i)) <> 0 And Val(temp(i)) <= UBound(EMonster) Then
        If EMonster(temp(i)).X > X Then
        Call PutFire
        Exit For
        End If
    End If
Next
End If

End If
End Sub

