VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LevelManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public World As Long, Level As Long
Public Step As Long, IsInProgress As Boolean
Public LastDayTime As Long
Public NewMsg As String
Function CanNew(ByVal X As Single, ByVal Y As Single, NewMType As Monster2_Types) As Boolean
Select Case World
    Case 0
    CanNew = NormalNew(X, Y, NewMType)
    Case 1
    CanNew = NormalNew(X, Y, NewMType)
    Case 2
    CanNew = SkyNew(X, Y, NewMType)
    Case 3
    CanNew = NormalNew(X, Y, NewMType)
    Case 4
    CanNew = NormalNew(X, Y, NewMType)
End Select
End Function
Function NormalNew(ByVal X As Single, ByVal Y As Single, NewMType As Monster2_Types) As Boolean
    NormalNew = True
    For i = 0 To UBound(FMonster)
    If FMonster(i).X = X And FMonster(i).Y = Y Then NormalNew = False: NewMsg = "你不能把它丢在" & FMonster(i).MonsterName & "身上啦！": Exit For
    Next
End Function
Function SkyNew(ByVal X As Single, ByVal Y As Single, NewMType As Monster2_Types) As Boolean
    For i = 0 To UBound(FMonster)
    If FMonster(i).X = X And FMonster(i).Y = Y And FMonster(i).MonsterType = cloud Then SkyNew = True: Exit For
    Next
    If SkyNew = False And NewMType <> fly And NewMType <> cloud Then NewMsg = "你的魔兽可不会飞，你要先放置云才行。": Exit Function
    If NewMType = fly Or NewMType = cloud Then SkyNew = True
    For i = 0 To UBound(FMonster)
    If FMonster(i).X = X And FMonster(i).Y = Y Then
        If FMonster(i).MonsterType <> cloud Then
        SkyNew = False: NewMsg = "你不能把它丢在" & FMonster(i).MonsterName & "身上啦！": Exit For
        ElseIf FMonster(i).MonsterType = cloud And NewMType = cloud Then
        SkyNew = False: NewMsg = "你不可以重叠放置" & FMonster(i).MonsterName & "啦！": Exit For
        End If
    End If
    Next
End Function
Sub Update(New_IsInProgress As Boolean)
IsInProgress = New_IsInProgress
Select Case World
    Case 0 '草原
    Call Grass
    Case 1 '夜晚
    Call Night
    Case 2 '天空
    Call Sky
    Case 3 '遗迹
    Call LostSpace
    Case 4 '沙漠
    Call Desert
End Select
End Sub
Sub Desert()
Select Case Level
    Case 0
    Call FirstDesert
End Select
End Sub
Sub Grass()
Select Case Level
    Case 0
    Call NewcomerLevel
    Case 7
    Call MonsterFlowerLevel
    Case 8
    Call MonsterFlowerLevel2
End Select
If Level >= 9 Then Call GrassEffect
End Sub
Sub Night()
Select Case Level
    Case 0
    Call FirstNight
    Case 1
    Call Night2
End Select
Call NightEffect
End Sub
Sub Sky()
Select Case Level
    Case 0
    Call FirstSky
    Case Else
    Call EverySky
End Select
End Sub
Sub LostSpace()
Select Case Level
    Case 0
    Call FirstLostSpace
End Select
Call LostSpaceEffect
End Sub
Sub LostSpaceEffect()
Randomize
If Int(Rnd * 100) < 7 And FightWindow.NowProgress <= FightWindow.MaxProgress Then
    If FightWindow.SpicalBack = "" Then
    NewEMonster "暗黑之眼", FightWindow.ScaleWidth * 0.6 + Int(Rnd * FightWindow.ScaleWidth * 0.2), FightWindow.Friends(Int(Rnd * 5) * 9).top
    Else
    NewEMonster "鬼炎", FightWindow.ScaleWidth * 0.6 + Int(Rnd * FightWindow.ScaleWidth * 0.2), FightWindow.Friends(Int(Rnd * 5) * 9).top
    End If
End If

If IsInProgress = True And FightWindow.BigProgress > 0 Then
    If FightWindow.NowProgress Mod FightWindow.BigProgress = 0 Then
        FightWindow.SpicalBack = "Back4-day"
        LastDayTime = GetTickCount
    End If
End If

If GetTickCount - LastDayTime >= 35000 Then FightWindow.SpicalBack = ""
End Sub
Sub NightEffect()
Randomize
If Int(Rnd * 100) < 7 And FightWindow.NowProgress <= FightWindow.MaxProgress Then
NewEMonster "电视", FightWindow.ScaleWidth * 0.6 + Int(Rnd * FightWindow.ScaleWidth * 0.2), FightWindow.Friends(Int(Rnd * 5) * 9).top
End If

If FightWindow.BigProgress <> 0 And IsInProgress = True Then
    If FightWindow.NowProgress Mod FightWindow.BigProgress = 0 And UBound(EMonster) > 0 Then
        For i = 1 To UBound(EMonster)
            If EMonster(i).MonsterType = TV Then
            NewEMonster "贞子", EMonster(i).X - 36, EMonster(i).Y
            End If
        Next
        If Level >= 5 Then
            For i = 1 To 5
                NewEMonster "蜘蛛怪", FightWindow.ScaleWidth * 0.6 + Int(Rnd * FightWindow.ScaleWidth * 0.2), FightWindow.Friends(Int(Rnd * 5) * 9).top
            Next
        End If
    End If
End If
End Sub
Sub GrassEffect()
If FightWindow.BigProgress <> 0 And IsInProgress = True Then
If FightWindow.NowProgress Mod FightWindow.BigProgress = 0 Then
For i = 1 To 5
NewEMonster "食人花", FightWindow.ScaleWidth * 0.4 + Int(Rnd * FightWindow.ScaleWidth * 0.3), FightWindow.Friends(Int(Rnd * 5) * 9).top
Next
End If
End If
End Sub
Sub Night2()
Select Case Step
    Case 0
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "黑嘴"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "冰棍"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "冰棍"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "黑嘴"
    FightWindow.NewFaceMsg "冰棍", "我New出了一个新的我，加入你们的战斗，拜。"
    Step = Step + 1
End Select
If Step <= 0 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub EverySky()
Select Case Step
    Case 0
    For i = 0 To 2
    For s = 0 To 36 Step 9
    NewFMonster FightWindow.Friends(i).Left, FightWindow.Friends(s).top, "云"
    Next
    Next
    Step = Step + 1
End Select
End Sub
Sub FirstSky()
Select Case Step
    Case 0
    For i = 0 To 3
    For s = 0 To 36 Step 9
    NewFMonster FightWindow.Friends(s).Left, FightWindow.Friends(i).top, "云"
    Next
    Next
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "食人鱼"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "食人鱼"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "食人鱼"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "食人鱼"
    FightWindow.NewFaceMsg "冰棍", "在天上敲代码总行了吧！！！呃....好像不行..."
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "冰棍", "对了，送你一朵云。"
    Step = Step + 1
    Case 2
    If OwnMonster("云") = False Then GetPresent "m 云"
    Step = Step + 1
End Select
If Step <= 1 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub FirstDesert()
Select Case Step
    Case 0
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "仙人掌"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "仙人掌"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "仙人掌"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "仙人掌"
    FightWindow.NewFaceMsg "冰棍", "好热啊....好热啊...好想吃冰棍啊...."
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "冰棍", "对了......我在这...热热...热.......里发现了四只仙人掌魔兽......"
    Step = Step + 1
    Case 2
    FightWindow.NewFaceMsg "冰棍", "还有.....四个.....热............热热热....滑稽果......."
    Ring "get"
    FunnyCounts = FunnyCounts + 4
    Step = Step + 1
    Case 3
    FightWindow.NewFaceMsg "冰棍", "我去吃冰棍了...........再见.......好热...."
    Step = Step + 1
End Select
If Step <= 1 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub FirstLostSpace()
Select Case Step
    Case 0
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "绯红之莲"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "绯红之莲"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "绯红之莲"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "绯红之莲"
    FightWindow.NewFaceMsg "冰棍", "又到晚上了，不过我很庆幸现在是晚上。"
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "冰棍", "为什么呢。因为夕阳下的敌对魔兽们都异常兴奋..."
    Step = Step + 1
    Case 2
    FightWindow.NewFaceMsg "冰棍", "帮你New好了四只绯红之莲，哦还有，给你四个滑稽果"
    Ring "get"
    FunnyCounts = FunnyCounts + 4
    Step = Step + 1
    Case 3
    FightWindow.NewFaceMsg "404", "受死吧喜欢按Ctrl+S的程序员诶！！"
    Step = Step + 1
End Select
If Step <= 1 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub FirstNight()
Select Case Step
    Case 0
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "黑嘴"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "黑嘴"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "黑嘴"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "黑嘴"
    FightWindow.NewFaceMsg "冰棍", "啊晚上了...404刚才让贞子给我打了个电话，妈的吓死我了..."
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "冰棍", "我...我要去敲代码冷静冷静，对了，你应该不怕鬼的，对吧？..."
    Step = Step + 1
End Select
If Step <= 1 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub MonsterFlowerLevel2()
Select Case Step
    Case 0
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "谜草"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "谜草"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "谜草"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "谜草"
    FightWindow.NewFaceMsg "冰棍", "404写了一个程序循环New魔花，我要疯了！！！"
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "冰棍", "不多说，这10个滑稽果给你！！！我疯了！！！"
    Ring "get"
    FunnyCounts = FunnyCounts + 10
    Step = Step + 1
End Select
If Step <= 1 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub MonsterFlowerLevel()
Select Case Step
    Case 0
    FightWindow.NewFaceMsg "冰棍", "啊，404疯了，他又给我打了电话，说昨天New了一大堆魔花，说要我去死。MMP。"
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "冰棍", "关于照片，我已经发给你了，你也看到了吧，一大堆的魔花。"
    Step = Step + 1
    Case 2
    FightWindow.NewFaceMsg "冰棍", "对了，谜草最讨厌魔花了，你可以让他试试。"
    Step = Step + 1
    Case 3
    FightWindow.NewFaceMsg "冰棍", "救助：6个滑稽果~"
    Ring "get"
    FunnyCounts = FunnyCounts + 6
    Step = Step + 1
End Select
If Step <= 3 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub NewcomerLevel()
Select Case Step
    Case 0
    FightWindow.SetCard 0, "滑稽之花", 0
    FightWindow.SetCard 1, "骑士", 0
    FightWindow.NewFaceMsg "冰棍", "有一天，404在敲代码...突然停电了。" & vbCrLf & "于是404发疯了，他召集了各种各样的魔兽，发誓要杀死喜欢按Ctrl + S的程序员。" & vbCrLf & "然后..."
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "404", "受死吧！喜欢按Ctrl+S的程序员内~~~！" & vbCrLf & vbCrLf & "游戏制作：Error 404 (QQ1361778219)，素材&音效：RPG MV&网络，音乐：滚动的天空（猎豹移动公司）&网络，图形加速：方程"
    Step = Step + 1
    Case 2
    FightWindow.NewFaceMsg "冰棍", "mmp说曹操曹操就到。话说你刚才念的那一堆是啥？算了不管他。我们说我们的吧。"
    Step = Step + 1
    Case 3
    FightWindow.NewFaceMsg "冰棍", "点击左边的魔兽，我的程序就会New出一个新的魔兽为你战斗啦，不过要足够的滑稽果才可以提供足够的能量让我的程序可以New出一个新的魔兽。"
    Step = Step + 1
    Case 4
    FightWindow.NewFaceMsg "冰棍", "开局的时候我会给你New出4个滑稽果，滑稽果到处都有，可能会从天上飘下来，滑稽果可不会停留，它会一直往下掉，开局的四个滑稽果可别浪费了，得到滑稽果也是需要一些时间的，你应该先聘请一些生产滑稽果类型的魔兽。"
    Step = Step + 1
    Case 5
    FightWindow.NewFaceMsg "冰棍", "遇到突发情况，你可以点击关卡进度条左边的 “滑稽能量”，给某一个魔兽，让这个魔兽发动必杀技。"
    Step = Step + 1
    Case 6
    FightWindow.NewFaceMsg "冰棍", "其它的应该不用我教你也会了吧。那我就先去敲代码了，你们加油，别让404把我绑走了。"
    Step = Step + 1
    Case 7
    FightWindow.NewFaceMsg "404", "hhh没错我就是这个游戏的 制作人 ，傻逼404。加油吧~"
    Step = Step + 1
    Case 8
    FightWindow.NewFaceMsg "冰棍", "啊，对了，你可以点击右上角的红色按钮“解雇”删除魔兽，点击另一个黑色按钮则可以暂停游戏。"
    Step = Step + 1
    Case 9
    FightWindow.NewFaceMsg "冰棍", "点击旁边的HP按钮可以查看魔兽的血量（打印文字耗性能，且有些用户不喜欢。）"
    Step = Step + 1
End Select
If Step <= 9 Then FightWindow.LastProgress = GetTickCount
End Sub
