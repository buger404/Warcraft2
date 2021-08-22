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
    If FMonster(i).X = X And FMonster(i).Y = Y Then NormalNew = False: NewMsg = "�㲻�ܰ�������" & FMonster(i).MonsterName & "��������": Exit For
    Next
End Function
Function SkyNew(ByVal X As Single, ByVal Y As Single, NewMType As Monster2_Types) As Boolean
    For i = 0 To UBound(FMonster)
    If FMonster(i).X = X And FMonster(i).Y = Y And FMonster(i).MonsterType = cloud Then SkyNew = True: Exit For
    Next
    If SkyNew = False And NewMType <> fly And NewMType <> cloud Then NewMsg = "���ħ�޿ɲ���ɣ���Ҫ�ȷ����Ʋ��С�": Exit Function
    If NewMType = fly Or NewMType = cloud Then SkyNew = True
    For i = 0 To UBound(FMonster)
    If FMonster(i).X = X And FMonster(i).Y = Y Then
        If FMonster(i).MonsterType <> cloud Then
        SkyNew = False: NewMsg = "�㲻�ܰ�������" & FMonster(i).MonsterName & "��������": Exit For
        ElseIf FMonster(i).MonsterType = cloud And NewMType = cloud Then
        SkyNew = False: NewMsg = "�㲻�����ص�����" & FMonster(i).MonsterName & "����": Exit For
        End If
    End If
    Next
End Function
Sub Update(New_IsInProgress As Boolean)
IsInProgress = New_IsInProgress
Select Case World
    Case 0 '��ԭ
    Call Grass
    Case 1 'ҹ��
    Call Night
    Case 2 '���
    Call Sky
    Case 3 '�ż�
    Call LostSpace
    Case 4 'ɳĮ
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
    NewEMonster "����֮��", FightWindow.ScaleWidth * 0.6 + Int(Rnd * FightWindow.ScaleWidth * 0.2), FightWindow.Friends(Int(Rnd * 5) * 9).top
    Else
    NewEMonster "����", FightWindow.ScaleWidth * 0.6 + Int(Rnd * FightWindow.ScaleWidth * 0.2), FightWindow.Friends(Int(Rnd * 5) * 9).top
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
NewEMonster "����", FightWindow.ScaleWidth * 0.6 + Int(Rnd * FightWindow.ScaleWidth * 0.2), FightWindow.Friends(Int(Rnd * 5) * 9).top
End If

If FightWindow.BigProgress <> 0 And IsInProgress = True Then
    If FightWindow.NowProgress Mod FightWindow.BigProgress = 0 And UBound(EMonster) > 0 Then
        For i = 1 To UBound(EMonster)
            If EMonster(i).MonsterType = TV Then
            NewEMonster "����", EMonster(i).X - 36, EMonster(i).Y
            End If
        Next
        If Level >= 5 Then
            For i = 1 To 5
                NewEMonster "֩���", FightWindow.ScaleWidth * 0.6 + Int(Rnd * FightWindow.ScaleWidth * 0.2), FightWindow.Friends(Int(Rnd * 5) * 9).top
            Next
        End If
    End If
End If
End Sub
Sub GrassEffect()
If FightWindow.BigProgress <> 0 And IsInProgress = True Then
If FightWindow.NowProgress Mod FightWindow.BigProgress = 0 Then
For i = 1 To 5
NewEMonster "ʳ�˻�", FightWindow.ScaleWidth * 0.4 + Int(Rnd * FightWindow.ScaleWidth * 0.3), FightWindow.Friends(Int(Rnd * 5) * 9).top
Next
End If
End If
End Sub
Sub Night2()
Select Case Step
    Case 0
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "����"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "����"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "����"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "����"
    FightWindow.NewFaceMsg "����", "��New����һ���µ��ң��������ǵ�ս�����ݡ�"
    Step = Step + 1
End Select
If Step <= 0 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub EverySky()
Select Case Step
    Case 0
    For i = 0 To 2
    For s = 0 To 36 Step 9
    NewFMonster FightWindow.Friends(i).Left, FightWindow.Friends(s).top, "��"
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
    NewFMonster FightWindow.Friends(s).Left, FightWindow.Friends(i).top, "��"
    Next
    Next
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "ʳ����"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "ʳ����"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "ʳ����"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "ʳ����"
    FightWindow.NewFaceMsg "����", "�������ô��������˰ɣ�������....������..."
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "����", "���ˣ�����һ���ơ�"
    Step = Step + 1
    Case 2
    If OwnMonster("��") = False Then GetPresent "m ��"
    Step = Step + 1
End Select
If Step <= 1 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub FirstDesert()
Select Case Step
    Case 0
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "������"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "������"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "������"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "������"
    FightWindow.NewFaceMsg "����", "���Ȱ�....���Ȱ�...����Ա�����...."
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "����", "����......������...����...��.......�﷢������ֻ������ħ��......"
    Step = Step + 1
    Case 2
    FightWindow.NewFaceMsg "����", "����.....�ĸ�.....��............������....������......."
    Ring "get"
    FunnyCounts = FunnyCounts + 4
    Step = Step + 1
    Case 3
    FightWindow.NewFaceMsg "����", "��ȥ�Ա�����...........�ټ�.......����...."
    Step = Step + 1
End Select
If Step <= 1 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub FirstLostSpace()
Select Case Step
    Case 0
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "糺�֮��"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "糺�֮��"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "糺�֮��"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "糺�֮��"
    FightWindow.NewFaceMsg "����", "�ֵ������ˣ������Һ��������������ϡ�"
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "����", "Ϊʲô�ء���ΪϦ���µĵж�ħ���Ƕ��쳣�˷�..."
    Step = Step + 1
    Case 2
    FightWindow.NewFaceMsg "����", "����New������ֻ糺�֮����Ŷ���У������ĸ�������"
    Ring "get"
    FunnyCounts = FunnyCounts + 4
    Step = Step + 1
    Case 3
    FightWindow.NewFaceMsg "404", "������ϲ����Ctrl+S�ĳ���Ա������"
    Step = Step + 1
End Select
If Step <= 1 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub FirstNight()
Select Case Step
    Case 0
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "����"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "����"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "����"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "����"
    FightWindow.NewFaceMsg "����", "��������...404�ղ������Ӹ��Ҵ��˸��绰�������������..."
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "����", "��...��Ҫȥ�ô����侲�侲�����ˣ���Ӧ�ò��¹��ģ��԰ɣ�..."
    Step = Step + 1
End Select
If Step <= 1 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub MonsterFlowerLevel2()
Select Case Step
    Case 0
    NewFMonster FightWindow.Friends(2).Left, FightWindow.Friends(2).top, "�ղ�"
    NewFMonster FightWindow.Friends(10).Left, FightWindow.Friends(10).top, "�ղ�"
    NewFMonster FightWindow.Friends(28).Left, FightWindow.Friends(28).top, "�ղ�"
    NewFMonster FightWindow.Friends(38).Left, FightWindow.Friends(38).top, "�ղ�"
    FightWindow.NewFaceMsg "����", "404д��һ������ѭ��Newħ������Ҫ���ˣ�����"
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "����", "����˵����10�����������㣡�����ҷ��ˣ�����"
    Ring "get"
    FunnyCounts = FunnyCounts + 10
    Step = Step + 1
End Select
If Step <= 1 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub MonsterFlowerLevel()
Select Case Step
    Case 0
    FightWindow.NewFaceMsg "����", "����404���ˣ����ָ��Ҵ��˵绰��˵����New��һ���ħ����˵Ҫ��ȥ����MMP��"
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "����", "������Ƭ�����Ѿ��������ˣ���Ҳ�����˰ɣ�һ��ѵ�ħ����"
    Step = Step + 1
    Case 2
    FightWindow.NewFaceMsg "����", "���ˣ��ղ�������ħ���ˣ�������������ԡ�"
    Step = Step + 1
    Case 3
    FightWindow.NewFaceMsg "����", "������6��������~"
    Ring "get"
    FunnyCounts = FunnyCounts + 6
    Step = Step + 1
End Select
If Step <= 3 Then FightWindow.LastProgress = GetTickCount
End Sub
Sub NewcomerLevel()
Select Case Step
    Case 0
    FightWindow.SetCard 0, "����֮��", 0
    FightWindow.SetCard 1, "��ʿ", 0
    FightWindow.NewFaceMsg "����", "��һ�죬404���ô���...ͻȻͣ���ˡ�" & vbCrLf & "����404�����ˣ����ټ��˸��ָ�����ħ�ޣ�����Ҫɱ��ϲ����Ctrl + S�ĳ���Ա��" & vbCrLf & "Ȼ��..."
    Step = Step + 1
    Case 1
    FightWindow.NewFaceMsg "404", "�����ɣ�ϲ����Ctrl+S�ĳ���Ա��~~~��" & vbCrLf & vbCrLf & "��Ϸ������Error 404 (QQ1361778219)���ز�&��Ч��RPG MV&���磬���֣���������գ��Ա��ƶ���˾��&���磬ͼ�μ��٣�����"
    Step = Step + 1
    Case 2
    FightWindow.NewFaceMsg "����", "mmp˵�ܲٲܲپ͵�����˵��ղ������һ����ɶ�����˲�����������˵���ǵİɡ�"
    Step = Step + 1
    Case 3
    FightWindow.NewFaceMsg "����", "�����ߵ�ħ�ޣ��ҵĳ���ͻ�New��һ���µ�ħ��Ϊ��ս����������Ҫ�㹻�Ļ������ſ����ṩ�㹻���������ҵĳ������New��һ���µ�ħ�ޡ�"
    Step = Step + 1
    Case 4
    FightWindow.NewFaceMsg "����", "���ֵ�ʱ���һ����New��4�����������������������У����ܻ������Ʈ�������������ɲ���ͣ��������һֱ���µ������ֵ��ĸ��������ɱ��˷��ˣ��õ�������Ҳ����ҪһЩʱ��ģ���Ӧ����Ƹ��һЩ�������������͵�ħ�ޡ�"
    Step = Step + 1
    Case 5
    FightWindow.NewFaceMsg "����", "����ͻ�����������Ե���ؿ���������ߵ� ����������������ĳһ��ħ�ޣ������ħ�޷�����ɱ����"
    Step = Step + 1
    Case 6
    FightWindow.NewFaceMsg "����", "������Ӧ�ò����ҽ���Ҳ���˰ɡ����Ҿ���ȥ�ô����ˣ����Ǽ��ͣ�����404���Ұ����ˡ�"
    Step = Step + 1
    Case 7
    FightWindow.NewFaceMsg "404", "hhhû���Ҿ��������Ϸ�� ������ ��ɵ��404�����Ͱ�~"
    Step = Step + 1
    Case 8
    FightWindow.NewFaceMsg "����", "�������ˣ�����Ե�����Ͻǵĺ�ɫ��ť����͡�ɾ��ħ�ޣ������һ����ɫ��ť�������ͣ��Ϸ��"
    Step = Step + 1
    Case 9
    FightWindow.NewFaceMsg "����", "����Աߵ�HP��ť���Բ鿴ħ�޵�Ѫ������ӡ���ֺ����ܣ�����Щ�û���ϲ������"
    Step = Step + 1
End Select
If Step <= 9 Then FightWindow.LastProgress = GetTickCount
End Sub