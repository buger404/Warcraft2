Attribute VB_Name = "modPresent"
Function GetPresent(PresentCode As String)
'激活战利品
Dim Codes() As String, PresentText As String
Codes = Split(PresentCode, " ")
Select Case Codes(0)
    Case "m" '获得魔兽
    ReDim Preserve MyMonster(UBound(MyMonster) + 1)
    ReDim Preserve MyMonsterLevel(UBound(MyMonsterLevel) + 1)
    MyMonster(UBound(MyMonster)) = Codes(1)
    PresentText = Codes(1) & "的声明&定义"
    Case "g" '获得游戏币
    Money = Money + Val(Codes(1))
    PresentText = "Money  " & Val(Codes(1))
    Case "f" '获得滑稽能量
    FightWindow.SuperCounts = FightWindow.SuperCounts + Val(Codes(1))
    PresentText = "SuperFunny x " & Val(Codes(1))
    Case "ff" '获得初始滑稽果
    AtFirstFunny = AtFirstFunny + Val(Codes(1))
    PresentText = "开局滑稽果数量+" & Val(Codes(1))
End Select

GetPresent = PresentText

If MainBGM = False Then
ShowInfo "恭喜获得 " & PresentText, PresentIcon(PresentCode)
End If
End Function
Sub NewPresent(PresentCode As String, X As Single, Y As Single)
'设置战利品
If FightWindow.DebugMode = True Then Exit Sub
ReDim Preserve Presents(UBound(Presents) + 1)
Presents(UBound(Presents)).X = X
Presents(UBound(Presents)).Y = Y
Presents(UBound(Presents)).Code = PresentCode
Dim Codes() As String
Codes = Split(PresentCode, " ")
Select Case Codes(0)
    Case "m" '获得魔兽
    Presents(UBound(Presents)).Icon = Codes(1) & "Icon"
    Case "g" '获得游戏币
    Presents(UBound(Presents)).Icon = "moneyIcon"
    Case "f" '获得滑稽能量
    Presents(UBound(Presents)).Icon = "SuperFunny"
End Select
End Sub
Function PresentIcon(ByVal PresentCode As String)
'设置战利品
Dim Codes() As String
Codes = Split(PresentCode, " ")
Select Case Codes(0)
    Case "m" '获得魔兽
    PresentIcon = Codes(1) & "Icon"
    Case "g" '获得游戏币
    PresentIcon = "moneyIcon"
    Case "f" '获得滑稽能量
    PresentIcon = "SuperFunny"
    Case "ff"
    PresentIcon = "funny"
End Select
End Function
