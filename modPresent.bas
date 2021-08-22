Attribute VB_Name = "modPresent"
Function GetPresent(PresentCode As String)
'����ս��Ʒ
Dim Codes() As String, PresentText As String
Codes = Split(PresentCode, " ")
Select Case Codes(0)
    Case "m" '���ħ��
    ReDim Preserve MyMonster(UBound(MyMonster) + 1)
    ReDim Preserve MyMonsterLevel(UBound(MyMonsterLevel) + 1)
    MyMonster(UBound(MyMonster)) = Codes(1)
    PresentText = Codes(1) & "������&����"
    Case "g" '�����Ϸ��
    Money = Money + Val(Codes(1))
    PresentText = "Money  " & Val(Codes(1))
    Case "f" '��û�������
    FightWindow.SuperCounts = FightWindow.SuperCounts + Val(Codes(1))
    PresentText = "SuperFunny x " & Val(Codes(1))
    Case "ff" '��ó�ʼ������
    AtFirstFunny = AtFirstFunny + Val(Codes(1))
    PresentText = "���ֻ���������+" & Val(Codes(1))
End Select

GetPresent = PresentText

If MainBGM = False Then
ShowInfo "��ϲ��� " & PresentText, PresentIcon(PresentCode)
End If
End Function
Sub NewPresent(PresentCode As String, X As Single, Y As Single)
'����ս��Ʒ
If FightWindow.DebugMode = True Then Exit Sub
ReDim Preserve Presents(UBound(Presents) + 1)
Presents(UBound(Presents)).X = X
Presents(UBound(Presents)).Y = Y
Presents(UBound(Presents)).Code = PresentCode
Dim Codes() As String
Codes = Split(PresentCode, " ")
Select Case Codes(0)
    Case "m" '���ħ��
    Presents(UBound(Presents)).Icon = Codes(1) & "Icon"
    Case "g" '�����Ϸ��
    Presents(UBound(Presents)).Icon = "moneyIcon"
    Case "f" '��û�������
    Presents(UBound(Presents)).Icon = "SuperFunny"
End Select
End Sub
Function PresentIcon(ByVal PresentCode As String)
'����ս��Ʒ
Dim Codes() As String
Codes = Split(PresentCode, " ")
Select Case Codes(0)
    Case "m" '���ħ��
    PresentIcon = Codes(1) & "Icon"
    Case "g" '�����Ϸ��
    PresentIcon = "moneyIcon"
    Case "f" '��û�������
    PresentIcon = "SuperFunny"
    Case "ff"
    PresentIcon = "funny"
End Select
End Function
