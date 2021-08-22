Attribute VB_Name = "modCore"
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" _
           (ByVal lpApplicationName As Long, _
            ByVal lpKeyName As Long, _
            ByVal lpDefault As Long, _
            ByVal lpReturnedString As Long, _
            ByVal nSize As Long, _
            ByVal lpFileName As Long) As Long
Public Const SND_ASYNC = &H1                                                    '  play asynchronously
Public Const SND_FILENAME = &H20000                                             '  name is a file name
Public Const SND_LOOP = &H8                                                     '  loop the sound until next sndPlaySound
Public Const FuckReg1 = "Ua4iwzzsZareGQqGt1SCq9ppQQGGm7maOwetKfTK4iwz8uBAqfrr5Fs7rS"
Public Const FuckReg2 = "MQ4EwB7XZRqErq"
Public Const ODClass = "������"
Public Const CEClass = "dBrofBqQz1Ub"
Public Const IsPreviewVersion As Boolean = False

Public GamePictures() As FrameManager
Public GameSounds()  As Monster2_Sounds
Public CheatTime As New VariableManager

Dim AssetsOK As Boolean

Public VoChangeTime As Long
Public RFPS As Integer
Public FPS As Integer
Public LoadOK As Boolean
Public BackMusic As String
Public MainBGM As Boolean

Public FocusPen As Long
Public Sets(8) As Boolean
Public HistoryCost As Long
Public DMode As Boolean 'is in debug , if is ,

Public GameVo As Single

Public OnLineMode As Integer '0=û�м��� 1=���� 2=��Ա

Public LastVersion As Long ' last version so that i can popup a update window later

Public Money As Double
Public NowWorld As Long
Public NowLevel(4) As Long
Public PlayerName As String
Public MyMonster() As String
Public MyMonsterLevel() As Long


Public WorldName(4) As String
Public SuperCount  As Long
Public FunnyCounts As Long
Public AtFirstFunny As Long

Public Presents() As Monster2_Present
Public FMonster() As New FriendlyMonster
Public EMonster() As New EvilMonster
Public GameEffect() As New EffectManager
Public FMonsterOnLine(4) As String
Public EMonsterOnLine(4) As String

Public LastGetPic As String
Public ActiveWindow As Form

Public OnLine As Winsock

Public MCards(11) As Monster2_Cards
Public Funnys() As Monster2_Funny '������
Public ErrCount As Long
Public BGM As New SongManager
Public Type Monster2_Present
    X As Long
    Y As Long
    Code As String
    Icon As String
End Type
Public Type Monster2_Funny
    X As Long
    Y As Long
    MoveX As Long
    MoveY As Long
    MoveCount As Byte
End Type
Public Type Monster2_Sounds
    name As String
    data As Long
End Type
Public Type Monster2_Cards
    name As String
    CDTime As Long
    NowCD As Long
    Spend As Long
    Level As Long
End Type
Public Enum Monster2_Effect '������Ч������
    MagicText = 0
    TaketurnsPic = 1
    Heartbeat = 2
    FadeInPic = 3
End Enum
Public Enum Monster2_Types '����ħ�޵�����
    Normal = 0 '��ͨ
    OhAngle = 1 '����ɨ��
    HateFlower = 2 '�ղ�ר��
    BeatOut = 3 '����
    Burn = 4 '����������
    RanAttack = 5 '���̶��˺�
    Chemical = 6 '����
    BOOM = 7 '��ը��
    boom2 = 8 '��͸��ը��
    CopyHP = 9 '��Ѫ
    TV = 10 '����
    ghost = 11 '��͸
    NewANew = 12 '�ٻ�
    Chemical2 = 13 '�ж�
    Define = 14 '����
    Ice = 15 '����
    GoBack = 16 '����
    boom3 = 17 '��ըEx
    BurnSuper = 18 '������������
    FireFire = 19 'ȼ��
    FireGhost = 20 '��͸ȼ��
    ThreeLine = 21 '���߹���
    IceAll = 22 '��������ħ��
    fly = 23 '����
    cloud = 24 '��
    ReSelf = 25 '�Զ���Ѫ
    MoveMon = 26 '�ƶ�ħ��
    MissAll = 27 '������ͨ����
    Becomess = 28 '��Ⱦ
    Gay = 29 'Gay��
    PurpleFunny = 30 '�ϻ���
    DarkNess = 31 '��ҹ
    Afraid = 32 '����
    TakeFunny = 33 '͵ȡ������
    ThrowAround = 34 '��ħ��
    NoFar = 35 '�̾��빥��
    Thread = 36 '���
    ThreadGhost = 37 '�����͸
    FireAll = 38 'ȼ��ȫ��
    CanMove = 39 '�ܹ��ƶ�
    FunnyAttack = 40 '�������������ҹ���
    Water = 41 'ˮ
    NoNeedWater = 42 '����Ҫˮ
    WaterAll = 43 '��������ιˮ
    Mirror = 44 '������
    Fly2 = 45 '˲��
End Enum
Function GetEffectName(Effect As Monster2_Types)
Effectname = Array("��ͨ", "��ɱ", "����", "����", "����������", "���", "����", "��ը", "��͸�ͱ�ը", "��Ѫ", "����", "��͸", _
                                "�ٻ�", "�ж�", "����", "����", "����", "��ըEx", "������������", "ȼ��", "��͸��ȼ��", "���߹���", "��������ħ��", _
                                "����", "��", "���һָ�", "����", "���߷Ǵ�͸����", "��Ⱦ", "Gay��", "������ǿ", "��ҹ", "����", "���ջ�����", "Ͷ��", "�̾���", "���", _
                                "��͸�͵��", "ȼ������ħ��", "�����ƶ�", "����+����������", "���", "�ͺ�", "ȫ�����", "������·", "˲��")
GetEffectName = Effectname(Effect)
End Function
Sub PostData(data As String)
On Error Resume Next

OnLine.SendData data & "��"
End Sub
Sub NewEMonster(ByVal name As String, ByVal X As Single, ByVal Y As Single)
If OnLineMode = 2 Then Exit Sub
ReDim Preserve EMonster(UBound(EMonster) + 1)
EMonster(UBound(EMonster)).X = X
EMonster(UBound(EMonster)).Y = Y
EMonster(UBound(EMonster)).MonsterName = name
EMonster(UBound(EMonster)).LoadMonster EMonster(UBound(EMonster)).MonsterName, UBound(EMonster)
If EMonster(UBound(EMonster)).MonsterType = Mirror Then
EMonster(UBound(EMonster)).X = FightWindow.Friends(1).Left
End If
If OnLineMode = 1 Then PostData "em " & name & " " & X & " " & Y
End Sub
Function GetLevelStr(TypeName As String, value As Variant, Optional EMonsterMa As Boolean = False)
Dim AttackLevel() As Variant, SpeedLevel() As Variant, DuringLevel() As Variant, HPLevel() As Variant, CDLevel() As Variant
AttackLevel = Array(5, 10, 15, 20)
If EMonsterMa = False Then
SpeedLevel = Array(3, 5, 10, 20)
Else
SpeedLevel = Array(0.8, 1, 1.5, 2)
End If
DuringLevel = Array(2500, 3000, 3500, 4000)
HPLevel = Array(50, 100, 200, 300)
CDLevel = Array(5, 10, 15, 20)

Select Case TypeName
    Case "Attack"
        GetLevelStr = LevelStr(AttackLevel, value)
    Case "Speed"
        GetLevelStr = LevelStr3(SpeedLevel, value)
    Case "During"
        GetLevelStr = LevelStr2(DuringLevel, value)
    Case "HP"
        GetLevelStr = LevelStr(HPLevel, value)
    Case "CD"
        GetLevelStr = LevelStr2(CDLevel, value)
End Select
End Function
Function LevelStr(LevelINI() As Variant, value As Variant)
Select Case value
    Case Is = 0
    LevelStr = "������"
    Case Is < LevelINI(0)
    LevelStr = "�ܵ�"
    Case Is < LevelINI(1)
    LevelStr = "�ϵ�"
    Case Is < LevelINI(2)
    LevelStr = "��ͨ"
    Case Is < LevelINI(3)
    LevelStr = "�ϸ�"
    Case Is < LevelINI(3) * 2
    LevelStr = "�ܸ�"
    Case Else
    LevelStr = "ը��"
End Select
End Function
Function LevelStr2(LevelINI() As Variant, value As Variant)
Select Case value
    Case Is = 0
    LevelStr2 = "������"
    Case Is < LevelINI(0)
    LevelStr2 = "�ܶ�"
    Case Is < LevelINI(1)
    LevelStr2 = "�϶�"
    Case Is < LevelINI(2)
    LevelStr2 = "��ͨ"
    Case Is < LevelINI(3)
    LevelStr2 = "�ϳ�"
    Case Is < LevelINI(3) * 2
    LevelStr2 = "�ܳ�"
    Case Else
    LevelStr2 = "ը��"
End Select
End Function
Function LevelStr3(LevelINI() As Variant, value As Variant)
Select Case value
    Case Is = 0
    LevelStr3 = "��ţ"
    Case Is < LevelINI(0)
    LevelStr3 = "����"
    Case Is < LevelINI(1)
    LevelStr3 = "����"
    Case Is < LevelINI(2)
    LevelStr3 = "��ͨ"
    Case Is < LevelINI(3)
    LevelStr3 = "�Ͽ�"
    Case Is < LevelINI(3) * 2
    LevelStr3 = "�ܿ�"
    Case Else
    LevelStr3 = "ը��"
End Select
End Function
Function GetPic(PicName As String) As Integer '���ͼƬ
For i = 1 To UBound(GamePictures)
If GamePictures(i).name = PicName Then GetPic = i: Exit For
Next
If GetPic = 0 Then Debug.Print "Error : Can't find the picture '" & PicName & "' ."
End Function
Sub Ring(ByVal filename As String)
If Sets(0) = True Then Exit Sub
Dim Find As Boolean
For i = 0 To UBound(GameSounds)
If GameSounds(i).name = filename & ".wav" Then
'If BGM.Rate <> 44100 Then BASS_ChannelSetAttribute GameSounds(i).data, BASS_ATTRIB_FREQ, BGM.Rate
BASS_ChannelPlay GameSounds(i).data, True
Find = True
Exit For
End If
Next
'If Find = False Then Err.Raise 4040, , "Failed to playsound:" & filename & " ."
End Sub
Sub LoadPic(PicName As String, Path As String) '���ͼƬ
Dim HeHe As New FrameManager
ReDim Preserve GamePictures(UBound(GamePictures) + 1)
HeHe.name = PicName
HeHe.LoadFromFile Dad.HDC, Path
Set GamePictures(UBound(GamePictures)) = HeHe
End Sub
Sub NewFMonster(ByVal X As Single, ByVal Y As Single, ByVal MonsterName As String, Optional ByVal NLevel As Long = 0)
ReDim Preserve FMonster(UBound(FMonster) + 1)
FMonster(UBound(FMonster)).X = X
FMonster(UBound(FMonster)).Y = Y
FMonster(UBound(FMonster)).Level = NLevel
FMonster(UBound(FMonster)).LoadMonster MonsterName, UBound(FMonster)
If NowWorld = 4 And FMonster(UBound(FMonster)).MonsterType <> NoNeedWater Then
FMonster(UBound(FMonster)).NeedWater = True
FMonster(UBound(FMonster)).Speed = FMonster(UBound(FMonster)).Speed / 2
FMonster(UBound(FMonster)).During = FMonster(UBound(FMonster)).During * 2
End If

If OnLineMode <> 0 Then PostData "fm " & MonsterName & " " & X & " " & Y
End Sub
Sub MakeFMonsterDead(ByVal Index As Integer)
If Index <= UBound(FMonster) Then
Set FMonster(Index) = FMonster(UBound(FMonster))
FMonster(Index).MonsterIndex = Index
ReDim Preserve FMonster(UBound(FMonster) - 1)
End If
End Sub
Function GetDoubleMonster(ByVal WithOutIndex As Integer, ByVal X As Single, ByVal Y As Single) As String   '�����ָ�������ص���ħ��
For i = 0 To UBound(FMonster)
If i <> WithOutIndex Then
If FMonster(i).X = X And FMonster(i).Y = Y Then GetDoubleMonster = FMonster(i).MonsterName: Exit For
End If
Next
End Function
Function GetDoubleMonsterIndex(ByVal WithOutIndex As Integer, ByVal X As Single, ByVal Y As Single) As Integer   '�����ָ�������ص���ħ�޵����
For i = 0 To UBound(FMonster)
If i <> WithOutIndex Then
If FMonster(i).X = X And FMonster(i).Y = Y And FMonster(i).MonsterType <> cloud Then GetDoubleMonsterIndex = i: Exit For
End If
Next
End Function
Sub LoadAllAssets() '��AssetsĿ¼�µ������ļ�����
Dim file As String, temp() As String
file = Dir(App.Path & "\assets\")
Do While file <> ""
DoEvents
temp = Split(file, ".")
LoadPic temp(0), App.Path & "\assets\" & file
file = Dir()
Loop
End Sub
Sub LoadAllSounds() '��soundsĿ¼�µ������ļ�����
Dim file As String, temp() As String
file = Dir(App.Path & "\sounds\")
Do While file <> ""
DoEvents
temp = Split(file, ".")
ReDim Preserve GameSounds(UBound(GameSounds) + 1)
GameSounds(UBound(GameSounds)).name = file
GameSounds(UBound(GameSounds)).data = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & "\sounds\" & file), 0, 0, 0)
If GameSounds(UBound(GameSounds)).data = 0 Then VBA.Msgbox "Failed to load sound:" & file & " .", 16
file = Dir()
Loop
End Sub
Sub LoadSounds(file As String)
'temp = Split(file, ".")
ReDim Preserve GameSounds(UBound(GameSounds) + 1)
GameSounds(UBound(GameSounds)).name = file
GameSounds(UBound(GameSounds)).data = BASS_StreamCreateFile(BASSFALSE, StrPtr(App.Path & "\sounds\" & file), 0, 0, 0)
If GameSounds(UBound(GameSounds)).data = 0 Then VBA.Msgbox "Failed to load sound:" & file & " .", 16
End Sub
Sub CreateAChild(Child As Form)
'Wait 2000
Set ActiveWindow = Child
SetParent Child.hWnd, Dad.hWnd
Child.Show
Child.Move 0, 0
End Sub
Sub ShowLastShow()
    '===============================================
    NowWorld = 4
    Dim tempUI As Long '��ʱ������
    Dim MonsterList() As String, file As String, temp() As String
    file = Dir(App.Path & "\monster\fmonster\")
    ReDim MonsterList(0)
    Do While file <> ""
    temp = Split(file, ".")
    MonsterList(UBound(MonsterList)) = temp(0)
    ReDim Preserve MonsterList(UBound(MonsterList) + 1)
    file = Dir()
    DoEvents
    Loop
    
    CreateAChild FightWindow: GdipCreateFromHDC FightWindow.HDC, tempUI
    FightWindow.UI = tempUI: FightWindow.DrawTimer.Enabled = True
    FightWindow.quickbutton.Visible = False
    FightWindow.ratebutton.Visible = False
    FightWindow.musicbutton.Visible = False
    FightWindow.SetLevel 50
    FightWindow.ProgressDuring = 0
    FightWindow.FirstDuring = 0
    BGM.SetPlayRate 2
    
    Dim MTY As Monster2_Types, MT As Integer, temp4 As New FriendlyMonster, temp5 As Integer
    For i = 1 To 200
        MT = Int(Rnd * (UBound(MonsterList) - 1))
        temp4.LoadMonster MonsterList(MT), 0
        MTY = temp4.MonsterType
            If i < 30 Then
            If MTY = BOOM Or MTY = boom2 Or MTY = boom3 Or MTY = cloud Or MTY = PurpleFunny Or MTY = BurnSuper Or _
            MTY = DarkNess Or MTY = GoBack Or MTY = FireAll Or MTY = IceAll Then
            ElseIf Dir(App.Path & "\assets\" & MonsterList(MT) & "Attack.png") <> "" And Dir(App.Path & "\assets\" & MonsterList(MT) & "Fire.png") <> "" And Dir(App.Path & "\assets\" & MonsterList(MT) & "1.png") <> "" Then
            temp5 = (i Mod 5) * 9
            NewFMonster FightWindow.Friends(Int(Rnd * 3)).Left, FightWindow.Friends(temp5).top, MonsterList(MT)
            End If
            End If
            
        If FightWindow.FirstDanger = True Then NewFMonster FightWindow.Friends(Int(Rnd * 44)).Left, FightWindow.Friends(temp5).top, "�֮Ů"
        If UBound(EMonster) > 120 Then NewFMonster FightWindow.Friends(Int(Rnd * 44)).Left, FightWindow.Friends(temp5).top, "�֮Ů"
        
        If i Mod 10 = 0 Then
        Ring "����"
        NowWorld = NowWorld + 1
        If NowWorld > 4 Then NowWorld = 0
        FightWindow.SetLevel 50
        FightWindow.ProgressDuring = 0
        FightWindow.FirstDuring = 0
        FightWindow.SpicalBack = ""
        End If
        Wait 300
        Debug.Print "Step " & i & "/100"
    Next
    
    Unload FightWindow
    '===============================================
End Sub
Sub Wait(DuringVal As Long)
temp = GetTickCount
Do While GetTickCount - temp < DuringVal
Sleep 10: DoEvents
Loop
End Sub
Sub Main()
    Dim strComputer, objWMIService, colItems, objItem, strOSversion As String
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    
    For Each objItem In colItems
        strOSversion = objItem.Version
    Next
    
    OSver = Split(strOSversion, ".")
    Select Case Left(strOSversion, 3)
    Case "10."
        strOSversion = "Windows 10"
    Case "5.2"
        strOSversion = "Windows Server 2003"
    Case "5.0"
        strOSversion = "Windows 2000"
    Case "5.1"
        strOSversion = "Windows XP"
    Case "6.0"
        strOSversion = "Windows Visita"
    Case "6.1"
        strOSversion = "Windows 7"
    Case Else
        strOSversion = "i don't know"
    End Select
If strOSversion = "Windows XP" Then VBA.Msgbox "����֧�ִ���Ϸ������Windows XP�ϣ�", 48, "ħ�޻�ս2": End

Set ActiveWindow = WelcomeWindow
GameVo = 0.03

For i = 0 To UBound(Sets)
Sets(i) = False
Next

If GetSetting(LockString(FuckReg1, "! you are cheatting !", 1), LockString(FuckReg2, "! you are cheatting !", 1), "0") <> "" Then CloseCEandOD: End

If App.PrevInstance = True Then VBA.Msgbox "���Ѿ�����һ��ħ�޻�ս�~�����ظ���Ӵ~", 64, "����ħ�޻�ս2����ܰ��ʾ"

WorldName(0) = "��ԭ": WorldName(1) = "ҹ��": WorldName(2) = "���": WorldName(3) = "�ż�": WorldName(4) = "ɳĮ"
ReDim GamePictures(0)
ReDim GameSounds(0)
ReDim GameEffect(0)

Set GamePictures(0) = New FrameManager

LoadPic "mouse", App.Path & "\assets\mouse.png"
LoadPic "frame", App.Path & "\assets\frame.png"
LoadPic "blueeffect", App.Path & "\assets\blueeffect.png"
LoadPic "MainBackground", App.Path & "\assets\MainBackground.png"
LoadPic "LOGO", App.Path & "\assets\LOGO.png"
charlist = "abcdefghijklnmopqrstuvwxyz0123456789 "
For i = 1 To Len(charlist)
LoadPic "char" & Mid(charlist, i, 1), App.Path & "\assets\" & "char" & Mid(charlist, i, 1) & ".png"
Next
LoadSounds "Welcome.wav"

GdipCreateSolidFill argb(0, 255, 255, 255), Brush1: GdipCreatePen1 argb(128, 255, 255, 255), 1, UnitPixel, Pen1

'===============================��������================================
GdipCreateFontFamilyFromName StrPtr("΢���ź�"), 0, fontfam
GdipCreateStringFormat 0, 0, strformat
GdipSetStringFormatAlign strformat, StringAlignmentNear
GdipCreateFont fontfam, 12, FontStyle.FontStyleRegular, UnitPixel, curFont
GdipCreateFont fontfam, 18, FontStyle.FontStyleRegular, UnitPixel, curFontBig
GdipSetTextRenderingHint Graphics, TextRenderingHintClearTypeGridFit
'====================================================================

With BGM
    .StopMusic
    .LoadMusic App.Path & "\music\" & "Background" & UBound(NowLevel) + 1 & ".mp3"
    .Play
End With

BGMBox.Show
BGMBox.Hide
Call ReadSet
Dad.Show
CreateAChild WelcomeWindow
Ring "Welcome"
Wait 1500
LoadAllAssets
LoadAllSounds '����������Դ

'===============================================
'Dim tempUI As Long '��ʱ������
'CreateAChild FightWindow: GdipCreateFromHDC FightWindow.Hdc, tempUI
'FightWindow.UI = tempUI: FightWindow.DrawTimer.Enabled = True
'===============================================
'Wait 2000

Set OnLine = BGMBox.Winsock1

If Dir("C:\Monster2\", vbDirectory) = "" Then MkDir "C:\Monster2\"

If Dir("C:\Monster2\save1.rsdata") = "" Then
ReDim MyMonster(1)
MyMonster(0) = "����֮��"
MyMonster(1) = "��ʿ"
Named:
PlayerName = Inputbox("�����Լ���һ��˧�������ְɣ�")
If PlayerName = "" Then Msgbox "����Ŷ����ø��Լ�������ְ���", , "����": GoTo Named
Money = 50
AtFirstFunny = 4
Call WriteSave
Else
Call ReadSave
Call UpdateSave
End If

LoadOK = True
'CreateAChild IntroWindow
'CreateAChild BookWindow
End Sub
Sub StartGame()
If NowLevel(0) = 0 Then
NowWorld = 0
'===============================================
Dim tempUI As Long '��ʱ������
CreateAChild FightWindow: GdipCreateFromHDC FightWindow.HDC, tempUI
FightWindow.UI = tempUI: FightWindow.DrawTimer.Enabled = True
'===============================================
'FadeIn 20, Dad
Else
CreateAChild MainWindow
'FadeIn 20, Dad
Wait 1000
Ring "play"

If Dir("C:\Monster2\Version.txt") <> "" Then LastVersion = Val(ReadFile("C:\Monster2\Version.txt"))
If LastVersion < 183310 Then '18���316����1�������
Open "C:\Monster2\Version.txt" For Output As #1
Print #1, 183310
Close #1
Msgbox ReadFile(App.Path & "\Update.txt"), , "��������"
End If

End If

Unload WelcomeWindow
End Sub
Public Function ReadINI(ByVal SectionName As String, ByVal KeyName As String, ByVal IniFileName As String) As String
    Dim strBuf As String
    strBuf = String(128, 0)
    GetPrivateProfileString StrPtr(SectionName), StrPtr(KeyName), StrPtr(""), StrPtr(strBuf), 128, StrPtr(IniFileName)
    strBuf = Replace(strBuf, Chr(0), "")
    ReadINI = strBuf
End Function
Function ReadFile(Path As String)
Open Path For Input As #1
Do While Not EOF(1)
Line Input #1, a
b = b & a & vbCrLf
Loop
Close #1
ReadFile = b
End Function
Sub ReadSave()
Dim temp As Variant, a As String
Open "C:\Monster2\save1.rsdata" For Input As #1
Line Input #1, a
PlayerName = a
Line Input #1, a
Money = Val(a)
Line Input #1, a
temp = Split(a, ";")
For i = 0 To UBound(NowLevel)
NowLevel(i) = Val(temp(i))
Next
Line Input #1, a
temp = Split(a, ";")
ReDim MyMonster(UBound(temp))
ReDim MyMonsterLevel(UBound(temp))
For i = 0 To UBound(MyMonster)
MyMonster(i) = temp(i)
Next
If EOF(1) Then
    AtFirstFunny = 4
Else
    Line Input #1, a
    AtFirstFunny = Val(a)
End If
If EOF(1) Then
    HistoryCost = 0
Else
    Line Input #1, a
    HistoryCost = Val(a)
End If
Close #1

If Dir("C:\Monster2\levels.rsdata") <> "" Then
i = 0
Open "C:\Monster2\levels.rsdata" For Input As #1
Do While Not EOF(1)
Line Input #1, tempLevel
MyMonsterLevel(i) = Val(tempLevel)
i = i + 1
Loop
Close #1
End If
End Sub
Sub WriteSave()
Dim MyMonsterForSave As String, LevelString As String
For i = 0 To UBound(MyMonster) - 1
MyMonsterForSave = MyMonsterForSave & MyMonster(i) & ";"
Next
For i = 0 To UBound(NowLevel)
LevelString = LevelString & NowLevel(i) & ";"
Next
MyMonsterForSave = MyMonsterForSave & MyMonster(UBound(MyMonster))
Open "C:\Monster2\save1.rsdata" For Output As #1
Print #1, PlayerName
Print #1, Money
Print #1, LevelString
Print #1, MyMonsterForSave
Print #1, AtFirstFunny
Print #1, HistoryCost
Close #1
Open "C:\Monster2\levels.rsdata" For Output As #1
For i = 0 To UBound(MyMonsterLevel)
Print #1, MyMonsterLevel(i)
Next
Close #1
End Sub
Sub UpdateSave()
For i = 0 To UBound(MyMonster)
If i > UBound(MyMonster) Then Exit For
If MyMonster(i) = "SWL" Then MyMonster(i) = "΢��"
If MyMonster(i) = "��԰��" Then MyMonster(i) = "��԰��"
If MyMonster(i) = "ǧ����" Then MyMonster(i) = "ǧ���"
If MyMonster(i) = "������" Then MyMonster(i) = "������"
Next
WriteSave
End Sub
Function OwnMonster(MName As String) As Boolean
For i = 0 To UBound(MyMonster)
If MyMonster(i) = MName Then OwnMonster = True: Exit For
Next
End Function
Function FindMonster(MName As String) As Integer
For i = 0 To UBound(MyMonster)
If MyMonster(i) = MName Then FindMonster = i: Exit For
Next
End Function
Function CanUpLevel(MName As String) As Integer
Dim s As Integer
For i = 0 To UBound(MyMonster)
If MyMonster(i) = MName Then s = s + 1
Next
CanUpLevel = s
End Function
Sub SaveSet()
Open "C:\Monster2\settings.rsdata" For Output As #1
For i = 0 To UBound(Sets)
Print #1, Sets(i)
Next
Close #1
End Sub
Sub ReadSet()
On Error Resume Next
Open "C:\Monster2\settings.rsdata" For Input As #1
For i = 0 To UBound(Sets)
If EOF(1) Then Exit For
Line Input #1, a
Sets(i) = a
Next
Close #1
End Sub
Sub CloseCEandOD()
Shell "TASKKILL /F /IM od.exe /T"
Shell "TASKKILL /F /IM ce.exe /T"
Shell "TASKKILL /F /IM �ᰮ�ƽ�.exe /T"
End Sub
