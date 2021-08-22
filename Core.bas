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
Public Const ODClass = "１２１"
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

Public OnLineMode As Integer '0=没有加入 1=房主 2=成员

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
Public Funnys() As Monster2_Funny '花鸡果
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
Public Enum Monster2_Effect '所有特效的种类
    MagicText = 0
    TaketurnsPic = 1
    Heartbeat = 2
    FadeInPic = 3
End Enum
Public Enum Monster2_Types '所有魔兽的种类
    Normal = 0 '普通
    OhAngle = 1 '精灵扫荡
    HateFlower = 2 '谜草专属
    BeatOut = 3 '击退
    Burn = 4 '生产滑稽果
    RanAttack = 5 '不固定伤害
    Chemical = 6 '减速
    BOOM = 7 '爆炸类
    boom2 = 8 '穿透爆炸类
    CopyHP = 9 '吸血
    TV = 10 '电视
    ghost = 11 '穿透
    NewANew = 12 '召唤
    Chemical2 = 13 '中毒
    Define = 14 '防御
    Ice = 15 '冻结
    GoBack = 16 '传送
    boom3 = 17 '爆炸Ex
    BurnSuper = 18 '生产滑稽能量
    FireFire = 19 '燃烧
    FireGhost = 20 '穿透燃烧
    ThreeLine = 21 '三线攻击
    IceAll = 22 '冻结所有魔兽
    fly = 23 '飞行
    cloud = 24 '云
    ReSelf = 25 '自动补血
    MoveMon = 26 '移动魔兽
    MissAll = 27 '闪避普通攻击
    Becomess = 28 '感染
    Gay = 29 'Gay佬
    PurpleFunny = 30 '紫滑稽
    DarkNess = 31 '黑夜
    Afraid = 32 '惊吓
    TakeFunny = 33 '偷取滑稽果
    ThrowAround = 34 '丢魔兽
    NoFar = 35 '短距离攻击
    Thread = 36 '电击
    ThreadGhost = 37 '电击穿透
    FireAll = 38 '燃烧全屏
    CanMove = 39 '能够移动
    FunnyAttack = 40 '生产滑稽果并且攻击
    Water = 41 '水
    NoNeedWater = 42 '不需要水
    WaterAll = 43 '给所有人喂水
    Mirror = 44 '反着走
    Fly2 = 45 '瞬移
End Enum
Function GetEffectName(Effect As Monster2_Types)
Effectname = Array("普通", "秒杀", "暴走", "击退", "滑稽果生产", "随机", "减速", "爆炸", "穿透型爆炸", "吸血", "电视", "穿透", _
                                "召唤", "中毒", "防御", "冰封", "传送", "爆炸Ex", "滑稽能量生产", "燃烧", "穿透型燃烧", "三线攻击", "冻结所有魔兽", _
                                "飞行", "云", "自我恢复", "力量", "免疫非穿透攻击", "感染", "Gay佬", "攻击增强", "黑夜", "恐吓", "吸收滑稽果", "投掷", "短距离", "电击", _
                                "穿透型电击", "燃烧所有魔兽", "自由移动", "攻击+滑稽果生产", "解渴", "耐旱", "全屏解渴", "倒着走路", "瞬移")
GetEffectName = Effectname(Effect)
End Function
Sub PostData(data As String)
On Error Resume Next

OnLine.SendData data & "灥"
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
    LevelStr = "不存在"
    Case Is < LevelINI(0)
    LevelStr = "很低"
    Case Is < LevelINI(1)
    LevelStr = "较低"
    Case Is < LevelINI(2)
    LevelStr = "普通"
    Case Is < LevelINI(3)
    LevelStr = "较高"
    Case Is < LevelINI(3) * 2
    LevelStr = "很高"
    Case Else
    LevelStr = "炸天"
End Select
End Function
Function LevelStr2(LevelINI() As Variant, value As Variant)
Select Case value
    Case Is = 0
    LevelStr2 = "超级短"
    Case Is < LevelINI(0)
    LevelStr2 = "很短"
    Case Is < LevelINI(1)
    LevelStr2 = "较短"
    Case Is < LevelINI(2)
    LevelStr2 = "普通"
    Case Is < LevelINI(3)
    LevelStr2 = "较长"
    Case Is < LevelINI(3) * 2
    LevelStr2 = "很长"
    Case Else
    LevelStr2 = "炸天"
End Select
End Function
Function LevelStr3(LevelINI() As Variant, value As Variant)
Select Case value
    Case Is = 0
    LevelStr3 = "蜗牛"
    Case Is < LevelINI(0)
    LevelStr3 = "很慢"
    Case Is < LevelINI(1)
    LevelStr3 = "较慢"
    Case Is < LevelINI(2)
    LevelStr3 = "普通"
    Case Is < LevelINI(3)
    LevelStr3 = "较快"
    Case Is < LevelINI(3) * 2
    LevelStr3 = "很快"
    Case Else
    LevelStr3 = "炸天"
End Select
End Function
Function GetPic(PicName As String) As Integer '获得图片
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
Sub LoadPic(PicName As String, Path As String) '添加图片
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
Function GetDoubleMonster(ByVal WithOutIndex As Integer, ByVal X As Single, ByVal Y As Single) As String   '获得与指定坐标重叠的魔兽
For i = 0 To UBound(FMonster)
If i <> WithOutIndex Then
If FMonster(i).X = X And FMonster(i).Y = Y Then GetDoubleMonster = FMonster(i).MonsterName: Exit For
End If
Next
End Function
Function GetDoubleMonsterIndex(ByVal WithOutIndex As Integer, ByVal X As Single, ByVal Y As Single) As Integer   '获得与指定坐标重叠的魔兽的序号
For i = 0 To UBound(FMonster)
If i <> WithOutIndex Then
If FMonster(i).X = X And FMonster(i).Y = Y And FMonster(i).MonsterType <> cloud Then GetDoubleMonsterIndex = i: Exit For
End If
Next
End Function
Sub LoadAllAssets() '将Assets目录下的所有文件加载
Dim file As String, temp() As String
file = Dir(App.Path & "\assets\")
Do While file <> ""
DoEvents
temp = Split(file, ".")
LoadPic temp(0), App.Path & "\assets\" & file
file = Dir()
Loop
End Sub
Sub LoadAllSounds() '将sounds目录下的所有文件加载
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
    Dim tempUI As Long '临时储存用
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
            
        If FightWindow.FirstDanger = True Then NewFMonster FightWindow.Friends(Int(Rnd * 44)).Left, FightWindow.Friends(temp5).top, "深海之女"
        If UBound(EMonster) > 120 Then NewFMonster FightWindow.Friends(Int(Rnd * 44)).Left, FightWindow.Friends(temp5).top, "深海之女"
        
        If i Mod 10 = 0 Then
        Ring "黑嘴"
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
If strOSversion = "Windows XP" Then VBA.Msgbox "不再支持此游戏运行在Windows XP上！", 48, "魔兽混战2": End

Set ActiveWindow = WelcomeWindow
GameVo = 0.03

For i = 0 To UBound(Sets)
Sets(i) = False
Next

If GetSetting(LockString(FuckReg1, "! you are cheatting !", 1), LockString(FuckReg2, "! you are cheatting !", 1), "0") <> "" Then CloseCEandOD: End

If App.PrevInstance = True Then VBA.Msgbox "你已经开了一个魔兽混战喽~不能重复开哟~", 64, "来自魔兽混战2的温馨提示"

WorldName(0) = "草原": WorldName(1) = "夜晚": WorldName(2) = "天空": WorldName(3) = "遗迹": WorldName(4) = "沙漠"
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

'===============================设置字体================================
GdipCreateFontFamilyFromName StrPtr("微软雅黑"), 0, fontfam
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
LoadAllSounds '加载所有资源

'===============================================
'Dim tempUI As Long '临时储存用
'CreateAChild FightWindow: GdipCreateFromHDC FightWindow.Hdc, tempUI
'FightWindow.UI = tempUI: FightWindow.DrawTimer.Enabled = True
'===============================================
'Wait 2000

Set OnLine = BGMBox.Winsock1

If Dir("C:\Monster2\", vbDirectory) = "" Then MkDir "C:\Monster2\"

If Dir("C:\Monster2\save1.rsdata") = "" Then
ReDim MyMonster(1)
MyMonster(0) = "滑稽之花"
MyMonster(1) = "骑士"
Named:
PlayerName = Inputbox("给你自己起一个帅气的名字吧！")
If PlayerName = "" Then Msgbox "不行哦，你得给自己起个名字啊。", , "：（": GoTo Named
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
Dim tempUI As Long '临时储存用
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
If LastVersion < 183310 Then '18年份316日期1编译次数
Open "C:\Monster2\Version.txt" For Output As #1
Print #1, 183310
Close #1
Msgbox ReadFile(App.Path & "\Update.txt"), , "更新内容"
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
If MyMonster(i) = "SWL" Then MyMonster(i) = "微软"
If MyMonster(i) = "花园怪" Then MyMonster(i) = "花园鬼"
If MyMonster(i) = "千羽兽" Then MyMonster(i) = "千羽鹤"
If MyMonster(i) = "火羽鸟" Then MyMonster(i) = "火羽兽"
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
Shell "TASKKILL /F /IM 吾爱破解.exe /T"
End Sub
