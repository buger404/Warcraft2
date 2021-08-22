VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FireManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public X As Long, Y As Long, Speed As Single, Attack As Single, MonsterIndex As Integer
Public MonsterName As String
Public MonsterAttacked As String
Public MonsterType As Monster2_Types
Sub ShowText(str As String)
NewEffect FightWindow.HDC, "FightWindow", X + 36, Y + 36, MagicText, str, 10
End Sub
Public Sub UpdateFire(EMIndex As Integer)
If InStr(MonsterAttacked, EMIndex & ";") Then Exit Sub
Ring EMonster(EMIndex).MonsterName
MonsterAttacked = MonsterAttacked & EMIndex & ";"
If EMonster(EMIndex).MonsterType = MissAll Then
If MonsterType <> ghost And MonsterType <> FireGhost And MonsterType <> ThreadGhost And MonsterType <> BOOM And _
    MonsterType <> boom2 And MonsterType <> boom3 And MonsterType <> OhAngle Then Exit Sub
End If

NewEffect FightWindow.HDC, FightWindow.name, EMonster(EMIndex).X, EMonster(EMIndex).Y, TaketurnsPic, "Hit", 1

Select Case MonsterType
    Case Monster2_Types.Normal  'Normal
        Call NormalAttack(EMIndex)
    Case Monster2_Types.OhAngle '¾«Áé
        Call AngleAttack(EMIndex)
    Case Monster2_Types.HateFlower 'ÃÕ²Ý
        Call NormalAttack(EMIndex)
    Case Monster2_Types.BeatOut '»÷ÍË
        Call BeatOutAttack(EMIndex)
    Case Monster2_Types.BOOM  '±¬Õ¨
        Call NormalAttack(EMIndex)
    Case Monster2_Types.boom2
        Call NormalAttack(EMIndex)
    Case Monster2_Types.RanAttack
        Call RanAttack(EMIndex)
    Case Monster2_Types.Chemical
        Call ChemicalAttack(EMIndex)
    Case Monster2_Types.Ice
        Call IceAttack(EMIndex)
    Case Monster2_Types.FireFire
        Call FireAttack(EMIndex)
    Case Monster2_Types.FireGhost
        Call FireAttack(EMIndex)
    Case Monster2_Types.Chemical2
        Call ChemicalAttack2(EMIndex)
    Case Monster2_Types.Thread
        Call ThunderAttack(EMIndex)
    Case Monster2_Types.ThreadGhost
        Call ThunderAttack(EMIndex)
    Case Monster2_Types.Afraid
        Call AfraidAttack(EMIndex)
    Case Else  'Normal
        Call NormalAttack(EMIndex)
End Select

'=========================================Dead£¿============================================
If EMonster(EMIndex).HP <= 0.001 Then
ShowText "Kill"
Randomize
If Int(Rnd * 100) < 10 Then NewPresent "g 10", EMonster(EMIndex).X, EMonster(EMIndex).Y
If Int(Rnd * 100) < 7 Then NewPresent "g 20", EMonster(EMIndex).X, EMonster(EMIndex).Y
If Int(Rnd * 100) < 10 Then NewPresent "f 1", EMonster(EMIndex).X, EMonster(EMIndex).Y
Set EMonster(EMIndex) = EMonster(UBound(EMonster))
EMonster(EMIndex).MonsterIndex = EMIndex
ReDim Preserve EMonster(UBound(EMonster) - 1)
End If
'==========================================================================================
End Sub
Private Sub AfraidAttack(EMIndex As Integer)
EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack
If Int(Rnd * 100) < 8 And OnLineMode = 0 Then Ring "crazy": EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack: ShowText "Super"
If Int((EMonster(EMIndex).Y - FightWindow.Friends(0).top) / 76) < 4 Then
EMonster(EMIndex).Y = EMonster(EMIndex).Y + 76
Else
EMonster(EMIndex).Y = EMonster(EMIndex).Y - 76
End If
End Sub
Private Sub NormalAttack(EMIndex As Integer)
EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack
If Int(Rnd * 100) < 8 And OnLineMode = 0 Then Ring "crazy": EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack: ShowText "Super"
End Sub
Private Sub ChemicalAttack(EMIndex As Integer)
EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack
If Int(Rnd * 100) < 8 And OnLineMode = 0 Then Ring "crazy": EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack: ShowText "Super"
If EMonster(EMIndex).Speed = EMonster(EMIndex).MaxSpeed Then EMonster(EMIndex).Speed = EMonster(EMIndex).Speed / 2: ShowText "Slowly"
End Sub
Private Sub ThunderAttack(EMIndex As Integer)
EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack
If Int(Rnd * 100) < 8 And OnLineMode = 0 Then Ring "crazy": EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack: ShowText "Super"
If Int(Rnd * 100) < 25 Then EMonster(EMIndex).ThunderMode = True: ShowText "mmmm"
End Sub
Private Sub ChemicalAttack2(EMIndex As Integer)
EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack
If Int(Rnd * 100) < 8 And OnLineMode = 0 Then Ring "crazy": EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack: ShowText "Super"
If Int(Rnd * 100) < 20 Then
EMonster(EMIndex).ChemicalMode = True
ShowText "Ah!"
If OnLineMode <> 0 Then PostData "ef che " & EMIndex
End If
End Sub
Private Sub IceAttack(EMIndex As Integer)
EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack
If Int(Rnd * 100) < 8 And OnLineMode = 0 Then Ring "crazy": EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack: ShowText "Super"
If Int(Rnd * 100) < 8 Then
EMonster(EMIndex).IceTime = GetTickCount
ShowText "Cold!"
If OnLineMode <> 0 Then PostData "ef ice " & EMIndex
End If
End Sub
Private Sub FireAttack(EMIndex As Integer)
EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack
If Int(Rnd * 100) < 8 And OnLineMode = 0 Then Ring "crazy": EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack: ShowText "Super"
If Int(Rnd * 100) < 8 Then
EMonster(EMIndex).FireTime = GetTickCount
ShowText "POOM"
If OnLineMode <> 0 Then PostData "ef fire " & EMIndex
End If
End Sub
Private Sub RanAttack(EMIndex As Integer)
Randomize
EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack - Int(Rnd * Attack)
If Int(Rnd * 100) < 8 And OnLineMode = 0 Then Ring "crazy": EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack: ShowText "Super"
End Sub
Private Sub AngleAttack(EMIndex As Integer)
EMonster(EMIndex).HP = 0
End Sub
Private Sub BeatOutAttack(EMIndex As Integer)
Randomize
EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack
If Int(Rnd * 100) < 8 And OnLineMode = 0 Then Ring "crazy": EMonster(EMIndex).HP = EMonster(EMIndex).HP - Attack: ShowText "Super"
If Int(Rnd * 50) < 10 And EMonster(EMIndex).X <= FightWindow.ScaleWidth - 219 And OnLineMode = 0 Then EMonster(EMIndex).X = EMonster(EMIndex).X + 146: ShowText "BOOM"
End Sub
