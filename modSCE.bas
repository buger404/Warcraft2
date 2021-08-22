Attribute VB_Name = "modSCE"
Public Const GiveYouMyDear As String = "喂喂 那边那个破解的，我知道你在看我。"
'Spicy chicken encryption 加密算法（简称SCE）
Function LockString(ByVal txt As String, ByVal key As String, flag As Integer)
Dim KeyP As Integer, temp As Variant
KeyP = 1
If flag = 0 Then
For i = 1 To Len(txt)
temp = Asc(Mid(txt, i, 1)) Xor Asc(Mid(key, KeyP, 1))
temp = temp + Asc(Mid(key, KeyP + 1, 1))
temp = Hex(temp)
If Len(temp) = 1 Then
temp = "=" & temp
End If
temp = Step3(temp, 0, Asc(Mid(key, KeyP, 1)) Mod 10)
LockString = LockString & temp
KeyP = KeyP + 1
If KeyP > Len(key) - 2 Then KeyP = 1
Next
Else
For i = 1 To Len(txt) Step 2
temp = Step3(Mid(txt, i, 2), 1, Asc(Mid(key, KeyP, 1)) Mod 10)
temp = CLng("&H" & Replace(temp, "=", ""))
temp = temp - Asc(Mid(key, KeyP + 1, 1))
temp = temp Xor Asc(Mid(key, KeyP, 1))
temp = Chr(temp)
LockString = LockString & temp
KeyP = KeyP + 1
If KeyP > Len(key) - 2 Then KeyP = 1
Next
End If
End Function
Function Step3(ByVal str As String, flag As Integer, Optional key)                    '0为加密
    Randomize
    
    a = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
    Dim RockChrs(9)
    RockChrs(0) = "41stuABQdefTUV6RSnopqrG5OPXwklvNYZabcx90HIjLM23ghiCDyzm78EFWJK"
    RockChrs(1) = "X1stuABQwzm78EFYZabcklv4defTUV6RSnOPNLM2HIjWJK3ghiCDyx9opqrG50"
    RockChrs(2) = "ghiCefSnopqr48EFWJKU90HIjLM23G51stuBQdTcxm7OPDyzAV6RXwklvNYZab"
    RockChrs(3) = "WJKQdefTUMYZabcqSn1sR4BvNOtuAopklVG5236PXrghiCDyzm78EFwx90HIjL"
    RockChrs(4) = "FaRSnopqrG546m78EOQdefTUVkltuABx9WJghivNYDyzbcPXZ1sCK0HIjLM23w"
    RockChrs(5) = "tuAdV64wRefOPXm78EG5vNYZkpBQqrlTUSDy1sJKzano23ghiCbcx90FWHIjLM"
    RockChrs(6) = "ZaUV6RSnm78EFYqrGW1stuAop4MdefOTJK5BQHIjLx902CDyzbcghi3PXwklvN"
    RockChrs(7) = "QNYZABqrG5a6RSnop41stuOdghiKPEFWJefTUVXwklvCDyzm78LM23x90HIjbc"
    RockChrs(8) = "Snklv23xp4cCDyXwgh6RUV0HIjbWJefTQNYZA9iKPEF1azm78stuOdoBqrG5LM"
    RockChrs(9) = "rGAFWJefBqstx90HIjbcTUVnop5a6M2XwuOd41QNYKPEghi378LCDyklvRSZmz"
    
    d = 0
    f = 1
    For p = 1 To Len(str)
    If p Mod 15 = 0 Then DoEvents
        key2 = Val(Mid(key, f, 1))
        d = 0
        For s = 1 To Len(a)
            If flag = 0 And Mid(a, s, 1) = Mid(str, p, 1) Then c = c & Mid(RockChrs(key2), s, 1): d = 1
            If flag = 1 And Mid(RockChrs(key2), s, 1) = Mid(str, p, 1) Then c = c & Mid(a, s, 1): d = 1
        Next
        f = f + 1
        If d = 0 Then c = c & Mid(str, p, 1)
        If f > Len(key) Then f = 1
    Next
    
    Step3 = c
    
End Function
Function GetKey()
    Randomize
    GetKey = Int(Rnd * 8999999 + 1000000)
End Function


