VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VariableManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Classkey As String
Private Classvalue As Variant
Private Sub Class_Initialize()
Randomize
Classkey = Hex(GetKey)
End Sub
Public Property Get Default() As Variant
Attribute Default.VB_UserMemId = 0
Default = LockString(Classvalue, CLng("&H" & Classkey), 1)
End Property
Public Property Let Default(ByVal vNewValue As Variant)
Classvalue = LockString(vNewValue, CLng("&H" & Classkey), 0)
End Property
