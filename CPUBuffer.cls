VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPUBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long '������ص�API
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Private Const SYSTEM_BASICINFORMATION = 0& '��صĳ���
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Const NO_ERROR = 0
 
'��VB���ɡ�VB���CPUռ����
Private Type LARGE_INTEGER '��ص���������
dwLow As Long
dwHigh As Long
End Type
 
'��VB���ɡ�VB���CPUռ����
Private Type SYSTEM_PERFORMANCE_INFORMATION
liIdleTime As LARGE_INTEGER
dwSpare(0 To 75) As Long
End Type
 
'��VB���ɡ�VB���CPUռ����
Private Type SYSTEM_BASIC_INFORMATION
dwUnknown1 As Long
uKeMaximumIncrement As Long
uPageSize As Long
uMmNumberOfPhysicalPages As Long
uMmLowestPhysicalPage As Long
uMmHighestPhysicalPage As Long
uAllocationGranularity As Long
pLowestUserAddress As Long
pMmHighestUserAddress As Long
uKeActiveProcessors As Long
bKeNumberProcessors As Byte
bUnknown2 As Byte
wUnknown3 As Integer
End Type
 
'��VB���ɡ�VB���CPUռ����
Private Type SYSTEM_TIME_INFORMATION
liKeBootTime As LARGE_INTEGER
liKeSystemTime As LARGE_INTEGER
liExpTimeZoneBias As LARGE_INTEGER
uCurrentTimeZoneId As Long
dwReserved As Long
End Type
 
'��VB���ɡ�VB���CPUռ����
Private lidOldIdle As LARGE_INTEGER
Private liOldSystem As LARGE_INTEGER
Function GetCPUUsage() As Long '���ǽӿڹ���
Dim sbSysBasicInfo As SYSTEM_BASIC_INFORMATION, spSysPerforfInfo As SYSTEM_PERFORMANCE_INFORMATION, stSysTimeInfo As SYSTEM_TIME_INFORMATION, curIdle As Currency, curSystem As Currency, lngResult As Long
 
'��VB���ɡ�VB���CPUռ����
GetCPUUsage = -1
lngResult = NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(sbSysBasicInfo), LenB(sbSysBasicInfo), 0&)
If lngResult <> NO_ERROR Then Exit Function
lngResult = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(stSysTimeInfo), LenB(stSysTimeInfo), 0&)
If lngResult <> NO_ERROR Then Exit Function
lngResult = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(spSysPerforfInfo), LenB(spSysPerforfInfo), ByVal 0&)
If lngResult <> NO_ERROR Then Exit Function
 
'��VB���ɡ�VB���CPUռ����
curIdle = ConvertLI(spSysPerforfInfo.liIdleTime) - ConvertLI(lidOldIdle) '����CPUռ����
curSystem = ConvertLI(stSysTimeInfo.liKeSystemTime) - ConvertLI(liOldSystem)
If curSystem <> 0 Then curIdle = curIdle / curSystem
curIdle = 100 - curIdle * 100 / sbSysBasicInfo.bKeNumberProcessors + 0.5
GetCPUUsage = Int(curIdle)
 
'��VB���ɡ�VB���CPUռ����
lidOldIdle = spSysPerforfInfo.liIdleTime
liOldSystem = stSysTimeInfo.liKeSystemTime
End Function
Private Function ConvertLI(liToConvert As LARGE_INTEGER) As Currency '��LARGE_INTEGER���͵�����ת����Currency����
CopyMemory ConvertLI, liToConvert, LenB(liToConvert)
End Function
Private Sub Class_Initialize() '���ʼ��
Dim stSysTimeInfo As SYSTEM_TIME_INFORMATION, spSysPerforfInfo As SYSTEM_PERFORMANCE_INFORMATION, lngResult As Long
lngResult = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(stSysTimeInfo), LenB(stSysTimeInfo), 0&)
If lngResult <> NO_ERROR Then Exit Sub
lngResult = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(spSysPerforfInfo), LenB(spSysPerforfInfo), ByVal 0&)
If lngResult <> NO_ERROR Then Exit Sub
lidOldIdle = spSysPerforfInfo.liIdleTime
liOldSystem = stSysTimeInfo.liKeSystemTime
End Sub
