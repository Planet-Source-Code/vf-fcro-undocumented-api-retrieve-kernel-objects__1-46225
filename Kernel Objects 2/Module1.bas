Attribute VB_Name = "Module1"
Option Explicit

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type FILE_BASIC_INFORMATION
CreationTime As FILETIME
LastAccessTime As FILETIME
LastWriteTime As FILETIME
ChangeTime As FILETIME
FileAttributes As Long
Unknown(5) As Long
End Type







Declare Function NtQueryInformationThread Lib "ntdll.dll" (ByVal ThreadH As Long, ByVal TypeOfInformation As Long, Buffer As Any, ByVal BfrLength As Long, ByRef BfrRequired As Long) As Long

Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const EVENT_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3)

Public Const THREAD_GET_CONTEXT = (&H8)
Public Const THREAD_SET_CONTEXT = (&H10)
Public Const THREAD_SUSPEND_RESUME = (&H2)
Public Const THREAD_TERMINATE = (&H1)
Public Const THREAD_SET_THREAD_TOKEN = (&H80)
Public Const THREAD_SET_INFORMATION = (&H20)
Public Const THREAD_IDLE_TIMEOUT = 10
Public Const THREAD_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)
Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function lstrlenW Lib "kernel32" (lpString As Any) As Long

Public Const DUPLICATE_SAME_ACCESS = &H2
Public Const PROCESS_ALL_ACCESS = &H1F0FFF
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Declare Function NtQueryInformationFile Lib "ntdll" (ByVal FileHandle As Long, IOStatusBlock As Any, FileInformation As Any, ByVal FileInfoLength As Long, ByVal ClassType As Long) As Long

Public Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal ClassType As Long, SYSINFO As Any, ByVal SYSINFOLEN As Long, SYSINFOLENBACK As Any) As Long
Public Declare Function NtQueryObject Lib "ntdll" (ByVal Handle As Long, ByVal ClassType As Long, OBJINFO As Any, ByVal OBJINFOLEN As Long, OBJINFOLENBACK As Any) As Long
Public Declare Function NtDuplicateObject Lib "ntdll" (ByVal SourcePHandle As Long, ByVal SourceHandle As Long, ByVal TargetPHandle As Long, ByRef DuplicateHandle As Long, ByVal Access As Long, ByVal InheritHandle As Long, ByVal Options As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Type OBJECT_BASIC_INFORMATION
Unknown1 As Long
DesiredAccess As Long
HandleCount As Long
ReferenceCount As Long
PagedPoolQuota As Long
NonPagedPoolQuota As Long
Unknown2(31) As Byte
End Type

Public Type SYSTEM_THREAD
KernelTimeLo As Long '       // 100 nsec units
KernelTimeHi As Long
UserTimeLo As Long ' // 100 nsec units
UserTimeHi As Long
CreateTimeLo As Long '       // relative to 01-01-1601
CreateTimeHi As Long
d18 As Long
StartAddress As Long
Tid As Long '               // process/thread ids
Priority As Long
BasePriority As Long
ContextSwitchesCount As Long
ThreadState As Long '      // 2=running, 5=waiting
WaitReason As Long
Reserved01 As Long
End Type

'NOT USED FOR THIS EXAMPLE
Public Function GetObjectCount(ByVal PObj As Long) As Long
Dim BO As OBJECT_BASIC_INFORMATION
Dim LLen As Long
NtQueryObject PObj, 0, BO, Len(BO), LLen
End Function
Public Function GetThreadStartAddress(ByVal ThreadHandle As Long) As Long
Dim Bfrlen As Long
Call NtQueryInformationThread(ThreadHandle, 9, GetThreadStartAddress, 4, Bfrlen)
End Function

Public Function InheritObjectHandle(ByVal PID As Long, ByVal PObj As Long) As Long
Dim PHandle As Long
Dim OUTH As Long
PHandle = OpenProcess(PROCESS_ALL_ACCESS, 1, PID)
NtDuplicateObject PHandle, PObj, -1, OUTH, 0, 0, 2
CloseHandle PHandle
InheritObjectHandle = OUTH
End Function
Public Function KObjectName(ByVal Handle As Long) As String
Dim OIRET As Long
Dim ODATA() As Byte
ReDim ODATA(10000)
Call NtQueryObject(Handle, &H1&, ODATA(0), 10000, OIRET)
If OIRET = 0 Then Exit Function
ReDim Preserve ODATA(OIRET - 1)
If OIRET = 0 Or OIRET < 9 Then Exit Function
KObjectName = Space(OIRET - 8)
CopyMemory ByVal StrPtr(KObjectName), ODATA(8), OIRET - 8
End Function
Public Function StingFromObjectType(ByVal Handle As Long) As String
Dim DTA(1000) As Long
Dim retlen As Long
NtQueryObject Handle, 2, DTA(0), 4000, retlen
retlen = lstrlenW(ByVal DTA(1))
StingFromObjectType = Space(retlen)
CopyMemory ByVal StrPtr(StingFromObjectType), ByVal DTA(1), retlen * 2
End Function
