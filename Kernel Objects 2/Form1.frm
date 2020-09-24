VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "KERNEL OBJECTS /USING UNDOCUMENTED API NTDLL by Vanja Fuckar"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   8460
   ScaleWidth      =   14190
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lw1 
      Height          =   7695
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   13573
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Populate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "WARNING:CLOSE DIRECT CD OR NERO IF RUNNING!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   14175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
lw1.ListItems.Clear




Dim NAMEOBJ As String
Dim SYSLEN As Long
Dim SYS() As Long
ReDim SYS(&H100000)
Dim ret As Long
ret = NtQuerySystemInformation(&H10&, SYS(0), &H100000, SYSLEN)
ReDim Preserve SYS(SYSLEN / 4 - 1)


Dim u As Long

Dim BFLAG As Byte
Dim TYPO As Byte

Dim HDL As Long
Dim InHDL As Long

Dim TYPEOBJ As String

For u = 0 To UBound(SYS) / 4 - 1


CopyMemory TYPO, SYS(u * 4 + 2), 1 'Get Type
CopyMemory BFLAG, ByVal VarPtr(SYS(u * 4 + 2)) + 1, 1 'Get Flag

CopyMemory HDL, ByVal VarPtr(SYS(u * 4 + 2)) + 2, 2 'GET Handle

InHDL = InheritObjectHandle(SYS(u * 4 + 1), HDL)


If InHDL <> 0 Then

    NAMEOBJ = KObjectName(InHDL): TYPEOBJ = StingFromObjectType(InHDL)
    If StrComp(UCase(TYPEOBJ), "THREAD") = 0 Then NAMEOBJ = "Thread Start At Address:" & Hex(GetThreadStartAddress(InHDL))

Else

    TYPEOBJ = "<CANNOT INHERIT OBJECT>"
    NAMEOBJ = "<CANNOT QUERY INFORMATION>"

End If

AddLV (SYS(u * 4 + 1)), HDL, TYPEOBJ, SYS(u * 4 + 3), SYS(u * 4 + 4), NAMEOBJ

CloseHandle InHDL


NAMEOBJ = ""

Next u





End Sub
Private Sub AddLV(ByVal PID As Long, ByVal Handle As Long, ByVal ObjectTP As String, ByVal MEMAdr As Long, ByVal AcFlag As Long, ByVal ObjectNM As String)
Dim LITM As ListItem
Set LITM = lw1.ListItems.Add(, , CStr(PID))
LITM.SubItems(1) = Hex(Handle)
LITM.SubItems(2) = ObjectTP
LITM.SubItems(3) = Hex(MEMAdr)
LITM.SubItems(4) = Hex(AcFlag)
LITM.SubItems(5) = ObjectNM
Set LITM = Nothing
End Sub







Private Sub Form_Load()
Top = (Screen.Height - Height) / 2
Left = (Screen.Width - Width) / 2

lw1.ColumnHeaders.Add , , "Process Id", 1200
lw1.ColumnHeaders.Add , , "Handle", 1200
lw1.ColumnHeaders.Add , , "Object Type", 4000
lw1.ColumnHeaders.Add , , "Memory Address", 1600
lw1.ColumnHeaders.Add , , "Access Flags", 1600
lw1.ColumnHeaders.Add , , "Object Name", 4000
End Sub
