VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.MDIForm Interface 
   BackColor       =   &H8000000C&
   Caption         =   "AGEP 2017"
   ClientHeight    =   10095
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16215
   Icon            =   "Interface.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "Interface.frx":0BC2
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   720
      Top             =   2520
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   15
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   16215
      TabIndex        =   1
      Top             =   0
      Width           =   16215
      Begin VB.Timer Timer10 
         Interval        =   500
         Left            =   0
         Top             =   0
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   615
         Left            =   720
         TabIndex        =   2
         Top             =   120
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1085
         _Version        =   393216
         Format          =   124977153
         CurrentDate     =   42637
      End
   End
   Begin MSComctlLib.StatusBar SBB1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9720
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1320
      OleObjectBlob   =   "Interface.frx":607F
      Top             =   1200
   End
End
Attribute VB_Name = "Interface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const FO_MOVE As Long = &H1
Private Const FO_COPY As Long = &H2
Private Const FO_DELETE As Long = &H3
Private Const FO_RENAME As Long = &H4
Private Const FOF_MULTIDESTFILES As Long = &H1
Private Const FOF_CONFIRMMOUSE As Long = &H2
Private Const FOF_SILENT As Long = &H4
Private Const FOF_RENAMEONCOLLISION As Long = &H8
Private Const FOF_NOCONFIRMATION As Long = &H10
Private Const FOF_WANTMAPPINGHANDLE As Long = &H20
Private Const FOF_CREATEPROGRESSDLG As Long = &H0
Private Const FOF_ALLOWUNDO As Long = &H40
Private Const FOF_FILESONLY As Long = &H80
Private Const FOF_SIMPLEPROGRESS As Long = &H100
Private Const FOF_NOCONFIRMMKDIR As Long = &H200

Private Type SHFILEOPSTRUCT
     hWnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Long
     fAnyOperationsAborted As Long
     hNameMappings As Long
     lpszProgressTitle As String
End Type

Private Declare Function SHFileOperation Lib "Shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Dim da As String
Dim dy As String
Dim mont As String
Dim ye As String
Dim myDate As String
Dim x As Integer
Dim yy As Integer
Private Sub MDIForm_Load()
On Error Resume Next
Dim j As Double
Me.Top = 100
Me.Left = 100
Skin1.LoadSkin App.Path & "\18.skn"
Skin1.ApplySkin Me.hWnd
Call chargepanels
Call dater
Call cont
Me.Caption = eb!gch
j = eb!act
Timer1.Enabled = j
login.Show
'Utilisateurs.Show
End Sub
Private Sub chargepanels()
On Error Resume Next
Call cont
SBB1.Panels(1).Width = 1300
'SBB1.Panels(1).Text = "2017-2018"
SBB1.Panels(1).Alignment = sbrRight
SBB1.Panels.Add 2
SBB1.Panels(2).Width = 1200
SBB1.Panels(2).Text = "«·”‰… «·œ—«”Ì…"
SBB1.Panels(2).Alignment = sbrRight
SBB1.Panels.Add 3
SBB1.Panels(3).Width = 4500
SBB1.Panels(3).Text = eb!eta
SBB1.Panels(3).Alignment = sbrRight
SBB1.Panels.Add 4
SBB1.Panels(4).Width = 800
SBB1.Panels(4).Text = "«·„ƒ””…"
SBB1.Panels(4).Alignment = sbrRight
SBB1.Panels.Add 5
SBB1.Panels(5).Width = 1500
SBB1.Panels(5).Text = ""
SBB1.Panels(5).Alignment = sbrRight
SBB1.Panels.Add 6
SBB1.Panels(6).Width = 800
SBB1.Panels(6).Text = ""
SBB1.Panels(6).Alignment = sbrRight
SBB1.Panels.Add 7
SBB1.Panels(7).Width = 4200
SBB1.Panels(7).Text = ""
SBB1.Panels(7).Alignment = sbrRight
SBB1.Panels.Add 8
SBB1.Panels(8).Width = 1000
SBB1.Panels(8).Text = "»—„Ã… Ê ’„Ì„: √»Ê»ﬂ— √Õ„œÊ «·€“«·Ì 22660920-33440920"
SBB1.Panels(8).Alignment = sbrRight
End Sub
Private Sub dater()
On Error Resume Next
da = DT1.DayOfWeek
dy = DT1.Day
mont = DT1.Month
ye = DT1.Year
'********** Days
If da = 1 Then
da = "«·«Õœ"
ElseIf da = 2 Then
da = "«·«À‰Ì‰"
ElseIf da = 3 Then
da = "«·À·«À«¡"
ElseIf da = 4 Then
da = "«·«—»⁄«¡"
ElseIf da = 5 Then
da = "«·Œ„Ì”"
ElseIf da = 6 Then
da = "«·Ã„⁄…"
ElseIf da = 7 Then
da = "«·”» "
End If
'********** Months
If mont = 1 Then
mont = "Ì‰«Ì—"
ElseIf mont = 2 Then
mont = "›»—«Ì—"
ElseIf mont = 3 Then
mont = "„«—”"
ElseIf mont = 4 Then
mont = "«»—Ì·"
ElseIf mont = 5 Then
mont = "„«ÌÊ"
ElseIf mont = 6 Then
mont = "ÌÊ‰ÌÊ"
ElseIf mont = 7 Then
mont = "ÌÊ·ÌÊ"
ElseIf mont = 8 Then
mont = "«€”ÿ”"
ElseIf mont = 9 Then
mont = "”» „»—"
ElseIf mont = 10 Then
mont = "«ﬂ Ê»—"
ElseIf mont = 11 Then
mont = "‰Ê›„»—"
ElseIf mont = 12 Then
mont = "œÌ”„»—"
End If
'********** Date
myDate = da + " " + dy + " " + mont + " " + ye
'Me.Caption = myDate + " : " + Time$
'SBB1.Panels(16).Text = Time$
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim Answer As Integer
   Answer = MsgBox("Â·  ÊœÊ‰ Õﬁ« «·Œ—ÊÃ ⁄‰ «·»—‰«„Ãø", _
   vbQuestion + vbYesNo, " √ﬂÌœ")
   If Answer = vbNo Then
   Cancel = -1
   Else
End
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim tx As String
Dim result As Long, fileop As SHFILEOPSTRUCT
Dim i As Double
Dim j As Double
Dim Security As SECURITY_ATTRIBUTES
Dim x$
Dim y$
i = eb!nbh
j = eb!dbt
tx = eb!prt
j = j - 1
If j = 0 Then
j = i
y$ = dir$(tx & "\*.mdb")
If y$ <> "" Then
Kill tx & "\*.mdb"
End If
x$ = ""
x$ = dir$(tx & "\*.mdb")
If x$ = "" Then
'Create a directory dossier images
Ret& = CreateDirectory(tx, Security)
End If
With fileop
        .hWnd = Me.hWnd
        .wFunc = FO_COPY
        .pFrom = App.Path & "\*.mdb" & vbNullChar & vbNullChar
        .pTo = tx & vbNullChar & vbNullChar
    'If Check2.Value = 1 Then
     '   .pFrom = App.Path & "\*.txt" & vbNullChar & vbNullChar
     '   .pTo = CommonDialog1.FileName & vbNullChar & vbNullChar
   ' End If
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)
If result <> 0 Then
      MsgBox "«·—Ã«¡ «· √ﬂœ „‰ ’Õ… „ﬂ«‰ «· Œ“Ì‰ «· ·ﬁ«∆Ì", vbExclamation
      Exit Sub
Else
        If fileop.fAnyOperationsAborted <> 0 Then
                    ' MsgBox "Operation Failed"
         End If
End If

End If
eb!dbt = j
eb.Update
 Exit Sub
P:
MsgBox "«·—Ã«¡ «· √ﬂœ „‰ ’Õ… „ﬂ«‰ «· Œ“Ì‰ «· ·ﬁ«∆Ì", vbExclamation
End Sub

Private Sub Timer10_Timer()
On Error Resume Next
DT1.Value = Date
Call dater

End Sub
