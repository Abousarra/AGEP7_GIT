VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Archives_AS 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   1935
      Left            =   1320
      ScaleHeight     =   1875
      ScaleWidth      =   2115
      TabIndex        =   22
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Text            =   "Text1"
         Top             =   720
         Width           =   3375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   255
         Left            =   0
         TabIndex        =   25
         Top             =   1200
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   12615
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   " √ﬂÌœ «·≈⁄«œ…"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2000
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   10440
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1600
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   " √ﬂÌœ «·‰”Œ «·«Õ Ì«ÿÌ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "«·€«¡ «·‰”Œ «·«Õ Ì«ÿÌ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         ItemData        =   "Archives_AS.frx":0000
         Left            =   240
         List            =   "Archives_AS.frx":004C
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.DriveListBox Drive2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4320
         TabIndex        =   14
         Top             =   300
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00008000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1980
         Width           =   3135
      End
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   " √ﬂÌœ › Õ ”‰… œ—«”Ì… ÃœÌœ…"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1980
         UseMaskColor    =   -1  'True
         Width           =   3135
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   " √ﬂÌœ «·‰”Œ «·«Õ Ì«ÿÌ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.DriveListBox Drive1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4320
         TabIndex        =   3
         Top             =   1400
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   " √ﬂÌœ «” —Ã«⁄ «·‰”Œ «·«Õ Ì«ÿÌ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1400
         UseMaskColor    =   -1  'True
         Width           =   3135
      End
      Begin VB.Shape Shape1 
         Height          =   1215
         Index           =   4
         Left            =   10320
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "≈⁄«œ… ”‰… œ—«”Ì…"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   10320
         TabIndex        =   26
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·œﬁ«∆ﬁ «·›«’· ⁄‰ «· Œ“Ì‰ «· «·Ì"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   18
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "›«—ﬁ «·”«⁄«  »Ì‰ ﬂ· ‰”ŒÌ‰ «Õ Ì«ÿÌ‰"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   360
         Width           =   3015
      End
      Begin VB.Line Line1 
         X1              =   8160
         X2              =   8160
         Y1              =   240
         Y2              =   1200
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÊÕœ… «· ﬁ”Ì„ „ﬂ«‰ «· Œ“Ì‰"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   15
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·‰”Œ «·«Õ Ì«ÿÌ «· ·ﬁ«∆Ì"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   13
         Top             =   600
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Index           =   3
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   10095
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Index           =   2
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   9255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "› Õ ”‰… œ—«”Ì… ÃœÌœ…"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7920
         TabIndex        =   9
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Index           =   0
         Left            =   10320
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Index           =   1
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   9255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "«·‰”Œ «·«Õ Ì«ÿÌ «·ÌœÊÌ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10440
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁÌ«„ »«” —Ã«⁄ «·‰”Œ «·«Õ Ì«ÿÌ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   7
         Top             =   1440
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   6375
      Left            =   4440
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      BackColor       =   32768
      BackColorFixed  =   32768
      BackColorBkg    =   32768
      RightToLeft     =   -1  'True
      FillStyle       =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   6375
      Left            =   8640
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      BackColor       =   32768
      BackColorFixed  =   32768
      BackColorBkg    =   32768
      RightToLeft     =   -1  'True
      FillStyle       =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grd3 
      Height          =   6375
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   11245
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      BackColor       =   32768
      BackColorFixed  =   32768
      BackColorBkg    =   32768
      RightToLeft     =   -1  'True
      FillStyle       =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "√—‘›… «·”‰Ì‰ «·œ—«”Ì…"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "Archives_AS"
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
Dim Text As String
Dim PicFilev As String
Dim strStream As ADODB.Stream
Dim fName As String
Public coo As ADODB.Connection
Public nn As ADODB.Recordset
Public cooo As ADODB.Connection
Public anb As ADODB.Recordset
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Function conn()
Set coo = New ADODB.Connection
Set nn = New ADODB.Recordset
coo.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
coo.ConnectionString = App.Path & "\ANNEES.mdb"
coo.Open
nn.Open "select*from Tannees", coo, adOpenKeyset, adLockOptimistic
End Function
Function connn()
Set cooo = New ADODB.Connection
Set anb = New ADODB.Recordset
cooo.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
cooo.ConnectionString = App.Path & "\AGEP7.mdb"
cooo.Open
anb.Open "select*from Etablissement", cooo, adOpenKeyset, adLockOptimistic
End Function
Private Sub bases()
Dim result As Long, fileop As SHFILEOPSTRUCT
With fileop
        .hWnd = Me.hWnd
        .wFunc = FO_COPY
        .pFrom = Text1.Text & "\*.mdb" & vbNullChar & vbNullChar
        .pTo = App.Path & vbNullChar & vbNullChar
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)

MsgBox "OpÈration est effectuÈe avec succÈs", vbInformation
End Sub
Private Sub chargcombo1()
Combo1.Clear
Call cont
Do While Not an.EOF
Combo1.AddItem an!ann
an.MoveNext
Loop
End Sub
Private Sub chargcombo3()
Combo3.Clear
Call conn
Do While Not nn.EOF
'If nn!act = "0" Then
Combo3.AddItem nn!ann
'End If
nn.MoveNext
Loop
End Sub

Private Sub Combo1_Change()
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
Else
Combo1.BackColor = &H8080FF
End If

End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Combo2_Change()
Dim i As Double
i = Combo2.Text
i = (i * 60)
Label8.Caption = i
End Sub

Private Sub Combo2_Click()
Combo2_Change
End Sub

Private Sub Command1_Click()
'On Error GoTo p
Dim result As Long, fileop As SHFILEOPSTRUCT
CommonDialog1.CancelError = True
CommonDialog1.DialogTitle = "Coisir la place d'enregistrement"
CommonDialog1.Filter = "|*.*|"
CommonDialog1.FileName = Interface.Caption
CommonDialog1.ShowSave
Text = CommonDialog1.FileName
'Replace the 'C:\MyDir' below with the name of the directory you want to delete.
'x = DelTree(CommonDialog1.FileName)
'Select Case x
'Case 0: MsgBox "Deleted"
'Case -1: MsgBox "Invalid Directory"
'Case Else: MsgBox "An Error was occured"
'End Select
With fileop
        .hWnd = Me.hWnd
        .wFunc = FO_COPY
        .pFrom = App.Path & "\*.mdb" & vbNullChar & vbNullChar
        .pTo = CommonDialog1.FileName & vbNullChar & vbNullChar
    'If Check2.Value = 1 Then
     '   .pFrom = App.Path & "\*.txt" & vbNullChar & vbNullChar
     '   .pTo = CommonDialog1.FileName & vbNullChar & vbNullChar
   ' End If
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)
If result <> 0 Then
      MsgBox "Vous avez annulÈ la rÈcupÈration des bases", vbExclamation
Else
        If fileop.fAnyOperationsAborted <> 0 Then
                    ' MsgBox "Operation Failed"
         End If
End If

MsgBox "OpÈration est effectuÈe avec succÈs", vbInformation
'Call Coder
 Exit Sub
P:
MsgBox "Erreur de souvegarder", vbExclamation


End Sub

Private Sub Command2_Click()
'On Error GoTo p
Text1.Text = Drive1.Drive & Interface.Caption
Text = Text1.Text
Call bases
'Call Coder
Exit Sub
P:
MsgBox "Êﬁ⁄ Œÿ√ «À‰«¡ «⁄«œ… «·‰”Œ «·«Õ Ì«ÿÌ —»„« ÌﬂÊ‰ «·„”«— Œÿ√, «·—Ã«¡ «⁄«œ… «·„Õ«Ê·…", vbExclamation

End Sub

Private Sub Command3_Click()
grd1.Visible = False
grd2.Visible = False
grd3.Visible = False
Call chargegrd1_5
Call chargegrd2_5
Call chargegrd3_T
grd1.Visible = True
grd2.Visible = True
grd3.Visible = True

End Sub

Private Sub Command4_Click()
Dim result As Long, fileop As SHFILEOPSTRUCT
Dim Security As SECURITY_ATTRIBUTES
Dim x$
Dim y$
Dim i As Double
If Combo2.Text = "" Then
MsgBox "«·—Ã«¡  ÕœÌœ ›«—ﬁ «·”«⁄«  »Ì‰ ﬂ· ‰”ŒÌ‰ «Õ Ì«ÿÌÌ‰", vbCritical
Exit Sub
End If
i = Combo2.Text
i = (i * 60)
Label9.Caption = Drive2.Drive
vg = Mid$(Label9.Caption, 1, 2)
Text1.Text = vg
Text1.Text = Text1.Text & "\" & Interface.Caption
Text = Text1.Text
y$ = dir$(Text1.Text & "\*.mdb")
If y$ <> "" Then
Kill Text1.Text & "\*.mdb"
End If
x$ = ""
x$ = dir$(Text1.Text & "\*.mdb")
If x$ = "" Then
'Create a directory dossier images
Ret& = CreateDirectory(Text1.Text, Security)
End If
'Text1.Text = Drive2.Drive & "\" & Interface.Caption
Call cont
eb!prt = Text1.Text
eb!nbh = i
eb!dbt = Label8.Caption
eb!act = "1"
eb.Update
Call cont
With fileop
        .hWnd = Me.hWnd
        .wFunc = FO_COPY
        .pFrom = App.Path & "\*.mdb" & vbNullChar & vbNullChar
        .pTo = Text1.Text & vbNullChar & vbNullChar
    'If Check2.Value = 1 Then
     '   .pFrom = App.Path & "\*.txt" & vbNullChar & vbNullChar
     '   .pTo = CommonDialog1.FileName & vbNullChar & vbNullChar
   ' End If
        .fFlags = FOF_SIMPLEPROGRESS Or FOF_FILESONLY
End With
result = SHFileOperation(fileop)
If result <> 0 Then
      MsgBox "Vous avez annulÈ la rÈcupÈration des bases", vbExclamation
      Exit Sub
Else
        If fileop.fAnyOperationsAborted <> 0 Then
                    ' MsgBox "Operation Failed"
         End If
End If
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation
Interface.Timer1.Enabled = True
 Exit Sub
P:
MsgBox "Erreur de souvegarder", vbExclamation


End Sub


Private Sub Command5_Click()
Call cont
eb!act = "0"
eb.Update
Interface.Timer1.Enabled = False
End Sub

Private Sub Command6_Click()
Call conn
Do While Not nn.EOF
If nn!ann = Combo3.Text Then
Start_UP.Label1.Caption = nn!ann
Interface.SBB1.Panels(1).Text = nn!ann
MsgBox "«‰  «·¬‰  ⁄„· ⁄·Ï «—‘Ì› ”‰… " + Combo3.Text
Archives_AS.Hide
Exit Sub
End If
nn.MoveNext
Loop
'Dim y$
'g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–« «·”‰…ø", vbInformation + vbYesNo + arabic, "AGEP7")
'If g = vbYes Then
'Call conn
'Do While Not nn.EOF
'If Combo3.Text = nn!ann Then
'nn.Delete
'y$ = App.Path & "\" & Combo3.Text & ".mdb"
'If y$ <> "" Then
'Kill App.Path & "\" & Combo3.Text & ".mdb"
'End If
'Call chargcombo3
'MsgBox " „ «·Õ–› »‰Ã«Õ", vbInformation
'Exit Sub
'End If
'nn.MoveNext
'Loop
'End If
End Sub

Private Sub Command9_Click()
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim tx5 As String
Dim tx6 As String
Dim tx7 As String
Dim tx8 As String
Dim tx9 As String
Dim tx10 As String
Dim tx11 As String
Dim tx12 As String
Dim tx13 As String
Dim tx14 As String
If Combo1.Text = "" Then
MsgBox "«·—Ã«¡ «Œ Ì«— «·”‰… «·œ—«”Ì… «· Ì ”Ì „ › ÕÂ«", vbCritical + arabic
Exit Sub
End If
Call conn
Do While Not nn.EOF
If Combo1.Text = nn!ann Then
MsgBox "€Ì— „„ﬂ‰...«·”‰… «·œ—«”Ì… «· Ì «œŒ· „  „ › ÕÂ« ”«»ﬁ«", vbCritical + arabic
Exit Sub
End If
nn.MoveNext
Loop
Call conn
Do While Not nn.EOF
nn!act = "0"
nn!sup = "0"
nn.Update
nn.MoveNext
Loop
nn.AddNew
nn!ann = Combo1.Text
nn!act = "1"
nn!sup = "0"
nn.Update
Call cont
tx1 = eb!eta
tx2 = eb!ann
tx3 = eb!moi
tx4 = eb!pce
tx5 = eb!pcp
tx6 = eb!ser
tx7 = eb!sri
tx8 = eb!rec
tx9 = eb!rcu
tx10 = eb!gch
tx11 = eb!prt
tx12 = eb!nbh
tx13 = eb!dbt
tx14 = eb!act
Call connn
anb!eta = tx1
anb!ann = Combo1.Text
anb!moi = tx3
anb!pce = tx4
anb!pcp = tx5
anb!ser = tx6
anb!sri = tx7
anb!rec = tx8
anb!rcu = tx9
anb!gch = tx10
anb!prt = tx11
anb!nbh = tx12
anb!dbt = tx13
anb!act = tx14
anb.Update
Call connn
cooo.Close
FileCopy App.Path & "\AGEP7.mdb", App.Path & "\" & Combo1.Text & ".mdb"
MsgBox " „ › Õ ”‰… œ—«”Ì… ÃœÌœ… »‰Ã«Õ , ‰”√· «··Â √‰  ﬂÊ‰ „ﬂ··… »«·‰Ã«Õ.... ”  „ ≈⁄«œ…  ‘€Ì· «·»—‰«„Ã „‰ ÃœÌœ", vbInformation
Call unloadforms
Interface.Hide
Unload Start_UP
Start_UP.Show
End Sub

Private Sub Form_Load()
Dim j As Double
Me.Top = 0
Me.Left = 0
Call chargcombo1
Call chargcombo3
Label9.Caption = eb!prt
j = eb!nbh
Combo2.Text = (j / 60)
Label8.Caption = eb!dbt
vg = Mid$(Label9.Caption, 1, 2)
Drive2.Drive = vg

End Sub


Private Sub chargegrd1_5()
Dim h As String
Dim i As Double
Dim j As Double
Dim tx1 As String
Dim tx2 As String
Dim P As Double
Dim s As Double
grd1.Clear
grd1.Cols = 3
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1000
grd1.ColWidth(2) = 2700
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "—.«· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«·„»·€"
i = 1
j = 0
s = 0
tx2 = ""
Call cont3
grd1.Rows = ce3.RecordCount + 3
Do While Not ce3.EOF
tx1 = ce3!ser
If j = 1 And tx1 <> tx2 Then
grd1.Row = i
grd1.Col = 1
grd1.Text = tx2
grd1.Col = 2
grd1.Text = s
i = i + 1
s = 0
End If
P = ce3!pay
s = s + P
tx2 = ce3!ser
j = 1
ce3.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
h = (i - 1)
MsgBox "1 ---- " + h
End Sub
Private Sub chargegrd2_5()
Dim h As String
Dim i As Double
Dim j As Double
Dim tx1 As String
Dim tx2 As String
Dim P As Double
Dim s As Double
grd2.Clear
grd2.Cols = 3
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1000
grd2.ColWidth(2) = 2700
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.Row = 0
grd2.Col = 1
grd2.Text = "—.«· ”·”·Ì"
grd2.Col = 2
grd2.Text = "«·„»·€"
i = 1
j = 0
s = 0
tx2 = ""
Call cont
grd2.Rows = ct.RecordCount + 3
Do While Not ct.EOF
tx1 = ct!sri
If j = 1 And tx1 <> tx2 Then
grd2.Row = i
grd2.Col = 1
grd2.Text = tx2
grd2.Col = 2
grd2.Text = s
i = i + 1
s = 0
End If
P = ct!pay
s = s + P
tx2 = ct!sri
j = 1
ct.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 1
h = (i - 1)
MsgBox "2 ---- " + h
End Sub
Private Sub chargegrd3_T()
Dim h As String
Dim i As Double
Dim j As Double
Dim n As Double
Dim m As Double
Dim k As Double
Dim q As Double
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
grd3.Clear
grd3.Cols = 4
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1000
grd3.ColWidth(2) = 1300
grd3.ColWidth(3) = 1400
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.Row = 0
grd3.Col = 1
grd3.Text = "—.«· ”·”·Ì"
grd3.Col = 2
grd3.Text = "«·„»·€ 1"
grd3.Col = 3
grd3.Text = "«·„»·€ 2"
n = grd1.Rows
j = 1
grd3.Rows = n + 3
For i = 1 To n - 1
grd1.Row = i
grd1.Col = 1
tx1 = grd1.Text
grd1.Col = 2
tx2 = grd1.Text
grd2.Row = i
grd2.Col = 1
tx3 = grd2.Text
grd2.Col = 2
tx4 = grd2.Text
If tx2 <> tx4 Then
grd3.Row = j
grd3.Col = 1
grd3.Text = tx1
grd3.Col = 2
grd3.Text = tx2
grd3.Col = 3
grd3.Text = tx4
j = j + 1
End If
Next i
grd3.Rows = j
grd3.Col = 1
grd3.Sort = 1
h = (j - 1)
MsgBox h
End Sub

