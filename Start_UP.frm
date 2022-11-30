VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Start_UP 
   Caption         =   "RÈalisÈ par Aboubekrine El Ghazali"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   120
      ScaleHeight     =   6255
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.ComboBox Combo4 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         ItemData        =   "Start_UP.frx":0000
         Left            =   5760
         List            =   "Start_UP.frx":003D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "«÷€ÿ Â‰« ··œŒÊ· «·Ï «·»—‰«„Ã"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Top             =   4920
         Width           =   7095
      End
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   240
         Top             =   480
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   0
         ScaleHeight     =   1815
         ScaleWidth      =   7575
         TabIndex        =   4
         Top             =   2400
         Width           =   7575
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   120
            Top             =   120
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   480
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   3720
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   2880
            TabIndex        =   5
            Text            =   "Text2"
            Top             =   3720
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            Height          =   1575
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   7335
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Õﬂ„… «·ÌÊ„"
            BeginProperty Font 
               Name            =   "Andalus"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   2400
            TabIndex        =   8
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·”‰… «·œ—«”Ì…"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   1095
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   7335
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   5640
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   930
         Left            =   360
         ScaleHeight     =   930
         ScaleWidth      =   6855
         TabIndex        =   1
         Top             =   1440
         Width           =   6855
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   5640
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·”‰… «·œ—«”Ì… «· Ì ”Ì „ «·⁄„· ⁄·ÌÂ«"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2280
         TabIndex        =   19
         Top             =   2760
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   18
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   4320
         Width           =   7335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÌÊÃœ ›Ì ﬁ«⁄œ… »Ì«‰«  «·”‰… «·œ—«”Ì… "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   4440
         Width           =   3135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2555333"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”Ã·«"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   4440
         Width           =   615
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   1335
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   4800
         Width           =   7335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   7335
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "√Ì «” Œœ«„ ·Â–Â «·‰”Œ… Œ«—Ã Â–« «·«ÿ«— Ì⁄ »—«‰ Â«ﬂ« ·ÕﬁÊﬁ «·„·ﬂÌ… , „„« ﬁœ Ì⁄—÷ «·„‰ Âﬂ ≈·Ï «·„”«¡·… «·ﬁ«‰Ê‰Ì…"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   7335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÕﬁÊﬁ ≈–‰ «” Œœ«„ Â–« «·»—‰«„Ã „Õ›ÊŸ… ··„»—„Ã ›ﬁÿ "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   7335
      End
      Begin VB.Label Label9 
         Caption         =   "Label9"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   3360
         Visible         =   0   'False
         Width           =   2295
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Start_UP.frx":0112
      Top             =   3360
   End
End
Attribute VB_Name = "Start_UP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public coo As ADODB.Connection
Public nn As ADODB.Recordset
Option Explicit
Dim Tips As New Collection
Const TIP_FILE = "TEST.doc"
Dim CurrentTip As Long

Public Sub DisplayCurrentTip()
On Error Resume Next
'On Error Resume Next
If Tips.Count > 0 Then
Label4.Caption = Tips.Item(CurrentTip)
End If
End Sub

Private Sub DoNextTip()
On Error Resume Next
'On Error Resume Next
CurrentTip = Int((Tips.Count * Rnd) + 1)
Start_UP.DisplayCurrentTip
End Sub

Function LoadTips(sFile As String) As Boolean
On Error Resume Next
'On Error Resume Next
Dim NextTip As String
Dim InFile As Integer
InFile = FreeFile
If sFile = "" Then
LoadTips = False
Exit Function
End If
If dir(sFile) = "" Then
LoadTips = False
Exit Function
End If
Open sFile For Input As InFile
While Not EOF(InFile)
Line Input #InFile, NextTip
Tips.Add NextTip
Wend
Close InFile
DoNextTip
LoadTips = True
End Function
Function conn()
On Error Resume Next
Set coo = New ADODB.Connection
Set nn = New ADODB.Recordset
coo.Provider = "microsoft.jet.oledb.4.0; jet oledb:database password=7346804"
coo.ConnectionString = App.Path & "\ANNEES.mdb"
coo.Open
nn.Open "select*from Tannees", coo, adOpenKeyset, adLockOptimistic
End Function


Private Sub Combo4_Change()
On Error Resume Next
'On Error Resume Next
Timer1.Enabled = True
End Sub

Private Sub Combo4_Click()
On Error Resume Next
'On Error Resume Next
Combo4_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
On Error Resume Next
Dim s As Double
If Combo4.Text = "" Then
MsgBox "ÌÃ»  ÕœÌœ «·”‰… «·œ—«”Ì…", vbCritical
Exit Sub
End If
Label1.Caption = Combo4.Text
If Label9.Caption = "TEST" Then
'If start.Combo4.Text <> "2012-2013" Then
'MsgBox "«·”‰… «·œ—«”Ì… «·„œŒ·… €Ì— „”„ÊÕ »Â«"
'End
'Exit Sub
'End If
'Label6.Caption = "500"
If Val(Label6.Caption) >= 5000 Then
MsgBox "Â–Â «·‰”Œ…  Ã—Ì»Ì… ›ﬁÿ , ≈–« √—œ „ «·«” „—«— ›Ì «·⁄„· »‘ﬂ· ’ÕÌÕ , Ì—ÃÏ «·« ’«· »«·—ﬁ„ 22660920 √Ê 33440920 · “ÊÌœﬂ„ »«·‰”Œ… «·√’·Ì… , Ê‘ﬂ—« ⁄·Ï  ›Â„ﬂ„", vbExclamation + arabic
End
End If
If Val(Label6.Caption) > 7000 And Val(Label6.Caption) < 9000 Then
MsgBox "ÊÃ» «· ‰»ÌÂ ≈·Ï √‰ Â–Â «·‰”Œ… ’«·Õ… ·”‰… Ê«Õœ… ›ﬁÿ , Ì„ﬂ‰ﬂ„ „“«Ê·… ⁄„·ﬂ„ «·¬‰..", vbInformation + arabic
End If
End If
Call conn
Do While Not nn.EOF
If nn!ann = Label1.Caption Then
If nn!act = "1" Then
'Interface.SBB1.Panels(2).Text = "«·”‰… «·œ—«”Ì…"
Else
'Interface.SBB1.Panels(2).Text = "«—‘Ì›"
End If
nn.MoveLast
End If
nn.MoveNext
Loop
Command1.Enabled = False
Command2.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
'On Error Resume Next
End
End Sub

Private Sub Form_Load()
On Error Resume Next
'On Error Resume Next
Dim annee As String
Timer1.Enabled = False
Me.Top = 1000
Me.Left = 5000
Skin1.LoadSkin App.Path & "\18.skn"
Skin1.ApplySkin Me.hWnd
Label6.Caption = ""
Dim ShowAtStartup As Long
ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
If ShowAtStartup = 0 Then
Unload Me
Exit Sub
End If
Randomize
If LoadTips(App.Path & "\" & TIP_FILE) = False Then
Label4.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
"Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
"Then place it in the same directory as the application. "
End If
DoNextTip
Combo4.Clear
'Call cont
Call conn
Do While Not nn.EOF
If nn!sup = "0" Then
Combo4.AddItem nn!ann
If nn!act = "1" Then
annee = nn!ann
End If
End If
nn.MoveNext
Loop
Command1.Enabled = False
Command2.Enabled = False
Combo4.Text = annee
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
'On Error Resume Next
Dim s As Double
Dim j As Double
Dim nb As Double
If Combo4.Text <> "" Then
Label1.Caption = Combo4.Text
End If
s = 0
Call cont
s = s + ut.RecordCount
s = s + ou.RecordCount
s = s + pr.RecordCount
s = s + eb.RecordCount
s = s + sr.RecordCount
s = s + an.RecordCount
s = s + fc.RecordCount
s = s + pf.RecordCount
s = s + cl.RecordCount
s = s + cr.RecordCount
s = s + et.RecordCount
s = s + ie.RecordCount
s = s + cf2.RecordCount
s = s + cf1.RecordCount
s = s + mt.RecordCount
s = s + nt.RecordCount
s = s + pe.RecordCount
s = s + pp.RecordCount
s = s + pc.RecordCount
s = s + em.RecordCount
s = s + pl.RecordCount
s = s + cp.RecordCount
s = s + cf.RecordCount
s = s + cs.RecordCount
s = s + ct.RecordCount
s = s + dp.RecordCount
s = s + bn.RecordCount
s = s + ca.RecordCount
s = s + rg.RecordCount
Label6.Caption = s
Text1.Text = s
'nb = s
j = Len(Text1.Text)
If j > 2 Then
j = j - 1
Text2.Text = Mid$(Text1.Text, j, 2)
nb = Text2.Text
End If
If s > 0 Then
If s = 1 Then
Label6.Caption = "”Ã·« "
Label7.Caption = "Ê«Õœ« "
ElseIf s = 2 Then
Label6.Caption = "”Ã·«‰ "
Label7.Caption = "«À‰«‰ "
ElseIf nb >= 3 And nb <= 10 Then
Label6.Caption = s
Label7.Caption = "”Ã·«  "
ElseIf nb >= 0 Then
Label6.Caption = s
Label7.Caption = "”Ã·« "
End If
End If
Command1.Enabled = True
Command2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
'On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 10
If ProgressBar1.Value > 90 Then
Interface.SBB1.Panels(1).Text = Combo4.Text
Interface.Show
Timer2.Enabled = False
Unload Me
End If
End Sub


