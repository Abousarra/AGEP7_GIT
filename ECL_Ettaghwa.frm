VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ECL_Ettaghwa 
   Caption         =   "Form1"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5535
      ScaleWidth      =   7575
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1005
         Left            =   360
         Picture         =   "ECL_Ettaghwa.frx":0000
         ScaleHeight     =   1005
         ScaleWidth      =   6855
         TabIndex        =   22
         Top             =   1800
         Width           =   6855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   15
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   " ›⁄Ì· «·»—‰«„Ã"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   12
         Top             =   3960
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "„”Õ «·ﬂÊœ «·Œÿ√"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   11
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "«·Œ—ÊÃ „‰ «·»—‰«Ã"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Text            =   "KW0N"
         Top             =   5640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Text            =   "T8H9"
         Top             =   5640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Text            =   "FCV3"
         Top             =   5640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   6
         Text            =   "2H7S"
         Top             =   5640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Text            =   "Text9"
         Top             =   6120
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Text            =   "Text10"
         Top             =   6840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Text            =   "Text11"
         Top             =   6480
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Text            =   "Text12"
         Top             =   6840
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00004000&
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   7335
         TabIndex        =   1
         Top             =   5520
         Width           =   7335
      End
      Begin ACTIVESKINLibCtl.Skin Skin2 
         Left            =   0
         OleObjectBlob   =   "ECL_Ettaghwa.frx":4AF1
         Top             =   0
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   615
         Left            =   1560
         Top             =   3240
         Width           =   4455
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Â–Â «·‰”Œ… Œ«’… »„œ«—” «· ﬁÊÏ «·Õ—… »ﬂ—›Ê— („ﬁ—— „‘ —ﬂ —ﬁ„ 01525)‹"
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
         TabIndex        =   21
         Top             =   720
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
         Left            =   1320
         TabIndex        =   20
         Top             =   1080
         Width           =   4935
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
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ  ›⁄Ì· «·»—‰«„Ã «·Œ«’ »„Ã„⁄ „œ«—” «· ﬁÊÏ «·Œ’Ê’Ì"
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
         Left            =   1200
         TabIndex        =   18
         Top             =   2880
         Width           =   5175
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   1575
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   7335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„œ«—” «· ﬁÊÏ «·Œ’Ê’Ì"
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
         Left            =   1680
         TabIndex        =   17
         Top             =   5160
         Width           =   4215
      End
   End
End
Attribute VB_Name = "ECL_Ettaghwa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Sub Command1_Click()
Dim Security As SECURITY_ATTRIBUTES
Dim x$
Dim vg As String
If Text1.BackColor = &H0& Or Text2.BackColor = &H0& Or Text3.BackColor = &H0& Or Text4.BackColor = &H0& Then
MsgBox "ÌÃ» «œŒ«· «·ﬂÊœ ﬂ«„·«", vbCritical
Exit Sub
End If
If Text1.BackColor = &H80& Or Text2.BackColor = &H80& Or Text3.BackColor = &H80& Or Text4.BackColor = &H80& Then
MsgBox "«·ﬂÊœ «·„œŒ· €Ì— ”·Ì„", vbCritical
Exit Sub
End If
If Text1.Text <> Text5.Text Then
MsgBox "«·Ã“¡ «·√Ê· „‰ «·ﬂÊœ «·„œŒ· €Ì— ’ÕÌÕ", vbCritical
Exit Sub
End If
If Text2.Text <> Text6.Text Then
MsgBox "«·Ã“¡ «·À«‰Ì „‰ «·ﬂÊœ «·„œŒ· €Ì— ’ÕÌÕ", vbCritical
Exit Sub
End If
If Text3.Text <> Text7.Text Then
MsgBox "«·Ã“¡ «·À«·À „‰ «·ﬂÊœ «·„œŒ· €Ì— ’ÕÌÕ", vbCritical
Exit Sub
End If
If Text4.Text <> Text8.Text Then
MsgBox "«·Ã“¡ «·—«»⁄ „‰ «·ﬂÊœ «·„œŒ· €Ì— ’ÕÌÕ", vbCritical
Exit Sub
End If
x$ = ""
x$ = dir$(Text10.Text & ":\CAP\CAP.TXT")
If x$ = "" Then
'Create a directory
Ret& = CreateDirectory(Text10.Text & ":\CAP", Security)
FileCopy App.Path & "\CAP.TXT", Text10.Text & ":\CAP\CAP.TXT"
'If CreateDirectory returns 0, the function has failed
'If Ret& = 0 Then MsgBox "Error : Couldn't create directory !", vbCritical + vbOKOnly
Text11.Text = Trim(Text11.Text)
Open Text10.Text & ":\CAP\CAP.TXT" For Append As #1
Print #1, Text11.Text
Close #1
Start_UP.Label8.Caption = Label8.Caption
Start_UP.Label11.Caption = Label11.Caption
Start_UP.Label10.Caption = Label10.Caption
Start_UP.Picture3.Picture = Picture2.Picture
'Interface.Caption = Label2.Caption
Start_UP.Show
Unload Me
Exit Sub
Else
FileCopy App.Path & "\CAP.TXT", Text10.Text & ":\CAP\CAP.TXT"
'If CreateDirectory returns 0, the function has failed
'If Ret& = 0 Then MsgBox "Error : Couldn't create directory !", vbCritical + vbOKOnly
Text11.Text = Trim(Text11.Text)
Open Text10.Text & ":\CAP\CAP.TXT" For Append As #1
Print #1, Text11.Text
Close #1
Start_UP.Label8.Caption = Label8.Caption
Start_UP.Label11.Caption = Label11.Caption
Start_UP.Label10.Caption = Label10.Caption
Start_UP.Picture3.Picture = Picture2.Picture
'face.Caption = Label2.Caption
Start_UP.Show
Unload Me
Exit Sub
End If

End Sub

Private Sub Command2_Click()
If Text4.BackColor = &H80& Then
Text4.BackColor = &H0&
Text4.Text = ""
Text4.SetFocus
End If
If Text3.BackColor = &H80& Then
Text3.BackColor = &H0&
Text3.Text = ""
Text3.SetFocus
End If
If Text2.BackColor = &H80& Then
Text2.BackColor = &H0&
Text2.Text = ""
Text2.SetFocus
End If
If Text1.BackColor = &H80& Then
Text1.BackColor = &H0&
Text1.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Command3_Click()
End
End Sub



Private Sub Form_Load()
Dim vg As String
Dim x$
Dim obj_FSO As Object, obj_Drive As Object
Me.Top = 100
Me.Left = 100
Skin2.LoadSkin App.Path & "\18.skn"
Skin2.ApplySkin Me.hWnd
x$ = ""
Text9.Text = App.Path
vg = Mid$(Text9.Text, 1, 1)
Text10.Text = vg
Set obj_FSO = CreateObject("Scripting.FileSystemObject")
Set obj_Drive = obj_FSO.GetDrive(Text10.Text & ":\")
Text11.Text = obj_Drive.SerialNumber
x$ = dir$(Text10.Text & ":\CAP\CAP.TXT")
If x$ = "" Then
Exit Sub
End If
Open Text10.Text & ":\CAP\CAP.TXT" For Input As #1
Text12.Text = Input(LOF(1), 1)
Close #1
Text12.Text = Trim(Text12.Text)
If Val(Text12.Text) <> Val(Text11.Text) Then
Exit Sub
End If
Start_UP.Label8.Caption = Label8.Caption
Start_UP.Label11.Caption = Label11.Caption
Start_UP.Label10.Caption = Label10.Caption
Start_UP.Picture3.Picture = Picture2.Picture
'Interface.Caption = Label2.Caption
Start_UP.Show
Unload Me

End Sub

Private Sub Text1_Change()
Dim l As Integer
l = Len(Text1.Text)
If l >= 4 Then
Text1.Text = Trim(Text1.Text)
Call lettrescapital
Text2.SetFocus
If Text1.Text = Text5.Text Then
Text1.BackColor = &H8000&
Else
Text1.BackColor = &H80&
End If
Else
Text1.BackColor = &H0&
End If
End Sub

Private Sub Text1_Click()
Text1_Change
End Sub

Private Sub Text2_Change()
Dim l As Integer
l = Len(Text2.Text)
If l >= 4 Then
Text2.Text = Trim(Text2.Text)
Call lettrescapital
Text3.SetFocus
If Text2.Text = Text6.Text Then
Text2.BackColor = &H8000&
Else
Text2.BackColor = &H80&
End If
Else
Text2.BackColor = &H0&
End If

End Sub

Private Sub Text2_Click()
Text2_Change
End Sub

Private Sub Text3_Change()
Dim l As Integer
l = Len(Text3.Text)
If l >= 4 Then
Text3.Text = Trim(Text3.Text)
Call lettrescapital
Text4.SetFocus
If Text3.Text = Text7.Text Then
Text3.BackColor = &H8000&
Else
Text3.BackColor = &H80&
End If
Else
Text3.BackColor = &H0&
End If

End Sub

Private Sub Text3_Click()
Text3_Change
End Sub

Private Sub Text4_Change()
Dim l As Integer
l = Len(Text4.Text)
If l >= 4 Then
Text4.Text = Trim(Text4.Text)
Call lettrescapital
Command1.SetFocus
If Text4.Text = Text8.Text Then
Text4.BackColor = &H8000&
Else
Text4.BackColor = &H80&
End If
Else
Text4.BackColor = &H0&
End If

End Sub

Private Sub Text4_Click()
Text4_Change
End Sub
Private Sub lettrescapital()
Text1.Text = UCase(Text1.Text)
Text2.Text = UCase(Text2.Text)
Text3.Text = UCase(Text3.Text)
Text4.Text = UCase(Text4.Text)

End Sub



