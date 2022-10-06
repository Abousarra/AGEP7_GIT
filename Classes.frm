VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Classes 
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6720
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "≈·€«¡"
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
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Õ›Ÿ «·»Ì«‰« "
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
      Left            =   1920
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H008080FF&
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
      ItemData        =   "Classes.frx":0000
      Left            =   6720
      List            =   "Classes.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3720
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3720
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   840
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "À«‰ÊÌ"
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
      Left            =   10920
      TabIndex        =   0
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "≈⁄œ«œÌ"
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
      Left            =   9720
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "—Ê÷…"
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
      Left            =   9840
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   3240
      ScaleHeight     =   2115
      ScaleWidth      =   5835
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1560
         Width           =   4335
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   120
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·„Ê«œ «·„”„ÊÕ »Â ›Ì Â–« «·ﬁ”„"
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
         Left            =   1800
         TabIndex        =   8
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Êﬁ⁄"
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
         Left            =   4080
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «· ·«„Ì– «·„”„ÊÕ »Â ›Ì Â–« «·ﬁ”„"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   3255
      End
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "«» œ«∆Ì"
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
      Left            =   10800
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   7455
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   13150
      _Version        =   393216
      BackColor       =   32768
      BackColorFixed  =   32768
      BackColorBkg    =   32768
      RightToLeft     =   -1  'True
      FillStyle       =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line3 
      X1              =   3600
      X2              =   3600
      Y1              =   720
      Y2              =   1680
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄Ì… «· œ—Ì”"
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
      Left            =   8280
      TabIndex        =   22
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—”Ê„ «· ”ÃÌ·"
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
      Left            =   5040
      TabIndex        =   20
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—”Ê„ «·‘Â—Ì…"
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
      Left            =   5040
      TabIndex        =   19
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   6600
      X2              =   6600
      Y1              =   720
      Y2              =   1680
   End
   Begin VB.Line Line1 
      X1              =   9600
      X2              =   9600
      Y1              =   720
      Y2              =   1680
   End
   Begin VB.Shape Shape1 
      Height          =   7695
      Index           =   0
      Left            =   120
      Top             =   1800
      Width           =   12615
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Index           =   9
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   12615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„” ÊÏ"
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
      Left            =   11880
      TabIndex        =   14
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·ﬁ”„"
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
      Left            =   8280
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰«  «·√ﬁ”«„"
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
      Left            =   5400
      TabIndex        =   12
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "Classes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
Else
Combo1.BackColor = &H8080FF
End If
If Combo1.Text = "„”«∆Ì" Then
Text2.Text = "0"
Text2.Enabled = False
Text3.SetFocus
Else
Text2.Text = ""
Text2.Enabled = True
End If
End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Command1_Click()
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
Text6.Text = Trim(Text6.Text)
If Option1.Value = False And Option2.Value = False And Option3.Value = False And Option4.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ «·„” ÊÏ", vbCritical
Exit Sub
End If
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Combo1.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text1.BackColor = &H8080FF Then
Text1.SetFocus
ElseIf Text2.BackColor = &H8080FF Then
Text2.SetFocus
ElseIf Text3.BackColor = &H8080FF Then
Text3.SetFocus
End If
Exit Sub
End If
Call cont
Do While Not cl.EOF
If Text1.Text = cl!cla And Label16.Caption <> cl!aut Then
MsgBox "€Ì— „„ﬂ‰... ·ﬁœ  „ ÕÃ“ Â–« «·«”„ ”«»ﬁ«", vbCritical
Exit Sub
End If
cl.MoveNext
Loop
If Label16.Caption <> "" Then
Call cont
Do While Not cl.EOF
If Label16.Caption = cl!aut Then
If Option1.Value = True Then
cl!niv = Option1.Caption
ElseIf Option2.Value = True Then
cl!niv = Option2.Caption
ElseIf Option3.Value = True Then
cl!niv = Option3.Caption
ElseIf Option4.Value = True Then
cl!niv = Option4.Caption
End If
cl!cla = Text1.Text
cl!typ = Combo1.Text
cl!fra = Text2.Text
cl!men = Text3.Text
cl!nmt = Text5.Text
cl!net = Text4.Text
cl!sit = Text6.Text
cl.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
cl.MoveNext
Loop
End If
cl.AddNew
If Option1.Value = True Then
cl!niv = Option1.Caption
ElseIf Option2.Value = True Then
cl!niv = Option2.Caption
ElseIf Option3.Value = True Then
cl!niv = Option3.Caption
ElseIf Option4.Value = True Then
cl!niv = Option4.Caption
End If
cl!cla = Text1.Text
cl!typ = Combo1.Text
cl!fra = Text2.Text
cl!men = Text3.Text
cl!nmt = Text5.Text
cl!net = Text4.Text
cl!sit = Text6.Text
cl!act = "1"
cl.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus
Text2.Text = ""
Text2.Enabled = True
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Label16.Caption = ""
grd1.Visible = False
grd1.Clear
grd1.Rows = 1
If Option1.Value = True Then
Call chargegrd1_1
ElseIf Option2.Value = True Then
Call chargegrd1_2
ElseIf Option3.Value = True Then
Call chargegrd1_3
ElseIf Option4.Value = True Then
Call chargegrd1_4
End If
grd1.Visible = True
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
End Sub

Private Sub grd1_Click()
Dim i As Double
Dim j As Double
Dim au As Double
Dim a As Double
Dim b As Double
Dim tx As String
i = grd1.Row
j = grd1.Col
If i > 0 Then
If j = 9 Then
grd1.Row = i
grd1.Col = 0
Label16.Caption = grd1.Text
grd1.Col = 1
tx = grd1.Text
If tx = "À«‰ÊÌ" Then
Option1.Value = True
ElseIf tx = "≈⁄œ«œÌ" Then
Option2.Value = True
ElseIf tx = "«» œ«∆Ì" Then
Option3.Value = True
ElseIf tx = "—Ê÷…" Then
Option4.Value = True
End If
grd1.Col = 2
Text1.Text = grd1.Text
grd1.Col = 3
Combo1.Text = grd1.Text
grd1.Col = 4
Text2.Text = grd1.Text
grd1.Col = 5
Text3.Text = grd1.Text
grd1.Col = 6
Text4.Text = grd1.Text
grd1.Col = 7
Text5.Text = grd1.Text
grd1.Col = 8
Text6.Text = grd1.Text
End If
If j = 10 Then
grd1.Row = i
grd1.Col = 0
Label16.Caption = grd1.Text
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–« «·ﬁ”„", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Call cont
Do While Not cl.EOF
If Label16.Caption = cl!aut Then
cl!act = "0"
cl.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
cl.MoveNext
Loop
Else
Label16.Caption = ""
End If
End If
End If

End Sub





Private Sub Option1_Click()
Command2_Click
End Sub

Private Sub Option2_Click()
Command2_Click
End Sub

Private Sub Option3_Click()
Command2_Click
End Sub

Private Sub Option4_Click()
Command2_Click
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Else
Text1.BackColor = &H8080FF
End If

End Sub

Private Sub Text1_Click()
Text1_Change
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) > 0 Then
Text2.BackColor = &HC000&
Else
Text2.BackColor = &H8080FF
End If

End Sub

Private Sub Text2_Click()
Text2_Change
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text3_Change()
If Len(Text3.Text) > 0 Then
Text3.BackColor = &HC000&
Else
Text3.BackColor = &H8080FF
End If

End Sub

Private Sub Text3_Click()
Text3_Change
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii <> 8 Then
If KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation + arabic
Command2_Click
End If

End Sub
Private Sub chargegrd1_1()
Dim a As Double
Dim b As Double
Dim j As Double
Dim i As Double
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 3400
grd1.ColWidth(3) = 1800
grd1.ColWidth(4) = 1800
grd1.ColWidth(5) = 1800
grd1.ColWidth(6) = 0
grd1.ColWidth(7) = 0
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 1000
grd1.ColWidth(10) = 1000
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 1
grd1.ColAlignment(9) = 3
grd1.ColAlignment(10) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«·„” ÊÏ"
grd1.Col = 2
grd1.Text = "«·ﬁ”„"
grd1.Col = 3
grd1.Text = "‰Ê⁄Ì… «· œ—Ì”"
grd1.Col = 4
grd1.Text = "—”Ê„ «· ”ÃÌ·"
grd1.Col = 5
grd1.Text = "«·—”Ê„ «·‘Â—Ì…"
grd1.Col = 6
grd1.Text = "⁄œœ «· ·«„Ì–"
grd1.Col = 7
grd1.Text = "⁄œœ «·„Ê«œ"
grd1.Col = 8
grd1.Text = "«·„Êﬁ⁄"
i = 1
Call cont
grd1.Rows = cl.RecordCount + 3
Do While Not cl.EOF
If Option1.Caption = cl!niv Then
If cl!act = "1" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = cl!aut
grd1.Col = 1
grd1.Text = cl!niv
grd1.Col = 2
grd1.Text = cl!cla
grd1.Col = 3
grd1.Text = cl!typ
grd1.Col = 4
grd1.Text = cl!fra
grd1.Col = 5
grd1.Text = cl!men
grd1.Col = 6
grd1.Text = cl!net
grd1.Col = 7
grd1.Text = cl!nmt
grd1.Col = 8
grd1.Text = cl!sit
grd1.Col = 9
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 10
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
End If
End If
cl.MoveNext
Loop
grd1.Rows = i
grd1.Col = 4
grd1.Sort = 2
End Sub
Private Sub chargegrd1_2()
Dim a As Double
Dim b As Double
Dim j As Double
Dim i As Double
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 3400
grd1.ColWidth(3) = 1800
grd1.ColWidth(4) = 1800
grd1.ColWidth(5) = 1800
grd1.ColWidth(6) = 0
grd1.ColWidth(7) = 0
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 1000
grd1.ColWidth(10) = 1000
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 1
grd1.ColAlignment(9) = 3
grd1.ColAlignment(10) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«·„” ÊÏ"
grd1.Col = 2
grd1.Text = "«·ﬁ”„"
grd1.Col = 3
grd1.Text = "‰Ê⁄Ì… «· œ—Ì”"
grd1.Col = 4
grd1.Text = "—”Ê„ «· ”ÃÌ·"
grd1.Col = 5
grd1.Text = "«·—”Ê„ «·‘Â—Ì…"
grd1.Col = 6
grd1.Text = "⁄œœ «· ·«„Ì–"
grd1.Col = 7
grd1.Text = "⁄œœ «·„Ê«œ"
grd1.Col = 8
grd1.Text = "«·„Êﬁ⁄"
i = 1
Call cont
grd1.Rows = cl.RecordCount + 3
Do While Not cl.EOF
If cl!act = "1" Then
If Option2.Caption = cl!niv Then
grd1.Row = i
grd1.Col = 0
grd1.Text = cl!aut
grd1.Col = 1
grd1.Text = cl!niv
grd1.Col = 2
grd1.Text = cl!cla
grd1.Col = 3
grd1.Text = cl!typ
grd1.Col = 4
grd1.Text = cl!fra
grd1.Col = 5
grd1.Text = cl!men
grd1.Col = 6
grd1.Text = cl!net
grd1.Col = 7
grd1.Text = cl!nmt
grd1.Col = 8
grd1.Text = cl!sit
grd1.Col = 9
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 10
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
End If
End If
cl.MoveNext
Loop
grd1.Rows = i
grd1.Col = 4
grd1.Sort = 2
End Sub
Private Sub chargegrd1_3()
Dim a As Double
Dim b As Double
Dim j As Double
Dim i As Double
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 3400
grd1.ColWidth(3) = 1800
grd1.ColWidth(4) = 1800
grd1.ColWidth(5) = 1800
grd1.ColWidth(6) = 0
grd1.ColWidth(7) = 0
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 1000
grd1.ColWidth(10) = 1000
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 1
grd1.ColAlignment(9) = 3
grd1.ColAlignment(10) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«·„” ÊÏ"
grd1.Col = 2
grd1.Text = "«·ﬁ”„"
grd1.Col = 3
grd1.Text = "‰Ê⁄Ì… «· œ—Ì”"
grd1.Col = 4
grd1.Text = "—”Ê„ «· ”ÃÌ·"
grd1.Col = 5
grd1.Text = "«·—”Ê„ «·‘Â—Ì…"
grd1.Col = 6
grd1.Text = "⁄œœ «· ·«„Ì–"
grd1.Col = 7
grd1.Text = "⁄œœ «·„Ê«œ"
grd1.Col = 8
grd1.Text = "«·„Êﬁ⁄"
i = 1
Call cont
grd1.Rows = cl.RecordCount + 3
Do While Not cl.EOF
If cl!act = "1" Then
If Option3.Caption = cl!niv Then
grd1.Row = i
grd1.Col = 0
grd1.Text = cl!aut
grd1.Col = 1
grd1.Text = cl!niv
grd1.Col = 2
grd1.Text = cl!cla
grd1.Col = 3
grd1.Text = cl!typ
grd1.Col = 4
grd1.Text = cl!fra
grd1.Col = 5
grd1.Text = cl!men
grd1.Col = 6
grd1.Text = cl!net
grd1.Col = 7
grd1.Text = cl!nmt
grd1.Col = 8
grd1.Text = cl!sit
grd1.Col = 9
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 10
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
End If
End If
cl.MoveNext
Loop
grd1.Rows = i
grd1.Col = 4
grd1.Sort = 2
End Sub
Private Sub chargegrd1_4()
Dim a As Double
Dim b As Double
Dim j As Double
Dim i As Double
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 3400
grd1.ColWidth(3) = 1800
grd1.ColWidth(4) = 1800
grd1.ColWidth(5) = 1800
grd1.ColWidth(6) = 0
grd1.ColWidth(7) = 0
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 1000
grd1.ColWidth(10) = 1000
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 1
grd1.ColAlignment(9) = 3
grd1.ColAlignment(10) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«·„” ÊÏ"
grd1.Col = 2
grd1.Text = "«·ﬁ”„"
grd1.Col = 3
grd1.Text = "‰Ê⁄Ì… «· œ—Ì”"
grd1.Col = 4
grd1.Text = "—”Ê„ «· ”ÃÌ·"
grd1.Col = 5
grd1.Text = "«·—”Ê„ «·‘Â—Ì…"
grd1.Col = 6
grd1.Text = "⁄œœ «· ·«„Ì–"
grd1.Col = 7
grd1.Text = "⁄œœ «·„Ê«œ"
grd1.Col = 8
grd1.Text = "«·„Êﬁ⁄"
i = 1
Call cont
grd1.Rows = cl.RecordCount + 3
Do While Not cl.EOF
If cl!act = "1" Then
If Option4.Caption = cl!niv Then
grd1.Row = i
grd1.Col = 0
grd1.Text = cl!aut
grd1.Col = 1
grd1.Text = cl!niv
grd1.Col = 2
grd1.Text = cl!cla
grd1.Col = 3
grd1.Text = cl!typ
grd1.Col = 4
grd1.Text = cl!fra
grd1.Col = 5
grd1.Text = cl!men
grd1.Col = 6
grd1.Text = cl!net
grd1.Col = 7
grd1.Text = cl!nmt
grd1.Col = 8
grd1.Text = cl!sit
grd1.Col = 9
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 10
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
End If
End If
cl.MoveNext
Loop
grd1.Rows = i
grd1.Col = 4
grd1.Sort = 2
End Sub

