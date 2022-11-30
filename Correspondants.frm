VERSION 5.00
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Correspondants 
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
      Left            =   8160
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   1560
      Width           =   3255
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
      Left            =   8160
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2040
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
      Left            =   10920
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "ﬂ·„… ”—  ·ﬁ«∆Ì…"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text4 
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
      Left            =   2280
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text5 
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
      Left            =   240
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
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
      Left            =   240
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
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
      Left            =   240
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   840
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   2640
      ScaleHeight     =   4635
      ScaleWidth      =   5835
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   0
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   1320
         TabIndex        =   25
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
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
      Left            =   5040
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
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
      Left            =   5040
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
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
      Left            =   5040
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
   Begin BARCODEXLib.BarcodeX BX1 
      Height          =   735
      Left            =   8160
      Top             =   840
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   1296
      _StockProps     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0000001"
      BarcodeType     =   6
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   6615
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   11668
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
      Left            =   5040
      TabIndex        =   6
      Top             =   2040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line2 
      X1              =   4920
      X2              =   4920
      Y1              =   720
      Y2              =   2520
   End
   Begin VB.Line Line1 
      X1              =   8040
      X2              =   8040
      Y1              =   720
      Y2              =   2520
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„ «·Êÿ‰Ì"
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
      Left            =   6720
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   6855
      Index           =   0
      Left            =   120
      Top             =   2640
      Width           =   12615
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Index           =   9
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   12615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·»—Ìœ «·≈·ﬂ —Ê‰Ì"
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
      Left            =   3240
      TabIndex        =   15
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄‰Ê«‰"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·Â« ›"
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
      Left            =   6720
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„ «· ”·”·Ì"
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
      Left            =   11040
      TabIndex        =   12
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "≈⁄«œ… ·Â«"
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
      Left            =   1200
      TabIndex        =   11
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ·„… «·”—"
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
      Left            =   3600
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·ÊﬂÌ·"
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
      Left            =   11280
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰«  «·Êﬂ·«¡"
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
      TabIndex        =   8
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÊŸÌ›…"
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
      Left            =   6720
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Correspondants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim seri As String

Private Sub Check1_Click()
On Error Resume Next
If Check1.Value = 1 Then
Text4.Text = "0000"
Text5.Text = "0000"
Else
Text4.Text = ""
Text5.Text = ""
End If

End Sub

Private Sub Command1_Click()
On Error Resume Next
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
Text6.Text = Trim(Text6.Text)
Text7.Text = Trim(Text7.Text)
Text8.Text = Trim(Text8.Text)
If Text1.Text = "" Or Text2.Text = "" Or Text5.Text = "" Or Text4.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text1.BackColor = &H8080FF Then
Text1.SetFocus
ElseIf Text2.BackColor = &H8080FF Then
Text2.SetFocus
ElseIf Text4.BackColor = &H8080FF Then
Text4.SetFocus
ElseIf Text5.BackColor = &H8080FF Then
Text5.SetFocus
End If
Exit Sub
End If
If Text5.Text <> Text4.Text Then
MsgBox "ﬂ·„ « «·”— €Ì— „ ÿ«»ﬁ Ì‰", vbCritical + arabic
Text5.Text = ""
Text5.SetFocus
Exit Sub
End If
Call cont
If Label16.Caption <> "" Then
Do While Not cr.EOF
If Label16.Caption = cr!aut Then
cr!sri = BX1.Caption
cr!nom = Text1.Text
cr!tel = Text2.Text
cr!nni = Text7.Text
cr!fon = Text8.Text
cr!adr = Text3.Text
cr!eml = Text6.Text
cr!mot = Text4.Text
cr.Update
Call changement_tel
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
cr.MoveNext
Loop
End If
cr.AddNew
cr!sri = BX1.Caption
cr!nom = Text1.Text
cr!tel = Text2.Text
cr!nni = Text7.Text
cr!fon = Text8.Text
cr!adr = Text3.Text
cr!eml = Text6.Text
cr!mot = Text4.Text
cr!act = "1"
cr.Update
eb!sri = Val(eb!sri) + 1
eb.Update
sr.AddNew
sr!sri = BX1.Caption
sr!eta = "ÊﬂÌ·"
sr.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
On Error Resume Next
Text1.Text = ""
Text1.SetFocus
Text2.Text = ""
Text3.Text = ""
Check1.Value = 1
Check1.Value = 0
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Label16.Caption = ""
Label6.Caption = ""
grd1.Visible = False
grd1.Clear
grd1.Rows = 1
Call chargegrd1
grd1.Visible = True
xe = eb!sri
Call Series
BX1.Caption = xs
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0
Call cont
xe = eb!sri
Call Series
BX1.Caption = xs
Call chargegrd1

End Sub

Private Sub grd1_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim au As Double
Dim a As Double
Dim b As Double
i = grd1.Row
j = grd1.Col
If i > 0 Then
If j = 8 Then
grd1.Row = i
grd1.Col = 0
Label16.Caption = grd1.Text
grd1.Col = 1
BX1.Caption = grd1.Text
grd1.Col = 2
Text1.Text = grd1.Text
grd1.Col = 3
Text2.Text = grd1.Text
Label6.Caption = grd1.Text
grd1.Col = 4
Text7.Text = grd1.Text
grd1.Col = 5
Text8.Text = grd1.Text
grd1.Col = 6
Text3.Text = grd1.Text
grd1.Col = 7
Text6.Text = grd1.Text
grd1.Col = 9
Text4.Text = grd1.Text
grd1.Col = 9
Text5.Text = grd1.Text
End If
If j = 10 Then
grd1.Row = i
grd1.Col = 0
Label16.Caption = grd1.Text
grd1.Row = i
grd1.Col = 1
seri = grd1.Text
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–« «·ÊﬂÌ·", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Call cont
Do While Not cr.EOF
If Label16.Caption = cr!aut Then
cr!act = "0"
cr.Update
'Call supression_series
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
cr.MoveNext
Loop
Else
Label16.Caption = ""
End If

End If
End If

End Sub




Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Else
Text1.BackColor = &H8080FF
End If

End Sub

Private Sub Text1_Click()
On Error Resume Next
Text1_Change
End Sub

Private Sub Text2_Change()
On Error Resume Next
If Len(Text2.Text) > 0 Then
Text2.BackColor = &HC000&
Else
Text2.BackColor = &H8080FF
End If

End Sub

Private Sub Text2_Click()
On Error Resume Next
Text2_Change
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text4_Change()
On Error Resume Next
If Len(Text4.Text) > 0 Then
Text4.BackColor = &HC000&
Else
Text4.BackColor = &H8080FF
End If

End Sub

Private Sub Text4_Click()
On Error Resume Next
Text4_Change
End Sub

Private Sub Text5_Change()
On Error Resume Next
If Len(Text5.Text) > 0 Then
Text5.BackColor = &HC000&
Else
Text5.BackColor = &H8080FF
End If

End Sub

Private Sub Text5_Click()
On Error Resume Next
Text5_Change
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation + arabic
Command2_Click
End If

End Sub
Private Sub chargegrd1()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim j As Double
Dim i As Double
Dim P As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1400
grd1.ColWidth(2) = 2000
grd1.ColWidth(3) = 1200
grd1.ColWidth(4) = 1300
grd1.ColWidth(5) = 1400
grd1.ColWidth(6) = 1800
grd1.ColWidth(7) = 1800
grd1.ColWidth(8) = 600
grd1.ColWidth(9) = 0
grd1.ColWidth(10) = 600
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 3
grd1.ColAlignment(10) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«·—ﬁ„ «· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«”„ «·ÊﬂÌ·"
grd1.Col = 3
grd1.Text = "—ﬁ„ «·Â« ›"
grd1.Col = 4
grd1.Text = "«·—ﬁ„ «·Êÿ‰Ì"
grd1.Col = 5
grd1.Text = "«·ÊŸÌ›…"
grd1.Col = 6
grd1.Text = "«·⁄‰Ê«‰"
grd1.Col = 7
grd1.Text = "«·»—Ìœ"
i = 1
Call cont
grd1.Rows = cr.RecordCount + 3
Do While Not cr.EOF
If cr!act = "1" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = cr!aut
grd1.Col = 1
grd1.Text = cr!sri
grd1.Col = 2
grd1.Text = cr!nom
grd1.Col = 3
grd1.Text = cr!tel
grd1.Col = 4
grd1.Text = cr!nni
grd1.Col = 5
grd1.Text = cr!fon
grd1.Col = 6
grd1.Text = cr!adr
grd1.Col = 7
grd1.Text = cr!eml
grd1.Col = 8
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 9
grd1.Text = cr!mot
grd1.Col = 10
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
End If
cr.MoveNext
Loop
grd1.Rows = i
grd1.Col = 4
grd1.Sort = 2
End Sub
Private Sub supression_series()
On Error Resume Next
Call cont
Do While Not sr.EOF
If seri = sr!sri Then
sr.Delete
If sr.RecordCount > 3 Then
sr.MoveLast
End If
End If
sr.MoveNext
Loop
End Sub
Private Sub changement_tel()
On Error Resume Next
Call cont
Do While Not et.EOF
If Label6.Caption = et!tel Then
et!tel = Text2.Text
et.Update
End If
et.MoveNext
Loop

End Sub



