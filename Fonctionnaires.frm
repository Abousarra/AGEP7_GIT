VERSION 5.00
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Fonctionnaires 
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
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   840
      Width           =   3255
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
      Left            =   3000
      TabIndex        =   21
      Top             =   2040
      Width           =   1455
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
      TabIndex        =   20
      Top             =   2040
      UseMaskColor    =   -1  'True
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
      Left            =   8160
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   5040
      ScaleHeight     =   4635
      ScaleWidth      =   5835
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton Command3 
         Caption         =   " Œ“Ì‰ «·„ÊŸ›Ì‰"
         Height          =   615
         Left            =   1080
         TabIndex        =   25
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   0
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
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
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
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
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
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
      Left            =   4680
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox Text7 
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
      Left            =   4680
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text8 
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4680
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1560
      Width           =   2175
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
      TabIndex        =   7
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
      Left            =   4680
      TabIndex        =   8
      Top             =   2040
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Line Line2 
      X1              =   4560
      X2              =   4560
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
      Left            =   3240
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   1200
      Width           =   1215
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
      Left            =   3240
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      Left            =   3240
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„ÊŸ›"
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
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
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
      TabIndex        =   10
      Top             =   0
      Width           =   12615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—« »"
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
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "Fonctionnaires"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim seri As String
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text4.Text = "0000"
Text5.Text = "0000"
Else
Text4.Text = ""
Text5.Text = ""
End If

End Sub

Private Sub Command1_Click()
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
Text6.Text = Trim(Text6.Text)
Text7.Text = Trim(Text7.Text)
Text8.Text = Trim(Text8.Text)
If Text1.Text = "" Or Text2.Text = "" Or Text5.Text = "" Or Text4.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text1.BackColor = &H8080FF Then
Text1.SetFocus
ElseIf Text2.BackColor = &H8080FF Then
Text2.SetFocus
ElseIf Text7.BackColor = &H8080FF Then
Text7.SetFocus
ElseIf Text8.BackColor = &H8080FF Then
Text8.SetFocus
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
Do While Not fc.EOF
If Label16.Caption = fc!aut Then
fc!sri = BX1.Caption
fc!nom = Text1.Text
fc!tel = Text2.Text
fc!fon = Text7.Text
fc!sal = Text8.Text
fc!adr = Text3.Text
fc!nni = Text6.Text
fc!mot = Text4.Text
fc.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
fc.MoveNext
Loop
End If
fc.AddNew
fc!sri = BX1.Caption
fc!nom = Text1.Text
fc!tel = Text2.Text
fc!fon = Text7.Text
fc!sal = Text8.Text
fc!adr = Text3.Text
fc!nni = Text6.Text
fc!mot = Text4.Text
fc!act = "1"
fc.Update
eb!sri = Val(eb!sri) + 1
eb.Update
sr.AddNew
sr!sri = BX1.Caption
sr!eta = "„ÊŸ›"
sr.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
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

Private Sub Command3_Click()
Dim j As Double
Dim tx As String
j = 4520
Call cont3
Do While Not fc3.EOF
Call cont
tx = j
tx = "000" + tx
BX1.Caption = tx
Text1.Text = fc3!nom
Text2.Text = fc3!tel
Text7.Text = fc3!foc
Text8.Text = "300000"
Check1.Value = 1
j = j + 1
Command1_Click
fc3.MoveNext
Loop
MsgBox "OK", vbInformation

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Call cont
xe = eb!sri
Call Series
BX1.Caption = xs
Call chargegrd1

End Sub

Private Sub grd1_Click()
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
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–« «·„ÊŸ›", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Call cont
Do While Not fc.EOF
If Label16.Caption = fc!aut Then
fc!act = "0"
fc.Update
'Call supression_series
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
fc.MoveNext
Loop
Else
Label16.Caption = ""
End If
End If
End If

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

Private Sub Text4_Change()
If Len(Text4.Text) > 0 Then
Text4.BackColor = &HC000&
Else
Text4.BackColor = &H8080FF
End If

End Sub

Private Sub Text4_Click()
Text4_Change
End Sub

Private Sub Text5_Change()
If Len(Text5.Text) > 0 Then
Text5.BackColor = &HC000&
Else
Text5.BackColor = &H8080FF
End If

End Sub

Private Sub Text5_Click()
Text5_Change
End Sub

Private Sub Text7_Change()
If Len(Text7.Text) > 0 Then
Text7.BackColor = &HC000&
Else
Text7.BackColor = &H8080FF
End If


End Sub

Private Sub Text7_Click()
Text7_Change
End Sub

Private Sub Text8_Change()
If Len(Text8.Text) > 0 Then
Text8.BackColor = &HC000&
Else
Text8.BackColor = &H8080FF
End If

End Sub

Private Sub Text8_Click()
Text8_Change
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
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
Private Sub chargegrd1()
Dim a As Double
Dim b As Double
b = 0
Dim j As Double
Dim i As Double
Dim P As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1300
grd1.ColWidth(2) = 2500
grd1.ColWidth(3) = 1100
grd1.ColWidth(4) = 1200
grd1.ColWidth(5) = 1300
grd1.ColWidth(6) = 1300
grd1.ColWidth(7) = 2000
grd1.ColWidth(8) = 700
grd1.ColWidth(9) = 0
grd1.ColWidth(10) = 700
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
grd1.Text = "«”„ «·„ÊŸ›"
grd1.Col = 3
grd1.Text = "—ﬁ„ «·Â« ›"
grd1.Col = 4
grd1.Text = "«·ÊŸÌ›…"
grd1.Col = 5
grd1.Text = "«·—« »"
grd1.Col = 6
grd1.Text = "«·—ﬁ„ «·Êÿ‰Ì"
grd1.Col = 7
grd1.Text = "«·⁄‰Ê«‰"
i = 1
Call cont
grd1.Rows = fc.RecordCount + 3
Do While Not fc.EOF
If fc!act = "1" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = fc!aut
grd1.Col = 1
grd1.Text = fc!sri
grd1.Col = 2
grd1.Text = fc!nom
grd1.Col = 3
grd1.Text = fc!tel
grd1.Col = 4
grd1.Text = fc!fon
grd1.Col = 5
grd1.Text = fc!sal
grd1.Col = 6
grd1.Text = fc!nni
grd1.Col = 7
grd1.Text = fc!adr
grd1.Col = 8
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 9
grd1.Text = fc!mot
grd1.Col = 10
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
End If
fc.MoveNext
Loop
grd1.Rows = i
grd1.Col = 4
grd1.Sort = 2
End Sub
Private Sub supression_series()
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

