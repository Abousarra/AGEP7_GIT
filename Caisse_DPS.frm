VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Caisse_DPS 
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
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   10320
      ScaleHeight     =   8745
      ScaleWidth      =   2505
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "€·ﬁ"
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
         Left            =   840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   8400
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid grd6 
         Height          =   8295
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   14631
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   8421631
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
   End
   Begin VB.CommandButton Command9 
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
      Left            =   1080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   780
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
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
      TabIndex        =   20
      Top             =   780
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CheckBox Check13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "⁄—÷ «·ﬂ·"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   11400
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
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
      Height          =   375
      Left            =   2400
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   780
      Width           =   4215
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "⁄—÷"
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
      Left            =   3240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1280
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "”Õ» "
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
      Left            =   2400
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1280
      Width           =   735
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
      Height          =   375
      Left            =   7440
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   780
      Width           =   1575
   End
   Begin VB.PictureBox Picture3 
      Height          =   3975
      Left            =   4920
      ScaleHeight     =   3915
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   360
         TabIndex        =   29
         Text            =   "Text5"
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   360
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Text            =   "Text4"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   0
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄ «·„»«·€ «·„’—Ê›…"
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
         Left            =   2160
         TabIndex        =   32
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   -240
         TabIndex        =   31
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Left            =   1080
         TabIndex        =   30
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„»«·€ «·„’—Ê›…"
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
         Left            =   1440
         TabIndex        =   23
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   -240
         TabIndex        =   22
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "0"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   255
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   345
      Left            =   9960
      TabIndex        =   8
      Top             =   780
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16744576
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16744576
      Format          =   124977153
      CurrentDate     =   42638
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   7455
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   13150
      _Version        =   393216
      Rows            =   1
      BackColor       =   32768
      BackColorFixed  =   32768
      BackColorBkg    =   32768
      RightToLeft     =   -1  'True
      FillStyle       =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DT2 
      Height          =   345
      Left            =   7200
      TabIndex        =   11
      Top             =   1275
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16744576
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16744576
      Format          =   124977153
      CurrentDate     =   42638
   End
   Begin MSComCtl2.DTPicker DT3 
      Height          =   345
      Left            =   4200
      TabIndex        =   12
      Top             =   1275
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16744576
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16744576
      Format          =   124977153
      CurrentDate     =   42638
   End
   Begin MSFlexGridLib.MSFlexGrid grd5 
      Height          =   615
      Left            =   10440
      TabIndex        =   27
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   8421631
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ›«’Ì·"
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
      Index           =   0
      Left            =   6000
      TabIndex        =   18
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄—÷ „‰  «—ÌŒ"
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
      Index           =   0
      Left            =   8400
      TabIndex        =   17
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï  «—ÌŒ"
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
      Index           =   0
      Left            =   5760
      TabIndex        =   16
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«· «—ÌŒ"
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
      Index           =   0
      Left            =   11280
      TabIndex        =   15
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„»·€"
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
      TabIndex        =   14
      Top             =   840
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   2
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   12615
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   12615
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "’‰œÊﬁ «·„’—Ê›« "
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
      Left            =   5280
      TabIndex        =   13
      Top             =   0
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      Height          =   7695
      Index           =   3
      Left            =   120
      Top             =   1800
      Width           =   12615
   End
End
Attribute VB_Name = "Caisse_DPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check13_Click()
On Error Resume Next
If Check13.Value = 1 Then
grd2.Visible = False
Call chargegrd2_T
grd2.Visible = True
Else
grd2.Visible = False
Call chargegrd2_M
grd2.Visible = True
End If
End Sub

Private Sub Command6_Click()
On Error Resume Next
Picture4.Visible = False

End Sub

Private Sub Command7_Click()
On Error Resume Next
If Check13.Value = 0 Then
grd2.Visible = False
Call chargegrd2_M
grd2.Visible = True
Else
Check13.Value = 0
End If
End Sub

Private Sub Command8_Click()
On Error Resume Next
Text3.Text = ""
DT1.Value = Date
Label17.Caption = ""
Label6.Caption = "0"
If Check13.Value = 1 Then
Check13.Value = 0
Else
grd2.Visible = False
Call chargegrd2_M
grd2.Visible = True
End If
Check13.Value = 0
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False
Text2.Text = ""
Text2.SetFocus
Call Operations

End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
If Text2.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· «·„»·€ «·„’—Ê›", vbCritical + arabic
Text2.SetFocus
Exit Sub
End If
'** controle caisse
Call cont
a = eb!cca
b = Text2.Text
c = Label6.Caption
d = (a + c) - b
If d < 0 Then
MsgBox "—’Ìœ «·’‰œÊﬁ €Ì— ﬂ«› ·≈ „«„ «·⁄„·Ì…", vbExclamation
Exit Sub
End If
eb!cca = d
eb.Update
'******* controle
'**** archive de caisse ajou et modif
Adat = Date
Aheu = Time$
If Label17.Caption = "" Then
Atyp = "≈÷«›…"
Else
Atyp = " ⁄œÌ·"
End If
Adet = Text3.Text
Amon = Text2.Text
Acom = "”Ã· «·„’—Ê›« "
Auti = directions.Label2.Caption
'****************************************
If Label17.Caption <> "" Then
Call cont
Do While Not dp.EOF
If Label17.Caption = dp!aut Then
dp!dat = DT1.Value
dp!mon = Text2.Text
dp!det = Text3.Text
dp!heu = Time$
If dp!act = "2" Then
dp!act = "3"
End If
dp.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
dp.MoveNext
Loop
End If
dp.AddNew
dp!dat = DT1.Value
dp!mon = Text2.Text
dp!det = Text3.Text
dp!heu = Time$
dp!act = "0"
dp!mtf = ""
dp.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0
DT1.Value = Date
DT2.Value = Date
DT3.Value = Date
Label17.Caption = ""
Check13.Value = 1
Call Operations
End Sub



Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim a As Double
Dim b As Double
Dim tx1 As String
i = grd2.Row
j = grd2.Col
If i > 0 Then
If j = 1 Then
grd2.Row = i
grd2.Col = 1
If grd2.CellBackColor = &HFF& Then
grd2.Row = i
grd2.Col = 0
tx1 = grd2.Text
Call cont
Do While Not dp.EOF
If dp!aut = tx1 Then
MsgBox dp!mtf
Exit Sub
End If
dp.MoveNext
Loop
End If
End If
If j = 5 Then
grd2.Row = i
grd2.Col = 0
Label17.Caption = grd2.Text
grd2.Col = 1
DT1.Value = grd2.Text
grd2.Col = 3
Text2.Text = grd2.Text
Label6.Caption = grd2.Text
grd2.Col = 4
Text3.Text = grd2.Text
End If
If j = 6 Then
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–Â «·⁄„·Ì…", vbInformation + vbYesNo + arabic, "AGEP7")
If g = vbYes Then
grd2.Row = i
grd2.Col = 0
Label17.Caption = grd2.Text
grd2.Col = 3
Label6.Caption = grd2.Text
a = eb!cca
b = Label6.Caption
a = a + b
eb!cca = a
eb.Update
Call cont
Do While Not dp.EOF
If Label17.Caption = dp!aut Then
'**** archive de caisse supp
Label8.Caption = dp!det
Adat = Date
Aheu = Time$
Atyp = "Õ–›"
Adet = Label8.Caption
Amon = b
Acom = "”Ã· «·„’—Ê›« "
Auti = directions.Label2.Caption
'****************************************
dp.Delete
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
dp.MoveNext
Loop
End If
End If
End If

End Sub

Private Sub grd5_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
k = 0
i = grd5.Row
If i = 0 Then
j = 0
k = 3
Else
j = 2
k = 2
End If
If i = 0 Or i = 1 Then
If grd5.RowHeight(i) > 0 Then
Text1.Text = j
Text5.Text = k
grd6.Visible = False
Call chargegrd6
grd6.Visible = True
Picture4.Visible = True
End If
End If

End Sub

Private Sub grd6_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
i = grd6.Row
j = grd6.Col
If i > 0 Then
grd6.Row = i
grd6.Col = 0
DT2.Value = grd6.Text
DT3.Value = grd6.Text
Command7_Click
Picture4.Visible = False

End If

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

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation + arabic
Command8_Click
Call archive_caisse
End If

End Sub

Private Sub chargegrd2_T()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim d As Double
Dim sd As Double
grd2.Clear
grd2.Cols = 7
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1500
grd2.ColWidth(2) = 1500
grd2.ColWidth(3) = 2500
grd2.ColWidth(4) = 5000
grd2.ColWidth(5) = 800
grd2.ColWidth(6) = 800
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 3
grd2.ColAlignment(6) = 3
grd2.Row = 0
grd2.Col = 1
grd2.Text = "«· «—ÌŒ"
grd2.Col = 2
grd2.Text = "«·”«⁄…"
grd2.Col = 3
grd2.Text = "«·„»·€"
grd2.Col = 4
grd2.Text = "«· ›«’Ì·"
i = 1
dat1 = DT2.Value
dat2 = DT3.Value
a = 0
sd = 0
Call cont
grd2.Rows = dp.RecordCount + 3
Do While Not dp.EOF
dat3 = dp!dat
If dp!act <> "1" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = dp!aut
grd2.Col = 1
grd2.Text = dp!dat
If dp!act = "2" Then
grd2.CellBackColor = &HFF&
End If
grd2.Col = 2
grd2.Text = dp!heu
grd2.Col = 3
grd2.Text = dp!mon
d = dp!mon
sd = sd + d
grd2.Col = 4
grd2.Text = dp!det
grd2.Col = 5
grd2.Text = " ⁄œÌ·"
grd2.CellBackColor = &HFFFF&
grd2.Col = 6
grd2.Text = "Õ–›"
grd2.CellBackColor = &HFF&
i = i + 1
End If
dp.MoveNext
Loop
grd2.Rows = i
Label1.Caption = sd
Label7.Caption = sd
End Sub
Private Sub chargegrd2_M()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim d As Double
Dim sd As Double
grd2.Clear
grd2.Cols = 7
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1500
grd2.ColWidth(2) = 1500
grd2.ColWidth(3) = 2500
grd2.ColWidth(4) = 5000
grd2.ColWidth(5) = 800
grd2.ColWidth(6) = 800
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 3
grd2.ColAlignment(6) = 3
grd2.Row = 0
grd2.Col = 1
grd2.Text = "«· «—ÌŒ"
grd2.Col = 2
grd2.Text = "«·”«⁄…"
grd2.Col = 3
grd2.Text = "«·„»·€"
grd2.Col = 4
grd2.Text = "«· ›«’Ì·"
i = 1
dat1 = DT2.Value
dat2 = DT3.Value
a = 0
sd = 0
Call cont
grd2.Rows = dp.RecordCount + 3
Do While Not dp.EOF
dat3 = dp!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If dp!act <> "1" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = dp!aut
grd2.Col = 1
grd2.Text = dp!dat
If dp!act = "2" Then
grd2.CellBackColor = &HFF&
End If
grd2.Col = 2
grd2.Text = dp!heu
grd2.Col = 3
grd2.Text = dp!mon
d = dp!mon
sd = sd + d
grd2.Col = 4
grd2.Text = dp!det
grd2.Col = 5
grd2.Text = " ⁄œÌ·"
grd2.CellBackColor = &HFFFF&
grd2.Col = 6
grd2.Text = "Õ–›"
grd2.CellBackColor = &HFF&
i = i + 1
End If
End If
dp.MoveNext
Loop
grd2.Rows = i
Label1.Caption = sd
End Sub
Private Sub Operations()
On Error Resume Next
Dim a As Double
Dim b As Double
grd5.Rows = 2
grd5.Cols = 2
grd5.ColWidth(0) = 1100
grd5.ColWidth(1) = 900
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 3
grd5.Col = 0
grd5.Row = 0
grd5.Text = "⁄„·Ì… ÃœÌœ…"
grd5.Row = 1
grd5.Text = "⁄„·Ì… „—›Ê÷…"
a = 0
b = 0
Call cont
Do While Not dp.EOF
If dp!act = "0" Or dp!act = "3" Then
a = a + 1
End If
If dp!act = "2" Then
b = b + 1
End If
dp.MoveNext
Loop
grd5.Col = 1
grd5.Row = 0
grd5.Text = a
grd5.CellBackColor = &HFFFF&
grd5.Row = 1
grd5.Text = b
grd5.CellBackColor = &HFF&
If a = 0 Then
grd5.RowHeight(0) = 0
Else
grd5.RowHeight(0) = 250
End If
If b = 0 Then
grd5.RowHeight(1) = 0
Else
grd5.RowHeight(1) = 250
End If
End Sub
Private Sub chargegrd6()
On Error Resume Next
Dim i As Double
Dim tx As String
grd6.Clear
grd6.Rows = 1
grd6.Cols = 2
grd6.ColWidth(0) = 1150
grd6.ColWidth(1) = 1100
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.Row = 0
grd6.Col = 0
grd6.Text = " «—ÌŒ «·⁄„·Ì…"
grd6.Col = 1
grd6.Text = "‰Ê⁄ «·⁄„·Ì…"
If Text1.Text = "2" Then
tx = "„—›Ê÷…"
Else
tx = "ÃœÌœ…"
End If
i = 1
Call cont
grd6.Rows = dp.RecordCount + 3
Do While Not dp.EOF
If dp!act = Text1.Text Or dp!act = Text5.Text Then
grd6.Row = i
grd6.Col = 0
grd6.Text = dp!dat
grd6.Col = 1
grd6.Text = tx
If Text1.Text = "0" Then
grd6.CellBackColor = &HFFFF&
Else
grd6.CellBackColor = &HFF&
End If
i = i + 1
End If
dp.MoveNext
Loop
grd6.Rows = i
End Sub


