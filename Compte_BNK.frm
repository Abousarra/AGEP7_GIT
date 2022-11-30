VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Compte_BNK 
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
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "⁄—÷"
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
      Left            =   1920
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   795
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "”Õ» "
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
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   795
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   7455
      Left            =   240
      TabIndex        =   3
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
      Left            =   6480
      TabIndex        =   4
      Top             =   795
      Width           =   1515
      _ExtentX        =   2672
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
      Left            =   3840
      TabIndex        =   5
      Top             =   795
      Width           =   1515
      _ExtentX        =   2672
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
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» «·»‰ﬂ"
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
      Left            =   4440
      TabIndex        =   16
      Top             =   0
      Width           =   3855
   End
   Begin VB.Shape Shape1 
      Height          =   7695
      Index           =   2
      Left            =   120
      Top             =   1800
      Width           =   12615
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   12615
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—’Ìœ «·⁄«„ ›Ì «·»‰ﬂ"
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
      TabIndex        =   14
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—’Ìœ"
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
      Left            =   5280
      TabIndex        =   13
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label12 
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
      Left            =   3960
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «·„»«·€ «·„”ÕÊ»…"
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
      Left            =   10320
      TabIndex        =   11
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label8 
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
      Left            =   9240
      TabIndex        =   10
      Top             =   1440
      Width           =   1335
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
      Left            =   6240
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ „»«·€ «·«Ìœ«⁄"
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
      Left            =   6960
      TabIndex        =   8
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   12615
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄—÷ ⁄„·Ì«  «·”Õ» Ê«·«Ìœ«⁄ „‰  «—ÌŒ"
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
      Left            =   7560
      TabIndex        =   7
      Top             =   840
      Width           =   3615
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
      Left            =   5280
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "Compte_BNK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
Me.Top = 0
Me.Left = 0
Call chargegrd2_T
End Sub
Private Sub chargegrd2_M()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim d As Double
Dim sd As Double
Dim r As Double
Dim sr As Double
Dim s As Double
grd2.Clear
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1300
grd2.ColWidth(2) = 1200
grd2.ColWidth(3) = 1500
grd2.ColWidth(4) = 1200
grd2.ColWidth(5) = 6900
grd2.ColWidth(6) = 0
grd2.ColWidth(7) = 0
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 3
grd2.ColAlignment(2) = 3
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 3
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 3
grd2.ColAlignment(7) = 3
grd2.Row = 0
grd2.Col = 1
grd2.Text = "«· «—ÌŒ"
grd2.Col = 2
grd2.Text = "«·”«⁄…"
grd2.Col = 3
grd2.Text = "«·„»·€"
grd2.Col = 4
grd2.Text = "‰Ê⁄ «·⁄„·Ì…"
grd2.Col = 5
grd2.Text = "«· ›«’Ì·"
i = 1
dat1 = DT2.Value
dat2 = DT3.Value
a = 0
sd = 0
sr = 0
Call cont
grd2.Rows = bn.RecordCount + 3
Do While Not bn.EOF
dat3 = bn!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If bn!act <> "1" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = bn!aut
grd2.Col = 1
grd2.Text = bn!dat
If bn!act = "2" Then
grd2.CellBackColor = &HFF&
End If
grd2.Col = 2
grd2.Text = bn!heu
grd2.Col = 3
grd2.Text = bn!mon
grd2.Col = 4
grd2.Text = bn!typ
If bn!typ = "«Ìœ«⁄" Then
d = bn!mon
sd = sd + d
Else
r = bn!mon
sr = sr + r
End If
grd2.Col = 5
grd2.Text = bn!det
i = i + 1
End If
End If
bn.MoveNext
Loop
grd2.Rows = i
Label7.Caption = sd
Label8.Caption = sr
s = (sd - sr)
Label12.Caption = s
End Sub
Private Sub chargegrd2_T()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim d As Double
Dim sd As Double
Dim r As Double
Dim sr As Double
Dim s As Double
grd2.Clear
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1300
grd2.ColWidth(2) = 1200
grd2.ColWidth(3) = 1500
grd2.ColWidth(4) = 1200
grd2.ColWidth(5) = 6900
grd2.ColWidth(6) = 0
grd2.ColWidth(7) = 0
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 3
grd2.ColAlignment(2) = 3
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 3
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 3
grd2.ColAlignment(7) = 3
grd2.Row = 0
grd2.Col = 1
grd2.Text = "«· «—ÌŒ"
grd2.Col = 2
grd2.Text = "«·”«⁄…"
grd2.Col = 3
grd2.Text = "«·„»·€"
grd2.Col = 4
grd2.Text = "‰Ê⁄ «·⁄„·Ì…"
grd2.Col = 5
grd2.Text = "«· ›«’Ì·"
i = 1
dat1 = DT2.Value
dat2 = DT3.Value
a = 0
sd = 0
sr = 0
Call cont
grd2.Rows = bn.RecordCount + 3
Do While Not bn.EOF
dat3 = bn!dat
If bn!act <> "1" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = bn!aut
grd2.Col = 1
grd2.Text = bn!dat
If bn!act = "2" Then
grd2.CellBackColor = &HFF&
End If
grd2.Col = 2
grd2.Text = bn!heu
grd2.Col = 3
grd2.Text = bn!mon
grd2.Col = 4
grd2.Text = bn!typ
grd2.Col = 5
grd2.Text = bn!det
i = i + 1
End If
If bn!typ = "«Ìœ«⁄" Then
d = bn!mon
sd = sd + d
Else
r = bn!mon
sr = sr + r
End If
bn.MoveNext
Loop
grd2.Rows = i
Label7.Caption = sd
Label8.Caption = sr
s = (sd - sr)
Label12.Caption = s
Label1.Caption = s
End Sub
