VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Compte_DPS 
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
      Left            =   10200
      TabIndex        =   8
      Top             =   840
      Width           =   1215
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
      Left            =   1440
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1335
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
      Left            =   2880
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   7455
      Left            =   240
      TabIndex        =   2
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
      TabIndex        =   3
      Top             =   840
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
      Format          =   22609921
      CurrentDate     =   42638
   End
   Begin MSComCtl2.DTPicker DT3 
      Height          =   345
      Left            =   4680
      TabIndex        =   4
      Top             =   840
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
      Format          =   22609921
      CurrentDate     =   42638
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "≈Ã„«·Ì «·„»«·€ «·„’—Ê›…"
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
      Left            =   9120
      TabIndex        =   12
      Top             =   1440
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
      Left            =   6720
      TabIndex        =   11
      Top             =   1440
      Width           =   2295
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
      Left            =   1320
      TabIndex        =   10
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄… «·„»«·€ «·„’—Ê›… ›Ì «· «—ÌŒ √⁄·«Â"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   0
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   10335
   End
   Begin VB.Shape Shape1 
      Height          =   7695
      Index           =   3
      Left            =   120
      Top             =   1800
      Width           =   12615
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» «·„’—Ê›« "
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
      Left            =   4560
      TabIndex        =   7
      Top             =   0
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   2
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   10335
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
      Left            =   5880
      TabIndex        =   6
      Top             =   840
      Width           =   1095
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
      Left            =   8160
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Compte_DPS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Call chargegrd2_T
End Sub
Private Sub chargegrd2_M()
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
grd2.ColWidth(4) = 6600
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 0
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
i = i + 1
End If
End If
dp.MoveNext
Loop
grd2.Rows = i
Label1.Caption = sd
End Sub
Private Sub chargegrd2_T()
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
grd2.ColWidth(4) = 6600
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 0
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
i = i + 1
End If
dp.MoveNext
Loop
grd2.Rows = i
Label1.Caption = sd
Label7.Caption = sd
End Sub
