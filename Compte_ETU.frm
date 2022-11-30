VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Compte_ETU 
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
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "”Õ»"
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
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "”Õ»"
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
      Left            =   4440
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "”Õ»"
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
      Left            =   8640
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9240
      Width           =   1095
   End
   Begin VB.OptionButton Option6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   " ·«„Ì– «·ﬁ”„"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
   End
   Begin VB.OptionButton Option5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   " ·«„Ì– «·ﬁ”„"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10560
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Ã„Ì⁄ «· ·«„Ì– ›Ì «·”‰… «·œ—«”Ì…"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   9
      Top             =   720
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Ã„Ì⁄ «· ·«„Ì– ›Ì ‘Â—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9840
      TabIndex        =   8
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox Combo7 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Compte_ETU.frx":0000
      Left            =   8040
      List            =   "Compte_ETU.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Compte_ETU.frx":0053
      Left            =   2640
      List            =   "Compte_ETU.frx":007B
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   720
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Compte_ETU.frx":00A6
      Left            =   9600
      List            =   "Compte_ETU.frx":00CE
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Compte_ETU.frx":00F9
      Left            =   8040
      List            =   "Compte_ETU.frx":0121
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.ComboBox Combo5 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Compte_ETU.frx":014C
      Left            =   5520
      List            =   "Compte_ETU.frx":0174
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.ComboBox Combo6 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Compte_ETU.frx":019F
      Left            =   2640
      List            =   "Compte_ETU.frx":01C7
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
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
      Height          =   825
      Left            =   1200
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   6975
      Left            =   4440
      TabIndex        =   12
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   12303
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
      Height          =   6975
      Left            =   8640
      TabIndex        =   15
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   12303
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
      Height          =   6975
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   12303
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
   Begin VB.Shape Shape1 
      Height          =   7935
      Index           =   5
      Left            =   120
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Shape Shape1 
      Height          =   7935
      Index           =   3
      Left            =   4320
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "·«∆Õ… «· ·«„Ì–"
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
      Index           =   6
      Left            =   8640
      TabIndex        =   28
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ã„Ê⁄"
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
      Index           =   5
      Left            =   3360
      TabIndex        =   26
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1440
      TabIndex        =   25
      Top             =   9240
      Width           =   1755
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ã„Ê⁄"
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
      Index           =   4
      Left            =   7560
      TabIndex        =   23
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5640
      TabIndex        =   22
      Top             =   9240
      Width           =   1755
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ã„Ê⁄"
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
      Index           =   3
      Left            =   11760
      TabIndex        =   20
      Top             =   9240
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9840
      TabIndex        =   19
      Top             =   9240
      Width           =   1755
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "·«∆Õ… «· ·«„Ì– «·–Ì‰ ·„ Ìœ›⁄Ê« »⁄œ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "·«∆Õ… «· ·«„Ì– «·–Ì‰ œ›⁄Ê«"
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
      Index           =   2
      Left            =   4440
      TabIndex        =   16
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "›Ì ‘Â—"
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
      Index           =   1
      Left            =   8880
      TabIndex        =   14
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "›Ì «·”‰… «·œ—«”Ì…"
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
      Index           =   2
      Left            =   3600
      TabIndex        =   13
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   0
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   9375
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   9375
   End
   Begin VB.Line Line1 
      X1              =   7920
      X2              =   7920
      Y1              =   600
      Y2              =   1560
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Index           =   2
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   7935
      Index           =   4
      Left            =   8520
      Top             =   1680
      Width           =   4215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» «· ·«„Ì–"
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
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "Compte_ETU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command7_Click()
On Error Resume Next
grd1.Visible = False
grd2.Visible = False
grd3.Visible = False
If Option2.Value = True Then
Call chargegrd1_2
Call chargegrd2_2
Call chargegrd3_T
End If
If Option3.Value = True Then
Call chargegrd1_3
Call chargegrd2_3
Call chargegrd3_T
End If
If Option5.Value = True Then
Call chargegrd1_5
Call chargegrd2_5
Call chargegrd3_T
End If
If Option6.Value = True Then
Call chargegrd1_6
Call chargegrd2_6
Call chargegrd3_T
End If
grd1.Visible = True
grd2.Visible = True
grd3.Visible = True
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 0
Me.Left = 0
Call chargcombo1_3_5
End Sub
Private Sub chargcombo1_3_5()
On Error Resume Next
Combo3.Clear
Combo5.Clear
Call cont
Do While Not cl.EOF
Combo3.AddItem cl!cla
Combo5.AddItem cl!cla
cl.MoveNext
Loop
End Sub
Private Sub chargegrd1_5()
On Error Resume Next
Dim i As Double
Dim h As String
grd1.Clear
grd1.Cols = 4
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1000
grd1.ColWidth(2) = 2000
grd1.ColWidth(3) = 700
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "—.«· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«·«”„"
grd1.Col = 3
grd1.Text = "«·ﬁ”„"
i = 1
Call cont
grd1.Rows = et.RecordCount + 3
Do While Not et.EOF
If et!cla = Combo3.Text Then
grd1.Row = i
grd1.Col = 1
grd1.Text = et!sri
grd1.Col = 2
grd1.Text = et!nom
grd1.Col = 3
grd1.Text = et!cla
i = i + 1
End If
et.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
Label4.Caption = (i - 1)
'MsgBox h
End Sub
Private Sub chargegrd2_5()
On Error Resume Next
Dim h As String
Dim i As Double
grd2.Clear
grd2.Cols = 4
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1000
grd2.ColWidth(2) = 2000
grd2.ColWidth(3) = 700
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.Row = 0
grd2.Col = 1
grd2.Text = "—.«· ”·”·Ì"
grd2.Col = 2
grd2.Text = "«·«”„"
grd2.Col = 3
grd2.Text = "«·ﬁ”„"
i = 1
Call cont
grd2.Rows = ct.RecordCount + 3
Do While Not ct.EOF
If ct!cla = Combo3.Text And ct!moi = Combo4.Text Then
grd2.Row = i
grd2.Col = 1
grd2.Text = ct!sri
grd2.Col = 2
grd2.Text = ct!nom
grd2.Col = 3
grd2.Text = ct!pay
i = i + 1
End If
ct.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 1
Label5.Caption = (i - 1)
'MsgBox h
End Sub
Private Sub chargegrd1_2()
On Error Resume Next
Dim i As Double
Dim h As String
grd1.Clear
grd1.Cols = 4
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1000
grd1.ColWidth(2) = 2000
grd1.ColWidth(3) = 700
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "—.«· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«·«”„"
grd1.Col = 3
grd1.Text = "«·ﬁ”„"
i = 1
Call cont
grd1.Rows = et.RecordCount + 3
Do While Not et.EOF
'If et!cla = Combo3.Text Then
grd1.Row = i
grd1.Col = 1
grd1.Text = et!sri
grd1.Col = 2
grd1.Text = et!nom
grd1.Col = 3
grd1.Text = et!cla
i = i + 1
'End If
et.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
Label4.Caption = (i - 1)
'MsgBox h
End Sub
Private Sub chargegrd2_2()
On Error Resume Next
Dim h As String
Dim i As Double
grd2.Clear
grd2.Cols = 4
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1000
grd2.ColWidth(2) = 2000
grd2.ColWidth(3) = 700
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.Row = 0
grd2.Col = 1
grd2.Text = "—.«· ”·”·Ì"
grd2.Col = 2
grd2.Text = "«·«”„"
grd2.Col = 3
grd2.Text = "«·ﬁ”„"
i = 1
Call cont
grd2.Rows = ct.RecordCount + 3
Do While Not ct.EOF
If ct!moi = Combo7.Text Then
grd2.Row = i
grd2.Col = 1
grd2.Text = ct!sri
grd2.Col = 2
grd2.Text = ct!nom
grd2.Col = 3
grd2.Text = ct!pay
i = i + 1
End If
ct.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 1
Label5.Caption = (i - 1)
'MsgBox h
End Sub
Private Sub chargegrd1_3()
On Error Resume Next
Dim i As Double
Dim h As String
grd1.Clear
grd1.Cols = 4
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1000
grd1.ColWidth(2) = 2000
grd1.ColWidth(3) = 700
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "—.«· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«·«”„"
grd1.Col = 3
grd1.Text = "«·ﬁ”„"
i = 1
Call cont
grd1.Rows = et.RecordCount + 3
Do While Not et.EOF
'If et!cla = Combo3.Text Then
grd1.Row = i
grd1.Col = 1
grd1.Text = et!sri
grd1.Col = 2
grd1.Text = et!nom
grd1.Col = 3
grd1.Text = et!cla
i = i + 1
'End If
et.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
Label4.Caption = (i - 1)
'MsgBox h
End Sub
Private Sub chargegrd2_3()
On Error Resume Next
Dim h As String
Dim i As Double
grd2.Clear
grd2.Cols = 4
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1000
grd2.ColWidth(2) = 2000
grd2.ColWidth(3) = 700
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.Row = 0
grd2.Col = 1
grd2.Text = "—.«· ”·”·Ì"
grd2.Col = 2
grd2.Text = "«·«”„"
grd2.Col = 3
grd2.Text = "«·ﬁ”„"
i = 1
Call cont
grd2.Rows = ct.RecordCount + 3
Do While Not ct.EOF
If ct!rcu <> "0" Then
grd2.Row = i
grd2.Col = 1
grd2.Text = ct!sri
grd2.Col = 2
grd2.Text = ct!nom
grd2.Col = 3
grd2.Text = ct!tpy
i = i + 1
End If
ct.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 1
Label5.Caption = (i - 1)
'MsgBox h
End Sub
Private Sub chargegrd1_6()
On Error Resume Next
Dim i As Double
Dim h As String
grd1.Clear
grd1.Cols = 4
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1000
grd1.ColWidth(2) = 2000
grd1.ColWidth(3) = 700
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "—.«· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«·«”„"
grd1.Col = 3
grd1.Text = "«·ﬁ”„"
i = 1
Call cont
grd1.Rows = et.RecordCount + 3
Do While Not et.EOF
If et!cla = Combo5.Text Then
grd1.Row = i
grd1.Col = 1
grd1.Text = et!sri
grd1.Col = 2
grd1.Text = et!nom
grd1.Col = 3
grd1.Text = et!cla
i = i + 1
End If
et.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
Label4.Caption = (i - 1)
'MsgBox h
End Sub
Private Sub chargegrd2_6()
On Error Resume Next
Dim h As String
Dim i As Double
grd2.Clear
grd2.Cols = 4
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1000
grd2.ColWidth(2) = 2000
grd2.ColWidth(3) = 700
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.Row = 0
grd2.Col = 1
grd2.Text = "—.«· ”·”·Ì"
grd2.Col = 2
grd2.Text = "«·«”„"
grd2.Col = 3
grd2.Text = "«·ﬁ”„"
i = 1
Call cont
grd2.Rows = ct.RecordCount + 3
Do While Not ct.EOF
If ct!cla = Combo5.Text Then
grd2.Row = i
grd2.Col = 1
grd2.Text = ct!sri
grd2.Col = 2
grd2.Text = ct!nom
grd2.Col = 3
grd2.Text = ct!pay
i = i + 1
End If
ct.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 1
Label5.Caption = (i - 1)
'MsgBox h
End Sub
Private Sub chargegrd3_T()
On Error Resume Next
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
grd3.ColWidth(2) = 2000
grd3.ColWidth(3) = 700
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.Row = 0
grd3.Col = 1
grd3.Text = "—.«· ”·”·Ì"
grd3.Col = 2
grd3.Text = "«·«”„"
grd3.Col = 3
grd3.Text = "«·ﬁ”„"
n = grd1.Rows
m = grd2.Rows
q = 1
grd3.Rows = n + 3
For i = 1 To n - 1
k = 0
grd1.Row = i
grd1.Col = 1
tx1 = grd1.Text
grd1.Col = 2
tx2 = grd1.Text
grd1.Col = 3
tx3 = grd1.Text
For j = 1 To m - 1
grd2.Row = j
grd2.Col = 1
tx4 = grd2.Text
If tx1 = tx4 Then
k = 1
j = m
End If
Next j
If k = 0 Then
grd3.Row = q
grd3.Col = 1
grd3.Text = tx1
grd3.Col = 2
grd3.Text = tx2
grd3.Col = 3
grd3.Text = tx3
q = q + 1
End If
Next i
grd3.Rows = q
grd3.Col = 1
grd3.Sort = 1
Label7.Caption = (q - 1)
'MsgBox h
End Sub

Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim m As Double
Dim s As Double
Dim t As String
n = grd2.Rows
s = 0
For i = 1 To n - 1
grd2.Row = i
grd2.Col = 3
m = grd2.Text
s = s + m
Next i
t = s
MsgBox t
End Sub
