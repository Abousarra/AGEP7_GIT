VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Compte_ARC 
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
      Left            =   4800
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9000
      Width           =   3255
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
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "ÌÊ„Ì"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   4
      Top             =   840
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0000FFFF&
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
      ItemData        =   "Compte_ARC.frx":0000
      Left            =   4800
      List            =   "Compte_ARC.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "‘Â—Ì"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "”‰ÊÌ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H0000FFFF&
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
      ItemData        =   "Compte_ARC.frx":0054
      Left            =   9120
      List            =   "Compte_ARC.frx":006E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   345
      Left            =   6480
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
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
      Format          =   104398849
      CurrentDate     =   42638
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   7455
      Left            =   240
      TabIndex        =   9
      Top             =   1440
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
   Begin VB.Shape Shape1 
      Height          =   8175
      Index           =   0
      Left            =   120
      Top             =   1320
      Width           =   12615
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   3
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   3495
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«—‘Ì› «·’‰œÊﬁ"
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
      Left            =   4800
      TabIndex        =   8
      Top             =   0
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   2
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   4
      Left            =   9000
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   3735
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·”Ã·"
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
      Index           =   5
      Left            =   11160
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "Compte_ARC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chargegrd1_D()
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
grd1.Clear
grd1.Cols = 7
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 1000
grd1.ColWidth(3) = 800
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 5800
grd1.ColWidth(6) = 1200
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«· «—ÌŒ"
grd1.Col = 2
grd1.Text = "«·”«⁄…"
grd1.Col = 3
grd1.Text = "«·‰Ê⁄"
grd1.Col = 4
grd1.Text = "«·„»·€"
grd1.Col = 5
grd1.Text = " ›«’Ì·"
grd1.Col = 6
grd1.Text = "«·„‰›–"
i = 1
dat1 = DT1.Value
Call cont
grd1.Rows = ca.RecordCount + 3
Do While Not ca.EOF
If ca!com = Combo3.Text Or Combo3.Text = "Ã„Ì⁄ «·”Ã·« " Then
dat2 = ca!dat
If dat1 = dat2 Then
grd1.Row = i
grd1.Col = 0
grd1.Text = ca!aut
grd1.Col = 1
grd1.Text = ca!dat
grd1.Col = 2
grd1.Text = ca!heu
grd1.Col = 3
grd1.Text = ca!typ
grd1.Col = 4
grd1.Text = ca!mon
grd1.Col = 5
grd1.Text = ca!det
grd1.Col = 6
grd1.Text = ca!uti
i = i + 1
End If
End If
ca.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 2
End Sub

Private Sub Combo1_Change()
'Call grds_clear

End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Combo3_Change()
'Call grds_clear

End Sub

Private Sub Combo3_Click()
Combo3_Change
End Sub

Private Sub Command7_Click()
If Combo3.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄ «·”Ã·", vbCritical
Exit Sub
End If
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  ⁄·Ï «·Ì„Ì‰", vbCritical
Exit Sub
End If
Command7.Enabled = False
grd1.Visible = False
'Call grds_clear
If Option1.Value = True Then
Call chargegrd1_D
End If
If Option2.Value = True Then
If Combo1.Text = "" Then
grd1.Visible = True
MsgBox "ﬁ„ » ÕœÌœ «·‘Â— „‰ Œ·«· «·ﬁ«∆„… «·„‰”œ·…", vbCritical
Command7.Enabled = True
Exit Sub
End If
'Call chargegrd1_M
End If
If Option3.Value = True Then
'Call chargegrd1_T
End If
grd1.Visible = True
Command7.Enabled = True

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
DT1.Value = Date
End Sub

Private Sub Option1_Click()
DT1.Visible = True
Combo1.Visible = False
'Call grds_clear

End Sub

Private Sub Option2_Click()
'Call grds_clear
DT1.Visible = False
Combo1.Visible = True


End Sub

Private Sub Option3_Click()
'Call grds_clear
DT1.Visible = False
Combo1.Visible = False

End Sub
