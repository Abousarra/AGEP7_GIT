VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Caisse_BNK 
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
      TabIndex        =   34
      Top             =   720
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
         TabIndex        =   35
         Top             =   8400
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid grd6 
         Height          =   8295
         Left            =   0
         TabIndex        =   36
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
      TabIndex        =   23
      Top             =   780
      UseMaskColor    =   -1  'True
      Width           =   735
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
      TabIndex        =   22
      Top             =   780
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.PictureBox Picture3 
      Height          =   4215
      Left            =   3480
      ScaleHeight     =   4155
      ScaleWidth      =   5475
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1440
         TabIndex        =   39
         Text            =   "Text5"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2880
         TabIndex        =   38
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   4320
         TabIndex        =   32
         Top             =   1560
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Text            =   "Text4"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Label19"
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   255
         Left            =   2400
         TabIndex        =   40
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "1"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1320
         Width           =   975
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
         Left            =   2520
         TabIndex        =   31
         Top             =   3720
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
         Left            =   360
         TabIndex        =   30
         Top             =   3720
         Width           =   2655
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
         Left            =   120
         TabIndex        =   29
         Top             =   3360
         Width           =   2655
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
         Left            =   2640
         TabIndex        =   28
         Top             =   3360
         Width           =   2295
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
         Left            =   -600
         TabIndex        =   27
         Top             =   2880
         Width           =   2895
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
         Left            =   0
         TabIndex        =   26
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·—’Ìœ ›Ì «·»‰ﬂ"
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
         Left            =   2400
         TabIndex        =   25
         Top             =   2400
         Width           =   1575
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
         Left            =   360
         TabIndex        =   24
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "0"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1815
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
      Height          =   375
      Left            =   8520
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   780
      Width           =   1335
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
      TabIndex        =   4
      Top             =   1400
      Width           =   735
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
      Left            =   1080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1400
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
      TabIndex        =   2
      Top             =   780
      Width           =   3135
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
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
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
      ItemData        =   "Caisse_BNK.frx":0000
      Left            =   6360
      List            =   "Caisse_BNK.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   780
      Width           =   975
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   345
      Left            =   10560
      TabIndex        =   10
      Top             =   780
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
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   9600
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   7575
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   13361
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
      Left            =   4800
      TabIndex        =   13
      Top             =   1395
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
      Left            =   2400
      TabIndex        =   14
      Top             =   1395
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
   Begin MSFlexGridLib.MSFlexGrid grd5 
      Height          =   615
      Left            =   10560
      TabIndex        =   37
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
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "’‰œÊﬁ «·»‰ﬂ"
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
      TabIndex        =   21
      Top             =   0
      Width           =   3735
   End
   Begin VB.Line Line1 
      X1              =   11280
      X2              =   11280
      Y1              =   1320
      Y2              =   1800
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
      Left            =   9600
      TabIndex        =   20
      Top             =   780
      Width           =   855
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
      TabIndex        =   19
      Top             =   780
      Width           =   1335
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
      Left            =   3600
      TabIndex        =   18
      Top             =   1440
      Width           =   1095
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
      Left            =   5880
      TabIndex        =   17
      Top             =   1440
      Width           =   3615
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
      Left            =   4920
      TabIndex        =   16
      Top             =   780
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«Ìœ«⁄/”Õ»"
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
      TabIndex        =   15
      Top             =   780
      Width           =   975
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
      Top             =   1320
      Width           =   12615
   End
End
Attribute VB_Name = "Caisse_BNK"
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

Private Sub Combo1_Change()
On Error Resume Next
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
Else
Combo1.BackColor = &H8080FF
End If

End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
Label16.Caption = "frdgtfrrrr"
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
Text2.Text = ""
Text3.Text = ""
Text2.SetFocus
DT1.Value = Date
Label17.Caption = ""
Label11.Caption = ""
Label16.Caption = "0"
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
Call Operations

End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
If Text2.Text = "" Or Combo1.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text2.BackColor = &H8080FF Then
Text2.SetFocus
End If
Exit Sub
End If
'**** controle bank
a = Label1.Caption
b = Text2.Text
c = 0
If Label11.Caption = "”Õ»" Then
c = Label16.Caption
End If
If Label11.Caption = "«Ìœ«⁄" Then
c = Label16.Caption
c = c * (-1)
End If
If Combo1.Text = "”Õ»" Then
d = (a + c) - b
Else
d = (a + c) + b
End If
If d < 0 Then
MsgBox "—’Ìœ «·»‰ﬂ €Ì— ﬂ«› ·≈ „«„ «·⁄„·Ì…", vbExclamation
Exit Sub
End If
'** controle caisse
Call cont
a = eb!cca
b = Text2.Text
c = 0
If Label11.Caption = "”Õ»" Then
c = Label16.Caption
c = c * (-1)
End If
If Label11.Caption = "«Ìœ«⁄" Then
c = Label16.Caption
End If
If Combo1.Text = "«Ìœ«⁄" Then
d = (a + c) - b
Else
d = (a + c) + b
End If
If d < 0 Then
MsgBox "—’Ìœ «·’‰œÊﬁ €Ì— ﬂ«› ·≈ „«„ «·⁄„·Ì…", vbExclamation
Exit Sub
End If
eb!cca = d
eb.Update
If Combo1.Text = "«Ìœ«⁄" Then
a = Label16.Caption
b = Text2.Text
c = Label1.Caption
c = (c - a) + b
Label1.Caption = c
End If
If Combo1.Text = "”Õ»" Then
a = Label16.Caption
b = Text2.Text
c = Label1.Caption
c = (c + a) - b
Label1.Caption = c
End If
'**** archive de caisse ajou et modif
Adat = Date
Aheu = Time$
If Label17.Caption = "" Then
Atyp = "≈÷«›…"
Else
Atyp = " ⁄œÌ·"
End If
Adet = Combo1.Text + "   ||   " + Text3.Text
Amon = Text2.Text
Acom = "”Ã· «·»‰ﬂ"
Auti = directions.Label2.Caption
'****************************************
If Label17.Caption <> "" Then
Call cont
Do While Not bn.EOF
If Label17.Caption = bn!aut Then
bn!dat = DT1.Value
bn!mon = Text2.Text
bn!typ = Combo1.Text
bn!det = Text3.Text
bn!heu = Time$
If bn!act = "2" Then
bn!act = "3"
End If
bn.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
bn.MoveNext
Loop
End If
bn.AddNew
bn!dat = DT1.Value
bn!mon = Text2.Text
bn!typ = Combo1.Text
bn!det = Text3.Text
bn!heu = Time$
bn!act = "0"
bn!mtf = ""
bn.Update
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
Do While Not bn.EOF
If bn!aut = tx1 Then
MsgBox bn!mtf
Exit Sub
End If
bn.MoveNext
Loop
End If
End If
If j = 6 Then
grd2.Row = i
grd2.Col = 0
Label17.Caption = grd2.Text
grd2.Col = 1
DT1.Value = grd2.Text
grd2.Col = 3
Text2.Text = grd2.Text
Label16.Caption = grd2.Text
grd2.Col = 4
Combo1.Text = grd2.Text
Label11.Caption = grd2.Text
grd2.Col = 5
Text3.Text = grd2.Text
End If
If j = 7 Then
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–Â «·⁄„·Ì…", vbInformation + vbYesNo + arabic, "AGEP7")
If g = vbYes Then
a = eb!cca
grd2.Row = i
grd2.Col = 0
Label17.Caption = grd2.Text
grd2.Col = 3
b = grd2.Text
grd2.Col = 4
Label11.Caption = grd2.Text
If Label11.Caption = "”Õ»" Then
'*** controle caisse
If b > a Then
MsgBox "—’Ìœ «·’‰œÊﬁ ·« Ì”„Õ »« „«„ «·⁄„·Ì…... Ì—ÃÏ ÷Œ „»·€ ÃœÌœ ›Ì «·’‰œÊﬁ", vbExclamation
Exit Sub
End If
'***
b = -b
End If
a = a + b
eb!cca = a
eb.Update
a = Label1.Caption
Label1.Caption = (a - b)
Call cont
Do While Not bn.EOF
If Label17.Caption = bn!aut Then
'**** archive de caisse supp
If b < 0 Then
b = -b
End If
Label18.Caption = bn!det
Label19.Caption = bn!typ
Adat = Date
Aheu = Time$
Atyp = "Õ–›"
Adet = Label19.Caption + "   ||   " + Label18.Caption
Amon = b
Acom = "”Ã· «·»‰ﬂ"
Auti = directions.Label2.Caption
'****************************************
bn.Delete
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
bn.MoveNext
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
grd2.ColWidth(5) = 5500
grd2.ColWidth(6) = 800
grd2.ColWidth(7) = 800
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
grd2.Col = 6
grd2.Text = " ⁄œÌ·"
grd2.CellBackColor = &HFFFF&
grd2.Col = 7
grd2.Text = "Õ–›"
grd2.CellBackColor = &HFF&
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
grd2.ColWidth(5) = 5500
grd2.ColWidth(6) = 800
grd2.ColWidth(7) = 800
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
grd2.Col = 6
grd2.Text = " ⁄œÌ·"
grd2.CellBackColor = &HFFFF&
grd2.Col = 7
grd2.Text = "Õ–›"
grd2.CellBackColor = &HFF&
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
Do While Not bn.EOF
If bn!act = "0" Or bn!act = "3" Then
a = a + 1
End If
If bn!act = "2" Then
b = b + 1
End If
bn.MoveNext
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
grd6.Rows = bn.RecordCount + 3
Do While Not bn.EOF
If bn!act = Text1.Text Or bn!act = Text5.Text Then
grd6.Row = i
grd6.Col = 0
grd6.Text = bn!dat
grd6.Col = 1
grd6.Text = tx
If Text1.Text = "0" Then
grd6.CellBackColor = &HFFFF&
Else
grd6.CellBackColor = &HFF&
End If
i = i + 1
End If
bn.MoveNext
Loop
grd6.Rows = i
End Sub
