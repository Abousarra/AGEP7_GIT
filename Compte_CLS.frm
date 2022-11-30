VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Compte_CLS 
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
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   5160
      ScaleHeight     =   2955
      ScaleWidth      =   3075
      TabIndex        =   28
      Top             =   4920
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   1920
         TabIndex        =   37
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "ﬁ”„"
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
         Left            =   3000
         TabIndex        =   33
         Top             =   1800
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         Caption         =   "Ã„Ì⁄ «·√ﬁ”«„ ›Ì ÌÊ„"
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
         Left            =   1680
         TabIndex        =   32
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
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
         ItemData        =   "Compte_CLS.frx":0000
         Left            =   2040
         List            =   "Compte_CLS.frx":0028
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1800
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   300
         Left            =   0
         TabIndex        =   34
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
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
      Begin MSComCtl2.DTPicker DT2 
         Height          =   300
         Left            =   0
         TabIndex        =   35
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
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
      Begin VB.Label Label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "›Ì ÌÊ„"
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
         Left            =   1320
         TabIndex        =   36
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Label19"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
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
      Left            =   5400
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9240
      Width           =   2055
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
      Left            =   1920
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   720
      Width           =   1095
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
      ItemData        =   "Compte_CLS.frx":0053
      Left            =   3360
      List            =   "Compte_CLS.frx":007B
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
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
      ItemData        =   "Compte_CLS.frx":00A6
      Left            =   6120
      List            =   "Compte_CLS.frx":00CE
      Style           =   2  'Dropdown List
      TabIndex        =   8
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
      ItemData        =   "Compte_CLS.frx":00F9
      Left            =   7920
      List            =   "Compte_CLS.frx":0121
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1200
      Width           =   735
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
      ItemData        =   "Compte_CLS.frx":014C
      Left            =   9480
      List            =   "Compte_CLS.frx":0174
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1200
      Width           =   855
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
      ItemData        =   "Compte_CLS.frx":019F
      Left            =   3360
      List            =   "Compte_CLS.frx":01C7
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   1215
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
      ItemData        =   "Compte_CLS.frx":01F2
      Left            =   7920
      List            =   "Compte_CLS.frx":021A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Ã„Ì⁄ «·√ﬁ”«„ ›Ì ‘Â—"
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
      Left            =   9120
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Ã„Ì⁄ «·√ﬁ”«„ ›Ì «·”‰… «·œ—«”Ì…"
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
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.OptionButton Option5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "ﬁ”„"
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
      Left            =   10440
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.OptionButton Option6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "ﬁ”„"
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
      Left            =   7080
      TabIndex        =   7
      Top             =   1200
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   6495
      Left            =   240
      TabIndex        =   26
      Top             =   2640
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   11456
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
      Height          =   7095
      Index           =   4
      Left            =   120
      Top             =   2520
      Width           =   12615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰’Ì» «·„ƒ””…"
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
      Index           =   1
      Left            =   2040
      TabIndex        =   25
      Top             =   1800
      Width           =   2055
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
      Height          =   375
      Left            =   1320
      TabIndex        =   24
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «·„»«·€"
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
      Left            =   2280
      TabIndex        =   23
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      Left            =   1320
      TabIndex        =   22
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      Height          =   855
      Index           =   3
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   10815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «·—”Ê„ «·‘Â—Ì…"
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
      Left            =   9720
      TabIndex        =   21
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ „” Õﬁ«  √”« –… ” ‘"
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
      Left            =   5640
      TabIndex        =   20
      Top             =   1800
      Width           =   2535
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
      Height          =   375
      Left            =   8400
      TabIndex        =   19
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ —”Ê„ «· ”ÃÌ·"
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
      Left            =   9960
      TabIndex        =   18
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label13 
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
      Left            =   4320
      TabIndex        =   17
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ „” Õﬁ«  √”« –… «·‰”»…"
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
      Left            =   5760
      TabIndex        =   16
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label20 
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
      Left            =   8400
      TabIndex        =   15
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label Label10 
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
      Left            =   4320
      TabIndex        =   14
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Index           =   2
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   7800
      X2              =   7800
      Y1              =   600
      Y2              =   1560
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   7935
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   0
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   7935
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
      Left            =   4200
      TabIndex        =   12
      Top             =   1200
      Width           =   1815
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
      Left            =   8760
      TabIndex        =   11
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» «·√ﬁ”«„"
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
      TabIndex        =   10
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "Compte_CLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo3_Change()
On Error Resume Next
Call grd2_clear

End Sub

Private Sub Combo3_Click()
On Error Resume Next
Combo3_Change
End Sub

Private Sub Combo4_Change()
On Error Resume Next
Call grd2_clear

End Sub

Private Sub Combo4_Click()
On Error Resume Next
Combo4_Change
End Sub

Private Sub Combo5_Change()
On Error Resume Next
Call grd2_clear

End Sub

Private Sub Combo5_Click()
On Error Resume Next
Combo5_Change
End Sub

Private Sub Combo7_Change()
On Error Resume Next
Call grd2_clear

End Sub

Private Sub Combo7_Click()
On Error Resume Next
Combo7_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
Call cont
Do While Not pr.EOF
pr.Delete
pr.MoveNext
Loop
Call cont
Do While Not sr.EOF
sr.Delete
sr.MoveNext
Loop
Call cont
Do While Not fc.EOF
fc.Delete
fc.MoveNext
Loop
Call cont
Do While Not pf.EOF
pf.Delete
pf.MoveNext
Loop
Call cont
Do While Not cl.EOF
cl.Delete
cl.MoveNext
Loop
Call cont
Do While Not cr.EOF
cr.Delete
cr.MoveNext
Loop
Call cont
Do While Not et.EOF
et.Delete
et.MoveNext
Loop
Call cont
Do While Not mt.EOF
mt.Delete
mt.MoveNext
Loop
Call cont
Do While Not nt.EOF
nt.Delete
nt.MoveNext
Loop
Call cont
Do While Not pp.EOF
pp.Delete
pp.MoveNext
Loop
Call cont
Do While Not pc.EOF
pc.Delete
pc.MoveNext
Loop
Call cont
Do While Not em.EOF
em.Delete
em.MoveNext
Loop
Call cont
Do While Not pl.EOF
pl.Delete
pl.MoveNext
Loop
Call cont
Do While Not cp.EOF
cp.Delete
cp.MoveNext
Loop
Call cont
Do While Not cf.EOF
cf.Delete
cf.MoveNext
Loop
Call cont
Do While Not cs.EOF
cs.Delete
cs.MoveNext
Loop
Call cont
Do While Not ct.EOF
ct.Delete
ct.MoveNext
Loop
Call cont
Do While Not dp.EOF
dp.Delete
dp.MoveNext
Loop
Call cont
Do While Not bn.EOF
bn.Delete
bn.MoveNext
Loop
Call cont
Do While Not ca.EOF
ca.Delete
ca.MoveNext
Loop
MsgBox "OK"
End Sub

Private Sub Command2_Click()
On Error Resume Next

End Sub

Private Sub Command7_Click()
On Error Resume Next
Label19.Caption = eb!pce
Label8.Caption = eb!moi
grd2.Visible = False
If Option2.Value = True Then
If Combo7.Text = "" Then
grd2.Visible = True
MsgBox "Ì—ÃÏ  ÕœÌœ «·‘Â—", vbCritical
Exit Sub
End If
Call chargegrd2_Tc_m
End If
If Option3.Value = True Then
Call chargegrd2_Tc_Tm
End If
If Option5.Value = True Then
If Combo3.Text = "" Then
grd2.Visible = True
MsgBox "Ì—ÃÏ  ÕœÌœ «·ﬁ”„", vbCritical
Exit Sub
End If
If Combo4.Text = "" Then
grd2.Visible = True
MsgBox "Ì—ÃÏ  ÕœÌœ «·‘Â—", vbCritical
Exit Sub
End If
Call chargegrd2_c_m
End If
If Option6.Value = True Then
If Combo5.Text = "" Then
grd2.Visible = True
MsgBox "Ì—ÃÏ  ÕœÌœ «·ﬁ”„", vbCritical
Exit Sub
End If
Call chargegrd2_c_Tm
End If
grd2.Visible = True
End Sub


Private Sub DT1_Change()
On Error Resume Next
Call grd2_clear

End Sub

Private Sub DT1_Click()
On Error Resume Next
DT1_Change
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 0
Me.Left = 0
Call chargcombo2_6
Call chargcombo1_3_5
Combo2.Text = Interface.SBB1.Panels(1).Text
Combo6.Text = Interface.SBB1.Panels(1).Text
Combo2.Enabled = False
Combo6.Enabled = False
DT1.Value = Date
DT2.Value = Date
End Sub
Private Sub chargcombo2_6()
On Error Resume Next
Combo2.Clear
Combo6.Clear
Call cont
Do While Not an.EOF
Combo2.AddItem an!ann
Combo6.AddItem an!ann
an.MoveNext
Loop
End Sub
Private Sub chargcombo1_3_5()
On Error Resume Next
Combo1.Clear
Combo3.Clear
Combo5.Clear
Call cont
Do While Not cl.EOF
Combo1.AddItem cl!cla
Combo3.AddItem cl!cla
Combo5.AddItem cl!cla
cl.MoveNext
Loop
End Sub
Private Sub chargegrd2_Tc_Tm()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim m As Double
Dim k As Double
Dim e As Double
Dim se As Double
Dim P As Double
Dim sp As Double
Dim r As Double
Dim c As Double
Dim l As Double
Dim sl As Double
Dim f As Double
Dim sf As Double
Dim n As Double
Dim sn As Double
Dim s As Double
i = 1
e = 0
se = 0
P = 0
sp = 0
r = 0
c = 0
l = 0
sl = 0
sf = 0
sn = 0
s = 0
k = Label2.Caption
Call cont
grd2.Rows = pc.RecordCount + 3
Do While Not pc.EOF
m = pc!moi
j = pc!nbr
e = 0
P = 0
r = 0
f = 0
t = 0
grd2.Row = i
grd2.Col = 0
If pc!moi = "0" Then
grd2.Text = "—. "
f = pc!etu
sf = sf + f
Else
grd2.Text = pc!moi
End If
e = pc!etu
se = se + e
grd2.Col = 1
grd2.Text = pc!cla
grd2.Col = 2
grd2.Text = pc!etu
grd2.Col = 3
grd2.Text = pc!pro
P = pc!pro
sp = sp + P
r = (e - P)
grd2.Col = 4
grd2.Text = r
If j > 0 Then
c = Label19.Caption
Else
c = 100
End If
grd2.Col = 5
grd2.Text = c
l = (r * c / 100)
MyNumber = Round(l, 0)
l = MyNumber
sl = sl + l
grd2.Col = 6
grd2.Text = l
c = (100 - c)
n = (r * c / 100)
MyNumber = Round(n, 0)
n = MyNumber
sn = sn + n
grd2.Col = 7
grd2.Text = c
grd2.Col = 8
grd2.Text = n
i = i + 1
pc.MoveNext
Loop
grd2.Rows = i
Label20.Caption = (se - sf)
Label5.Caption = sf
Label10.Caption = sp
Label2.Caption = sl
Label13.Caption = sn
Label4.Caption = (sp + sl + sn)
End Sub
Private Sub chargegrd2_Tc_m()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim m As Double
Dim k As Double
Dim e As Double
Dim se As Double
Dim P As Double
Dim sp As Double
Dim r As Double
Dim c As Double
Dim l As Double
Dim sl As Double
Dim f As Double
Dim sf As Double
Dim n As Double
Dim sn As Double
Dim s As Double
i = 1
e = 0
se = 0
P = 0
sp = 0
r = 0
c = 0
l = 0
sl = 0
sf = 0
sn = 0
s = 0
k = Label2.Caption
Call cont
grd2.Rows = pc.RecordCount + 3
Do While Not pc.EOF
If Combo7.Text = pc!moi Or pc!moi = "0" And Combo7.Text = Label8.Caption Then
m = pc!moi
j = pc!nbr
e = 0
P = 0
r = 0
f = 0
t = 0
grd2.Row = i
grd2.Col = 0
If pc!moi = "0" Then
grd2.Text = "—. "
f = pc!etu
sf = sf + f
Else
grd2.Text = pc!moi
End If
e = pc!etu
se = se + e
grd2.Col = 1
grd2.Text = pc!cla
grd2.Col = 2
grd2.Text = pc!etu
grd2.Col = 3
grd2.Text = pc!pro
P = pc!pro
sp = sp + P
r = (e - P)
grd2.Col = 4
grd2.Text = r
If j > 0 Then
c = Label19.Caption
Else
c = 100
End If
grd2.Col = 5
grd2.Text = c
l = (r * c / 100)
MyNumber = Round(l, 0)
l = MyNumber
sl = sl + l
grd2.Col = 6
grd2.Text = l
c = (100 - c)
n = (r * c / 100)
MyNumber = Round(n, 0)
n = MyNumber
sn = sn + n
grd2.Col = 7
grd2.Text = c
grd2.Col = 8
grd2.Text = n
i = i + 1
End If
pc.MoveNext
Loop
grd2.Rows = i
Label20.Caption = (se - sf)
Label5.Caption = sf
Label10.Caption = sp
Label2.Caption = sl
Label13.Caption = sn
Label4.Caption = (sp + sl + sn)
End Sub
Private Sub chargegrd2_c_m()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim m As Double
Dim k As Double
Dim e As Double
Dim se As Double
Dim P As Double
Dim sp As Double
Dim r As Double
Dim c As Double
Dim l As Double
Dim sl As Double
Dim f As Double
Dim sf As Double
Dim n As Double
Dim sn As Double
Dim s As Double
i = 1
e = 0
se = 0
P = 0
sp = 0
r = 0
c = 0
l = 0
sl = 0
sf = 0
sn = 0
s = 0
k = Label2.Caption
Call cont
grd2.Rows = pc.RecordCount + 3
Do While Not pc.EOF
If Combo3.Text = pc!cla Then
If Combo4.Text = pc!moi Or pc!moi = "0" And Combo4.Text = Label8.Caption Then
m = pc!moi
j = pc!nbr
e = 0
P = 0
r = 0
f = 0
t = 0
grd2.Row = i
grd2.Col = 0
If pc!moi = "0" Then
grd2.Text = "—. "
f = pc!etu
sf = sf + f
Else
grd2.Text = pc!moi
End If
e = pc!etu
se = se + e
grd2.Col = 1
grd2.Text = pc!cla
grd2.Col = 2
grd2.Text = pc!etu
grd2.Col = 3
grd2.Text = pc!pro
P = pc!pro
sp = sp + P
r = (e - P)
grd2.Col = 4
grd2.Text = r
If j > 0 Then
c = Label19.Caption
Else
c = 100
End If
grd2.Col = 5
grd2.Text = c
l = (r * c / 100)
MyNumber = Round(l, 0)
l = MyNumber
sl = sl + l
grd2.Col = 6
grd2.Text = l
c = (100 - c)
n = (r * c / 100)
MyNumber = Round(n, 0)
n = MyNumber
sn = sn + n
grd2.Col = 7
grd2.Text = c
grd2.Col = 8
grd2.Text = n
i = i + 1
End If
End If
pc.MoveNext
Loop
grd2.Rows = i
Label20.Caption = (se - sf)
Label5.Caption = sf
Label10.Caption = sp
Label2.Caption = sl
Label13.Caption = sn
Label4.Caption = (sp + sl + sn)
End Sub
Private Sub chargegrd2_c_Tm()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim m As Double
Dim k As Double
Dim e As Double
Dim se As Double
Dim P As Double
Dim sp As Double
Dim r As Double
Dim c As Double
Dim l As Double
Dim sl As Double
Dim f As Double
Dim sf As Double
Dim n As Double
Dim sn As Double
Dim s As Double
i = 1
e = 0
se = 0
P = 0
sp = 0
r = 0
c = 0
l = 0
sl = 0
sf = 0
sn = 0
s = 0
Call cont
grd2.Rows = pc.RecordCount + 3
Do While Not pc.EOF
If Combo5.Text = pc!cla Then
m = pc!moi
j = pc!nbr
e = 0
P = 0
r = 0
f = 0
t = 0
grd2.Row = i
grd2.Col = 0
If pc!moi = "0" Then
grd2.Text = "—. "
f = pc!etu
sf = sf + f
Else
grd2.Text = pc!moi
End If
e = pc!etu
se = se + e
grd2.Col = 1
grd2.Text = pc!cla
grd2.Col = 2
grd2.Text = pc!etu
grd2.Col = 3
grd2.Text = pc!pro
P = pc!pro
sp = sp + P
r = (e - P)
grd2.Col = 4
grd2.Text = r
If j > 0 Then
c = Label19.Caption
Else
c = 100
End If
grd2.Col = 5
grd2.Text = c
l = (r * c / 100)
MyNumber = Round(l, 0)
l = MyNumber
sl = sl + l
grd2.Col = 6
grd2.Text = l
c = (100 - c)
n = (r * c / 100)
MyNumber = Round(n, 0)
n = MyNumber
sn = sn + n
grd2.Col = 7
grd2.Text = c
grd2.Col = 8
grd2.Text = n
i = i + 1
End If
pc.MoveNext
Loop
grd2.Rows = i
Label20.Caption = (se - sf)
Label5.Caption = sf
Label10.Caption = sp
Label2.Caption = sl
Label13.Caption = sn
Label4.Caption = (sp + sl + sn)
End Sub
Private Sub chargegrd2_Tc_d()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim m As Double
Dim k As Double
Dim e As Double
Dim se As Double
Dim P As Double
Dim sp As Double
Dim r As Double
Dim c As Double
Dim l As Double
Dim sl As Double
Dim f As Double
Dim sf As Double
Dim n As Double
Dim sn As Double
Dim s As Double
i = 1
e = 0
se = 0
P = 0
sp = 0
r = 0
c = 0
l = 0
sl = 0
sf = 0
sn = 0
s = DT1.Month
Call cont
grd2.Rows = pc.RecordCount + 3
Do While Not pc.EOF
m = pc!moi
j = pc!nbr
If m = s Or m = 0 And Label8.Caption = s Then
grd2.Row = i
grd2.Col = 0
If pc!moi = "0" Then
grd2.Text = "—. "
Else
grd2.Text = pc!moi
End If
grd2.Col = 1
grd2.Text = pc!cla
grd2.Col = 2
grd2.Text = "0"
grd2.Col = 3
grd2.Text = "0"
grd2.Col = 4
grd2.Text = "0"
If j > 0 Then
c = Label19.Caption
Else
c = 100
End If
grd2.Col = 5
grd2.Text = c
grd2.Col = 6
grd2.Text = "0"
c = (100 - c)
grd2.Col = 7
grd2.Text = c
grd2.Col = 8
grd2.Text = "0"
i = i + 1
End If
pc.MoveNext
Loop
grd2.Rows = i
End Sub
Private Sub Option1_Click()
On Error Resume Next
Call grd2_clear
DT1.Enabled = True
Combo7.Enabled = False
Combo1.Enabled = False
DT2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = False
End Sub

Private Sub Option2_Click()
On Error Resume Next
Call grd2_clear
DT1.Enabled = False
Combo7.Enabled = True
Combo1.Enabled = False
DT2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = False

End Sub

Private Sub Option3_Click()
Call grd2_clear
DT1.Enabled = False
Combo7.Enabled = False
Combo1.Enabled = False
DT2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = False

End Sub

Private Sub Option4_Click()
On Error Resume Next
Call grd2_clear
DT1.Enabled = False
Combo7.Enabled = False
Combo1.Enabled = True
DT2.Enabled = True
Combo3.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = False

End Sub

Private Sub Option5_Click()
On Error Resume Next
Call grd2_clear
DT1.Enabled = False
Combo7.Enabled = False
Combo1.Enabled = False
DT2.Enabled = False
Combo3.Enabled = True
Combo4.Enabled = True
Combo5.Enabled = False

End Sub

Private Sub Option6_Click()
On Error Resume Next
Call grd2_clear
DT1.Enabled = False
Combo7.Enabled = False
Combo1.Enabled = False
DT2.Enabled = False
Combo3.Enabled = False
Combo4.Enabled = False
Combo5.Enabled = True

End Sub
Private Sub grd2_clear()
On Error Resume Next
grd2.Clear
grd2.Cols = 9
grd2.Rows = 1
grd2.ColWidth(0) = 600
grd2.ColWidth(1) = 1000
grd2.ColWidth(2) = 1500
grd2.ColWidth(3) = 1500
grd2.ColWidth(4) = 1500
grd2.ColWidth(5) = 1500
grd2.ColWidth(6) = 1500
grd2.ColWidth(7) = 1500
grd2.ColWidth(8) = 1500
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
grd2.ColAlignment(7) = 1
grd2.ColAlignment(8) = 1
grd2.Row = 0
grd2.Col = 0
grd2.Text = "«·‘Â—"
grd2.Col = 1
grd2.Text = "«·ﬁ”„"
grd2.Col = 2
grd2.Text = "«·—”Ê„"
grd2.Col = 3
grd2.Text = "„” Õﬁ«  √”« –… ”"
grd2.Col = 4
grd2.Text = "«·»«ﬁÌ"
grd2.Col = 5
grd2.Text = "‰”»… «·„ƒ””…"
grd2.Col = 6
grd2.Text = "‰’Ì» «·„ƒ””…"
grd2.Col = 7
grd2.Text = "‰”»… «·√”« –…"
grd2.Col = 8
grd2.Text = "‰’Ì» «·√”« –…"

End Sub
Private Sub recettes()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim t As Double
Dim b As Double
Dim dat1 As Date
Dim dat2 As Date
Dim cl1 As String
Dim cl2 As String
Dim m1 As String
Dim m2 As String
dat1 = DT1.Value
Call cont
Do While Not ct.EOF
dat2 = ct!dat
cl1 = ct!cla
m1 = ct!moi
b = ct!pay
If dat1 = dat2 Then
n = grd2.Rows
For i = 1 To n - 1
grd2.Row = i
grd2.Col = 0
m2 = grd2.Text
grd2.Col = 1
cl2 = grd2.Text
If cl1 = cl2 And m2 <> "—. " And m1 <> "0" Then
grd2.Row = i
grd2.Col = 2
t = grd2.Text
t = (t + b)
grd2.Row = i
grd2.Col = 2
grd2.Text = t
i = n
End If
Next i
End If
ct.MoveNext
Loop

End Sub

