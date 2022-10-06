VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form login 
   BorderStyle     =   0  'None
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "login.frx":0000
   ScaleHeight     =   9645
   ScaleWidth      =   16155
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   0
      Picture         =   "login.frx":22CFD
      ScaleHeight     =   9585
      ScaleWidth      =   16425
      TabIndex        =   3
      Top             =   0
      Width           =   16455
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   7440
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   8040
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00004000&
         Height          =   345
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   7560
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "œŒÊ·"
         Default         =   -1  'True
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
         Left            =   8280
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   8520
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„” Œœ„"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   8520
         TabIndex        =   5
         Top             =   7560
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ·„… «·”—"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   8520
         TabIndex        =   4
         Top             =   8040
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         Height          =   1695
         Left            =   7200
         Shape           =   4  'Rounded Rectangle
         Top             =   7320
         Width           =   3735
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   480
      OleObjectBlob   =   "login.frx":458BC
      Top             =   2640
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
Text1.SetFocus
End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Command1_Click()
If Combo1.Text = "" Or Text1.Text = "" Then
MsgBox " ›÷·Ê« »«Œ Ì«— «”„ «·„” Œœ„ Ê«œŒ«· ﬂ·„… «·”—", vbCritical
Exit Sub
End If
Call cont
Do While Not ut.EOF
If Combo1.Text = ut!nom And Text1.Text = ut!mot Then
directions.Label2.Caption = Combo1.Text
Unload Me
directions.Show
'Utilisateurs.Show
Exit Sub
End If
'***********
ut.MoveNext
Loop
MsgBox "«”„ «·„” Œœ„ Ê ﬂ·„… «·”— €Ì— „ ÿ«»ﬁÌ‰", vbCritical
Text1.Text = ""
Text1.SetFocus

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Call chargcombo1
End Sub
Private Sub chargcombo1()
Combo1.Clear
Call cont
Do While Not ut.EOF
Combo1.AddItem ut!nom
ut.MoveNext
Loop

End Sub

