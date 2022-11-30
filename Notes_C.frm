VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Notes_C 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   3720
      ScaleHeight     =   2595
      ScaleWidth      =   2475
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   480
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "ÇáÑÊÈÉ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   360
         TabIndex        =   16
         Text            =   "Text6"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   360
         TabIndex        =   15
         Text            =   "Text5"
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command8 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ÍÝÙ ÇáÈíÇäÇÊ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11040
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ÚÑÖ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11280
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
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
      ItemData        =   "Notes_C.frx":0000
      Left            =   11040
      List            =   "Notes_C.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "Notes_C.frx":001B
      Left            =   11040
      List            =   "Notes_C.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "Notes_C.frx":0044
      Left            =   11040
      List            =   "Notes_C.frx":004E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3240
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   7935
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   13996
      _Version        =   393216
      Rows            =   1
      BackColor       =   32768
      BackColorFixed  =   32768
      BackColorBkg    =   32768
      Enabled         =   -1  'True
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÊÓÌíá äÊÇÆÌ ÞÓã ÞÓã"
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
      TabIndex        =   11
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáãÓÊæì"
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
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáÞÓã"
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
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáãÇÏÉ"
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
      TabIndex        =   8
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   8175
      Index           =   0
      Left            =   10920
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      TabIndex        =   7
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáÖÇÑÈ"
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
      Left            =   11400
      TabIndex        =   5
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ãÚÏá ÇáãÇÏÉ"
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
      TabIndex        =   4
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   8175
      Index           =   1
      Left            =   120
      Top             =   840
      Width           =   10695
   End
End
Attribute VB_Name = "Notes_C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Change()
On Error Resume Next
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
Call chargcombo3
Combo3.BackColor = &H8080FF
If Combo2.Text = "ÇÈÊÏÇÆí" Then
Call chargegrd2_tete_pr
Else
Call chargegrd2_tete
End If
'grd1.Visible = False
'Call chargegrd1_clear
'grd1.Visible = True
Else
Combo1.BackColor = &H8080FF
End If

End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Combo2_Change()
On Error Resume Next
If Len(Combo2.Text) > 0 Then
Combo2.BackColor = &HC000&
Call chargcombo1
Combo1.BackColor = &H8080FF
If Combo2.Text = "ÇÈÊÏÇÆí" Then
Call chargegrd2_tete_pr
Else
Call chargegrd2_tete
End If
Else
Combo2.BackColor = &H8080FF
End If

End Sub

Private Sub Combo2_Click()
On Error Resume Next
Combo2_Change
End Sub

Private Sub Combo3_Change()
On Error Resume Next
If Len(Combo3.Text) > 0 Then
Combo3.BackColor = &HC000&
Else
Combo3.BackColor = &H8080FF
End If
End Sub

Private Sub Combo3_Click()
On Error Resume Next
Combo3_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
grd2.Visible = False
Call calcule_moyenne_lc
grd2.Visible = True
End Sub

Private Sub Command12_Click()
On Error Resume Next
Command12.Enabled = False
If Combo2.Text = "ÇÈÊÏÇÆí" Then
Call rangs_P
Else
Call rangs
End If
Command12.Enabled = True

End Sub

Private Sub Command5_Click()
On Error Resume Next
Call cont
Do While Not mt.EOF
If mt!mat = Combo3.Text Then
Label3.Caption = mt!cof
Label4.Caption = mt!moy
grd2.Visible = False
If Combo2.Text = "ÇÈÊÏÇÆí" Then
Call chargegrd2_tete_pr
Else
Call chargegrd2_tete
End If
Call chargegrd2
Call coff_dv_ex
Call chargegrd2_notes
Call calcule_moyenne_lc
grd2.Visible = True
Exit Sub
End If
mt.MoveNext
Loop

End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
If Combo2.Text = "" Or Combo1.Text = "" Or Combo3.Text = "" Then
MsgBox "ÇáÑÌÇÁ ãáÁ ÌãíÚ ÇáÍÞæá ÇáãáæäÉ ÈÇááæä ÇáÃÍãÑ", vbCritical + arabic
Exit Sub
End If
n = grd2.Rows
If n = 1 Then
MsgBox "ÛíÑ ããßä.. áÇ ÊæÌÏ ÈíÇäÇÊ , íÌÈ ÅÏÎÇá ÈíÇäÇÊ ÃæáÇ", vbCritical
Exit Sub
End If
Call cont
Do While Not nt.EOF
If nt!cla = Combo1.Text And nt!mat = Combo3.Text Then
nt.Delete
End If
nt.MoveNext
Loop
grd2.Visible = False
Call calcule_moyenne_lc
Call cont
For i = 1 To n - 1
nt.AddNew
grd2.Row = i
grd2.Col = 19
nt!sri = grd2.Text
nt!niv = Combo2.Text
nt!cla = Combo1.Text
nt!mat = Combo3.Text
grd2.Row = i
grd2.Col = 20
nt!num = grd2.Text
grd2.Row = i
grd2.Col = 0
nt!nom = grd2.Text
grd2.Col = 1
nt!cmt = grd2.Text
grd2.Col = 2
nt!mmt = grd2.Text
grd2.Col = 3
nt!dv1 = grd2.Text
grd2.Col = 4
nt!dv2 = grd2.Text
grd2.Col = 5
nt!dv3 = grd2.Text
grd2.Col = 6
nt!dv4 = grd2.Text
grd2.Col = 7
nt!dv5 = grd2.Text
grd2.Col = 8
nt!dv6 = grd2.Text
grd2.Col = 9
nt!mdv = grd2.Text
grd2.Col = 10
nt!cdv = grd2.Text
grd2.Col = 11
nt!ex1 = grd2.Text
grd2.Col = 12
nt!cx1 = grd2.Text
grd2.Col = 13
nt!ex2 = grd2.Text
grd2.Col = 14
nt!cx2 = grd2.Text
grd2.Col = 15
nt!ex3 = grd2.Text
grd2.Col = 16
nt!cx3 = grd2.Text
grd2.Col = 17
nt!mym = grd2.Text
grd2.Col = 18
nt!tot = grd2.Text
nt!moy = "0"
nt!tto = ""
nt!tcf = ""
nt!men = ""
nt!ran = ""
nt!dat = Date
nt!Abs = ""
nt!obs = ""
nt!tt1 = ""
nt!mo1 = ""
nt!tt2 = ""
nt!mo2 = ""
nt!tt3 = ""
nt!mo3 = ""
nt!tt4 = ""
nt!mo4 = ""
nt.Update
Next i
grd2.Visible = True
MsgBox "Êã ÍÝÙ ÇáÈíÇÊÇÊ", vbInformation

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 480
Call chargegrd2_tete
'Call chargegrd1_clear
'Text1.SetFocus
End Sub
Private Sub chargcombo1()
On Error Resume Next
Combo1.Clear
Call cont
Do While Not cl.EOF
If Combo2.Text = cl!niv And cl!act = "1" Then
Combo1.AddItem cl!cla
End If
cl.MoveNext
Loop
End Sub
Private Sub chargcombo3()
On Error Resume Next
Combo3.Clear
Call cont
Do While Not mt.EOF
If Combo1.Text = mt!cla Then
Combo3.AddItem mt!mat
End If
mt.MoveNext
Loop
End Sub
Private Sub chargegrd2_tete()
On Error Resume Next
Dim i As Double
Dim j As Double
grd2.Clear
grd2.Cols = 21
grd2.Rows = 1
grd2.ColWidth(0) = 2200
grd2.ColWidth(1) = 0
grd2.ColWidth(2) = 0
grd2.ColWidth(3) = 650
grd2.ColWidth(4) = 650
grd2.ColWidth(5) = 650
grd2.ColWidth(6) = 650
grd2.ColWidth(7) = 650
grd2.ColWidth(8) = 650
grd2.ColWidth(9) = 650
grd2.ColWidth(10) = 0
grd2.ColWidth(11) = 650
grd2.ColWidth(12) = 0
grd2.ColWidth(13) = 650
grd2.ColWidth(14) = 0
grd2.ColWidth(15) = 650
grd2.ColWidth(16) = 0
grd2.ColWidth(17) = 800
grd2.ColWidth(18) = 800
grd2.ColWidth(19) = 0
grd2.ColWidth(20) = 0
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
grd2.ColAlignment(7) = 1
grd2.ColAlignment(8) = 1
grd2.ColAlignment(9) = 1
grd2.ColAlignment(10) = 1
grd2.ColAlignment(11) = 1
grd2.ColAlignment(12) = 1
grd2.ColAlignment(13) = 1
grd2.ColAlignment(14) = 1
grd2.ColAlignment(15) = 1
grd2.ColAlignment(16) = 1
grd2.ColAlignment(17) = 1
grd2.ColAlignment(18) = 1
grd2.ColAlignment(19) = 1
grd2.Row = 0
grd2.Col = 0
grd2.Text = "ÇÓã ÇáÊáãíÐ"
grd2.Col = 1
grd2.Text = "ÖÜ"
grd2.Col = 2
grd2.Text = "ã . ã"
grd2.Col = 3
grd2.Text = "ÇÎÜ 1"
grd2.Col = 4
grd2.Text = "ÇÎÜ 2"
grd2.Col = 5
grd2.Text = "ÇÎÜ 3"
grd2.Col = 6
grd2.Text = "ÇÎÜ 4"
grd2.Col = 7
grd2.Text = "ÇÎÜ 5"
grd2.Col = 8
grd2.Text = "ÇÎÜ 6"
grd2.Col = 9
grd2.Text = "ãÚÏá ÇÎÜ"
grd2.Col = 10
grd2.Text = "ÖÜ"
grd2.Col = 11
grd2.Text = "ÇãÊÍÜ 1"
grd2.Col = 12
grd2.Text = "ÖÜ"
grd2.Col = 13
grd2.Text = "ÇãÊÍÜ 2"
grd2.Col = 14
grd2.Text = "ÖÜ"
grd2.Col = 15
grd2.Text = "ÇãÊÍÜ 3"
grd2.Col = 16
grd2.Text = "ÖÜ"
grd2.Col = 17
grd2.Text = "ÇáãÚÏá"
grd2.Col = 18
grd2.Text = "ÇáãÌãæÚ"
End Sub
Private Sub chargegrd2_tete_pr()
On Error Resume Next
Dim i As Double
Dim j As Double
grd2.Clear
grd2.Cols = 21
grd2.Rows = 1
grd2.ColWidth(0) = 4250
grd2.ColWidth(1) = 0
grd2.ColWidth(2) = 0
grd2.ColWidth(3) = 0
grd2.ColWidth(4) = 0
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 0
grd2.ColWidth(7) = 0
grd2.ColWidth(8) = 0
grd2.ColWidth(9) = 0
grd2.ColWidth(10) = 0
grd2.ColWidth(11) = 1200
grd2.ColWidth(12) = 0
grd2.ColWidth(13) = 1200
grd2.ColWidth(14) = 0
grd2.ColWidth(15) = 1200
grd2.ColWidth(16) = 0
grd2.ColWidth(17) = 1200
grd2.ColWidth(18) = 1200
grd2.ColWidth(19) = 0
grd2.ColWidth(20) = 0
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
grd2.ColAlignment(7) = 1
grd2.ColAlignment(8) = 1
grd2.ColAlignment(9) = 1
grd2.ColAlignment(10) = 1
grd2.ColAlignment(11) = 1
grd2.ColAlignment(12) = 1
grd2.ColAlignment(13) = 1
grd2.ColAlignment(14) = 1
grd2.ColAlignment(15) = 1
grd2.ColAlignment(16) = 1
grd2.ColAlignment(17) = 1
grd2.ColAlignment(18) = 1
grd2.ColAlignment(19) = 1
grd2.Row = 0
grd2.Col = 0
grd2.Text = "ÇÓã ÇáÊáãíÐ"
grd2.Col = 1
grd2.Text = "ÖÜ"
grd2.Col = 2
grd2.Text = "ã . ã"
grd2.Col = 3
grd2.Text = "ÇÎÜ 1"
grd2.Col = 4
grd2.Text = "ÇÎÜ 2"
grd2.Col = 5
grd2.Text = "ÇÎÜ 3"
grd2.Col = 6
grd2.Text = "ÇÎÜ 4"
grd2.Col = 7
grd2.Text = "ÇÎÜ 5"
grd2.Col = 8
grd2.Text = "ÇÎÜ 6"
grd2.Col = 9
grd2.Text = "ãÚÏá ÇÎÜ"
grd2.Col = 10
grd2.Text = "ÖÜ"
grd2.Col = 11
grd2.Text = "ÇãÊÍÜ 1"
grd2.Col = 12
grd2.Text = "ÖÜ"
grd2.Col = 13
grd2.Text = "ÇãÊÍÜ 2"
grd2.Col = 14
grd2.Text = "ÖÜ"
grd2.Col = 15
grd2.Text = "ÇãÊÍÜ 3"
grd2.Col = 16
grd2.Text = "ÖÜ"
grd2.Col = 17
grd2.Text = "ÇáãÚÏá"
grd2.Col = 18
grd2.Text = "ÇáãÌãæÚ"
End Sub
Private Sub chargegrd2()
On Error Resume Next
Dim i As Double
i = 1
Call cont
grd2.Rows = et.RecordCount + 4
Do While Not et.EOF
If et!cla = Combo1.Text Then
grd2.Row = i
grd2.Col = 0
grd2.Text = et!nom
grd2.Col = 1
grd2.Text = Label3.Caption
grd2.CellBackColor = &H808080
grd2.Col = 2
grd2.Text = Label4.Caption
grd2.CellBackColor = &H808080
grd2.Col = 19
grd2.Text = et!sri
grd2.Col = 20
grd2.Text = et!num
i = i + 1
End If
et.MoveNext
Loop
grd2.Rows = i
End Sub
Private Sub chargegrd2_notes()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim tx1 As String
Dim tx2 As String
n = grd2.Rows
Call cont
Do While Not nt.EOF
If nt!cla = Combo1.Text And nt!mat = Combo3.Text Then
tx1 = nt!sri
For i = 1 To n - 1
grd2.Row = i
grd2.Col = 19
tx2 = grd2.Text
If tx1 = tx2 Or Val(tx1) = Val(tx2) Then
grd2.Col = 3
grd2.Text = nt!dv1
grd2.Col = 4
grd2.Text = nt!dv2
grd2.Col = 5
grd2.Text = nt!dv3
grd2.Col = 6
grd2.Text = nt!dv4
grd2.Col = 7
grd2.Text = nt!dv5
grd2.Col = 8
grd2.Text = nt!dv6
grd2.Col = 9
grd2.Text = nt!mdv
grd2.CellBackColor = &H808080
grd2.Col = 11
grd2.Text = nt!ex1
grd2.Col = 13
grd2.Text = nt!ex2
grd2.Col = 15
grd2.Text = nt!ex3
grd2.Col = 17
grd2.Text = nt!mym
grd2.CellBackColor = &H808080
grd2.Col = 18
grd2.Text = nt!tot
grd2.CellBackColor = &H808080
End If
Next i
End If
nt.MoveNext
Loop
End Sub
Private Sub coff_dv_ex()
On Error Resume Next
Dim n As Double
Dim i As Double
Dim j As Double
Dim k As Double
Dim tx As String
Call cont
n = grd2.Rows
For i = 1 To n - 1
If Combo2.Text = "ÇÈÊÏÇÆí" Then
grd2.Row = i
grd2.Col = 12
grd2.Text = cf2!cof16
grd2.CellBackColor = &H80C0FF
grd2.Col = 14
grd2.Text = cf2!cof17
grd2.CellBackColor = &H80C0FF
grd2.Col = 16
grd2.Text = cf2!cof18
grd2.CellBackColor = &H80C0FF
Else
grd2.Row = i
grd2.Col = 10
grd2.Text = cf2!cof0
grd2.CellBackColor = &H80C0FF
grd2.Col = 12
grd2.Text = cf2!cof1
grd2.CellBackColor = &H80C0FF
grd2.Col = 14
grd2.Text = cf2!cof2
grd2.CellBackColor = &H80C0FF
grd2.Col = 16
grd2.Text = cf2!cof3
grd2.CellBackColor = &H80C0FF
End If
Next i
End Sub

Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
i = grd2.Row
j = grd2.Col
If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j = 7 Or j = 8 Or j = 11 Or j = 13 Or j = 15 Then
grd2.Row = i
grd2.Col = j
grd2.CellBackColor = &HFFFF&
End If

End Sub

Private Sub grd2_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Double
Dim j As Double
Dim n As Double
Dim a As Double
Dim b As Double
Dim k As Double
Dim P As Double
i = grd2.Row
j = grd2.Col
If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j = 7 Or j = 8 Or j = 11 Or j = 13 Or j = 15 Then
If KeyAscii = 8 Then
grd2.Row = i
grd2.Col = j
grd2.Text = ""
Exit Sub
End If
grd2.Row = i
grd2.Col = j
Text3.Text = grd2.Text
n = Len(Text3.Text)
If n > 4 Then
KeyAscii = 0
Exit Sub
End If
If n = 0 And KeyAscii = 46 Then
KeyAscii = 0
Exit Sub
End If
For k = 1 To n
vg = Mid$(Text3.Text, k, 1)
r = Asc(vg)
If r = 46 Then
P = k + 2
End If
If P > 2 And KeyAscii = 46 Then
KeyAscii = 0
End If
If k = P And KeyAscii <> 8 Then
KeyAscii = 0
End If
If k = P Then
k = n
End If
Next k
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
Exit Sub
End If
With grd2
        Select Case .Col
            Case 0, 3:
             .Text = .Text + Chr$(KeyAscii)
             Case 0, 4:
             .Text = .Text + Chr$(KeyAscii)
              Case 0, 5:
             .Text = .Text + Chr$(KeyAscii)
             Case 0, 6:
             .Text = .Text + Chr$(KeyAscii)
             Case 0, 7:
             .Text = .Text + Chr$(KeyAscii)
             Case 0, 8:
             .Text = .Text + Chr$(KeyAscii)
             Case 0, 11:
             .Text = .Text + Chr$(KeyAscii)
             Case 0, 13:
             .Text = .Text + Chr$(KeyAscii)
             Case 0, 15:
             .Text = .Text + Chr$(KeyAscii)
            Case Else:
        End Select
    End With
grd2.Row = i
grd2.Col = 2
b = grd2.Text
grd2.Row = i
grd2.Col = j
a = grd2.Text
If a > b Then
grd2.Row = i
grd2.Col = j
grd2.Text = ""
End If
'Call calcule_moyenne_lc
'grd2.Row = i
'grd2.Col = j
End If

End Sub

Public Sub calcule_moyenne_lc()
On Error Resume Next
Dim d1 As Double
Dim sd As Double
Dim nd As Double
Dim md As Double
Dim cd As Double
Dim c1 As Double
Dim c2 As Double
Dim c3 As Double
Dim cm As Double
Dim sc As Double
Dim mm As Double
Dim e1 As Double
Dim e2 As Double
Dim e3 As Double
Dim t As Double
Dim i As Double
Dim j As Double
Dim n As Double
Dim scm As Double
Dim st As Double
Dim moy As Double
Dim tx As String
Dim tx2 As String
n = grd2.Rows
scm = 0
st = 0
moy = 0
For i = 1 To n - 1
d1 = 0
nd = 0
sd = 0
sc = 0
cd = 0
cm = 0
c1 = 0
c2 = 0
c3 = 0
e1 = 0
e2 = 0
e3 = 0
md = 0
nd = 0
mm = 0
t = 0
sc = 0
    For j = 1 To 19
    ' cof mat
        If j = 1 Then
        grd2.Row = i
        grd2.Col = j
        cm = grd2.Text
        End If
        'not dev
        If j > 2 And j < 9 Then
        grd2.Row = i
        grd2.Col = j
        tx = grd2.Text
        If tx <> "" Then
        d1 = tx
        nd = nd + 1
        sd = sd + d1
        End If
        End If
        'moy dev
        If j = 9 Then
        If nd > 0 Then
        md = sd / nd
        MyNumber = Round(md, 2)
        md = MyNumber
        End If
        grd2.Row = i
        grd2.Col = j
        grd2.Text = md
        End If
        'cof dev
        If j = 10 Then
        If nd > 0 Then
        grd2.Row = i
        grd2.Col = j
        cd = grd2.Text
        sc = sc + cd
        End If
        End If
        'not ex1
        If j = 11 Then
        grd2.Row = i
        grd2.Col = j
        tx = grd2.Text
        If tx <> "" Then
        e1 = tx
        End If
        End If
        'cof ex1
        If j = 12 Then
        If tx <> "" Then
        grd2.Row = i
        grd2.Col = j
        c1 = grd2.Text
        sc = sc + c1
        End If
        End If
        'not ex2
        If j = 13 Then
        grd2.Row = i
        grd2.Col = j
        tx = grd2.Text
        If tx <> "" Then
        e2 = tx
        End If
        End If
        'cof ex2
        If j = 14 Then
        If tx <> "" Then
        grd2.Row = i
        grd2.Col = j
        c2 = grd2.Text
        sc = sc + c2
        End If
        End If
        'not ex3
        If j = 15 Then
        grd2.Row = i
        grd2.Col = j
        tx = grd2.Text
        If tx <> "" Then
        e3 = tx
        End If
        End If
        'cof ex3
        If j = 16 Then
        If tx <> "" Then
        grd2.Row = i
        grd2.Col = j
        c3 = grd2.Text
        sc = sc + c3
        End If
        End If
        'moy mat
        If j = 17 Then
        If sc > 0 Then
        mm = ((md * cd) + (e1 * c1) + (e2 * c2) + (e3 * c3)) / sc
        MyNumber = Round(mm, 2)
        mm = MyNumber
        scm = scm + cm
        End If
        grd2.Row = i
        grd2.Col = j
        grd2.Text = mm
        End If
        t = (mm * cm)
        MyNumber = Round(t, 2)
        t = MyNumber
        'tot mat
        If j = 18 Then
        grd2.Row = i
        grd2.Col = j
        grd2.Text = t
        End If
    Next j
     st = st + t
Next i
End Sub

Private Sub rangs()
On Error GoTo P
Dim j As Double
Dim i As Double
Dim n As Double
Dim sql
Call cont
n = nt.RecordCount
Do While Not nt.EOF
If Combo1.Text = nt!cla Then
nt!ran = "0"
nt.Update
End If
nt.MoveNext
Loop
j = 0
For i = 1 To n
Call cont
sql = "select max(moy) from Notes where ran ='0'"
If rg.State = adStateOpen Then rg.Close
rg.Open sql, co, adOpenKeyset, adLockOptimistic
Text5.Text = rg.Fields(0)
Do While Not nt.EOF
If nt!moy = Text5.Text And Combo1.Text = nt!cla Then
If Text5.Text <> Text6.Text Then
j = j + 1
End If
nt!ran = j
nt.Update
Text6.Text = Text5.Text
End If
nt.MoveNext
Loop
Next i
Exit Sub
P:
Exit Sub
End Sub
Private Sub rangs_P()
On Error GoTo P
Dim j As Double
Dim i As Double
Dim n As Double
Dim sql
Call cont
n = nt.RecordCount
Do While Not nt.EOF
If Combo1.Text = nt!cla Then
nt!ran = "0"
nt.Update
End If
nt.MoveNext
Loop
j = 0
For i = 1 To n
Call cont
sql = "select max(moy) from Notes where ran ='0'"
If rg.State = adStateOpen Then rg.Close
rg.Open sql, co, adOpenKeyset, adLockOptimistic
Text5.Text = rg.Fields(0)
Do While Not nt.EOF
If nt!moy = Text5.Text And Combo1.Text = nt!cla Then
If Text5.Text <> Text6.Text Then
j = j + 1
End If
nt!ran = j
nt.Update
Text6.Text = Text5.Text
End If
nt.MoveNext
Loop
Next i
Exit Sub
P:
Exit Sub
End Sub

