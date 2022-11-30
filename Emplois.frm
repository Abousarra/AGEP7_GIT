VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Emplois 
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
      Left            =   10800
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "”Õ»"
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
      TabIndex        =   19
      Top             =   9000
      UseMaskColor    =   -1  'True
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
      ItemData        =   "Emplois.frx":0000
      Left            =   10680
      List            =   "Emplois.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3600
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   1200
      ScaleHeight     =   3195
      ScaleWidth      =   5835
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   1800
         Top             =   120
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
      ItemData        =   "Emplois.frx":0004
      Left            =   10680
      List            =   "Emplois.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
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
      ItemData        =   "Emplois.frx":0008
      Left            =   10680
      List            =   "Emplois.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4680
      Width           =   975
   End
   Begin VB.ComboBox Combo4 
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
      ItemData        =   "Emplois.frx":000C
      Left            =   10680
      List            =   "Emplois.frx":0040
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin VB.ComboBox Combo5 
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
      ItemData        =   "Emplois.frx":00B4
      Left            =   10680
      List            =   "Emplois.frx":00C4
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ComboBox Combo6 
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
      ItemData        =   "Emplois.frx":00E6
      Left            =   10680
      List            =   "Emplois.frx":00E8
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   6240
      Width           =   1935
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   10680
      TabIndex        =   8
      Top             =   7680
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   7575
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   13361
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
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line5 
      X1              =   10560
      X2              =   12720
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ÿ«„ ”«⁄ Ì‰"
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
      TabIndex        =   18
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ÿ«„ ”«⁄…"
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
      TabIndex        =   17
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Line Line4 
      X1              =   10560
      X2              =   12720
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line3 
      X1              =   10560
      X2              =   12720
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line2 
      X1              =   10560
      X2              =   12720
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      X1              =   10560
      X2              =   12720
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„«œ…"
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
      Left            =   11760
      TabIndex        =   16
      Top             =   3240
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   8655
      Index           =   10
      Left            =   120
      Top             =   840
      Width           =   10335
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ãœ«Ê· «·“„‰"
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
      TabIndex        =   15
      Top             =   120
      Width           =   10335
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁ”„"
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
      TabIndex        =   14
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   8655
      Index           =   9
      Left            =   10560
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«· ÊﬁÌ "
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
      TabIndex        =   13
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„” ÊÏ"
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
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÌÊ„"
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
      TabIndex        =   11
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„·«ÕŸ…: ·Õ–› √Ì „«œ… „‰ ÃœÊ· «·“„‰ Ì—ÃÏ «·÷€ÿ ⁄·Ï «·„«œ… Ê √ﬂÌœ «·Õ–›"
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
      TabIndex        =   10
      Top             =   960
      Width           =   8295
   End
End
Attribute VB_Name = "Emplois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chargecombo3()
Combo3.Clear
Combo3.AddItem "8 - 9"
Combo3.AddItem "9 - 10"
Combo3.AddItem "10 - 11"
Combo3.AddItem "11 - 12"
Combo3.AddItem "12 - 13"
Combo3.AddItem "13 - 14"
Combo3.AddItem "14 - 15"
Combo3.AddItem "15 - 16"
Combo3.AddItem "16 - 17"
Combo3.AddItem "17 - 18"
Combo3.AddItem "18 - 19"
Combo3.AddItem "19 - 20"
Combo3.AddItem "20 - 21"
Combo3.AddItem "21 - 22"
Combo3.AddItem "22 - 23"
Combo3.AddItem "23 - 00"

End Sub
Private Sub chargecombo4()
On Error Resume Next
Combo4.Clear
Combo4.AddItem "8 - 10"
Combo4.AddItem "10 - 12"
Combo4.AddItem "12 - 14"
Combo4.AddItem "14 - 16"
Combo4.AddItem "16 - 18"
Combo4.AddItem "18 - 20"
Combo4.AddItem "20 - 22"
Combo4.AddItem "22 - 00"

End Sub
Private Sub chargecombo6()
On Error Resume Next
Combo6.Clear
Combo6.AddItem "«·«À‰Ì‰"
Combo6.AddItem "«·À·«À«¡"
Combo6.AddItem "«·√—»⁄«¡"
Combo6.AddItem "«·Œ„Ì”"
Combo6.AddItem "«·Ã„⁄…"
Combo6.AddItem "«·”» "
Combo6.AddItem "«·√Õœ"
End Sub

Private Sub Combo1_Change()
On Error Resume Next
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
Call chargecombo2
Call chargegrd1
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
On Error Resume Next
If Len(Combo2.Text) > 0 Then
Combo2.BackColor = &HC000&
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
Combo4.BackColor = &HC000&
Call chargecombo4
Else
Combo3.BackColor = &H8080FF
End If

End Sub

Private Sub Combo3_Click()
On Error Resume Next
Combo3_Change
End Sub

Private Sub Combo4_Change()
On Error Resume Next
If Len(Combo4.Text) > 0 Then
Combo4.BackColor = &HC000&
Combo3.BackColor = &HC000&
Call chargecombo3
Else
Combo4.BackColor = &H8080FF
End If

End Sub

Private Sub Combo4_Click()
On Error Resume Next
Combo4_Change
End Sub

Private Sub Combo5_Change()
On Error Resume Next
If Len(Combo5.Text) > 0 Then
Combo5.BackColor = &HC000&
Call chargecombo1
Call chargegrd1_clear
Else
Combo5.BackColor = &H8080FF
End If

End Sub

Private Sub Combo5_Click()
On Error Resume Next
Combo5_Change
End Sub

Private Sub Combo6_Change()
On Error Resume Next
If Len(Combo6.Text) > 0 Then
Combo6.BackColor = &HC000&
Else
Combo6.BackColor = &H8080FF
End If

End Sub

Private Sub Combo6_Click()
On Error Resume Next
Combo6_Change
End Sub

Private Sub Command3_Click()
On Error Resume Next
If Combo5.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or Combo6.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
Exit Sub
End If
If Combo3.Text = "" And Combo4.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
Exit Sub
End If
Call cont
Do While Not em.EOF
If Combo5.Text = em!niv And Combo1.Text = em!cla Then
If Combo3.Text = em!heu Or Combo4.Text = em!heu Then
If Combo6.Text = "«·«À‰Ì‰" Then
em!lun = Combo2.Text
ElseIf Combo6.Text = "«·À·«À«¡" Then
em!mar = Combo2.Text
ElseIf Combo6.Text = "«·√—»⁄«¡" Then
em!mer = Combo2.Text
ElseIf Combo6.Text = "«·Œ„Ì”" Then
em!jeu = Combo2.Text
ElseIf Combo6.Text = "«·Ã„⁄…" Then
em!ven = Combo2.Text
ElseIf Combo6.Text = "«·”» " Then
em!sam = Combo2.Text
ElseIf Combo6.Text = "«·√Õœ" Then
em!Dim = Combo2.Text
End If
em.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
End If
em.MoveNext
Loop
em.AddNew
em!niv = Combo5.Text
em!cla = Combo1.Text
If Combo3.Text <> "" Then
em!heu = Combo3.Text
Else
em!heu = Combo4.Text
End If
em!lun = ""
em!mar = ""
em!mer = ""
em!jeu = ""
em!ven = ""
em!sam = ""
em!Dim = ""
If Combo6.Text = "«·«À‰Ì‰" Then
em!lun = Combo2.Text
ElseIf Combo6.Text = "«·À·«À«¡" Then
em!mar = Combo2.Text
ElseIf Combo6.Text = "«·√—»⁄«¡" Then
em!mer = Combo2.Text
ElseIf Combo6.Text = "«·Œ„Ì”" Then
em!jeu = Combo2.Text
ElseIf Combo6.Text = "«·Ã„⁄…" Then
em!ven = Combo2.Text
ElseIf Combo6.Text = "«·”» " Then
em!sam = Combo2.Text
ElseIf Combo6.Text = "«·√Õœ" Then
em!Dim = Combo2.Text
End If
em.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0
Call chargecombo3
Call chargecombo4
Call chargecombo6
Call chargegrd1_clear
End Sub
Private Sub chargecombo1()
On Error Resume Next
Combo1.Clear
Call cont
Do While Not cl.EOF
If Combo5.Text = cl!niv And cl!act = "1" Then
Combo1.AddItem cl!cla
End If
cl.MoveNext
Loop
End Sub
Private Sub chargecombo2()
On Error Resume Next
Combo2.Clear
Call cont
Do While Not mt.EOF
If Combo1.Text = mt!cla Then
Combo2.AddItem mt!mat
End If
mt.MoveNext
Loop
End Sub

Private Sub grd1_Click()
On Error Resume Next
Dim h As String
Dim j As String
Dim r As Double
Dim c As Double
r = grd1.Row
c = grd1.Col
grd1.Row = r
grd1.Col = 0
h = grd1.Text
grd1.Row = 0
grd1.Col = c
j = grd1.Text
If r > 0 Then
If c > 0 Then
grd1.Row = r
grd1.Col = c
grd1.CellBackColor = 741
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–Â «·„«œ…ø", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Call cont
Do While Not em.EOF
If Combo5.Text = em!niv And Combo1.Text = em!cla Then
If h = em!heu Or h = em!heu Then
If j = "«·«À‰Ì‰" Then
em!lun = ""
ElseIf j = "«·À·«À«¡" Then
em!mar = ""
ElseIf j = "«·√—»⁄«¡" Then
em!mer = ""
ElseIf j = "«·Œ„Ì”" Then
em!jeu = ""
ElseIf j = "«·Ã„⁄…" Then
em!ven = ""
ElseIf j = "«·”» " Then
em!sam = ""
ElseIf j = "«·√Õœ" Then
em!Dim = ""
End If
em.Update
If em!lun = "" And em!mar = "" And em!mer = "" And em!jeu = "" And em!ven = "" And em!sam = "" And em!Dim = "" Then
em.Delete
End If
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
End If
em.MoveNext
Loop
End If
grd1.Row = r
grd1.Col = c
grd1.CellBackColor = &HFF8080
End If
End If
End Sub

Private Sub chargegrd1()
On Error Resume Next
Dim a As Double
Dim b As Double
b = 0
Dim j As Double
Dim i As Double
Dim P As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Visible = False
grd1.Cols = 8
grd1.Rows = 1
grd1.ColWidth(0) = 1100
grd1.ColWidth(1) = 1250
grd1.ColWidth(2) = 1250
grd1.ColWidth(3) = 1250
grd1.ColWidth(4) = 1250
grd1.ColWidth(5) = 1250
grd1.ColWidth(6) = 1250
grd1.ColWidth(7) = 1250
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.Row = 0
grd1.Col = 0
grd1.Text = "«· ÊﬁÌ "
grd1.Col = 1
grd1.Text = "«·«À‰Ì‰"
grd1.Col = 2
grd1.Text = "«·À·«À«¡"
grd1.Col = 3
grd1.Text = "«·√—»⁄«¡"
grd1.Col = 4
grd1.Text = "«·Œ„Ì”"
grd1.Col = 5
grd1.Text = "«·Ã„⁄…"
grd1.Col = 6
grd1.Text = "«·”» "
grd1.Col = 7
grd1.Text = "«·√Õœ"
i = 1
Call cont
grd1.Rows = em.RecordCount + 3
Do While Not em.EOF
If Combo5.Text = em!niv And Combo1.Text = em!cla Then
grd1.Row = i
grd1.Col = 0
grd1.Text = em!heu
grd1.Col = 1
grd1.Text = em!lun
grd1.Col = 2
grd1.Text = em!mar
grd1.Col = 3
grd1.Text = em!mer
grd1.Col = 4
grd1.Text = em!jeu
grd1.Col = 5
grd1.Text = em!ven
grd1.Col = 6
grd1.Text = em!sam
grd1.Col = 7
grd1.Text = em!Dim
i = i + 1
End If
em.MoveNext
Loop
grd1.Rows = i
grd1.Col = 0
grd1.Sort = 1
grd1.Visible = True
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation + arabic
grd1.Visible = False
grd1.Clear
grd1.Rows = 1
Call chargegrd1
grd1.Visible = True
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False
End If

End Sub
Private Sub chargegrd1_clear()
On Error Resume Next
Dim a As Double
Dim b As Double
b = 0
Dim j As Double
Dim i As Double
Dim P As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Cols = 8
grd1.Rows = 1
grd1.ColWidth(0) = 1100
grd1.ColWidth(1) = 1250
grd1.ColWidth(2) = 1250
grd1.ColWidth(3) = 1250
grd1.ColWidth(4) = 1250
grd1.ColWidth(5) = 1250
grd1.ColWidth(6) = 1250
grd1.ColWidth(7) = 1250
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.Row = 0
grd1.Col = 0
grd1.Text = "«· ÊﬁÌ "
grd1.Col = 1
grd1.Text = "«·«À‰Ì‰"
grd1.Col = 2
grd1.Text = "«·À·«À«¡"
grd1.Col = 3
grd1.Text = "«·√—»⁄«¡"
grd1.Col = 4
grd1.Text = "«·Œ„Ì”"
grd1.Col = 5
grd1.Text = "«·Ã„⁄…"
grd1.Col = 6
grd1.Text = "«·”» "
grd1.Col = 7
grd1.Text = "«·√Õœ"
End Sub

