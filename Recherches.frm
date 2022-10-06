VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Recherches 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   ClientHeight    =   9645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "»ÕÀ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   3615
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      IMEMode         =   3  'DISABLE
      Left            =   240
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   1635
      Width           =   1575
   End
   Begin VB.TextBox Text9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      IMEMode         =   3  'DISABLE
      Left            =   2640
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   1635
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      IMEMode         =   3  'DISABLE
      Left            =   5400
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1635
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      IMEMode         =   3  'DISABLE
      Left            =   7680
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   1635
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      IMEMode         =   3  'DISABLE
      Left            =   9840
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1635
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      IMEMode         =   3  'DISABLE
      Left            =   240
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1150
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      IMEMode         =   3  'DISABLE
      Left            =   5400
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1150
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      IMEMode         =   3  'DISABLE
      Left            =   7680
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   1150
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      IMEMode         =   3  'DISABLE
      Left            =   9840
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1150
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      IMEMode         =   3  'DISABLE
      Left            =   2640
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1150
      Width           =   1575
   End
   Begin VB.OptionButton Option5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "⁄‰  ·„Ì–"
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
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "⁄‰ ÊﬂÌ·"
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
      Left            =   4680
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "⁄‰ √” «–"
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
      Left            =   5880
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "⁄‰ ‘—Ìﬂ"
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
      Left            =   8400
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "⁄‰ „ÊŸ›"
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
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   6855
      Left            =   240
      TabIndex        =   27
      Top             =   2640
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   12091
      _Version        =   393216
      BackColor       =   32768
      ForeColor       =   0
      BackColorFixed  =   32768
      BackColorBkg    =   32768
      RightToLeft     =   -1  'True
      FillStyle       =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÊŸÌ›…"
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
      Left            =   1920
      TabIndex        =   25
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·‰œ«¡"
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
      Left            =   4080
      TabIndex        =   24
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·Ã‰”"
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
      Left            =   6360
      TabIndex        =   23
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label7 
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
      Left            =   9120
      TabIndex        =   22
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„«œ… «· œ—Ì”"
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
      TabIndex        =   21
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   2
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   12615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄‰Ê«‰"
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
      Left            =   1920
      TabIndex        =   20
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„ «·Êÿ‰Ì"
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
      Left            =   4080
      TabIndex        =   19
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·«”„"
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
      TabIndex        =   18
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„ «· ”·”·Ì"
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
      Left            =   11280
      TabIndex        =   17
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·Â« ›"
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
      Left            =   6360
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   12615
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»ÕÀ ⁄«„"
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
      TabIndex        =   7
      Top             =   0
      Width           =   12615
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   0
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   6375
   End
End
Attribute VB_Name = "Recherches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command2_Click()
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
Text6.Text = Trim(Text6.Text)
Text7.Text = Trim(Text7.Text)
Text8.Text = Trim(Text8.Text)
Text9.Text = Trim(Text9.Text)
Text10.Text = Trim(Text10.Text)
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text6.Text = "" And Text7.Text = "" And Text8.Text = "" And Text9.Text = "" And Text10.Text = "" Then
MsgBox "Ì—ÃÏ «œŒ«· „Õœœ«  «·»ÕÀ", vbCritical
Text1.SetFocus
Exit Sub
End If
grd1.Visible = False
Call grd1_clear
If Option1.Value = True Then
Call chargegrd1_1
End If
If Option2.Value = True Then
Call chargegrd1_2
End If
If Option3.Value = True Then
Call chargegrd1_3
End If
If Option4.Value = True Then
Call chargegrd1_4
End If
If Option5.Value = True Then
Call chargegrd1_5
End If
grd1.Visible = True
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Option1_Click()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True
Label5.Visible = True
Text5.Visible = True
Label6.Visible = False
Text6.Visible = False
Label7.Visible = False
Text7.Visible = False
Label8.Visible = False
Text8.Visible = False
Label9.Visible = False
Text9.Visible = False
Label10.Visible = False
Text10.Visible = False
End Sub

Private Sub Option2_Click()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True
Label5.Visible = True
Text5.Visible = True
Label6.Visible = False
Text6.Visible = False
Label7.Visible = False
Text7.Visible = False
Label8.Visible = False
Text8.Visible = False
Label9.Visible = False
Text9.Visible = False
Label10.Visible = True
Text10.Visible = True

End Sub

Private Sub Option3_Click()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True
Label5.Visible = True
Text5.Visible = True
Label6.Visible = True
Text6.Visible = True
Label7.Visible = False
Text7.Visible = False
Label8.Visible = False
Text8.Visible = False
Label9.Visible = False
Text9.Visible = False
Label10.Visible = False
Text10.Visible = False

End Sub

Private Sub Option4_Click()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True
Label5.Visible = True
Text5.Visible = True
Label6.Visible = False
Text6.Visible = False
Label7.Visible = False
Text7.Visible = False
Label8.Visible = False
Text8.Visible = False
Label9.Visible = False
Text9.Visible = False
Label10.Visible = True
Text10.Visible = True

End Sub

Private Sub Option5_Click()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True
Label5.Visible = False
Text5.Visible = False
Label6.Visible = False
Text6.Visible = False
Label7.Visible = True
Text7.Visible = True
Label8.Visible = True
Text8.Visible = True
Label9.Visible = True
Text9.Visible = True
Label10.Visible = False
Text10.Visible = False

End Sub
Private Sub chargegrd1_1()
Dim i As Double
i = 1
Call cont
'cr.Filter = "[tel]" & "Like '*" & Text3 & "*'" 'entre
grd1.Rows = pr.RecordCount + 3
Do While Not pr.EOF
If pr!act = "1" Then
If pr!sri Like "*" & Text1 & "*" And pr!nom Like "*" & Text2 & "*" And pr!tel Like "*" & Text3 & "*" And pr!nni Like "*" & Text4 & "*" And pr!adr Like "*" & Text5 & "*" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = pr!aut
grd1.Col = 1
grd1.Text = pr!sri
grd1.Col = 2
grd1.Text = pr!nom
grd1.Col = 3
grd1.Text = pr!tel
grd1.Col = 4
grd1.Text = pr!nni
grd1.Col = 5
grd1.Text = pr!adr
i = i + 1
End If
End If
pr.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
End Sub
Private Sub chargegrd1_2()
Dim i As Double
i = 1
Call cont
'cr.Filter = "[tel]" & "Like '*" & Text3 & "*'" 'entre
grd1.Rows = fc.RecordCount + 3
Do While Not fc.EOF
If fc!act = "1" Then
If fc!sri Like "*" & Text1 & "*" And fc!nom Like "*" & Text2 & "*" And fc!tel Like "*" & Text3 & "*" And fc!nni Like "*" & Text4 & "*" And fc!adr Like "*" & Text5 & "*" And fc!fon Like "*" & Text10 & "*" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = fc!aut
grd1.Col = 1
grd1.Text = fc!sri
grd1.Col = 2
grd1.Text = fc!nom
grd1.Col = 3
grd1.Text = fc!tel
grd1.Col = 4
grd1.Text = fc!nni
grd1.Col = 5
grd1.Text = fc!adr
grd1.Col = 10
grd1.Text = fc!fon
i = i + 1
End If
End If
fc.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
End Sub
Private Sub chargegrd1_3()
Dim i As Double
i = 1
Call cont
'cr.Filter = "[tel]" & "Like '*" & Text3 & "*'" 'entre
grd1.Rows = pf.RecordCount + 3
Do While Not pf.EOF
If pf!act = "1" Then
If pf!sri Like "*" & Text1 & "*" And pf!nom Like "*" & Text2 & "*" And pf!tel Like "*" & Text3 & "*" And pf!nni Like "*" & Text4 & "*" And pf!adr Like "*" & Text5 & "*" And pf!mat Like "*" & Text6 & "*" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = pf!aut
grd1.Col = 1
grd1.Text = pf!sri
grd1.Col = 2
grd1.Text = pf!nom
grd1.Col = 3
grd1.Text = pf!tel
grd1.Col = 4
grd1.Text = pf!nni
grd1.Col = 5
grd1.Text = pf!adr
grd1.Col = 6
grd1.Text = pf!mat
i = i + 1
End If
End If
pf.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
End Sub
Private Sub chargegrd1_4()
Dim i As Double
i = 1
Call cont
'cr.Filter = "[tel]" & "Like '*" & Text3 & "*'" 'entre
grd1.Rows = cr.RecordCount + 3
Do While Not cr.EOF
If cr!act = "1" Then
If cr!sri Like "*" & Text1 & "*" And cr!nom Like "*" & Text2 & "*" And cr!tel Like "*" & Text3 & "*" And cr!nni Like "*" & Text4 & "*" And cr!adr Like "*" & Text5 & "*" And cr!fon Like "*" & Text10 & "*" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = cr!aut
grd1.Col = 1
grd1.Text = cr!sri
grd1.Col = 2
grd1.Text = cr!nom
grd1.Col = 3
grd1.Text = cr!tel
grd1.Col = 4
grd1.Text = cr!nni
grd1.Col = 5
grd1.Text = cr!adr
grd1.Col = 10
grd1.Text = cr!fon
i = i + 1
End If
End If
cr.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
End Sub
Private Sub chargegrd1_5()
Dim i As Double
i = 1
Call cont
'cr.Filter = "[tel]" & "Like '*" & Text3 & "*'" 'entre
grd1.Rows = et.RecordCount + 3
Do While Not et.EOF
If et!act = "1" Then
If et!sri Like "*" & Text1 & "*" And et!nom Like "*" & Text2 & "*" And et!tel Like "*" & Text3 & "*" And et!nni Like "*" & Text4 & "*" And et!cla Like "*" & Text7 & "*" And et!sex Like "*" & Text8 & "*" And et!num Like "*" & Text9 & "*" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = et!aut
grd1.Col = 1
grd1.Text = et!sri
grd1.Col = 2
grd1.Text = et!nom
grd1.Col = 3
grd1.Text = et!tel
grd1.Col = 4
grd1.Text = et!nni
grd1.Col = 7
grd1.Text = et!cla
grd1.Col = 8
grd1.Text = et!sex
grd1.Col = 9
grd1.Text = et!num
i = i + 1
End If
End If
et.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
End Sub
Private Sub grd1_clear()
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
If Option1.Value = True Then
grd1.ColWidth(1) = 2300
grd1.ColWidth(2) = 3000
grd1.ColWidth(3) = 2100
grd1.ColWidth(4) = 2300
grd1.ColWidth(5) = 2400
grd1.ColWidth(6) = 0
grd1.ColWidth(7) = 0
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 0
grd1.ColWidth(10) = 0
End If
If Option2.Value = True Then
grd1.ColWidth(1) = 2000
grd1.ColWidth(2) = 3000
grd1.ColWidth(3) = 2000
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 1900
grd1.ColWidth(6) = 0
grd1.ColWidth(7) = 0
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 0
grd1.ColWidth(10) = 1200
End If
If Option3.Value = True Then
grd1.ColWidth(1) = 2000
grd1.ColWidth(2) = 3000
grd1.ColWidth(3) = 2000
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 1900
grd1.ColWidth(6) = 1200
grd1.ColWidth(7) = 0
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 0
grd1.ColWidth(10) = 0
End If
If Option4.Value = True Then
grd1.ColWidth(1) = 2000
grd1.ColWidth(2) = 3000
grd1.ColWidth(3) = 2000
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 1900
grd1.ColWidth(6) = 0
grd1.ColWidth(7) = 0
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 0
grd1.ColWidth(10) = 1200
End If
If Option5.Value = True Then
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 3800
grd1.ColWidth(3) = 2000
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 0
grd1.ColWidth(6) = 0
grd1.ColWidth(7) = 1000
grd1.ColWidth(8) = 1000
grd1.ColWidth(9) = 1100
grd1.ColWidth(10) = 0
End If
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(9) = 1
grd1.ColAlignment(10) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«·«”„"
grd1.Col = 3
grd1.Text = "«·Â« ›"
grd1.Col = 4
grd1.Text = "—. «·Êÿ‰Ì"
grd1.Col = 5
grd1.Text = "«·⁄‰Ê«‰"
grd1.Col = 6
grd1.Text = "«·„«œ…"
grd1.Col = 7
grd1.Text = "«·ﬁ”„"
grd1.Col = 8
grd1.Text = "«·Ã‰”"
grd1.Col = 9
grd1.Text = "—.«·‰œ«¡"
grd1.Col = 10
grd1.Text = "«·ÊŸÌ›…"

End Sub

Private Sub Text1_Change()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True

End Sub

Private Sub Text1_Click()
Text1_Change
End Sub

Private Sub Text10_Change()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True

End Sub

Private Sub Text10_Click()
Text10_Change
End Sub

Private Sub Text2_Change()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True

End Sub

Private Sub Text2_Click()
Text2_Change
End Sub

Private Sub Text3_Change()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True

End Sub

Private Sub Text3_Click()
Text3_Change
End Sub

Private Sub Text4_Change()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True

End Sub

Private Sub Text4_Click()
Text4_Change
End Sub

Private Sub Text5_Change()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True

End Sub

Private Sub Text5_Click()
Text5_Change
End Sub

Private Sub Text6_Change()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True

End Sub

Private Sub Text6_Click()
Text6_Change
End Sub

Private Sub Text7_Change()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True

End Sub

Private Sub Text7_Click()
Text7_Change
End Sub

Private Sub Text8_Change()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True

End Sub

Private Sub Text8_Click()
Text8_Change
End Sub

Private Sub Text9_Change()
grd1.Visible = False
Call grd1_clear
grd1.Visible = True

End Sub

Private Sub Text9_Click()
Text9_Change
End Sub
