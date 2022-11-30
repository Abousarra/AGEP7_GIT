VERSION 5.00
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Etablissement 
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
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
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
      Left            =   6000
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "«—›«ﬁ ’Ê—… «·—√”Ì… ··’›Õ«  «·ﬁ«∆„…"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8040
      MaskColor       =   &H00FF0000&
      Picture         =   "Etablissement.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7800
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "«—›«ﬁ ’Ê—… «·—√”Ì… ··’›Õ«  «·√›ﬁÌ…"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10200
      MaskColor       =   &H00FF0000&
      Picture         =   "Etablissement.frx":035F
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
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
      Height          =   495
      Left            =   8280
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.TextBox Text1 
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
      Height          =   555
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1080
      Width           =   5655
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
      Height          =   345
      Left            =   6840
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
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
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
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
      Left            =   2520
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      ItemData        =   "Etablissement.frx":06BE
      Left            =   2520
      List            =   "Etablissement.frx":06E6
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Text4 
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
      Height          =   345
      Left            =   6000
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Text5 
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
      Height          =   345
      Left            =   6000
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3600
      Width           =   2175
   End
   Begin BARCODEXLib.BarcodeX BX1 
      Height          =   735
      Left            =   2520
      Top             =   3120
      Width           =   3255
      _Version        =   65536
      _ExtentX        =   5741
      _ExtentY        =   1296
      _StockProps     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0000001"
      BarcodeType     =   6
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   4440
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„ƒ””… »«·›—‰”Ì…"
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
      Left            =   8400
      TabIndex        =   21
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      DragMode        =   1  'Automatic
      Height          =   1215
      Left            =   240
      Stretch         =   -1  'True
      Top             =   8160
      Width           =   7815
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   240
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   11655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   63.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   17
      Top             =   4920
      Width           =   11535
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„ƒ””…"
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
      Left            =   8760
      TabIndex        =   16
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·”‰… «·œ—«”Ì…"
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
      Left            =   8760
      TabIndex        =   15
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " »œ√ »‘Â—"
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
      TabIndex        =   14
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰”»… «·„ƒ””… ›Ì «·⁄«∆œ"
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
      Left            =   7800
      TabIndex        =   13
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰”»… √”« –… «·‰”»…"
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
      Left            =   4320
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   3855
      Index           =   0
      Left            =   2280
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   8415
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„·«ÕŸ…:  Õ”» —”Ê„ «· ”ÃÌ· ›Ì Õ”«»«  «·‘Â— «·–Ì  »œ√ »Â «·”‰…"
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
      Left            =   2880
      TabIndex        =   11
      Top             =   4080
      Width           =   5295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰«  «·„ƒ””…"
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
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   8415
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ì»œ√ «· —ﬁÌ„ «· ”·”·Ì »«·—ﬁ„"
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
      Left            =   7800
      TabIndex        =   9
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ì»œ√ Ê’· «·œ›⁄ »«·—ﬁ„"
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
      Left            =   7800
      TabIndex        =   8
      Top             =   3600
      Width           =   2655
   End
End
Attribute VB_Name = "Etablissement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub Combo2_Change()
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

Private Sub Command1_Click()
On Error Resume Next
Dim a As Double
Dim b As Double
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
Text4.Text = Trim(Text4.Text)
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text1.BackColor = &H8080FF Then
Text1.SetFocus
ElseIf Text2.BackColor = &H8080FF Then
Text2.SetFocus
ElseIf Text3.BackColor = &H8080FF Then
Text3.SetFocus
ElseIf Text4.BackColor = &H8080FF Then
Text4.SetFocus
End If
Exit Sub
End If
Call cont
eb!eta = Text1.Text
eb!ann = Combo1.Text
eb!moi = Combo2.Text
a = Text3.Text
b = (100 - a)
eb!pce = b
eb!pcp = Text3.Text
eb!ser = Text4.Text
eb!rec = Text5.Text
eb!gch = Text6.Text
If Text4.Enabled = True Then
eb!sri = Text4.Text
eb!rcu = Text4.Text
End If
eb.Update
MsgBox " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ", vbInformation
Call cont
Text1.Text = eb!eta
Combo1.Text = eb!ann
Combo2.Text = eb!moi
Text2.Text = eb!pce
Text3.Text = eb!pcp
Text4.Text = eb!ser
Text6.Text = eb!gch
Interface.Caption = eb!gch
Interface.SBB1.Panels(3).Text = eb!eta

End Sub

Private Sub Command2_Click()
On Error Resume Next
On Error GoTo P
PicFile = ""
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Picture JPG |*.jpg|Picture Gif |*.Gif|Picture Bmp |*.Bmp|Picture Icon |*.ICO|All Picture |*.*"
    CommonDialog1.DialogTitle = "Picture"
    CommonDialog1.ShowOpen
    PicFile = CommonDialog1.FileName 'lien d'image
    Image1.Picture = LoadPicture(PicFile) 'Afficher l'image
    SavePicture Image1.Picture, App.Path & "\Tete_Long2266.jpg"
Exit Sub
P:
MsgBox "Êﬁ⁄ Œÿ√ ›Ì  Õ„Ì· «·ÊÀÌﬁ… , «·—Ã«¡ «⁄«œ… «·„Õ«Ê·…", vbExclamation
PicFile = ""
   Image1.Picture = LoadPicture(PicFile) 'Afficher l'image

End Sub

Private Sub Command3_Click()
On Error Resume Next
On Error GoTo P
PicFile = ""
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Picture JPG |*.jpg|Picture Gif |*.Gif|Picture Bmp |*.Bmp|Picture Icon |*.ICO|All Picture |*.*"
    CommonDialog1.DialogTitle = "Picture"
    CommonDialog1.ShowOpen
    PicFile = CommonDialog1.FileName 'lien d'image
    Image2.Picture = LoadPicture(PicFile) 'Afficher l'image
    SavePicture Image2.Picture, App.Path & "\Tete_Short0920.jpg"
Exit Sub
P:
MsgBox "Êﬁ⁄ Œÿ√ ›Ì  Õ„Ì· «·ÊÀÌﬁ… , «·—Ã«¡ «⁄«œ… «·„Õ«Ê·…", vbExclamation
PicFile = ""
   Image2.Picture = LoadPicture(PicFile) 'Afficher l'image

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim x$
Me.Left = 0
Me.Top = 0
Call active
Call chargcombo1
x$ = ""
x$ = dir$(App.Path & "\Tete_Long2266.jpg")
If x$ <> "" Then
PicFile = App.Path & "\Tete_Long2266.jpg"
Image1.Picture = LoadPicture(PicFile)
End If
x$ = ""
x$ = dir$(App.Path & "\Tete_Short0920.jpg")
If x$ <> "" Then
PicFile = App.Path & "\Tete_Short0920.jpg"
Image2.Picture = LoadPicture(PicFile)
End If
End Sub
Private Sub chargcombo1()
On Error Resume Next
Combo1.Clear
Call cont
Do While Not an.EOF
Combo1.AddItem an!ann
an.MoveNext
Loop
Text1.Text = eb!eta
Combo1.Text = eb!ann
Combo2.Text = eb!moi
Text2.Text = eb!pce
Text3.Text = eb!pcp
Text4.Text = eb!ser
Text5.Text = eb!rec
Text6.Text = eb!gch
If eb!ser = eb!sri Then
Text4.Enabled = True
Text5.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Else
Text4.Enabled = False
Text5.Enabled = False
Combo1.Enabled = False
Combo2.Enabled = False
End If

End Sub



Private Sub Text1_Change()
On Error Resume Next
Label9.Caption = Text1.Text
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Else
Text1.BackColor = &H8080FF
End If

End Sub

Private Sub Text1_Click()
On Error Resume Next
Text1_Change
End Sub

Private Sub Text2_Change()
On Error Resume Next
Dim a As Double
Dim b As Double
a = Val(Text2.Text)
b = 100 - a
Text3.Text = b
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
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text3_Change()
On Error Resume Next
If Len(Text3.Text) > 0 Then
Text3.BackColor = &HC000&
Else
Text3.BackColor = &H8080FF
End If

End Sub

Private Sub Text3_Click()
On Error Resume Next
Text3_Change
End Sub

Private Sub Text4_Change()
On Error Resume Next
If Len(Text4.Text) > 0 Then
xe = Text4.Text
Call Series
BX1.Caption = xs
Text4.BackColor = &HC000&
Else
Text4.BackColor = &H8080FF
End If

End Sub

Private Sub Text4_Click()
On Error Resume Next
Text4_Change
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If Len(Text4.Text) > 8 Then
KeyAscii = 0
End If
If KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text5_Change()
On Error Resume Next
If Len(Text5.Text) > 0 Then
Text5.BackColor = &HC000&
Else
Text5.BackColor = &H8080FF
End If

End Sub

Private Sub Text5_Click()
On Error Resume Next
Text5_Change
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If Len(Text4.Text) > 8 Then
KeyAscii = 0
End If
If KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub active()
On Error Resume Next
Call cont
If sr.RecordCount = 0 Then
eb!sri = eb!ser
eb.Update
End If
End Sub
