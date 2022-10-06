VERSION 5.00
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Begin VB.Form Cartes 
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
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ÚÑÖ"
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
      Left            =   7200
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ÓÍÈ"
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
      Left            =   360
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ÓÍÈ"
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
      Left            =   3240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ÓÍÈ"
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
      Left            =   5640
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   1575
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
      ItemData        =   "Cartes.frx":0000
      Left            =   4920
      List            =   "Cartes.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1200
      Width           =   1455
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
      Height          =   345
      Left            =   8880
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   1200
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5295
      Left            =   1200
      Picture         =   "Cartes.frx":001B
      ScaleHeight     =   5295
      ScaleWidth      =   10575
      TabIndex        =   0
      Top             =   2760
      Width           =   10575
      Begin BARCODEXLib.BarcodeX BX1 
         Height          =   495
         Left            =   3960
         Top             =   3720
         Width           =   3135
         _Version        =   65536
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   13
         BackColor       =   16777215
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÑÞã ÇáÊÓáÓáí"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   30
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Line Line3 
         Index           =   3
         X1              =   3120
         X2              =   8040
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   17
         Top             =   840
         Width           =   6375
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2055
         Left            =   1320
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Image Image3 
         Appearance      =   0  'Flat
         Height          =   615
         Left            =   1680
         Picture         =   "Cartes.frx":18F9E
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Line Line3 
         Index           =   0
         X1              =   3120
         X2              =   8040
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÇÓã ÇáßÇãá"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   3120
         TabIndex        =   15
         Top             =   1560
         Width           =   4935
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÌäÓ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   3480
         TabIndex        =   13
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÑÞã ÇáæØäí"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   12
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   6360
         TabIndex        =   11
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Line Line3 
         Index           =   1
         X1              =   3120
         X2              =   8040
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáãÓÊæì"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8280
         TabIndex        =   10
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   7080
         TabIndex        =   9
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÞÓã"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   5280
         TabIndex        =   7
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÑÞã ÇáäÏÇÁ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   6
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   3120
         TabIndex        =   5
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "åÇÊÝ Çáæßíá"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   4
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   3120
         TabIndex        =   3
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÊÇÑíÎ ÇáÊÓÌíá"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   2
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   6360
         TabIndex        =   1
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Line Line3 
         Index           =   2
         X1              =   3120
         X2              =   8040
         Y1              =   2880
         Y2              =   2880
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   375
      Left            =   4440
      TabIndex        =   31
      Top             =   2280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÈØÇÞÇÊ ÇáÏÎæá"
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
      Left            =   5400
      TabIndex        =   25
      Top             =   0
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Index           =   9
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   12615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÈØÇÞÉ ÏÎæá ÊáÇãíÐ ÞÓã ãÚíä"
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
      TabIndex        =   24
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÈØÇÞÉ ÏÎæá ÊáãíÐ ãÚíä"
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
      Left            =   9720
      TabIndex        =   23
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÇáÑÞã ÇáÊÓáÓáí"
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
      Left            =   9720
      TabIndex        =   22
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Left            =   6240
      TabIndex        =   21
      Top             =   1200
      Width           =   735
   End
   Begin VB.Line Line1 
      X1              =   7080
      X2              =   7080
      Y1              =   720
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   3120
      X2              =   3120
      Y1              =   720
      Y2              =   1680
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÈØÇÞÉ ÏÎæá ÊáÇãíÐ ÇáãÄÓÓÉ ÌãíÚÇ"
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
      Left            =   240
      TabIndex        =   20
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "Cartes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim data As New Access.Application
Private Sub chargcombo1()
Combo1.Clear
Call cont
Do While Not cl.EOF
Combo1.AddItem cl!cla
cl.MoveNext
Loop
End Sub

Private Sub Combo1_Change()
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
Else
Combo1.BackColor = &H8080FF
End If
Call cont
Do While Not cl.EOF
If Combo1.Text = cl!cla Then
Label5.Caption = cl!aut
Exit Sub
End If
cl.MoveNext
Loop
End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub


Private Sub Command2_Click()
Dim a As Double
a = Val(Label5.Caption)
Call cont
data.OpenCurrentDatabase App.Path & "\" & Interface.SBB1.Panels(1).Text & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
data.DoCmd.OpenReport "Cartes_Etudiants", acViewPreview, , "aut =" & a, acWindowNormal, OpenArgs
'data.DoCmd.OpenReport "Cartes", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing

End Sub

Private Sub Command3_Click()
Dim a As Double
a = Val(Label5.Caption)
Call cont
data.OpenCurrentDatabase App.Path & "\" & Interface.SBB1.Panels(1).Text & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
data.DoCmd.OpenReport "Cartes_Etudiants", acViewPreview, , "ncl =" & a, acWindowNormal, OpenArgs
'data.DoCmd.OpenReport "Cartes", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing

End Sub

Private Sub Command4_Click()
Dim a As Double
a = Val(Label5.Caption)
Call cont
data.OpenCurrentDatabase App.Path & "\" & Interface.SBB1.Panels(1).Text & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
data.DoCmd.OpenReport "Cartes_Etudiants", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.OpenReport "Cartes", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing

End Sub

Private Sub Command5_Click()
Dim x$
Call cont
Do While Not et.EOF
If Text1.Text = et!sri Or Val(Text1.Text) = Val(et!sri) Then
BX1.Caption = et!sri
Label5.Caption = et!aut
Label11.Caption = et!nom
Label16.Caption = et!nni
Label14.Caption = et!sex
Label18.Caption = et!niv
Label20.Caption = et!cla
Label22.Caption = et!num
Label26.Caption = et!dat
Label24.Caption = et!tel
x$ = ""
PicFile = ""
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & Label20.Caption & "\" & BX1.Caption & ".jpg")
If x$ <> "" Then
PicFile = App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & Label20.Caption & "\" & BX1.Caption & ".jpg"
End If
Image1.Picture = LoadPicture(PicFile)
Exit Sub
End If
et.MoveNext
Loop

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Call cont
Label9.Caption = eb!eta
Call chargcombo1
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Else
Text1.BackColor = &H8080FF
End If

End Sub

Private Sub Text1_Click()
Text1_Change
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Text1.Text <> "" Then
If KeyCode = 13 Then
Command5_Click
End If
End If

End Sub
