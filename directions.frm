VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form directions 
   BackColor       =   &H000000C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ÎÑæÌ"
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
      Left            =   0
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9270
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   8655
      Left            =   90
      TabIndex        =   0
      Top             =   480
      Width           =   3755
      _ExtentX        =   6615
      _ExtentY        =   15266
      _Version        =   327682
      HideSelection   =   0   'False
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      BorderStyle     =   1
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
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   9240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÓã ÇáãÓÊÎÏã"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÓã ÇáãÓÊÎÏã"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "directions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const GWL_STYLE = -16&
Private Const TVM_SETBKCOLOR = 4381&
Private Const TVM_GETBKCOLOR = 4383&
Private Const TVS_HASLINES = 2&
Dim frmlastForm As Form
'**** right TreeView
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYOUTRTL = &H400000
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
'**** right TreeView
Private Sub MakeTreeViewRTL()
On Error Resume Next
Dim rClientRect As RECT
Dim ReturnStyle As Long
ReturnStyle = GetWindowLong(TreeView1.hWnd, GWL_EXSTYLE)
SetWindowLong TreeView1.hWnd, GWL_EXSTYLE, ReturnStyle Or WS_EX_LAYOUTRTL
GetClientRect TreeView1.hWnd, rClientRect
InvalidateRect TreeView1.hWnd, rClientRect, True
End Sub
Private Sub couleur_treeview1()
On Error Resume Next
Dim lngStyle As Long
Call SendMessage(TreeView1.hWnd, TVM_SETBKCOLOR, 0, ByVal RGB(250, 247, 13))    'Change the background 'color to red.
    ' Now reset the style so that the tree lines appear properly
    lngStyle = GetWindowLong(TreeView1.hWnd, GWL_STYLE)
    Call SetWindowLong(TreeView1.hWnd, GWL_STYLE, lngStyle - TVS_HASLINES)
    Call SetWindowLong(TreeView1.hWnd, GWL_STYLE, lngStyle)
TreeView1.Sorted = True
End Sub

Private Sub Command1_Click()
On Error Resume Next
Call unloadforms
Unload Me
login.Show
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 0
Me.Left = 12880
Call MakeTreeViewRTL
Call couleur_treeview1
Call chargetreeview1



End Sub
Private Sub chargetreeview1()
On Error Resume Next
Dim id1 As String
Dim id2 As String
Dim i1 As Double
Dim i2 As Double
Dim i3 As Double
Dim i4 As Double
Dim i5 As Double
Dim i6 As Double
Dim n As Double
TreeView1.Nodes.Clear
TreeView1.Nodes.Add , , "OT", "ÕáÇÍíÇÊ ÇáãÓÊÎÏã"
i1 = 0
i2 = 0
i3 = 0
i4 = 0
i5 = 0
i6 = 0
Call cont
Do While Not ou.EOF
If ou!nom = login.Combo1.Text Then
If ou!act = "1" Then
id1 = ou!frm
id2 = ou!dir
If id2 = "1" Then
If i1 = 0 Then
TreeView1.Nodes.Add "OT", tvwChild, "DA", "ÇáÅÏÇÑÉ ÇáÚÇãÉ"
i1 = 1
End If
TreeView1.Nodes.Add "DA", tvwChild, id1, ou!div
End If
If id2 = "2" Then
If i2 = 0 Then
TreeView1.Nodes.Add "OT", tvwChild, "DL", "ÅÏÇÑÉ ÇáÑÞÇÈÉ"
i2 = 1
End If
TreeView1.Nodes.Add "DL", tvwChild, id1, ou!div
End If
If id2 = "3" Then
If i3 = 0 Then
TreeView1.Nodes.Add "OT", tvwChild, "DR", "ÅÏÇÑÉ ÇáÏÑæÓ"
i3 = 1
End If
TreeView1.Nodes.Add "DR", tvwChild, id1, ou!div
End If
If id2 = "4" Then
If i4 = 0 Then
TreeView1.Nodes.Add "OT", tvwChild, "DC", "ÅÏÇÑÉ ÇáÕäÏæÞ"
i4 = 1
End If
TreeView1.Nodes.Add "DC", tvwChild, id1, ou!div
End If
If id2 = "5" Then
If i5 = 0 Then
TreeView1.Nodes.Add "OT", tvwChild, "DT", "ÅÏÇÑÉ ÇáãÍÇÓÈÉ"
i5 = 1
End If
TreeView1.Nodes.Add "DT", tvwChild, id1, ou!div
End If
If id2 = "6" Then
If i6 = 0 Then
TreeView1.Nodes.Add "OT", tvwChild, "DH", "ÅÏÇÑÉ ÇáÇÑÔíÝ"
i6 = 1
End If
TreeView1.Nodes.Add "DH", tvwChild, id1, ou!div
End If
End If
End If
ou.MoveNext
Loop
TreeView1.Nodes(1).Expanded = True
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
On Error Resume Next
Dim n As Double
n = Node.Index
Label4.Caption = "OT"
Label3.Caption = Node.Key
If n > 1 Then
Label4.Caption = Node.Parent.Key
End If
Label5.Caption = n
If Label4.Caption <> "OT" Then
Call unloadforms
If Label3.Caption = "uti" Then
Utilisateurs.Show
End If
If Label3.Caption = "prt" Then
Partenaires.Show
End If
If Label3.Caption = "dir" Then
Etablissement.Show
End If
If Label3.Caption = "fnc" Then
Fonctionnaires.Show
End If
If Label3.Caption = "prf" Then
Professeurs.Show
End If
If Label3.Caption = "cla" Then
Classes.Show
End If
If Label3.Caption = "agn" Then
Correspondants.Show
End If
If Label3.Caption = "etu" Then
Etudiants.Show
End If
If Label3.Caption = "bil" Then
Cartes.Show
End If
If Label3.Caption = "mat" Then
Matieres.Show
End If
If Label3.Caption = "note" Then
Notes.Show
End If
If Label3.Caption = "pet" Then
Pointage_E.Show
End If
If Label3.Caption = "ppr" Then
Pointage_P.Show
End If
If Label3.Caption = "emp" Then
Emplois.Show
End If
If Label3.Caption = "cpr" Then
Caisse_PRT.Show
End If
If Label3.Caption = "cfn" Then
Caisse_FNC.Show
End If
If Label3.Caption = "cpf" Then
Caisse_PRF.Show
End If
If Label3.Caption = "cet" Then
Caisse_ETU.Show
End If
If Label3.Caption = "cdp" Then
Caisse_DPS.Show
End If
If Label3.Caption = "cca" Then
Caisse_SLD.Show
End If
If Label3.Caption = "cbn" Then
Caisse_BNK.Show
End If
If Label3.Caption = "trc" Then
Compte_TRS.Show
End If
If Label3.Caption = "tca" Then
Compte_ARC.Show
End If
If Label3.Caption = "tpr" Then
Compte_PRT.Show
End If
If Label3.Caption = "tfn" Then
Compte_FNC.Show
End If
If Label3.Caption = "tpf" Then
Compte_PRF.Show
End If
If Label3.Caption = "tet" Then
Compte_ETU.Show
End If
If Label3.Caption = "tdp" Then
Compte_DPS.Show
End If
If Label3.Caption = "tbn" Then
Compte_BNK.Show
End If
If Label3.Caption = "tcl" Then
Compte_CLS.Show
End If
If Label3.Caption = "spn" Then
Archives_AS.Show
End If
If Label3.Caption = "tjr" Then
Coin_CRS.Show
End If
If Label3.Caption = "rch" Then
Recherches.Show
End If

End If
End Sub

