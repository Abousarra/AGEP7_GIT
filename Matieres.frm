VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Matieres 
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
   Begin VB.CommandButton Command4 
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
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   79
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
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
      TabIndex        =   78
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   2175
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
      ItemData        =   "Matieres.frx":0000
      Left            =   10200
      List            =   "Matieres.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   41
      Top             =   1200
      Width           =   1695
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
      Left            =   5760
      ScrollBars      =   2  'Vertical
      TabIndex        =   40
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Text7 
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
      IMEMode         =   3  'DISABLE
      Left            =   7680
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Top             =   1200
      Width           =   1215
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
      IMEMode         =   3  'DISABLE
      Left            =   2640
      ScrollBars      =   2  'Vertical
      TabIndex        =   38
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox mens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   37
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   4
      Left            =   5400
      TabIndex        =   36
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   315
      Index           =   5
      Left            =   4320
      TabIndex        =   35
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   300
      Index           =   6
      Left            =   5400
      TabIndex        =   34
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   315
      Index           =   7
      Left            =   4320
      TabIndex        =   33
      Top             =   7680
      Width           =   735
   End
   Begin VB.TextBox mens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   32
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox mens 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   6
      Left            =   240
      TabIndex        =   31
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox mens 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   7
      Left            =   240
      TabIndex        =   30
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   8
      Left            =   5400
      TabIndex        =   29
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   300
      Index           =   9
      Left            =   4320
      TabIndex        =   28
      Top             =   8040
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   315
      Index           =   10
      Left            =   5400
      TabIndex        =   27
      Top             =   8400
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   300
      Index           =   11
      Left            =   4320
      TabIndex        =   26
      Top             =   8400
      Width           =   735
   End
   Begin VB.TextBox mens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   25
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox mens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   24
      Top             =   8400
      Width           =   1695
   End
   Begin VB.TextBox mens 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   8
      Left            =   240
      TabIndex        =   23
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox mens 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   9
      Left            =   240
      TabIndex        =   22
      Top             =   8400
      Width           =   1695
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   300
      Index           =   12
      Left            =   5400
      TabIndex        =   21
      Top             =   8760
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   13
      Left            =   4320
      TabIndex        =   20
      Top             =   8760
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   300
      Index           =   14
      Left            =   5400
      TabIndex        =   19
      Top             =   9120
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
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
      Height          =   315
      Index           =   15
      Left            =   4320
      TabIndex        =   18
      Top             =   9120
      Width           =   735
   End
   Begin VB.TextBox mens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   17
      Top             =   8760
      Width           =   1695
   End
   Begin VB.TextBox mens 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   16
      Top             =   9120
      Width           =   1695
   End
   Begin VB.TextBox mens 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   10
      Left            =   240
      TabIndex        =   15
      Top             =   8760
      Width           =   1695
   End
   Begin VB.TextBox mens 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
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
      Height          =   285
      Index           =   11
      Left            =   240
      TabIndex        =   14
      Top             =   9120
      Width           =   1695
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   9720
      TabIndex        =   13
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   8760
      TabIndex        =   12
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   7800
      TabIndex        =   11
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   6840
      TabIndex        =   10
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   17
      Left            =   7800
      TabIndex        =   9
      Top             =   8160
      Width           =   735
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   18
      Left            =   6840
      TabIndex        =   8
      Top             =   8160
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ÃœÊ·  ·ﬁ«∆Ì"
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
      Left            =   8520
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Õ›Ÿ «· €ÌÌ—« "
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
      Left            =   10320
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8760
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   3240
      ScaleHeight     =   795
      ScaleWidth      =   5835
      TabIndex        =   4
      Top             =   2520
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
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
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
      ItemData        =   "Matieres.frx":0004
      Left            =   10200
      List            =   "Matieres.frx":0011
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox coff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   16
      Left            =   8760
      TabIndex        =   2
      Top             =   8160
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "„”Õ «·ﬂ·"
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
      Left            =   6840
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8760
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
      ItemData        =   "Matieres.frx":002D
      Left            =   2640
      List            =   "Matieres.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   5760
      TabIndex        =   42
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   4695
      Left            =   240
      TabIndex        =   43
      Top             =   1920
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8281
      _Version        =   393216
      BackColor       =   32768
      BackColorFixed  =   32768
      BackColorSel    =   32768
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
   Begin VB.Line Line3 
      X1              =   10080
      X2              =   10080
      Y1              =   1680
      Y2              =   720
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Index           =   1
      Left            =   6720
      Top             =   6840
      Width           =   6015
   End
   Begin VB.Line Line2 
      X1              =   2520
      X2              =   2520
      Y1              =   1680
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   5640
      X2              =   5640
      Y1              =   720
      Y2              =   1680
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "÷«—» «·„«œ…"
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
      TabIndex        =   77
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Index           =   9
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   12615
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
      TabIndex        =   76
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„«œ…"
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
      TabIndex        =   75
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ê«œ Ê«·÷Ê«—»"
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
      TabIndex        =   74
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄œ· «·„«œ… „‰"
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
      Left            =   3960
      TabIndex        =   73
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label55 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«· ﬁœÌ—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   2880
      TabIndex        =   72
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label Label55 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«· ﬁœÌ—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   71
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label Label60 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„‰"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6000
      TabIndex        =   70
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label Label59 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   69
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„‰"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6000
      TabIndex        =   68
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label Label56 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   67
      Top             =   7680
      Width           =   855
   End
   Begin VB.Label Label55 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«· ﬁœÌ—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   66
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Label Label55 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«· ﬁœÌ—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   65
      Top             =   8040
      Width           =   2295
   End
   Begin VB.Label Label54 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„‰"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6000
      TabIndex        =   64
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   63
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label51 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„‰"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6000
      TabIndex        =   62
      Top             =   8040
      Width           =   495
   End
   Begin VB.Label Label50 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   61
      Top             =   8400
      Width           =   855
   End
   Begin VB.Label Label55 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«· ﬁœÌ—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   60
      Top             =   9120
      Width           =   2295
   End
   Begin VB.Label Label55 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«· ﬁœÌ—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   1920
      TabIndex        =   59
      Top             =   8760
      Width           =   2295
   End
   Begin VB.Label Label48 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„‰"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6000
      TabIndex        =   58
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label Label47 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   57
      Top             =   8760
      Width           =   855
   End
   Begin VB.Label Label43 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„‰"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6000
      TabIndex        =   56
      Top             =   8760
      Width           =   495
   End
   Begin VB.Label Label44 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   55
      Top             =   9120
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Index           =   0
      Left            =   120
      Top             =   6840
      Width           =   6495
   End
   Begin VB.Label Label66 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "÷Ê«—»"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   10440
      TabIndex        =   54
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label65 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·«„ Õ«‰ 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8640
      TabIndex        =   53
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label62 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·«„ Õ«‰ 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7560
      TabIndex        =   52
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label61 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·«„ Õ«‰ 3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6720
      TabIndex        =   51
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÃœÊ· «· ﬁœÌ—«  ··„” ÊÏ «·«⁄œ«œÌ Ê«·À«‰ÊÌ"
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
      Left            =   2520
      TabIndex        =   50
      Top             =   6960
      Width           =   3975
   End
   Begin VB.Label Label64 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„”‰ÊÏ «·«⁄œ«œÌ Ê«·À«‰ÊÌ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9720
      TabIndex        =   49
      Top             =   7800
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„”‰ÊÏ «·«» œ«∆Ì "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9840
      TabIndex        =   48
      Top             =   8160
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·«Œ »«—« "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9600
      TabIndex        =   47
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÃœÊ· «·÷Ê«—» "
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
      Left            =   10080
      TabIndex        =   46
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      Height          =   4935
      Index           =   10
      Left            =   120
      Top             =   1800
      Width           =   12615
   End
   Begin VB.Label Label4 
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
      TabIndex        =   45
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "·€…  œ—Ì” «·„«œ…"
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
      Left            =   3840
      TabIndex        =   44
      Top             =   1200
      Width           =   1695
   End
End
Attribute VB_Name = "Matieres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub coff_Change(Index As Integer)
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim a As Double
Dim b As Double
Dim c As String
Dim d As String
i = Index
If Len(coff(i).Text) > 0 Then
coff(i).BackColor = &HC000&
Else
coff(i).BackColor = &H8080FF
End If
For i = 4 To 15
j = j + 1
d = j
If coff(i).Text <> "" Then
If i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Then
a = coff(i).Text
b = coff(i - 1).Text
c = b
coff(i + 1).Text = coff(i).Text
If a >= b Then
MsgBox " «·⁄œœ «·„œŒ· ÌÃ» √‰ ÌﬂÊ‰ √’€— „‰ " + c, vbCritical
coff(i).Text = ""
coff(i).SetFocus
Exit Sub
End If
End If
End If
Next i
End Sub
Public Sub coffes()
On Error Resume Next
Call cont
coff(0).Text = cf2!cof0
coff(1).Text = cf2!cof1
coff(2).Text = cf2!cof2
coff(3).Text = cf2!cof3
coff(16).Text = cf2!cof16
coff(17).Text = cf2!cof17
coff(18).Text = cf2!cof18
coff(4).Text = cf2!cof4
coff(5).Text = cf2!cof5
coff(6).Text = cf2!cof6
coff(7).Text = cf2!cof7
coff(8).Text = cf2!cof8
coff(9).Text = cf2!cof9
coff(10).Text = cf2!cof10
coff(11).Text = cf2!cof11
coff(12).Text = cf2!cof12
coff(13).Text = cf2!cof13
coff(14).Text = cf2!cof14
coff(15).Text = cf2!cof15
mens(0).Text = cf2!tex9
mens(1).Text = cf2!tex12
mens(2).Text = cf2!tex15
mens(3).Text = cf2!tex18
mens(4).Text = cf2!tex19
mens(5).Text = cf2!tex20
mens(6).Text = cf2!tex21
mens(7).Text = cf2!tex22
mens(8).Text = cf2!tex23
mens(9).Text = cf2!tex24
mens(10).Text = cf2!tex25
mens(11).Text = cf2!tex26
End Sub

Private Sub Combo1_Change()
On Error Resume Next
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
grd1.Visible = False
Call chargegrd1
grd1.Visible = True
Text1.SetFocus
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

Private Sub Combo4_Change()
On Error Resume Next
If Len(Combo4.Text) > 0 Then
Combo4.BackColor = &HC000&
Combo1.BackColor = &H8080FF
Call chargcombo1
Call chargegrd1_clear
If Combo4.Text <> "«» œ«∆Ì" Then
Combo2.Text = "⁄—»Ì…"
Combo2.Enabled = False
Else
Combo2.Enabled = True
Combo2.Clear
Combo2.AddItem "⁄—»Ì…"
Combo2.AddItem "›—‰”Ì…"
Combo2.BackColor = &H8080FF
End If
Else
Combo4.BackColor = &H8080FF
End If
End Sub

Private Sub Combo4_Click()
On Error Resume Next
Combo4_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim i As Integer
Dim j As Integer
Dim a As Double
Dim b As Double
Dim c As String
Dim d As String
For i = 0 To 18
If coff(i).Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
coff(i).SetFocus
Exit Sub
End If
Next i
For i = 0 To 11
If mens(i).Text = "" Or coff(i).Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
mens(i).SetFocus
Exit Sub
End If
Next i
j = 0
For i = 4 To 15
j = j + 1
d = j
If coff(i).Text <> "" Then
If i = 5 Or i = 7 Or i = 9 Or i = 11 Or i = 13 Then
a = coff(i).Text
b = coff(i - 1).Text
c = b
coff(i + 1).Text = coff(i).Text
If a >= b Then
MsgBox " «·⁄œœ «·„œŒ· ÌÃ» √‰ ÌﬂÊ‰ √’€— „‰ " + c, vbCritical
coff(i).Text = ""
coff(i).SetFocus
Exit Sub
End If
End If
End If
Next i
Call cont
cf2!cof0 = coff(0).Text
cf2!cof1 = coff(1).Text
cf2!cof2 = coff(2).Text
cf2!cof3 = coff(3).Text
cf2!cof4 = coff(4).Text
cf2!cof5 = coff(5).Text
cf2!cof6 = coff(6).Text
cf2!cof7 = coff(7).Text
cf2!cof8 = coff(8).Text
cf2!cof9 = coff(9).Text
cf2!cof10 = coff(10).Text
cf2!cof11 = coff(11).Text
cf2!cof12 = coff(12).Text
cf2!cof13 = coff(13).Text
cf2!cof14 = coff(14).Text
cf2!cof15 = coff(15).Text
cf2!cof16 = coff(16).Text
cf2!cof17 = coff(17).Text
cf2!cof18 = coff(18).Text
cf2!tex9 = mens(0).Text
cf2!tex12 = mens(1).Text
cf2!tex15 = mens(2).Text
cf2!tex18 = mens(3).Text
cf2!tex19 = mens(4).Text
cf2!tex20 = mens(5).Text
cf2!tex21 = mens(6).Text
cf2!tex22 = mens(7).Text
cf2!tex23 = mens(8).Text
cf2!tex24 = mens(9).Text
cf2!tex25 = mens(10).Text
cf2!tex26 = mens(11).Text
cf2.Update
'Call changementdecoffcients
MsgBox " „ Õ›Ÿ «· €ÌÌ—« ", vbInformation
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim i As Integer
Dim j As Integer
For j = 0 To 11
mens(j).Text = ""
Next j
i = 18
Do While i > -1
coff(i).Text = ""
i = i - 1
Loop
For j = 6 To 14
coff(j).Text = ""
Next j
coff(15).Text = "0"
coff(0).SetFocus
End Sub

Private Sub Command3_Click()
On Error Resume Next
Text1.Text = ""
Text1.SetFocus
Text7.Text = ""
'Text4.Text = ""
Label16.Caption = ""
'Call num_etudiants
grd1.Visible = False
grd1.Clear
grd1.Rows = 1
Call chargegrd1
grd1.Visible = True
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False

End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim n As Double
Text1.Text = Trim(Text1.Text)
Text7.Text = Trim(Text7.Text)
Text4.Text = Trim(Text4.Text)
If Text1.Text = "" Or Text7.Text = "" Or Text4.Text = "" Or Combo1.Text = "" Or Combo4.Text = "" Or Combo2.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text1.BackColor = &H8080FF Then
Text1.SetFocus
ElseIf Text7.BackColor = &H8080FF Then
Text7.SetFocus
ElseIf Text4.BackColor = &H8080FF Then
Text4.SetFocus
End If
Exit Sub
End If
If Text1.Text = "„Ã„Ê⁄ „Ê«œ «·⁄—»Ì…" Or Text1.Text = "Total MatiÈres FR" Then
MsgBox "€Ì— „„ﬂ‰... Ì⁄ »— Â–« «·«”„ „ÕÃÊ“«", vbCritical
Exit Sub
End If
Call cont
Do While Not mt.EOF
If Combo1.Text = mt!cla And Text1.Text = mt!mat And Label16.Caption <> mt!aut Then
MsgBox "€Ì— „„ﬂ‰... ·ﬁœ  „ ÕÃ“ Â–« «·«”„ ”«»ﬁ«", vbCritical
Exit Sub
End If
mt.MoveNext
Loop
If Label16.Caption <> "" Then
Call cont
Do While Not mt.EOF
If Label16.Caption = mt!aut Then
mt!niv = Combo4.Text
mt!cla = Combo1.Text
mt!mat = Text1.Text
mt!cof = Text7.Text
mt!moy = Text4.Text
mt!lng = Combo2.Text
mt.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
mt.MoveNext
Loop
End If
n = grd1.Rows
If n = 1 And Combo4.Text = "«» œ«∆Ì" Then
Call cont
Do While Not mt.EOF
If Combo1.Text = mt!cla Then
mt.Delete
End If
mt.MoveNext
Loop
mt.AddNew
mt!niv = Combo4.Text
mt!cla = Combo1.Text
mt!mat = "„Ã„Ê⁄ „Ê«œ «·⁄—»Ì…"
mt!cof = Text7.Text
mt!moy = Text4.Text
mt!lng = "⁄—»Ì…"
mt.Update
mt.AddNew
mt!niv = Combo4.Text
mt!cla = Combo1.Text
mt!mat = "Total MatiÈres FR"
mt!cof = Text7.Text
mt!moy = Text4.Text
mt!lng = "›—‰”Ì…"
mt.Update
End If
mt.AddNew
mt!niv = Combo4.Text
mt!cla = Combo1.Text
mt!mat = Text1.Text
mt!cof = Text7.Text
mt!moy = Text4.Text
mt!lng = Combo2.Text
mt.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Command7_Click()
On Error Resume Next
Call coffes1

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0
Call coffes
chargegrd1_clear
End Sub

Private Sub grd1_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim au As Double
Dim a As Double
Dim b As Double
Dim tx As String
i = grd1.Row
j = grd1.Col
If i > 0 Then
If j = 5 Then
grd1.Row = i
grd1.Col = 0
Label16.Caption = grd1.Text
grd1.Col = 1
Combo2.Text = grd1.Text
grd1.Col = 2
Text1.Text = grd1.Text
grd1.Col = 3
Text7.Text = grd1.Text
grd1.Col = 4
Text4.Text = grd1.Text
End If
If j = 6 Then
grd1.Row = i
grd1.Col = 0
Label16.Caption = grd1.Text
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–Â «·„«œ…", vbInformation + vbYesNo + arabic, "AGEP7")
If g = vbYes Then
Call cont
Do While Not mt.EOF
If Label16.Caption = mt!aut Then
mt.Delete
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
mt.MoveNext
Loop
Else
Label16.Caption = ""
End If

End If
End If

End Sub


Private Sub mens_Change(Index As Integer)
On Error Resume Next
Dim i As Double
i = Index
If Len(mens(i).Text) > 0 Then
mens(i).BackColor = &HC000&
Else
mens(i).BackColor = &H8080FF
End If

End Sub
Public Sub coffes1()
On Error Resume Next
Call cont
coff(0).Text = cf1!cof0
coff(1).Text = cf1!cof1
coff(2).Text = cf1!cof2
coff(3).Text = cf1!cof3
coff(16).Text = cf1!cof1
coff(17).Text = cf1!cof2
coff(18).Text = cf1!cof3
coff(4).Text = cf1!cof4
coff(5).Text = cf1!cof5
coff(6).Text = cf1!cof6
coff(7).Text = cf1!cof7
coff(8).Text = cf1!cof8
coff(9).Text = cf1!cof9
coff(10).Text = cf1!cof10
coff(11).Text = cf1!cof11
coff(12).Text = cf1!cof12
coff(13).Text = cf1!cof13
coff(14).Text = cf1!cof14
coff(15).Text = cf1!cof15
mens(0).Text = cf1!tex9
mens(1).Text = cf1!tex12
mens(2).Text = cf1!tex15
mens(3).Text = cf1!tex18
mens(4).Text = cf1!tex19
mens(5).Text = cf1!tex20
mens(6).Text = cf1!tex21
mens(7).Text = cf1!tex22
mens(8).Text = cf1!tex23
mens(9).Text = cf1!tex24
mens(10).Text = cf1!tex25
mens(11).Text = cf1!tex26
End Sub
Private Sub chargcombo1()
On Error Resume Next
Combo1.Clear
Call cont
Do While Not cl.EOF
If Combo4.Text = cl!niv And cl!act = "1" Then
Combo1.AddItem cl!cla
End If
cl.MoveNext
Loop
End Sub

Private Sub Text1_Change()
On Error Resume Next
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

Private Sub Text4_Change()
On Error Resume Next
If Len(Text4.Text) > 0 Then
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
If KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If

End Sub

Private Sub Text7_Change()
On Error Resume Next
If Len(Text7.Text) > 0 Then
Text7.BackColor = &HC000&
Else
Text7.BackColor = &H8080FF
End If

End Sub

Private Sub Text7_Click()
On Error Resume Next
Text7_Change
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
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
Command3_Click
End If

End Sub
Private Sub chargegrd1()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim j As Double
Dim i As Double
Dim P As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Cols = 7
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 2000
grd1.ColWidth(2) = 4400
grd1.ColWidth(3) = 2000
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 800
grd1.ColWidth(6) = 800
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 3
grd1.ColAlignment(6) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "·€… «· œ—Ì”"
grd1.Col = 2
grd1.Text = "«·„«œ…"
grd1.Col = 3
grd1.Text = "÷«—» «·„«œ…"
grd1.Col = 4
grd1.Text = "„⁄œ· «·„«œ…"
i = 1
Call cont
grd1.Rows = mt.RecordCount + 3
Do While Not mt.EOF
If Combo1.Text = mt!cla And mt!mat <> "„Ã„Ê⁄ „Ê«œ «·⁄—»Ì…" And mt!mat <> "Total MatiÈres FR" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = mt!aut
grd1.Col = 1
grd1.Text = mt!lng
grd1.Col = 2
grd1.Text = mt!mat
grd1.Col = 3
grd1.Text = mt!cof
grd1.Col = 4
grd1.Text = mt!moy
grd1.Col = 5
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 6
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
End If
mt.MoveNext
Loop
grd1.Rows = i
'grd1.Col = 4
'grd1.Sort = 2
End Sub
Private Sub chargegrd1_clear()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim j As Double
Dim i As Double
Dim P As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Cols = 7
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 2000
grd1.ColWidth(2) = 4400
grd1.ColWidth(3) = 2000
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 800
grd1.ColWidth(6) = 800
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 3
grd1.ColAlignment(6) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "·€… «· œ—Ì”"
grd1.Col = 2
grd1.Text = "«·„«œ…"
grd1.Col = 3
grd1.Text = "÷«—» «·„«œ…"
grd1.Col = 4
grd1.Text = "„⁄œ· «·„«œ…"
End Sub


