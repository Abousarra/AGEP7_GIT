VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Caisse_ETU 
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
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   -9360
      ScaleHeight     =   7815
      ScaleWidth      =   10095
      TabIndex        =   82
      Top             =   6360
      Width           =   10095
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   10320
      ScaleHeight     =   8745
      ScaleWidth      =   2505
      TabIndex        =   75
      Top             =   840
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
         TabIndex        =   77
         Top             =   8400
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid grd6 
         Height          =   8295
         Left            =   0
         TabIndex        =   76
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
   Begin VB.CommandButton Command5 
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
      Height          =   345
      Left            =   1920
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Õ›Ÿ «·»Ì«‰« "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5220
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   8580
      Width           =   1200
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   10320
      ScaleHeight     =   8745
      ScaleWidth      =   2505
      TabIndex        =   67
      Top             =   840
      Width           =   2535
      Begin VB.CheckBox Check13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "≈„ﬂ«‰Ì… ⁄—÷ √ﬂÀ— „‰ ﬁ”„ ›Ì ¬‰ Ê«Õœ"
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
         Height          =   255
         Left            =   0
         TabIndex        =   80
         Top             =   8400
         Width           =   2535
      End
      Begin ComctlLib.TreeView TreeView1 
         Height          =   8415
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   14843
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
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7800
      ScrollBars      =   2  'Vertical
      TabIndex        =   66
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   345
      Left            =   5760
      ScrollBars      =   2  'Vertical
      TabIndex        =   65
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   345
      Left            =   5760
      ScrollBars      =   2  'Vertical
      TabIndex        =   64
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
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
      Height          =   330
      Left            =   6240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   960
      Width           =   735
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
      Left            =   7080
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   960
      Width           =   1695
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00008000&
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
      ItemData        =   "Caisse_ETU.frx":0000
      Left            =   2880
      List            =   "Caisse_ETU.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   72
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "«· ”ÃÌ·"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   32
      Top             =   2400
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Ì‰«Ì—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   31
      Top             =   1800
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "›»—«Ì—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   30
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "„«—”"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   29
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "«»—Ì·"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   28
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "„«ÌÊ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   27
      Top             =   2160
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "ÌÊ‰ÌÊ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   26
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "ÌÊ·ÌÊ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   2400
      TabIndex        =   25
      Top             =   2520
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "√€”ÿ”"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1200
      TabIndex        =   24
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "”» „»—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   23
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "«ﬂ Ê»—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   3240
      TabIndex        =   22
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "‰Ê›„»—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   2280
      TabIndex        =   21
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "œÌ”„»—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   1320
      TabIndex        =   20
      Top             =   1800
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Ã„Ì⁄ «·—”Ê„"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   4200
      TabIndex        =   19
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox Text5 
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   7800
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Text6 
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
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   7800
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Text            =   "0"
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "”Õ»"
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
      Left            =   1080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "Õ–›"
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
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3000
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3000
      Top             =   120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3480
      Top             =   120
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   7200
      ScaleHeight     =   555
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   8520
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton Command7 
         Caption         =   " Œ“Ì‰ «·«Ê’«·"
         Height          =   375
         Left            =   120
         TabIndex        =   83
         Top             =   120
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2040
         TabIndex        =   78
         Text            =   "Text8"
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox Combo3 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Caisse_ETU.frx":0027
         Left            =   1080
         List            =   "Caisse_ETU.frx":0031
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "Caisse_ETU.frx":0042
         Left            =   1800
         List            =   "Caisse_ETU.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Text            =   "Text7"
         Top             =   2160
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid grd1 
         Height          =   8175
         Left            =   960
         TabIndex        =   3
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   14420
         _Version        =   393216
         BackColor       =   16744576
         BackColorFixed  =   16744576
         BackColorBkg    =   16744576
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
      Begin VB.Label Label37 
         Caption         =   "Label37"
         Height          =   375
         Left            =   3120
         TabIndex        =   79
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Label27"
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "Label28"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label29 
         Caption         =   "Label29"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label30 
         Caption         =   "Label30"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label32 
         Caption         =   "Label32"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label33 
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   0
         Width           =   735
      End
      Begin VB.Label Label34 
         Caption         =   "0"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„” ÊÏ"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
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
         Left            =   1680
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "0"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   1575
      End
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   5280
      TabIndex        =   14
      Top             =   9240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   4695
      Left            =   5280
      TabIndex        =   35
      Top             =   3360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   8281
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
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grd3 
      Height          =   3735
      Left            =   240
      TabIndex        =   36
      Top             =   3360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   1
      Cols            =   4
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
   Begin MSFlexGridLib.MSFlexGrid grd4 
      Height          =   1335
      Left            =   240
      TabIndex        =   37
      Top             =   8040
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   2355
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
   Begin MSComCtl2.DTPicker DT1 
      Height          =   345
      Left            =   5280
      TabIndex        =   38
      Top             =   3000
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
   Begin MSFlexGridLib.MSFlexGrid grd5 
      Height          =   615
      Left            =   10440
      TabIndex        =   74
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
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·Ê’·"
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
      TabIndex        =   81
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   1320
      X2              =   1320
      Y1              =   7200
      Y2              =   7560
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
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
      ForeColor       =   &H00000000&
      Height          =   320
      Left            =   200
      TabIndex        =   71
      Top             =   7230
      Width           =   1095
   End
   Begin VB.Shape Shape6 
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      Left            =   8520
      TabIndex        =   63
      Top             =   960
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   375
      Index           =   1
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   8535
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   62
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label11 
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
      Left            =   4920
      TabIndex        =   61
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3000
      TabIndex        =   60
      Top             =   1395
      Width           =   1455
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
      Left            =   4200
      TabIndex        =   59
      Top             =   1395
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5760
      TabIndex        =   58
      Top             =   1395
      Width           =   1455
   End
   Begin VB.Label Label6 
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
      Left            =   6600
      TabIndex        =   57
      Top             =   1395
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8040
      TabIndex        =   56
      Top             =   1395
      Width           =   1335
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   8880
      TabIndex        =   55
      Top             =   1395
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   2
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   8535
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8640
      TabIndex        =   54
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8640
      TabIndex        =   53
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·Ê’·"
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
      TabIndex        =   52
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «·œ›⁄"
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
      Left            =   6600
      TabIndex        =   51
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      Height          =   1095
      Index           =   3
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1750
      Width           =   10095
   End
   Begin VB.Line Line1 
      X1              =   5640
      X2              =   5640
      Y1              =   1800
      Y2              =   2880
   End
   Begin VB.Shape Shape1 
      Height          =   5655
      Index           =   4
      Left            =   5160
      Top             =   2880
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      Height          =   4335
      Index           =   5
      Left            =   120
      Top             =   2880
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "«·„œ›Ê⁄"
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
      Height          =   375
      Left            =   9240
      TabIndex        =   50
      Top             =   8160
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7680
      TabIndex        =   49
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "«·»«ﬁÌ"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   48
      Top             =   8160
      Width           =   855
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5280
      TabIndex        =   47
      Top             =   8160
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      Height          =   1335
      Left            =   5160
      Top             =   8160
      Width           =   5055
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "«·„œ›Ê⁄"
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
      Height          =   615
      Left            =   6480
      TabIndex        =   46
      Top             =   8520
      Width           =   3615
   End
   Begin VB.Line Line3 
      X1              =   6480
      X2              =   6480
      Y1              =   8520
      Y2              =   9120
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "«·—”Ê„ «·√’·Ì…    ‰”»… «· Œ›Ì÷   «·—”Ê„ »⁄œ «· Œ›Ì÷"
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
      Left            =   5760
      TabIndex        =   45
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1440
      TabIndex        =   44
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "«·»«ﬁÌ"
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
      Height          =   375
      Left            =   2280
      TabIndex        =   43
      Top             =   7200
      Width           =   855
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3240
      TabIndex        =   42
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "«·„œ›Ê⁄"
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
      Height          =   375
      Left            =   4200
      TabIndex        =   41
      Top             =   7200
      Width           =   855
   End
   Begin VB.Shape Shape3 
      Height          =   735
      Left            =   120
      Top             =   7200
      Width           =   5055
   End
   Begin VB.Shape Shape4 
      Height          =   1935
      Left            =   120
      Top             =   7560
      Width           =   5055
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‘ÂÊ—  Õ„· œÌÊ‰«"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   14
      Left            =   840
      TabIndex        =   40
      Top             =   7560
      Width           =   3615
   End
   Begin VB.Line Line4 
      X1              =   5160
      X2              =   10200
      Y1              =   9120
      Y2              =   9120
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "’‰œÊﬁ «· ·«„Ì–"
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
      TabIndex        =   39
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Caisse_ETU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
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
Dim data As New Access.Application
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
Private Sub Check1_Click(Index As Integer)
On Error Resume Next
Dim i As Double
Dim k As Double
Dim j As Double
Dim n As Double
Dim tx As String
Dim a As Double
Dim P As Double
Dim r As Double
j = Index
If j = 13 Then
For k = 0 To 12
If Check1(k).Enabled = True Then
Check1(k).Value = Check1(13).Value
End If
Next k
End If
If j < 13 Then
grd2.Row = 0
grd2.Col = 0
grd2.Text = "—”Ê„"
grd2.Col = 1
grd2.Text = "«·„” Õﬁ"
grd2.Col = 2
grd2.Text = "«·„œ›Ê⁄"
grd2.Col = 3
grd2.Text = "«·»«ﬁÌ"
grd2.Visible = False
grd2.Rows = 14
For k = 1 To 13
If k < 4 Then
grd2.ColWidth(k) = 1100
grd2.ColAlignment(k) = 1
grd2.ColWidth(4) = 0
grd2.ColAlignment(0) = 1
End If
grd2.RowHeight(k) = 330
Next k
k = 1
If j > 0 And j < 10 Then
k = j + 4
End If
If j > 9 Then
k = j - 8
End If
If Check1(j).Value = 1 And Check1(j).Enabled = True Then
grd2.Row = k
grd2.Col = 0
grd2.Text = Check1(j).Caption
grd2.Col = 1
If j = 0 Then
grd2.Text = Text3.Text
Else
grd2.Text = Text4.Text
End If
grd2.Col = 2
If j = 0 Then
grd2.Text = Text3.Text
Else
grd2.Text = Text4.Text
End If
grd2.Col = 3
grd2.Text = "0"
grd2.Col = 4
grd2.Text = j
End If
If Check1(j).Value = 0 Then
grd2.Row = k
grd2.Col = 0
grd2.Text = ""
grd2.Col = 1
grd2.Text = ""
grd2.Col = 2
grd2.Text = ""
grd2.Col = 3
grd2.Text = ""
End If
Label19.Caption = ""
P = 0
r = 0
For k = 1 To 13
grd2.Row = k
grd2.Col = 0
tx = grd2.Text
If Len(grd2.Text) = 0 Then
grd2.Row = k
grd2.Col = 0
grd2.RowHeight(k) = 0
Else
grd2.Row = k
grd2.Col = 2
a = grd2.Text
P = P + a
grd2.Row = k
grd2.Col = 3
a = grd2.Text
r = r + a
Label19.Caption = Label19.Caption + tx + "  "
End If
Next k
grd2.Visible = True
Label7.Caption = P
Label17.Caption = r
End If
Label36.Caption = "0"
'MsgBox k
End Sub

Private Sub Check13_Click()
On Error Resume Next
Dim i As Double
For i = 1 To TreeView1.Nodes.Count
TreeView1.Nodes(i).Expanded = False
Next i
End Sub

Private Sub Combo1_Change()
On Error Resume Next
If Len(Combo1.Text) > 0 Then
'Combo1.BackColor = &HC000&
Call chargcombo3
'Combo3.BackColor = &H8080FF
Call chargegrd1_clear
Else
'Combo1.BackColor = &H8080FF
End If

End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub
Private Sub chargcombo3()
On Error Resume Next
Combo3.Clear
Call cont
Do While Not cl.EOF
If Combo1.Text = cl!niv Then
Combo3.AddItem cl!cla
End If
cl.MoveNext
Loop
End Sub
Private Sub chargegrd1()
On Error Resume Next
Dim i As Double
grd1.Clear
grd1.Cols = 2
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 3500
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.Row = 0
grd1.Col = 0
grd1.Text = "«·—ﬁ„ «· ”·”·Ì"
grd1.Col = 1
grd1.Text = "«”„ «· ·„Ì–"
i = 1
Call cont
grd1.Rows = et.RecordCount + 3
Do While Not et.EOF
If Combo3.Text = et!cla Then
grd1.Row = i
grd1.Col = 0
grd1.Text = et!sri
grd1.Col = 1
grd1.Text = et!nom
i = i + 1
End If
et.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 2
End Sub
Private Sub Combo2_Change()
On Error Resume Next
Label18.BackColor = &H8000&
Label18.Caption = ""
Label23.Caption = "0"
Label21.Caption = "0"
Command3.Enabled = False
Command1.Enabled = False
grd3.Visible = False
grd3.Clear
grd3.Cols = 5
grd3.Rows = 1
grd3.ColWidth(0) = 1100
grd3.ColWidth(1) = 1100
grd3.ColWidth(2) = 1100
grd3.ColWidth(3) = 1100
grd3.ColWidth(4) = 0
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.Row = 0
grd3.Col = 0
grd3.Text = "—”Ê„"
grd3.Col = 1
grd3.Text = "«·„” Õﬁ"
grd3.Col = 2
grd3.Text = "«·„œ›Ê⁄"
grd3.Col = 3
grd3.Text = "«·»«ﬁÌ"
grd3.Col = 4
grd3.Text = ""
grd3.Visible = True
End Sub

Private Sub Combo2_Click()
On Error Resume Next
Combo2_Change
End Sub

Private Sub Combo3_Change()
On Error Resume Next
If Len(Combo3.Text) > 0 Then
'Combo3.BackColor = &HC000&
grd1.Visible = False
Call chargegrd1
grd1.Visible = True
Else
'Combo3.BackColor = &H8080FF
End If
End Sub

Private Sub Combo3_Click()
On Error Resume Next
Combo3_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim a As Double
If Combo2.Text = "" Then
MsgBox "ÌÃ» «Œ Ì«— —ﬁ„ «·Ê’·", vbCritical
Exit Sub
End If
a = Val(Combo2.Text)
Call cont
data.OpenCurrentDatabase App.Path & "\" & Interface.SBB1.Panels(1).Text & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
data.DoCmd.OpenReport "Recus", acViewPreview, , "rcu =" & a, acWindowNormal, OpenArgs
'data.DoCmd.OpenReport "List_Etudiants", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing

End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim x$
Text1.Text = Trim(Text1.Text)
If Text1.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· «·—ﬁ„ «· ”·”·Ì À„ ⁄—÷ «·»Ì«‰« ", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
'*** verif n s
vtx1 = Text1.Text
Call verif_n_serie
Text1.Text = vtx2
'*** end verif n s
Label36.Caption = "0"
Call cont
Do While Not et.EOF
If et!sri = Text1.Text Or Val(et!sri) = Val(Text1.Text) Then
If et!act = "1" Then
Label12.Caption = et!nom
Label4.Caption = et!niv
Label8.Caption = et!cla
Label10.Caption = et!num
Call Frais_mensuel
grd4.Visible = False
Call chargegrd4
grd4.Visible = True
PicFile = ""
Image1.Picture = LoadPicture(PicFile)
x$ = ""
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & Label8.Caption & "\" & Text1.Text & ".jpg")
If x$ <> "" Then
PicFile = App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & Label8.Caption & "\" & Text1.Text & ".jpg"
Image1.Picture = LoadPicture(PicFile)
End If
Picture2.Visible = False
Exit Sub
Else
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· · ·„Ì–  „ Õ–›Â", vbCritical + arabic
Exit Sub
End If
End If
et.MoveNext
Loop
Call cont
Do While Not sr.EOF
If sr!sri = Text1.Text Or Val(sr!sri) = Val(Text1.Text) Then
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· ·Ì” —ﬁ„  ”·”·Ì · ·„Ì– Ê≈‰„« —ﬁ„  ”·”·Ì ·" + sr!eta, vbExclamation
Text1.Text = ""
Text1.SetFocus
Exit Sub
End If
sr.MoveNext
Loop
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· €Ì— ’ÕÌÕ", vbExclamation
Text1.Text = ""
Text1.SetFocus

End Sub


Private Sub Command3_Click()
On Error Resume Next
Dim tx1 As String
Dim i As Double
Dim n As Double
Dim m1 As Double
Dim m2 As Double
Dim a As Double
Dim b As Double
If Combo2.Text = "" Then
MsgBox "ÌÃ» «Œ Ì«— —ﬁ„ «·Ê’·", vbCritical
Exit Sub
End If
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–« «·Ê’·ø", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Call cont
Do While Not ct.EOF
If Combo2.Text = ct!rec Then
tx1 = ct!yes
If Val(ct!yes) > 0 Then
MsgBox "·« Ì„ﬂ‰ Õ–› Â–« «·Ê’· Õ Ï Ì „ Õ–› «·Ê’· —ﬁ„ " + tx1, vbExclamation + arabic
Exit Sub
End If
End If
ct.MoveNext
Loop
a = eb!cca
b = Label23.Caption
'*** controle caisse
If b > a Then
MsgBox "—’Ìœ «·’‰œÊﬁ ·« Ì”„Õ »« „«„ «·⁄„·Ì…... Ì—ÃÏ ÷Œ „»·€ ÃœÌœ ›Ì «·’‰œÊﬁ", vbExclamation
Exit Sub
End If
'***
a = (a - b)
eb!cca = a
eb.Update
Call cont
Do While Not ct.EOF
If Combo2.Text = ct!rec Then
ct.Delete
End If
ct.MoveNext
Loop
Call cont
Do While Not ct.EOF
If Combo2.Text = ct!yes Then
ct!yes = "-1"
ct.Update
End If
ct.MoveNext
Loop
n = grd3.Rows
grd3.Visible = False
Call cont
Do While Not pc.EOF
Label27.Caption = pc!moi
Label30.Caption = pc!cla
m2 = pc!etu
For i = 1 To n - 1
grd3.Row = i
grd3.Col = 2
Label32.Caption = grd3.Text
grd3.Col = 4
Label28.Caption = grd3.Text
If Label27.Caption = Label28.Caption And Label30.Caption = Label8.Caption Then
m1 = Label32.Caption
pc!etu = (m2 - m1)
pc.Update
End If
Next i
pc.MoveNext
Loop
grd3.Visible = True
'**** archive de caisse supp
Adat = Date
Aheu = Time$
Atyp = "Õ–›"
Adet = "Ê’· —ﬁ„ " + Combo2.Text + " „‰ ÿ—› «· ·„Ì– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + Text1.Text
Amon = Label23.Caption
Acom = "”Ã· «· ·«„Ì–"
Auti = directions.Label2.Caption
'****************************************
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer2.Enabled = True
grd5.Col = 1
grd5.Row = 0
a = grd5.Text
grd5.Row = 1
b = grd5.Text
If Label18.Caption = "ÃœÌœ" Then
a = a - 1
End If
If Label18.Caption = "„—›Ê÷" Then
b = b - 1
End If
grd5.Col = 1
grd5.Row = 0
grd5.Text = a
grd5.Row = 1
grd5.Text = b
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
Label18.BackColor = &H8000&
Label18.Caption = ""
Picture4.Visible = False
Command3.Enabled = False

End If

End Sub


Private Sub Command4_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim n As Double
Dim k As Double
Dim a As Double
Dim b As Double
Dim tx1 As String
Dim tx2 As String
Dim w As Double
Dim u As Double
Dim y As Double
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
Text4.Text = Trim(Text4.Text)
If Text1.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· «·—ﬁ„ «· ”·”·Ì À„ ⁄—÷ «·»Ì«‰« ", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
If Label12.Caption = "" Then
MsgBox "«·—Ã«¡ «·÷€ÿ ⁄·Ï “— ⁄—÷ √Ê «·÷€ÿ ⁄·Ï ENTER", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text2.BackColor = &H8080FF Then
Text2.SetFocus
ElseIf Text3.BackColor = &H8080FF Then
Text3.SetFocus
ElseIf Text4.BackColor = &H8080FF Then
Text4.SetFocus
End If
Exit Sub
End If
Command4.Enabled = False
ProgressBar1.Value = 20
grd2.Visible = False
Call calcul_pay_res
n = grd2.Rows
If Label36.Caption = "1" Then
Call cont
Do While Not ct.EOF
If ct!sri = Text1.Text Or Val(ct!sri) = Val(Text1.Text) Then
tx1 = ct!mos
For i = 1 To n - 1
grd2.Row = i
grd2.Col = 0
tx2 = grd2.Text
If tx1 = tx2 And ct!yes = "-1" Then
ct!yes = Text2.Text
ct.Update
i = n
End If
Next i
End If
ct.MoveNext
Loop
End If
n = grd2.Rows
k = 0
u = 0
Call cont
ProgressBar1.Value = 0
For i = 1 To n - 1
grd2.Row = i
grd2.Col = 0
If Len(grd2.Text) > 1 Then
ct.AddNew
If k = 0 Then
ct!rcu = Text2.Text
k = 1
Else
ct!rcu = "0"
End If
ct!ann = Interface.SBB1.Panels(1).Text
ct!sri = Text1.Text
ct!nom = Label12.Caption
ct!niv = Label4.Caption
ct!cla = Label8.Caption
ct!num = Label10.Caption
ct!fra1 = Label16.Caption
ct!men1 = Label14.Caption
ct!edf = Text5.Text
ct!edm = Text6.Text
ct!fra2 = Text3.Text
ct!men2 = Text3.Text
ct!rec = Text2.Text
grd2.Row = i
grd2.Col = 0
ct!mos = grd2.Text
grd2.Col = 1
ct!men = grd2.Text
grd2.Col = 2
ct!pay = grd2.Text
Label29.Caption = grd2.Text
grd2.Col = 3
ct!res = grd2.Text
grd2.Col = 4
ct!moi = grd2.Text
Label27.Caption = grd2.Text
grd2.Col = 3
a = grd2.Text
If a > 0 Then
ct!yes = "-1"
Else
ct!yes = "0"
End If
ct!tpy = Label7.Caption
ct!trs = Label17.Caption
ct!mois = Label19.Caption
ct!dat = DT1.Value
ct!act = "0"
ct!mtf = ""
ct!sim = App.Path & "\Tete_Long2266.jpg"
ct.Update
Call calcul_prc_E_A
u = u + 1
w = (u * 100 / 13)
ProgressBar1.Value = w
End If
Next i
grd2.Visible = True
If k = 0 Then
MsgBox "·„ Ì „ œ›⁄ √Ì —”Ê„", vbCritical
Command4.Enabled = True
Exit Sub
End If
a = eb!cca
b = Label7.Caption
a = (a + b)
eb!rcu = Val(eb!rcu) + 1
eb!cca = a
eb.Update
'**** archive de caisse ajou et modif
Adat = Date
Aheu = Time$
Atyp = "≈÷«›…"
Adet = "Ê’· —ﬁ„ " + Text2.Text + " „‰ ÿ—› «· ·„Ì– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + Text1.Text
Amon = Label7.Caption
Acom = "”Ã· «· ·«„Ì–"
Auti = directions.Label2.Caption
'****************************************

ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
grd5.Col = 1
grd5.Row = 0
a = grd5.Text
grd5.Row = 1
b = grd5.Text
a = a + 1
grd5.Col = 1
grd5.Row = 0
grd5.Text = a
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

Private Sub Command5_Click()
On Error Resume Next
Dim i As Double
Dim k As Double
Dim P As Double
Dim r As Double
Dim m As Double
If Combo2.Text = "" Then
MsgBox "ÌÃ» «Œ Ì«— —ﬁ„ «·Ê’·", vbCritical
Exit Sub
End If
grd3.Visible = False
grd3.Clear
grd3.Cols = 5
grd3.Rows = 1
grd3.ColWidth(0) = 1100
grd3.ColWidth(1) = 1100
grd3.ColWidth(2) = 1100
grd3.ColWidth(3) = 1100
grd3.ColWidth(4) = 0
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.Row = 0
grd3.Col = 0
grd3.Text = "—”Ê„"
grd3.Col = 1
grd3.Text = "«·„” Õﬁ"
grd3.Col = 2
grd3.Text = "«·„œ›Ê⁄"
grd3.Col = 3
grd3.Text = "«·»«ﬁÌ"
grd3.Col = 4
grd3.Text = ""
i = 1
P = 0
r = 0
Call cont
grd3.Rows = ct.RecordCount + 3
Do While Not ct.EOF
If ct!rec = Combo2.Text Then
grd3.Row = i
grd3.Col = 0
grd3.Text = ct!mos
grd3.Col = 1
grd3.Text = ct!men
grd3.Col = 2
grd3.Text = ct!pay
grd3.Col = 3
grd3.Text = ct!res
grd3.Col = 4
grd3.Text = ct!moi
m = ct!pay
P = P + m
m = ct!res
r = r + m
If ct!rcu <> "0" Then
If ct!act = "0" Then
Label18.BackColor = &HFFFF&
Label18.Caption = "ÃœÌœ"
Command3.Enabled = True
Command1.Enabled = True
End If
If ct!act = "1" Then
Label18.BackColor = &HFF0000
Label18.Caption = "„ƒﬂœ"
Command3.Enabled = False
Command1.Enabled = True
End If
If ct!act = "2" Then
Label18.BackColor = &HFF&
Label18.Caption = "„—›Ê÷"
Command3.Enabled = True
Command1.Enabled = True
End If
End If
i = i + 1
End If
ct.MoveNext
Loop
grd3.Rows = i
grd3.Visible = True
Label23.Caption = P
Label21.Caption = r

End Sub

Private Sub Command6_Click()
On Error Resume Next
Picture4.Visible = False
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim j As Double
Dim r As Double
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
tx1 = ""
tx2 = ""
j = 1
Call cont3
Do While Not ce3.EOF
Call cont
tx1 = ce3!ser
If tx1 <> tx2 And tx2 <> "" Then
Command4_Click
End If
If tx1 <> tx2 Then
Text1.Text = ce3!ser
Command2_Click
End If
Text3.Text = ce3!fra
Text4.Text = ce3!man
Text2.Text = ce3!rec
DT1.Value = ce3!dat
r = ce3!res
tx3 = ce3!moi
For i = 0 To 12
If Check1(i).Caption = ce3!moi Then
Check1(i).Value = 1
i = 12
End If
Next i
If r > 0 Then
n = grd2.Rows
For j = 1 To n - 1
grd2.Row = j
grd2.Col = 0
tx4 = grd2.Text
If tx4 = tx3 Then
'MsgBox tx3
'Exit Sub
grd2.Row = j
grd2.Col = 2
grd2.Text = ce3!pay
grd2.Col = 3
grd2.Text = ce3!res
j = n
End If
Next j
End If
tx2 = ce3!ser
'If j = 10 Then
'MsgBox "OK  10", vbInformation
'Exit Sub
'End If
ce3.MoveNext
Loop
Command4_Click
MsgBox "OK", vbInformation


End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0
Call cont
Text2.Text = eb!rcu
DT1.Value = Date
Call Operations
Call MakeTreeViewRTL
Call chargetreeview1
Call couleur_treeview1
End Sub

Private Sub grd1_Click()
On Error Resume Next
Dim i As Double
i = grd1.Row
If i > 0 Then
grd1.Row = i
grd1.Col = 0
Text1.Text = grd1.Text
Command2_Click
End If
End Sub

Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
i = grd2.Row
j = grd2.Col
If j = 2 Then
grd2.Row = i
grd2.Col = j
grd2.CellBackColor = &HFFFF&
'Call calcul_pay_res
End If
End Sub

Private Sub grd2_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim i As Double
Dim j As Double
Dim n As Double
Dim a As Double
Dim b As Double
Dim c As Double
i = grd2.Row
j = grd2.Col
If j = 2 Then
If KeyAscii = 8 Then
grd2.Row = i
grd2.Col = 1
b = grd2.Text
grd2.Row = i
grd2.Col = 3
grd2.Text = b
grd2.Row = i
grd2.Col = j
grd2.Text = ""
Exit Sub
End If
If KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
Exit Sub
End If
With grd2
        Select Case .Col
            Case 0, 2:
             .Text = .Text + Chr$(KeyAscii)
            Case Else:
        End Select
    End With
grd2.Row = i
grd2.Col = 1
b = grd2.Text
grd2.Row = i
grd2.Col = j
a = grd2.Text
c = (b - a)
grd2.Row = i
grd2.Col = 3
grd2.Text = c
If a > b Then
grd2.Row = i
grd2.Col = 3
grd2.Text = b
grd2.Row = i
grd2.Col = j
grd2.Text = ""
End If
grd2.Row = i
grd2.Col = j
End If

End Sub

Private Sub grd4_Click()
On Error Resume Next
Dim n As Double
Dim i As Double
Dim j As Double
Dim tx As String
Dim tx2 As String
Dim s As Double
Dim P As Double
Dim r As Double
Dim m As Double
Dim k As Double
i = grd4.Row
j = grd4.Col
If i > 0 And j > 0 Then
grd4.Row = i
grd4.Col = 0
tx = grd4.Text
grd4.Col = 1
s = grd4.Text
grd4.Col = 2
P = grd4.Text
grd4.Col = 3
r = grd4.Text
grd4.Col = 4
m = grd4.Text
n = grd2.Rows
For k = 1 To n - 1
grd2.Row = k
grd2.Col = 0
tx2 = grd2.Text
If tx = tx2 Then
Exit Sub
End If
Next k
grd2.Rows = n + 1
grd2.Row = n
grd2.Col = 0
grd2.Text = tx
grd2.Col = 1
grd2.Text = r
grd2.Col = 2
grd2.Text = r
grd2.Col = 3
grd2.Text = "0"
grd2.Col = 4
grd2.Text = m
Label36.Caption = "1"
End If
End Sub


Private Sub grd5_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
i = grd5.Row
If i = 0 Then
j = i
Else
j = i + 1
End If
If i = 0 Or i = 1 Then
If grd5.RowHeight(i) > 0 Then
Text8.Text = j
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
Text1.Text = grd6.Text
grd6.Col = 1
Label37.Caption = grd6.Text
Command2_Click
Combo2.Text = Label37.Caption
Command5_Click
End If
End Sub

Private Sub Label18_Click()
On Error Resume Next
If Label18.BackColor = &HFF& Then
Call cont
Do While Not ct.EOF
If ct!rcu = Combo2.Text Then
MsgBox ct!mtf
Exit Sub
End If
ct.MoveNext
Loop
End If

End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Label36.Caption = "0"
Label18.BackColor = &H8000&
Label18.Caption = ""
Label23.Caption = "0"
Label21.Caption = "0"
Command3.Enabled = False
Command1.Enabled = False
grd2.Visible = False
grd3.Visible = False
grd4.Visible = False
Call chargegrd_clear
grd2.Visible = True
grd3.Visible = True
grd4.Visible = True
Text3.Text = ""
Text4.Text = ""
Label16.Caption = ""
Label12.Caption = ""
Label14.Caption = ""
Label23.Caption = ""
Label21.Caption = ""
Else
Text1.BackColor = &H8080FF
End If
Picture2.Visible = True
PicFile = ""
Image1.Picture = LoadPicture(PicFile)

End Sub

Private Sub Text1_Click()
On Error Resume Next
Text1_Change
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
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

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Text1.Text <> "" Then
If KeyCode = 13 Then
Command2_Click
End If
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

Private Sub Text3_KeyPress(KeyAscii As Integer)
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

Private Sub Text5_Change()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
If Len(Text5.Text) > 0 Then
a = Label16.Caption
b = Text5.Text
If b > 100 Then
Text5.Text = "100"
End If
If b < 0 Then
Text5.Text = "0"
End If
b = Text5.Text
c = (b * a) / 100
d = a - c
Text3.Text = d
Else
Text3.Text = Label16.Caption
End If
End Sub

Private Sub Text5_Click()
On Error Resume Next
Text5_Change
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
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

Private Sub Text6_Change()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
If Len(Text6.Text) > 0 Then
a = Label14.Caption
b = Text6.Text
If b > 100 Then
Text6.Text = "100"
End If
If b < 0 Then
Text6.Text = "0"
End If
b = Text6.Text
c = (b * a) / 100
d = a - c
Text4.Text = d
Else
Text4.Text = Label14.Caption
End If

End Sub

Private Sub Text6_Click()
On Error Resume Next
Text6_Change
End Sub
Private Sub chargegrd1_clear()
On Error Resume Next
Dim i As Double
grd1.Clear
grd1.Cols = 2
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 3500
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.Row = 0
grd1.Col = 0
grd1.Text = "«·—ﬁ„ «· ”·”·Ì"
grd1.Col = 1
grd1.Text = "«”„ «· ·„Ì–"
End Sub
Private Sub chargegrd_clear()
On Error Resume Next
Dim k As Double
For k = 0 To 13
Check1(k).Enabled = True
Check1(k).Value = 0
Next k
Combo2.Clear
grd2.Clear
grd2.Cols = 5
grd2.Rows = 1
grd2.ColWidth(0) = 1100
grd2.ColWidth(1) = 1100
grd2.ColWidth(2) = 1100
grd2.ColWidth(3) = 1100
grd2.ColWidth(4) = 0
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.Row = 0
grd2.Col = 0
grd2.Text = "—”Ê„"
grd2.Col = 1
grd2.Text = "«·„” Õﬁ"
grd2.Col = 2
grd2.Text = "«·„œ›Ê⁄"
grd2.Col = 3
grd2.Text = "«·»«ﬁÌ"
grd3.Clear
grd3.Cols = 5
grd3.Rows = 1
grd3.ColWidth(0) = 1100
grd3.ColWidth(1) = 1100
grd3.ColWidth(2) = 1100
grd3.ColWidth(3) = 1100
grd3.ColWidth(4) = 0
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.Row = 0
grd3.Col = 0
grd3.Text = "—”Ê„"
grd3.Col = 1
grd3.Text = "«·„” Õﬁ"
grd3.Col = 2
grd3.Text = "«·„œ›Ê⁄"
grd3.Col = 3
grd3.Text = "«·»«ﬁÌ"
grd4.Clear
grd4.Cols = 5
grd4.Rows = 1
grd4.ColWidth(0) = 1100
grd4.ColWidth(1) = 1100
grd4.ColWidth(2) = 1100
grd4.ColWidth(3) = 1100
grd4.ColWidth(4) = 0
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.Row = 0
grd4.Col = 0
grd4.Text = "—”Ê„"
grd4.Col = 1
grd4.Text = "«·„” Õﬁ"
grd4.Col = 2
grd4.Text = "«·„œ›Ê⁄"
grd4.Col = 3
grd4.Text = "«·»«ﬁÌ"
End Sub
Private Sub chargegrd4()
On Error Resume Next
Dim i As Double
Dim k As Double
grd4.Clear
grd4.Cols = 5
grd4.Rows = 1
grd4.ColWidth(0) = 1100
grd4.ColWidth(1) = 1100
grd4.ColWidth(2) = 1100
grd4.ColWidth(3) = 1100
grd4.ColWidth(4) = 0
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.Row = 0
grd4.Col = 0
grd4.Text = "«·‘Â—"
grd4.Col = 1
grd4.Text = "«·„” Õﬁ"
grd4.Col = 2
grd4.Text = "«·„œ›Ê⁄"
grd4.Col = 3
grd4.Text = "«·»«ﬁÌ"
grd4.Col = 4
grd4.Text = ""
i = 1
Combo2.Clear
Call cont
grd4.Rows = ct.RecordCount + 3
Do While Not ct.EOF
If ct!sri = Text1.Text Or Val(ct!sri) = Val(Text1.Text) Then
If ct!rcu <> "0" Then
Combo2.AddItem ct!rcu
End If
If ct!yes = "-1" Then
grd4.Row = i
grd4.Col = 0
grd4.Text = ct!mos
grd4.Col = 1
grd4.Text = ct!men
grd4.Col = 2
grd4.Text = ct!pay
grd4.Col = 3
grd4.Text = ct!res
grd4.Col = 4
grd4.Text = ct!moi
i = i + 1
End If
k = ct!moi
Check1(k).Enabled = False
Check1(k).Value = 1
End If
ct.MoveNext
Loop
grd4.Rows = i
End Sub
Private Sub Frais_mensuel()
On Error Resume Next
Call cont
Do While Not cl.EOF
If cl!cla = Label8.Caption Then
Label16.Caption = cl!fra
Label14.Caption = cl!men
Text3.Text = cl!fra
Text4.Text = cl!men
Exit Sub
End If
cl.MoveNext
Loop
End Sub
Private Sub calcul_pay_res()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim P As Double
Dim r As Double
Dim m As Double
Dim tx As String
P = 0
r = 0
Label19.Caption = ""
n = grd2.Rows
For i = 1 To n - 1
grd2.Row = i
grd2.Col = 0
tx = grd2.Text
If Len(grd2.Text) > 1 Then
grd2.Row = i
grd2.Col = 2
If Len(grd2.Text) < 1 Then
grd2.Text = "0"
End If
m = grd2.Text
P = P + m
grd2.Row = i
grd2.Col = 3
m = grd2.Text
r = r + m
Label19.Caption = Label19.Caption + tx + "  "
End If
Next i
Label7.Caption = P
Label17.Caption = r
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
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
'ProgressBar1.Value = ProgressBar1.Value + 8
'If ProgressBar1.Value > 90 Then
Label7.Caption = ""
Label17.Caption = ""
Label23.Caption = ""
Label21.Caption = ""
Timer1.Enabled = False
ProgressBar1.Value = 0
ProgressBar1.Visible = True
grd2.Visible = False
grd3.Visible = False
grd4.Visible = False
Call chargegrd_clear
Call chargegrd4
grd2.Visible = True
grd3.Visible = True
grd4.Visible = True
'Combo2.Text = Text2.Text
Text2.Text = Val(Text2.Text) + 1
'MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation + arabic
Call archive_caisse

Command4.Enabled = True
'End If

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation + arabic
Label7.Caption = ""
Label17.Caption = ""
Label23.Caption = ""
Label21.Caption = ""
Timer2.Enabled = False
ProgressBar1.Value = 0
ProgressBar1.Visible = True
grd2.Visible = False
grd3.Visible = False
grd4.Visible = False
Call chargegrd_clear
Call chargegrd4
grd2.Visible = True
grd3.Visible = True
grd4.Visible = True
Call archive_caisse
End If

End Sub
Private Sub calcul_prc_E_A()
On Error Resume Next
Dim m1 As Double
Dim m2 As Double
Dim m3 As Double
Dim m As Double
m1 = Label29.Caption
m = Label27.Caption
Call cont
Do While Not pc.EOF
If Label8.Caption = pc!cla And m = pc!moi Then
m2 = pc!etu
m3 = (m1 + m2)
pc!etu = m3
pc.Update
Exit Sub
End If
pc.MoveNext
Loop
pc.AddNew
pc!moi = Label27.Caption
pc!niv = Label4.Caption
pc!cla = Label8.Caption
pc!etu = Label29.Caption
pc!pro = "0"
pc!nbr = "0"
pc.Update
End Sub
Private Sub chargetreeview1()
On Error Resume Next
On Error Resume Next
Dim id1 As String
Dim id2 As String
Dim id3 As String
Dim id4 As String
Dim i As Double
Dim n As Double
TreeView1.Nodes.Clear
Call cont
Do While Not cl.EOF
If cl!act = "1" Then
id2 = cl!cla
TreeView1.Nodes.Add , tvwChild, id2, cl!cla
End If
cl.MoveNext
Loop
Call cont
Do While Not et.EOF
If et!act = "1" Then
id1 = et!sri
id2 = "E" + id1
id3 = et!cla
TreeView1.Nodes.Add id3, tvwChild, id2, et!nom
End If
et.MoveNext
Loop
End Sub

Private Sub TreeView1_Expand(ByVal Node As ComctlLib.Node)
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Double
If Check13.Value = 0 Then
j = Node.Index
For i = 1 To TreeView1.Nodes.Count
If i <> j Then
TreeView1.Nodes(i).Expanded = False
End If
Next i
End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
On Error Resume Next
'Dim vg As String
Dim n As Double
Text1.Text = ""
Text7.Text = Node.Key
n = Len(Text7.Text)
If n > 2 Then
n = (n - 1)
vg = Mid$(Text7.Text, 2, n)
If Val(vg) = 0 Then
Exit Sub
End If
Text1.Text = vg
Command2_Click
End If

End Sub
Private Sub Operations()
On Error Resume Next
Dim a As Double
Dim b As Double
grd5.Rows = 2
grd5.Cols = 2
grd5.ColWidth(0) = 1000
grd5.ColWidth(1) = 1000
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 3
grd5.Col = 0
grd5.Row = 0
grd5.Text = "Ê’· ÃœÌœ"
grd5.Row = 1
grd5.Text = "Ê’· „—›Ê÷"
a = 0
b = 0
Call cont
Do While Not ct.EOF
If ct!rcu <> "0" Then
If ct!act = "0" Then
a = a + 1
End If
If ct!act = "2" Then
b = b + 1
End If
End If
ct.MoveNext
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
grd6.Clear
grd6.Rows = 1
grd6.Cols = 2
grd6.ColWidth(0) = 1150
grd6.ColWidth(1) = 1100
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.Row = 0
grd6.Col = 0
grd6.Text = "«·—ﬁ„ «· ”·”·Ì"
grd6.Col = 1
grd6.Text = "—ﬁ„ «·Ê’·"
i = 1
Call cont
grd6.Rows = ct.RecordCount + 3
Do While Not ct.EOF
If ct!act = Text8.Text And ct!rcu <> "0" Then
grd6.Row = i
grd6.Col = 0
grd6.Text = ct!sri
grd6.Col = 1
grd6.Text = ct!rcu
If Text8.Text = "0" Then
grd6.CellBackColor = &HFFFF&
Else
grd6.CellBackColor = &HFF&
End If
i = i + 1
End If
ct.MoveNext
Loop
grd6.Rows = i
End Sub

