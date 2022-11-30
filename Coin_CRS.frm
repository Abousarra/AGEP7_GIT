VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Coin_CRS 
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
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8415
      ScaleWidth      =   10215
      TabIndex        =   90
      Top             =   1200
      Width           =   10215
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   10440
      ScaleHeight     =   4695
      ScaleWidth      =   2295
      TabIndex        =   89
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
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
      Left            =   9600
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "⁄—÷ «·Õ”«»"
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
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox Text2 
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
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Left            =   8640
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   720
      Width           =   855
   End
   Begin VB.PictureBox Picture111 
      Height          =   3255
      Left            =   6960
      ScaleHeight     =   3195
      ScaleWidth      =   1875
      TabIndex        =   6
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Text            =   "Text3"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "Text5"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
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
         Left            =   840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Command11"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label41 
         Caption         =   "Label41"
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Label23"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label40 
         Caption         =   "Label40"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label50 
         Caption         =   "Label50"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label51 
         Caption         =   "Label51"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture222 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   10440
      ScaleHeight     =   3105
      ScaleWidth      =   2265
      TabIndex        =   4
      Top             =   4920
      Width           =   2295
      Begin ComctlLib.TreeView TreeView2 
         Height          =   3135
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   5530
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
   Begin VB.PictureBox Picture333 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   10440
      ScaleHeight     =   3585
      ScaleWidth      =   2265
      TabIndex        =   2
      Top             =   1200
      Width           =   2295
      Begin ComctlLib.TreeView TreeView1 
         Height          =   3615
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   6376
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
      Left            =   4080
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "⁄—÷"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "«·Õ”«» Ê«·Õ÷Ê—"
      TabPicture(0)   =   "Coin_CRS.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "«·‰ «∆Ã Ê«·ﬂ‘Ê›"
      TabPicture(1)   =   "Coin_CRS.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture444"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.PictureBox Picture444 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   6975
         Left            =   120
         ScaleHeight     =   6975
         ScaleWidth      =   9975
         TabIndex        =   39
         Top             =   360
         Width           =   9975
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   9855
            TabIndex        =   57
            Top             =   6000
            Width           =   9855
            Begin VB.CommandButton Command9 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "ﬂ‘› «·œ—Ã« "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   0
               MaskColor       =   &H00FF0000&
               Style           =   1  'Graphical
               TabIndex        =   59
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   1215
            End
            Begin VB.CommandButton Command5 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "«·„⁄œ· «·⁄«„"
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
               Height          =   330
               Left            =   6120
               MaskColor       =   &H00FF0000&
               Style           =   1  'Graphical
               TabIndex        =   58
               Top             =   120
               Width           =   1095
            End
            Begin VB.Line Line6 
               X1              =   4080
               X2              =   4080
               Y1              =   480
               Y2              =   840
            End
            Begin VB.Line Line5 
               X1              =   3120
               X2              =   3120
               Y1              =   480
               Y2              =   840
            End
            Begin VB.Line Line4 
               X1              =   1200
               X2              =   6000
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "320"
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
               Left            =   1200
               TabIndex        =   71
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·€Ì«»"
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
               TabIndex        =   70
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label25 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "320"
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
               TabIndex        =   69
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·— »…"
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
               Left            =   5160
               TabIndex        =   68
               Top             =   480
               Width           =   735
            End
            Begin VB.Line Line3 
               X1              =   1200
               X2              =   1200
               Y1              =   0
               Y2              =   840
            End
            Begin VB.Line Line2 
               X1              =   6000
               X2              =   6000
               Y1              =   0
               Y2              =   840
            End
            Begin VB.Line Line1 
               X1              =   7320
               X2              =   7320
               Y1              =   0
               Y2              =   960
            End
            Begin VB.Shape Shape1 
               Height          =   855
               Index           =   2
               Left            =   -120
               Top             =   0
               Width           =   9855
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Mention"
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
               Left            =   1320
               TabIndex        =   67
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label21"
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
               Left            =   7320
               TabIndex        =   66
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«· ﬁœÌ—"
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
               Left            =   4440
               TabIndex        =   65
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label19"
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
               Left            =   2160
               TabIndex        =   64
               Top             =   120
               Width           =   3135
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label18"
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
               Left            =   6120
               TabIndex        =   63
               Top             =   480
               Width           =   1095
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label16"
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
               Left            =   7320
               TabIndex        =   62
               Top             =   480
               Width           =   855
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„Ã„Ê⁄ «·÷Ê«—»"
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
               Left            =   8280
               TabIndex        =   61
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "„Ã„Ê⁄ «·‰ﬁ«ÿ"
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
               Left            =   8280
               TabIndex        =   60
               Top             =   120
               Width           =   1455
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            Height          =   855
            Left            =   120
            ScaleHeight     =   855
            ScaleWidth      =   9855
            TabIndex        =   40
            Top             =   6000
            Width           =   9855
            Begin VB.CommandButton Command10 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "ﬂ‘› «·œ—Ã« "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   0
               MaskColor       =   &H00FF0000&
               Style           =   1  'Graphical
               TabIndex        =   42
               Top             =   120
               UseMaskColor    =   -1  'True
               Width           =   1215
            End
            Begin VB.CommandButton Command4 
               Appearance      =   0  'Flat
               BackColor       =   &H0000FFFF&
               Caption         =   "«· ﬁœÌ—"
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
               Height          =   330
               Left            =   8400
               MaskColor       =   &H00FF0000&
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   480
               Width           =   1335
            End
            Begin VB.Line Line19 
               X1              =   8280
               X2              =   8280
               Y1              =   960
               Y2              =   480
            End
            Begin VB.Line Line17 
               X1              =   3120
               X2              =   3120
               Y1              =   480
               Y2              =   960
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„Ã„Ê⁄"
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
               Left            =   8280
               TabIndex        =   56
               Top             =   0
               Width           =   1455
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·„⁄œ·"
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
               Left            =   8280
               TabIndex        =   55
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label33 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label33"
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
               Left            =   6960
               TabIndex        =   54
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label34 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label34"
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
               TabIndex        =   53
               Top             =   480
               Width           =   3975
            End
            Begin VB.Label Label36 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label36"
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
               Left            =   6960
               TabIndex        =   52
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label Label37 
               BackStyle       =   0  'Transparent
               Caption         =   "Mention"
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
               Left            =   3240
               TabIndex        =   51
               Top             =   480
               Width           =   1575
            End
            Begin VB.Shape Shape1 
               Height          =   855
               Index           =   3
               Left            =   0
               Top             =   0
               Width           =   9855
            End
            Begin VB.Line Line8 
               X1              =   8280
               X2              =   8280
               Y1              =   0
               Y2              =   480
            End
            Begin VB.Line Line9 
               X1              =   6960
               X2              =   6960
               Y1              =   0
               Y2              =   480
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·— »…"
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
               TabIndex        =   50
               Top             =   480
               Width           =   735
            End
            Begin VB.Label Label42 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "320"
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
               Left            =   1320
               TabIndex        =   49
               Top             =   480
               Width           =   975
            End
            Begin VB.Line Line10 
               X1              =   1200
               X2              =   9840
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line Line13 
               X1              =   1200
               X2              =   9840
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Line Line11 
               X1              =   6360
               X2              =   6360
               Y1              =   480
               Y2              =   0
            End
            Begin VB.Label Label43 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label43"
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
               Left            =   5040
               TabIndex        =   48
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label Label44 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label44"
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
               Left            =   5040
               TabIndex        =   47
               Top             =   240
               Width           =   1335
            End
            Begin VB.Line Line12 
               X1              =   5040
               X2              =   5040
               Y1              =   0
               Y2              =   480
            End
            Begin VB.Line Line14 
               X1              =   4440
               X2              =   4440
               Y1              =   480
               Y2              =   0
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label45"
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
               Left            =   3120
               TabIndex        =   46
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label46"
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
               Left            =   3120
               TabIndex        =   45
               Top             =   240
               Width           =   1335
            End
            Begin VB.Line Line15 
               X1              =   3120
               X2              =   3120
               Y1              =   0
               Y2              =   480
            End
            Begin VB.Line Line16 
               X1              =   2520
               X2              =   2520
               Y1              =   0
               Y2              =   480
            End
            Begin VB.Label Label32 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label32"
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
               Left            =   1200
               TabIndex        =   44
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Label38"
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
               Left            =   1200
               TabIndex        =   43
               Top             =   240
               Width           =   1335
            End
            Begin VB.Line Line18 
               X1              =   1200
               X2              =   1200
               Y1              =   960
               Y2              =   0
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grd2 
            Height          =   4935
            Left            =   120
            TabIndex        =   72
            Top             =   0
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   8705
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
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   6975
         Left            =   -74880
         ScaleHeight     =   6975
         ScaleWidth      =   9975
         TabIndex        =   21
         Top             =   360
         Width           =   9975
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
            ItemData        =   "Coin_CRS.frx":0038
            Left            =   6720
            List            =   "Coin_CRS.frx":0045
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
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
            Left            =   5760
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   120
            Width           =   735
         End
         Begin VB.CommandButton Command8 
            Appearance      =   0  'Flat
            BackColor       =   &H0000FFFF&
            Caption         =   "”Õ» ”Ã· «·Õ÷Ê— «·ÌÊ„Ì"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3600
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   6480
            Width           =   2895
         End
         Begin MSFlexGridLib.MSFlexGrid grd3 
            Height          =   4815
            Left            =   5160
            TabIndex        =   25
            Top             =   480
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   8493
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
            Height          =   5055
            Left            =   240
            TabIndex        =   26
            Top             =   480
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   8916
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
         Begin VB.Label Label2 
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
            TabIndex        =   38
            Top             =   5400
            Width           =   1335
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
            Left            =   6480
            TabIndex        =   37
            Top             =   5400
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
            Left            =   7560
            TabIndex        =   36
            Top             =   5400
            Width           =   1335
         End
         Begin VB.Label Label13 
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
            Left            =   8760
            TabIndex        =   35
            Top             =   5400
            Width           =   855
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
            Left            =   720
            TabIndex        =   34
            Top             =   120
            Width           =   3615
         End
         Begin VB.Label Label29 
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
            Left            =   7800
            TabIndex        =   33
            Top             =   120
            Width           =   1215
         End
         Begin VB.Shape Shape1 
            Height          =   5775
            Index           =   4
            Left            =   5040
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   4815
         End
         Begin VB.Shape Shape1 
            Height          =   855
            Index           =   5
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   6000
            Width           =   9735
         End
         Begin VB.Shape Shape1 
            Height          =   5775
            Index           =   6
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   4815
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "-"
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
            Height          =   255
            Left            =   6600
            TabIndex        =   32
            Top             =   6120
            Width           =   1335
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "-"
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
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   6120
            Width           =   1335
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·€Ì«»"
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
            Left            =   1200
            TabIndex        =   30
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Label Label47 
            Alignment       =   2  'Center
            BackColor       =   &H0000FF00&
            Caption         =   "-"
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
            Height          =   255
            Left            =   2520
            TabIndex        =   29
            Top             =   6120
            Width           =   1335
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„Ê⁄ ⁄œœ «·Õ’’ «· Ì Õ÷—Â«"
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
            Left            =   3600
            TabIndex        =   28
            Top             =   6120
            Width           =   2895
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ã„Ê⁄ Õ’’ «·ﬁ”„"
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
            Left            =   7800
            TabIndex        =   27
            Top             =   6120
            Width           =   1935
         End
      End
   End
   Begin VB.Label Label0 
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
      Left            =   3960
      TabIndex        =   88
      Top             =   720
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   12615
   End
   Begin VB.Label Label21000 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬂ‰ «·Êﬂ·«¡"
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
      TabIndex        =   87
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label00 
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
      Left            =   7800
      TabIndex        =   86
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„ «· ”·”·Ì"
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
      Index           =   0
      Left            =   6120
      TabIndex        =   85
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Left            =   9000
      TabIndex        =   84
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label6 
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
      Left            =   6840
      TabIndex        =   83
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "-"
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
      Height          =   255
      Left            =   6240
      TabIndex        =   82
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·‰œ«¡"
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
      Left            =   4920
      TabIndex        =   81
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "-"
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
      Height          =   255
      Left            =   4320
      TabIndex        =   80
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·«”„"
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
      Left            =   3000
      TabIndex        =   79
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "-"
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
      Height          =   255
      Left            =   240
      TabIndex        =   78
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   10215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„ «· ”·”·Ì"
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
      Index           =   1
      Left            =   11160
      TabIndex        =   77
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·«”„"
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
      Index           =   2
      Left            =   7080
      TabIndex        =   76
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ·„… «·”—"
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
      Index           =   3
      Left            =   2520
      TabIndex        =   75
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·Â« ›"
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
      Index           =   4
      Left            =   8880
      TabIndex        =   74
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "-"
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
      Height          =   255
      Left            =   8280
      TabIndex        =   73
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   10800
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Shape Shape6 
      Height          =   1575
      Left            =   10800
      Top             =   8040
      Width           =   1575
   End
End
Attribute VB_Name = "Coin_CRS"
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
Dim data As New Access.Application
Private Sub MakeTreeViewRTL()
On Error Resume Next
Dim rClientRect As RECT
Dim ReturnStyle As Long
ReturnStyle = GetWindowLong(TreeView1.hWnd, GWL_EXSTYLE)
ReturnStyle = GetWindowLong(TreeView2.hWnd, GWL_EXSTYLE)
SetWindowLong TreeView1.hWnd, GWL_EXSTYLE, ReturnStyle Or WS_EX_LAYOUTRTL
SetWindowLong TreeView2.hWnd, GWL_EXSTYLE, ReturnStyle Or WS_EX_LAYOUTRTL
GetClientRect TreeView1.hWnd, rClientRect
GetClientRect TreeView2.hWnd, rClientRect
InvalidateRect TreeView1.hWnd, rClientRect, True
InvalidateRect TreeView2.hWnd, rClientRect, True
End Sub
Private Sub couleur_treeview1()
On Error Resume Next
Dim lngStyle As Long
Call SendMessage(TreeView1.hWnd, TVM_SETBKCOLOR, 0, ByVal RGB(250, 247, 13))    'Change the background 'color to red.
Call SendMessage(TreeView2.hWnd, TVM_SETBKCOLOR, 0, ByVal RGB(250, 247, 13))    'Change the background 'color to red.
    ' Now reset the style so that the tree lines appear properly
    lngStyle = GetWindowLong(TreeView1.hWnd, GWL_STYLE)
    lngStyle = GetWindowLong(TreeView2.hWnd, GWL_STYLE)
    Call SetWindowLong(TreeView1.hWnd, GWL_STYLE, lngStyle - TVS_HASLINES)
    Call SetWindowLong(TreeView2.hWnd, GWL_STYLE, lngStyle - TVS_HASLINES)
    Call SetWindowLong(TreeView1.hWnd, GWL_STYLE, lngStyle)
    Call SetWindowLong(TreeView2.hWnd, GWL_STYLE, lngStyle)
TreeView1.Sorted = True
TreeView2.Sorted = True
End Sub

Private Sub chargetreeview1()
On Error Resume Next
Dim id1 As String
Dim id2 As String
Dim i As Double
Dim n As Double
TreeView1.Nodes.Clear
'TreeView1.Nodes.Add , , "CR", "√”„«¡ «·Êﬂ·«¡"
Call cont
Do While Not cr.EOF
If cr!act = "1" Then
id1 = cr!sri
id2 = "C" + id1
TreeView1.Nodes.Add , tvwChild, id2, cr!nom
End If
cr.MoveNext
Loop
End Sub
Private Sub chargetreeview2()
On Error Resume Next
Dim id1 As String
Dim id2 As String
Dim i As Double
Dim n As Double
TreeView2.Nodes.Clear
'TreeView2.Nodes.Add , , "ET", "√”„«¡ «·Êﬂ·«¡"
Call cont
Do While Not et.EOF
If et!tel = Label00.Caption Then
id1 = et!sri
id2 = "E" + id1
TreeView2.Nodes.Add , tvwChild, id2, et!nom
End If
et.MoveNext
Loop
End Sub


Private Sub Combo2_Change()
On Error Resume Next
Label7.Caption = "0"
Label2.Caption = "0"
Command6.Enabled = False
Call grd2_clear
Command3_Click
End Sub
Private Sub grd2_clear()
On Error Resume Next
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

Private Sub Command1_Click()
On Error Resume Next
Text1.Text = Trim(Text1.Text)
If Text1.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· «·—ﬁ„ «· ”·”·Ì ", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
Call cont
Do While Not cr.EOF
If Text1.Text = cr!sri Or Val(Text1.Text) = Val(cr!sri) Then
Label0.Caption = cr!nom
Label00.Caption = cr!tel
Label51.Caption = cr!mot
Call chargetreeview2
Exit Sub
End If
cr.MoveNext
Loop
Call cont
Do While Not sr.EOF
If sr!sri = Text1.Text Or Val(sr!sri) = Val(Text1.Text) Then
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· ·Ì” —ﬁ„  ”·”·Ì ·„ÊŸ› Ê≈‰„« —ﬁ„  ”·”·Ì ·" + sr!eta, vbExclamation
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

Private Sub Command11_Click()
On Error Resume Next
On Error GoTo P
Dim n1 As Double
Dim n2 As Double
Dim k1 As Double
Dim k2 As Double
Dim s As Double
 Dim y$
n1 = 0
n2 = 0
s = 0
Label5.Caption = n1         'nombre de cours classe
     Label47.Caption = n2           'nombre de cours etudiant
     Label28.Caption = (n1 - n2)    'nombre de cours absence
     y$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Label8.Caption & ".txt")
If y$ <> "" Then
Open App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Label8.Caption & ".txt" For Input As #1
Do While Not EOF(1)
        s = s + 1
        Line Input #1, x
        If s > 2 And s Mod 2 <> 0 Then
        Label50.Caption = x
        vg = Mid$(Label50.Caption, 13, 1)
        k1 = vg
        n1 = n1 + k1
        End If
        
Loop
Close #1
End If

    y$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Text4.Text & ".txt")
If y$ <> "" Then
 Open App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Text4.Text & ".txt" For Input As #2
s = 0
    Do While Not EOF(2)
        s = s + 1
        Line Input #2, x
        If s > 2 And s Mod 2 <> 0 Then
        Label50.Caption = x
        vg = Mid$(Label50.Caption, 13, 1)
        k2 = vg
        n2 = n2 + k2
        End If
    Loop
    Close #2
End If
        
     Label5.Caption = n1         'nombre de cours classe
     Label47.Caption = n2           'nombre de cours etudiant
     Label28.Caption = (n1 - n2)    'nombre de cours absence
     Exit Sub
P:
     Exit Sub
    Close #1
     Close #2
End Sub

Private Sub Command2_Click()
On Error Resume Next
If Text1.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· «·—ﬁ„ «· ”·”·Ì À„ «·÷€ÿ ⁄·Ï ⁄—÷", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
If Label26.Caption = "" Then
MsgBox "«·—Ã«¡ «·÷€ÿ ⁄·Ï “— ⁄—÷ √Ê «·÷€ÿ ⁄·Ï ENTER", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· ﬂ·„… «·”— À„ «·÷€ÿ ⁄·Ï ⁄—÷ «·Õ”«»", vbCritical + arabic
Text2.SetFocus
Exit Sub
End If
If Text2.Text = Label51.Caption Then
Picture4.Visible = False
Picture5.Visible = False
Else
MsgBox "ﬂ·„… «·”— «· Ì √œŒ· „ €Ì— ’ÕÌÕ…", vbExclamation + arabic
Text2.Text = ""
Text2.SetFocus
End If

End Sub

Private Sub Command3_Click()
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
Command6.Enabled = True
i = i + 1
End If
ct.MoveNext
Loop
grd3.Rows = i
grd3.Visible = True
Label7.Caption = P
Label2.Caption = r

End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim x$
Text4.Text = Trim(Text4.Text)
If Text4.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· «·—ﬁ„ «· ”·”·Ì À„ ⁄—÷ «·»Ì«‰« ", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
Call cont
Do While Not et.EOF
If et!sri = Text4.Text Or Val(et!sri) = Val(Text4.Text) Then
Label4.Caption = et!niv
Label8.Caption = et!cla
Label10.Caption = et!num
Label12.Caption = et!nom
Label41.Caption = et!nfr
Label23.Caption = et!ncl
Label4.BackColor = &HC000&
Label8.BackColor = &HC000&
Label10.BackColor = &HC000&
Label12.BackColor = &HC000&
Call grd2_clear
grd2.Visible = False
If Label4.Caption = "«» œ«∆Ì" Then
Call chargegrd2_tete_pr
Else
Call chargegrd2_tete
End If
Call chargegrd2
Call coff_dv_ex
Call chargegrd2_notes
Call calcule_moyenne_lc
grd2.Visible = True
grd4.Visible = False
Call chargegrd4
grd4.Visible = True
Command11_Click
PicFile = ""
Image1.Picture = LoadPicture(PicFile)
x$ = ""
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & Label8.Caption & "\" & Text4.Text & ".jpg")
If x$ <> "" Then
PicFile = App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & Label8.Caption & "\" & Text4.Text & ".jpg"
Image1.Picture = LoadPicture(PicFile)
End If
Exit Sub
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

Private Sub Command8_Click()
On Error Resume Next
Dim x$
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Text4.Text & ".txt")
If x$ = "" Then
MsgBox "«·„·› «·„ÿ·Ê» €Ì— „ÊÃÊœ", vbExclamation
Exit Sub
End If
FileCopy App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Text4.Text & ".txt", App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\AGEP7.txt"
Shell "notepad.exe" & " " & App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\AGEP7.txt", vbNormalFocus

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Top = 0
Me.Left = 0
Call MakeTreeViewRTL
Call chargetreeview1
Call couleur_treeview1
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

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
On Error Resume Next
Dim n As Double
Text3.Text = Node.Key
n = Len(Text3.Text)
If n > 2 Then
n = (n - 1)
vg = Mid$(Text3.Text, 2, n)
Text1.Text = vg
Command1_Click
End If

End Sub


Private Sub TreeView2_NodeClick(ByVal Node As ComctlLib.Node)
On Error Resume Next
Dim i As Double
Dim j As Double
Dim n As Double
Text5.Text = Node.Key
n = Len(Text5.Text)
If n > 2 Then
n = (n - 1)
vg = Mid$(Text5.Text, 2, n)
Text4.Text = vg
Command7_Click
End If

End Sub
Private Sub chargegrd2_tete()
On Error Resume Next
Dim i As Double
Dim j As Double
grd2.Clear
grd2.Cols = 20
grd2.Rows = 1
grd2.ColWidth(0) = 1900
grd2.ColWidth(1) = 350
grd2.ColWidth(2) = 0
grd2.ColWidth(3) = 600
grd2.ColWidth(4) = 600
grd2.ColWidth(5) = 600
grd2.ColWidth(6) = 600
grd2.ColWidth(7) = 600
grd2.ColWidth(8) = 600
grd2.ColWidth(9) = 600
grd2.ColWidth(10) = 0
grd2.ColWidth(11) = 600
grd2.ColWidth(12) = 0
grd2.ColWidth(13) = 600
grd2.ColWidth(14) = 0
grd2.ColWidth(15) = 600
grd2.ColWidth(16) = 0
grd2.ColWidth(17) = 600
grd2.ColWidth(18) = 700
grd2.ColWidth(19) = 0
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
grd2.Text = "«·„«œ…"
grd2.Col = 1
grd2.Text = "÷‹"
grd2.Col = 2
grd2.Text = "„ . „"
grd2.Col = 3
grd2.Text = "«Œ‹ 1"
grd2.Col = 4
grd2.Text = "«Œ‹ 2"
grd2.Col = 5
grd2.Text = "«Œ‹ 3"
grd2.Col = 6
grd2.Text = "«Œ‹ 4"
grd2.Col = 7
grd2.Text = "«Œ‹ 5"
grd2.Col = 8
grd2.Text = "«Œ‹ 6"
grd2.Col = 9
grd2.Text = "„⁄œ· «Œ‹"
grd2.Col = 10
grd2.Text = "÷‹"
grd2.Col = 11
grd2.Text = "«„ Õ‹ 1"
grd2.Col = 12
grd2.Text = "÷‹"
grd2.Col = 13
grd2.Text = "«„ Õ‹ 2"
grd2.Col = 14
grd2.Text = "÷‹"
grd2.Col = 15
grd2.Text = "«„ Õ‹ 3"
grd2.Col = 16
grd2.Text = "÷‹"
grd2.Col = 17
grd2.Text = "«·„⁄œ·"
grd2.Col = 18
grd2.Text = "«·„Ã„Ê⁄"
End Sub

Private Sub chargegrd2_tete_pr()
On Error Resume Next
Dim i As Double
Dim j As Double
grd2.Clear
grd2.Cols = 20
grd2.Rows = 1
grd2.ColWidth(0) = 3200
grd2.ColWidth(1) = 0
grd2.ColWidth(2) = 550
grd2.ColWidth(3) = 0
grd2.ColWidth(4) = 0
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 0
grd2.ColWidth(7) = 0
grd2.ColWidth(8) = 0
grd2.ColWidth(9) = 0
grd2.ColWidth(10) = 0
grd2.ColWidth(11) = 1200
grd2.ColWidth(12) = 350
grd2.ColWidth(13) = 1200
grd2.ColWidth(14) = 350
grd2.ColWidth(15) = 1200
grd2.ColWidth(16) = 350
grd2.ColWidth(17) = 1200
grd2.ColWidth(18) = 0
grd2.ColWidth(19) = 0
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
grd2.Text = "«·„«œ…"
grd2.Col = 1
grd2.Text = "÷‹"
grd2.Col = 2
grd2.Text = "„ . „"
grd2.Col = 3
grd2.Text = "«Œ‹ 1"
grd2.Col = 4
grd2.Text = "«Œ‹ 2"
grd2.Col = 5
grd2.Text = "«Œ‹ 3"
grd2.Col = 6
grd2.Text = "«Œ‹ 4"
grd2.Col = 7
grd2.Text = "«Œ‹ 5"
grd2.Col = 8
grd2.Text = "«Œ‹ 6"
grd2.Col = 9
grd2.Text = "„⁄œ· «Œ‹"
grd2.Col = 10
grd2.Text = "÷‹"
grd2.Col = 11
grd2.Text = "«„ Õ‹ 1"
grd2.Col = 12
grd2.Text = "÷‹"
grd2.Col = 13
grd2.Text = "«„ Õ‹ 2"
grd2.Col = 14
grd2.Text = "÷‹"
grd2.Col = 15
grd2.Text = "«„ Õ‹ 3"
grd2.Col = 16
grd2.Text = "÷‹"
grd2.Col = 17
grd2.Text = "«·„⁄œ·"
grd2.Col = 18
grd2.Text = "«·„Ã„Ê⁄"
End Sub
Private Sub chargegrd2()
On Error Resume Next
Dim i As Double
Dim tx1 As String
Dim tx2 As String
i = 1
tx1 = "⁄—»Ì…"
Call cont
grd2.Rows = mt.RecordCount + 4
Do While Not mt.EOF
If mt!cla = Label8.Caption And mt!mat <> "„Ã„Ê⁄ „Ê«œ «·⁄—»Ì…" And mt!mat <> "Total MatiÈres FR" Then
If Label4.Caption = "«» œ«∆Ì" Then
tx2 = mt!lng
If tx1 <> tx2 Then
grd2.Row = i
grd2.Col = 0
grd2.Text = "„Ã„Ê⁄ „Ê«œ «·⁄—»Ì…"
grd2.CellBackColor = &H808080
i = i + 1
End If
End If
grd2.Row = i
grd2.Col = 0
grd2.Text = mt!mat
grd2.Col = 1
grd2.Text = mt!cof
grd2.CellBackColor = &H808080
grd2.Col = 2
grd2.Text = mt!moy
grd2.CellBackColor = &H808080
grd2.Col = 19
grd2.Text = mt!lng
tx1 = mt!lng
i = i + 1
End If
mt.MoveNext
Loop
If Label4.Caption = "«» œ«∆Ì" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = "Total MatiÈres FR"
grd2.CellBackColor = &H808080
i = i + 1
Picture3.Visible = True
Picture2.Visible = False
Else
Picture2.Visible = True
Picture3.Visible = False
End If
grd2.Rows = i
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
grd2.Row = i
grd2.Col = 19
tx = grd2.Text
If tx = "⁄—»Ì…" Or tx = "›—‰”Ì…" Then
If Label4.Caption = "«» œ«∆Ì" Then
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
Else
For k = 1 To 18
grd2.Row = i
grd2.Col = k
grd2.CellBackColor = &H808080
Next k
End If
Next i
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
If nt!sri = Text4.Text Or Val(nt!sri) = Val(Text4.Text) Then
tx1 = nt!mat
For i = 1 To n - 1
grd2.Row = i
grd2.Col = 0
tx2 = grd2.Text
If tx1 = tx2 Then
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
Label18.Caption = nt!moy
Label40.Caption = nt!moy
Label19.Caption = nt!men
Label34.Caption = nt!men
Label21.Caption = nt!tto
Label16.Caption = nt!tcf
Label19.Caption = nt!men
Label36.Caption = nt!tt1
Label33.Caption = nt!mo1
Label43.Caption = nt!tt2
Label44.Caption = nt!mo2
Label45.Caption = nt!tt3
Label46.Caption = nt!mo3
Label32.Caption = nt!tt4
Label38.Caption = nt!mo4
End If
Next i
End If
nt.MoveNext
Loop
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
grd2.Row = i
grd2.Col = 19
tx2 = grd2.Text
If tx2 = "⁄—»Ì…" Or tx2 = "›—‰”Ì…" Then
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
End If
Next i
If scm > 0 Then
moy = st / scm
MyNumber = Round(moy, 2)
moy = MyNumber
End If
Label16.Caption = scm
Label21.Caption = st
Label18.Caption = moy
If Label4.Caption = "«» œ«∆Ì" Then
Call bulletin_primaire
End If
If Val(Label16.Caption) > 0 Then
Call mention
End If
End Sub
Private Sub bulletin_primaire()
On Error Resume Next
'On Error Resume Next
Dim i As Double
Dim k As Double
Dim n As Double
Dim tx1 As String
Dim tx2 As String
Dim nex1 As Double
Dim snex1 As Double
Dim nex2 As Double
Dim snex2 As Double
Dim nex3 As Double
Dim snex3 As Double
Dim nmmt As Double
Dim snmmt1 As Double
Dim snmmt2 As Double
Dim snmmt3 As Double
Dim snmmt4 As Double
Dim nmyt As Double
Dim snmyt As Double
Dim tsnex1 As String
Dim tsnex2 As String
Dim tsnex3 As String
Dim tsnmmt As String
Dim tsnmyt As String
Dim a1 As Double
Dim a2 As Double
Dim a3 As Double
Dim a4 As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim ta1 As String
Dim ta2 As String
Dim ta3 As String
Dim ta4 As String
Dim tb As String
Dim tc As String
Dim td As String
Dim te As String
Dim m1 As Double
Dim m2 As Double
Dim m3 As Double
Dim m4 As Double
n = grd2.Rows
snex1 = 0
snex2 = 0
snex3 = 0
snmmt1 = 0
snmmt2 = 0
snmmt3 = 0
snmmt4 = 0
snmyt = 0
a1 = 0
a2 = 0
a3 = 0
a4 = 0
b = 0
c = 0
d = 0
e = 0
For i = 1 To n - 1
grd2.Row = i
grd2.Col = 0
tx1 = grd2.Text
If tx1 = "„Ã„Ê⁄ „Ê«œ «·⁄—»Ì…" Then
a1 = a1 + snmmt1
a2 = a2 + snmmt2
a3 = a3 + snmmt3
a4 = a4 + snmmt4
b = b + snex1
c = c + snex2
d = d + snex3
e = e + snmyt
tsnmmt = snmmt1
tsnex1 = snex1
tsnex1 = tsnex1 + "/" + tsnmmt
tsnmmt = snmmt2
tsnex2 = snex2
tsnex2 = tsnex2 + "/" + tsnmmt
tsnmmt = snmmt3
tsnex3 = snex3
tsnex3 = tsnex3 + "/" + tsnmmt
tsnmmt = snmmt4
tsnmyt = snmyt
tsnmyt = tsnmyt + "/" + tsnmmt
grd2.Row = i
grd2.Col = 11
grd2.Text = tsnex1
grd2.Col = 13
grd2.Text = tsnex2
grd2.Col = 15
grd2.Text = tsnex3
grd2.Col = 17
grd2.Text = tsnmyt
snex1 = 0
snex2 = 0
snex3 = 0
snmmt1 = 0
snmmt2 = 0
snmmt3 = 0
snmmt4 = 0
snmyt = 0
End If
If tx1 = "Total MatiÈres FR" Then
a1 = a1 + snmmt1
a2 = a2 + snmmt2
a3 = a3 + snmmt3
a4 = a4 + snmmt4
b = b + snex1
c = c + snex2
d = d + snex3
e = e + snmyt
tsnmmt = snmmt1
tsnex1 = snex1
tsnex1 = tsnex1 + "/" + tsnmmt
tsnmmt = snmmt2
tsnex2 = snex2
tsnex2 = tsnex2 + "/" + tsnmmt
tsnmmt = snmmt3
tsnex3 = snex3
tsnex3 = tsnex3 + "/" + tsnmmt
tsnmmt = snmmt4
tsnmyt = snmyt
tsnmyt = tsnmyt + "/" + tsnmmt
grd2.Row = i
grd2.Col = 11
grd2.Text = tsnex1
grd2.Col = 13
grd2.Text = tsnex2
grd2.Col = 15
grd2.Text = tsnex3
grd2.Col = 17
grd2.Text = tsnmyt
snex1 = 0
snex2 = 0
snex3 = 0
snmmt1 = 0
snmmt2 = 0
snmmt3 = 0
snmmt4 = 0
snmyt = 0
End If
grd2.Row = i
grd2.Col = 19
tx2 = grd2.Text
If tx2 = "⁄—»Ì…" Or tx2 = "›—‰”Ì…" Then
nex1 = 0
nex2 = 0
nex3 = 0
nmmt = 0
k = 0
grd2.Row = i
grd2.Col = 2
If Len(grd2.Text) > 0 Then
nmmt = grd2.Text
End If
grd2.Col = 11
If Len(grd2.Text) > 0 Then
nex1 = grd2.Text
snex1 = snex1 + nex1
snmmt1 = snmmt1 + nmmt
k = 1
End If
grd2.Col = 13
If Len(grd2.Text) > 0 Then
nex2 = grd2.Text
snex2 = snex2 + nex2
snmmt2 = snmmt2 + nmmt
k = 1
End If
grd2.Row = i
grd2.Col = 15
If Len(grd2.Text) > 0 Then
nex3 = grd2.Text
snex3 = snex3 + nex3
snmmt3 = snmmt3 + nmmt
k = 1
End If
grd2.Col = 17
If Len(grd2.Text) > 0 Then
nmyt = grd2.Text
snmyt = snmyt + nmyt
End If
If k = 1 Then
snmmt4 = snmmt4 + nmmt
End If
End If
Next i
ta1 = a1
ta2 = a2
ta3 = a3
ta4 = a4
tb = b
tb = tb + "/" + ta1
tc = c
tc = tc + "/" + ta2
td = d
td = td + "/" + ta3
te = e
te = te + "/" + ta4
Label36.Caption = tb
Label43.Caption = tc
Label45.Caption = td
Label32.Caption = te
m1 = 0
m2 = 0
m3 = 0
m4 = 0
If a1 > 0 Then
m1 = b / a1 * 10
MyNumber = Round(m1, 2)
m1 = MyNumber
End If
If a2 > 0 Then
m2 = c / a2 * 10
MyNumber = Round(m2, 2)
m2 = MyNumber
End If
If a3 > 0 Then
m3 = d / a3 * 10
MyNumber = Round(m3, 2)
m3 = MyNumber
End If
If a4 > 0 Then
m4 = e / a4 * 10
MyNumber = Round(m4, 2)
m4 = MyNumber
End If
Label40.Caption = (m4 * 2)
ta1 = "10"
tb = m1
tb = tb + "/" + ta1
tc = m2
tc = tc + "/" + ta1
td = m3
td = td + "/" + ta1
te = m4
te = te + "/" + ta1
Label33.Caption = tb
Label44.Caption = tc
Label46.Caption = td
Label38.Caption = te

End Sub
Private Sub mention()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim g As Double
Dim f As Double
Dim ma1 As String
Dim MA2 As String
Dim MA3 As String
Dim MA4 As String
Dim MA5 As String
Dim MA6 As String
Dim mf1 As String
Dim mf2 As String
Dim mf3 As String
Dim mf4 As String
Dim mf5 As String
Dim mf6 As String
Dim m As Double
Dim MA As String
Dim MF As String
Call cont
a = cf2!cof4
b = cf2!cof6
c = cf2!cof8
d = cf2!cof10
e = cf2!cof12
f = cf2!cof14
g = cf2!cof15
ma1 = cf2!tex9
MA2 = cf2!tex12
MA3 = cf2!tex15
MA4 = cf2!tex18
MA5 = cf2!tex19
MA6 = cf2!tex20
mf1 = cf2!tex21
mf2 = cf2!tex22
mf3 = cf2!tex23
mf4 = cf2!tex24
mf5 = cf2!tex25
mf6 = cf2!tex26
If Label4.Caption = "«» œ«∆Ì" Then
m = Label40.Caption
Else
m = Label18.Caption
End If
If m <= a And m > b Then
MA = ma1
MF = mf1
ElseIf m <= b And m > c Then
MA = MA2
MF = mf2
ElseIf m <= c And m > d Then
MA = MA3
MF = mf3
ElseIf m <= d And m > e Then
MA = MA4
MF = mf4
ElseIf m <= e And m > f Then
MA = MA5
MF = mf5
ElseIf m <= f And m >= g Then
MA = MA6
MF = mf6
End If
If Label4.Caption = "«» œ«∆Ì" Then
Label34.Caption = MF + "   " + MA
Else
Label19.Caption = MF + "   " + MA
End If
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
If ct!sri = Text4.Text Or Val(ct!sri) = Val(Text4.Text) Then
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
End If
ct.MoveNext
Loop
grd4.Rows = i
End Sub



