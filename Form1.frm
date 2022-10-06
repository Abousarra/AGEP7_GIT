VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   10050
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10050
   ScaleWidth      =   13050
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4575
      ScaleWidth      =   12375
      TabIndex        =   48
      Top             =   2760
      Width           =   12375
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
      ItemData        =   "Form1.frx":0000
      Left            =   10200
      List            =   "Form1.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   47
      Top             =   1200
      Width           =   1455
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
      ItemData        =   "Form1.frx":0029
      Left            =   10200
      List            =   "Form1.frx":0033
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   1560
      Width           =   1455
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
      TabIndex        =   45
      Text            =   "004522317"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
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
      Left            =   8760
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   2280
      Width           =   735
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   960
      ScaleHeight     =   2955
      ScaleWidth      =   6555
      TabIndex        =   32
      Top             =   3480
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1320
         TabIndex        =   38
         Text            =   "Text3"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   375
         Left            =   0
         TabIndex        =   36
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   1800
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid grd3 
         Height          =   4095
         Left            =   3000
         TabIndex        =   39
         Top             =   120
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   7223
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid grd1 
         Height          =   1095
         Left            =   0
         TabIndex        =   40
         Top             =   1920
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   1931
         _Version        =   393216
         Rows            =   1
         FixedRows       =   0
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
      Begin VB.Label Label28 
         Caption         =   "Label28"
         Height          =   375
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label40 
         Caption         =   "Label40"
         Height          =   255
         Left            =   1320
         TabIndex        =   42
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label23 
         Caption         =   "Label23"
         Height          =   255
         Left            =   1320
         TabIndex        =   41
         Top             =   1440
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   12615
      TabIndex        =   17
      Top             =   7560
      Width           =   12615
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Caption         =   "«·„⁄œ· «·⁄«„"
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
         Left            =   7080
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         Width           =   1335
      End
      Begin VB.Line Line6 
         X1              =   4920
         X2              =   4920
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line5 
         X1              =   2640
         X2              =   2640
         Y1              =   360
         Y2              =   720
      End
      Begin VB.Line Line4 
         X1              =   1320
         X2              =   6960
         Y1              =   360
         Y2              =   360
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
         Left            =   1440
         TabIndex        =   31
         Top             =   360
         Width           =   1335
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
         Left            =   2760
         TabIndex        =   30
         Top             =   360
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
         Left            =   4920
         TabIndex        =   29
         Top             =   360
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
         Left            =   6120
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
      Begin VB.Line Line3 
         X1              =   1320
         X2              =   1320
         Y1              =   -120
         Y2              =   720
      End
      Begin VB.Line Line2 
         X1              =   6960
         X2              =   6960
         Y1              =   -120
         Y2              =   720
      End
      Begin VB.Line Line1 
         X1              =   8520
         X2              =   8520
         Y1              =   -120
         Y2              =   720
      End
      Begin VB.Shape Shape1 
         Height          =   735
         Index           =   2
         Left            =   0
         Top             =   0
         Width           =   12615
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
         Left            =   1440
         TabIndex        =   27
         Top             =   0
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
         Left            =   8640
         TabIndex        =   26
         Top             =   0
         Width           =   1335
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
         Left            =   5400
         TabIndex        =   25
         Top             =   0
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
         Left            =   2400
         TabIndex        =   24
         Top             =   0
         Width           =   3735
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
         Left            =   7080
         TabIndex        =   23
         Top             =   360
         Width           =   1335
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
         Left            =   8640
         TabIndex        =   22
         Top             =   360
         Width           =   1335
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
         Left            =   11040
         TabIndex        =   21
         Top             =   360
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
         Left            =   11040
         TabIndex        =   20
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Õ›Ÿ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      ScaleHeight     =   975
      ScaleWidth      =   12615
      TabIndex        =   0
      Top             =   8400
      Width           =   12615
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   8400
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   1335
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
         TabIndex        =   16
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
         Left            =   8280
         TabIndex        =   15
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
         TabIndex        =   14
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
         Left            =   2280
         TabIndex        =   13
         Top             =   600
         Width           =   6015
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
         TabIndex        =   12
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
         Left            =   1320
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Index           =   3
         Left            =   0
         Top             =   0
         Width           =   12615
      End
      Begin VB.Line Line7 
         X1              =   9840
         X2              =   9840
         Y1              =   0
         Y2              =   960
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
         Left            =   120
         TabIndex        =   10
         Top             =   120
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.Line Line10 
         X1              =   1200
         X2              =   9840
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Õ›Ÿ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   9840
         TabIndex        =   8
         Top             =   480
         Width           =   1815
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
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
      Height          =   4575
      Left            =   240
      TabIndex        =   49
      Top             =   2760
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   1
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
   Begin MSFlexGridLib.MSFlexGrid grd4 
      Height          =   375
      Left            =   240
      TabIndex        =   50
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
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
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2160
      TabIndex        =   51
      Top             =   1320
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   2160
      TabIndex        =   52
      Top             =   1560
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Label Label35 
      Caption         =   "Label35"
      Height          =   375
      Left            =   480
      TabIndex        =   66
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”Õ»"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   8760
      TabIndex        =   65
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
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
      TabIndex        =   64
      Top             =   1560
      Width           =   1215
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
      Left            =   11160
      TabIndex        =   63
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ‘› «·œ—Ã«  "
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
      TabIndex        =   62
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label5 
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
      TabIndex        =   61
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   12615
   End
   Begin VB.Shape Shape1 
      Height          =   5295
      Index           =   1
      Left            =   120
      Top             =   2160
      Width           =   12615
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
      Left            =   7440
      TabIndex        =   60
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      Height          =   375
      Left            =   6840
      TabIndex        =   59
      Top             =   2280
      Width           =   975
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
      Left            =   5520
      TabIndex        =   58
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      Height          =   375
      Left            =   5160
      TabIndex        =   57
      Top             =   2280
      Width           =   975
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
      Left            =   3840
      TabIndex        =   56
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
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
      Height          =   375
      Left            =   3480
      TabIndex        =   55
      Top             =   2280
      Width           =   735
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
      Left            =   2160
      TabIndex        =   54
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
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
      Height          =   375
      Left            =   240
      TabIndex        =   53
      Top             =   2280
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim data As New Access.Application

Private Sub Combo1_Change()
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
Call cont
Do While Not cl.EOF
If Combo1.Text = cl!cla Then
Label28.Caption = cl!aut
Exit Sub
End If
cl.MoveNext
Loop
Else
Combo1.BackColor = &H8080FF
End If

End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Combo2_Change()
If Len(Combo2.Text) > 0 Then
Combo2.BackColor = &HC000&
Call chargcombo1
Combo1.BackColor = &H8080FF
grd1.Visible = False
Call chargegrd1_clear
grd1.Visible = True
Else
Combo2.BackColor = &H8080FF
End If

End Sub

Private Sub Combo2_Click()
Combo2_Change
End Sub

Private Sub Command1_Click()
With grd2
    Text2.Visible = False
    'Set Text2.Font = .Font
    Text2.RightToLeft = .RightToLeft
    Text2.Alignment = .CellAlignment
    Text2.Left = .Left + .ColPos(.Col) + .BorderStyle * 30
    Text2.Top = .Top + .RowPos(.Row) + .BorderStyle * 30
    Text2.Width = .ColWidth(.Col)
    Text2.Height = .RowHeight(.Row)
    Text2.Appearance = vbFlat
    Text2.Text = .Text
    Text2.Visible = True
    Text2.SetFocus
End With
End Sub


Private Sub Command3_Click()
MsgBox Chr$(8)
End Sub

Private Sub Command4_Click()
grd2.Visible = False
Call calcule_moyenne_lc
grd2.Visible = True

End Sub

Private Sub Command5_Click()
grd2.Visible = False
Call calcule_moyenne_lc
grd2.Visible = True
End Sub

Private Sub Command6_Click()
r = (27 - n)
For i = 1 To r
nt.AddNew
nt!sri = Text1.Text
nt!niv = Label4.Caption
nt!cla = Label8.Caption
nt!num = Label10.Caption
nt!nom = Label12.Caption
nt!nof = Label35.Caption
nt!mat = ""
nt!cmt = ""
nt!mmt = ""
nt!dv1 = ""
nt!dv2 = ""
nt!dv3 = ""
nt!dv4 = ""
nt!dv5 = ""
nt!dv6 = ""
nt!mdv = ""
nt!cdv = ""
nt!ex1 = ""
nt!cx1 = ""
nt!ex2 = ""
nt!cx2 = ""
nt!ex3 = ""
nt!cx3 = ""
nt!mym = ""
nt!tot = ""
nt!moy = Label40.Caption
nt!tto = ""
nt!tcf = ""
nt!men = Label34.Caption
nt!ran = Label42.Caption
nt!dat = Date
nt!Abs = ""
nt!obs = ""
nt!tt1 = Label36.Caption
nt!mo1 = Label33.Caption
nt!tt2 = Label43.Caption
nt!mo2 = Label44.Caption
nt!tt3 = Label45.Caption
nt!mo3 = Label46.Caption
nt!tt4 = Label32.Caption
nt!mo4 = Label38.Caption
nt!ncl = Label23.Caption
nt.Update
Next i

End Sub

Private Sub Command7_Click()
Text1.Text = Trim(Text1.Text)
If Text1.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· «·—ﬁ„ «· ”·”·Ì À„ ⁄—÷ «·»Ì«‰« ", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
Call cont
Do While Not et.EOF
If et!sri = Text1.Text Or Val(et!sri) = Val(Text1.Text) Then
Label4.Caption = et!niv
Label8.Caption = et!cla
Label10.Caption = et!num
Label12.Caption = et!nom
Label35.Caption = et!nfr
Label23.Caption = et!ncl
Label4.BackColor = &HC000&
Label8.BackColor = &HC000&
Label10.BackColor = &HC000&
Label12.BackColor = &HC000&
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

Private Sub Form_Load()
Me.Left = 0
Me.Top = 480
Call chargegrd2_tete
Call chargegrd1_clear
Picture2.Visible = False
Picture3.Visible = False
End Sub
Private Sub chargcombo1()
Combo1.Clear
Call cont
Do While Not cl.EOF
If Combo2.Text = cl!niv Then
Combo1.AddItem cl!cla
End If
cl.MoveNext
Loop
End Sub
Private Sub chargegrd1_clear()
grd1.Clear
grd1.Cols = 1
grd1.Rows = 2
'grd1.RowHeight(1) = 400
grd1.ColWidth(0) = 1500
grd1.ColAlignment(0) = 1
grd1.Row = 1
grd1.Col = 0
grd1.Text = "«”„ «· ·„Ì–"
grd1.Row = 0
grd1.Col = 0
grd1.Text = "«·—ﬁ„ «· ”·”·Ì"
End Sub
Private Sub chargegrd1()
Dim i As Double
i = 1
Call cont
grd1.Cols = et.RecordCount + 3
Do While Not et.EOF
If Combo1.Text = et!cla Then
grd1.ColWidth(i) = 2000
grd1.ColAlignment(i) = 1
grd1.Col = i
grd1.Row = 1
grd1.Text = et!nom
grd1.Row = 0
grd1.Text = et!sri
i = i + 1
End If
et.MoveNext
Loop
grd1.Cols = i
End Sub
Private Sub chargegrd2_tete()
Dim i As Double
Dim j As Double
grd2.Clear
grd2.Cols = 20
grd2.Rows = 1
grd2.ColWidth(0) = 3000
grd2.ColWidth(1) = 350
grd2.ColWidth(2) = 400
grd2.ColWidth(3) = 700
grd2.ColWidth(4) = 700
grd2.ColWidth(5) = 700
grd2.ColWidth(6) = 700
grd2.ColWidth(7) = 700
grd2.ColWidth(8) = 700
grd2.ColWidth(9) = 800
grd2.ColWidth(10) = 350
grd2.ColWidth(11) = 650
grd2.ColWidth(12) = 350
grd2.ColWidth(13) = 650
grd2.ColWidth(14) = 350
grd2.ColWidth(15) = 650
grd2.ColWidth(16) = 350
grd2.ColWidth(17) = 900
grd2.ColWidth(18) = 800
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
Dim i As Double
Dim tx1 As String
Dim tx2 As String
i = 1
tx1 = "⁄—»Ì…"
Call cont
grd2.Rows = mt.RecordCount + 4
Do While Not mt.EOF
If mt!cla = Label8.Caption Then
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
Private Sub chargegrd2_tete_pr()
Dim i As Double
Dim j As Double
grd2.Clear
grd2.Cols = 20
grd2.Rows = 1
grd2.ColWidth(0) = 4000
grd2.ColWidth(1) = 950
grd2.ColWidth(2) = 950
grd2.ColWidth(3) = 0
grd2.ColWidth(4) = 0
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 0
grd2.ColWidth(7) = 0
grd2.ColWidth(8) = 0
grd2.ColWidth(9) = 0
grd2.ColWidth(10) = 0
grd2.ColWidth(11) = 1350
grd2.ColWidth(12) = 550
grd2.ColWidth(13) = 1350
grd2.ColWidth(14) = 550
grd2.ColWidth(15) = 1350
grd2.ColWidth(16) = 550
grd2.ColWidth(17) = 1150
grd2.ColWidth(18) = 1000
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

Private Sub grd1_Click()
Dim i As Double
Dim j As Double
i = grd1.Row
j = grd1.Col
If j > 0 Then
grd1.Col = j
grd1.Row = 0
Text1.Text = grd1.Text
Command7_Click
End If
End Sub



Private Sub grd2_Click()
Dim i As Double
Dim j As Double
i = grd2.Row
j = grd2.Col
If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j = 7 Or j = 8 Or j = 11 Or j = 13 Or j = 15 Then
grd2.Row = i
grd2.Col = j
grd2.CellBackColor = &HFFC0C0
End If
End Sub

Private Sub grd2_KeyPress(KeyAscii As Integer)
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

Private Sub Label13_Click()
Dim i As Double
Dim n As Double
Dim r As Double
Text1.Text = Trim(Text1.Text)
n = grd2.Rows
If n = 1 Then
Exit Sub
End If
grd2.Visible = False
Call calcule_moyenne_lc
grd2.Visible = True
If Val(Label16.Caption) = 0 Then
Exit Sub
End If
Call cont
Do While Not nt.EOF
If nt!sri = Text1.Text Or Val(nt!sri) = Val(Text1.Text) Then
nt.Delete
End If
nt.MoveNext
Loop
grd2.Visible = False
Call cont
For i = 1 To n - 1
nt.AddNew
nt!sri = Text1.Text
nt!niv = Label4.Caption
nt!cla = Label8.Caption
nt!num = Label10.Caption
nt!nom = Label12.Caption
nt!nof = Label35.Caption
grd2.Row = i
grd2.Col = 0
nt!mat = grd2.Text
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
nt!moy = Label18.Caption
nt!tto = Label21.Caption
nt!tcf = Label16.Caption
nt!men = Label19.Caption
nt!ran = Label25.Caption
nt!dat = Date
nt!Abs = Label27.Caption
nt!obs = ""
nt!tt1 = ""
nt!mo1 = ""
nt!tt2 = ""
nt!mo2 = ""
nt!tt3 = ""
nt!mo3 = ""
nt!tt4 = ""
nt!mo4 = ""
nt!ncl = Label23.Caption
nt.Update
Next i
r = (19 - n)
For i = 1 To r
nt.AddNew
nt!ann = Interface.SBB1.Panels(1).Text
nt!sri = Text1.Text
nt!niv = Label4.Caption
nt!cla = Label8.Caption
nt!num = Label10.Caption
nt!nom = Label12.Caption
nt!nof = Label35.Caption
nt!mat = ""
nt!cmt = ""
nt!mmt = ""
nt!dv1 = ""
nt!dv2 = ""
nt!dv3 = ""
nt!dv4 = ""
nt!dv5 = ""
nt!dv6 = ""
nt!mdv = ""
nt!cdv = ""
nt!ex1 = ""
nt!cx1 = ""
nt!ex2 = ""
nt!cx2 = ""
nt!ex3 = ""
nt!cx3 = ""
nt!mym = ""
nt!tot = ""
nt!moy = Label40.Caption
nt!tto = ""
nt!tcf = ""
nt!men = Label34.Caption
nt!ran = Label42.Caption
nt!dat = Date
nt!Abs = ""
nt!obs = ""
nt!tt1 = Label36.Caption
nt!mo1 = Label33.Caption
nt!tt2 = Label43.Caption
nt!mo2 = Label44.Caption
nt!tt3 = Label45.Caption
nt!mo3 = Label46.Caption
nt!tt4 = Label32.Caption
nt!mo4 = Label38.Caption
nt!ncl = Label23.Caption
nt!img = App.Path & "\IMAGES\" & Text1.Text & ".jpg"
nt.Update
Next i
grd2.Visible = True
'MsgBox " „ Õ›Ÿ «·»Ì« « ", vbInformation
End Sub

Private Sub Label28_Click()
a = Val(Label10.Caption)
Call cont
data.OpenCurrentDatabase App.Path & "\2016-2017.mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
data.DoCmd.OpenReport "IM_Notes_pr", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
'data.DoCmd.OpenReport "List_Etudiants", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing

End Sub

Private Sub Label22_Click()
Dim n As Double
Dim P As Double
Dim r As Double
grd4.Row = 0
grd4.Col = 0
grd4.ColWidth(0) = 6000
grd4.Text = "«·—Ã«¡ «·«‰ Ÿ«—.... —ÌÀ„« Ì „  ÕœÌÀ ‰ «∆Ã  ·«„Ì– «·ﬁ”„ " + Combo1.Text + " Ã«—Ì «·≈Ã—«¡...‹"
P = 0
Call cont2
n = et2.RecordCount
Do While Not et2.EOF
If Combo1.Text = et2!cla Then
Text1.Text = et2!sri
Command7_Click
End If
If Combo2.Text = "«» œ«∆Ì" Then
Label29_Click
Else
Label13_Click
End If
P = P + 1
r = (P * 100 / n)
ProgressBar2 = r
et2.MoveNext
Loop
a = Val(Label28.Caption)
Call cont
data.OpenCurrentDatabase App.Path & "\" & Interface.SBB1.Panels(1).Text & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
data.DoCmd.OpenReport "IM_Notes_pr", acViewPreview, , "ncl =" & a, acWindowNormal, OpenArgs
'data.DoCmd.OpenReport "List_Etudiants", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing

End Sub

Private Sub Label29_Click()
Dim i As Double
Dim n As Double
Dim r As Double
Text1.Text = Trim(Text1.Text)
n = grd2.Rows
If n = 1 Then
Exit Sub
End If
grd2.Visible = False
Call calcule_moyenne_lc
grd2.Visible = True
Call cont
Do While Not nt.EOF
If nt!sri = Text1.Text Or Val(nt!sri) = Val(Text1.Text) Then
nt.Delete
End If
nt.MoveNext
Loop
grd2.Visible = False
Call cont
For i = 1 To n - 1
nt.AddNew
nt!ann = Interface.SBB1.Panels(1).Text
nt!sri = Text1.Text
nt!niv = Label4.Caption
nt!cla = Label8.Caption
nt!num = Label10.Caption
nt!nom = Label12.Caption
nt!nof = Label35.Caption
grd2.Row = i
grd2.Col = 0
nt!mat = grd2.Text
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
nt!moy = Label40.Caption
nt!tto = ""
nt!tcf = ""
nt!men = Label34.Caption
nt!ran = Label42.Caption
nt!dat = Date
nt!Abs = ""
nt!obs = ""
nt!tt1 = Label36.Caption
nt!mo1 = Label33.Caption
nt!tt2 = Label43.Caption
nt!mo2 = Label44.Caption
nt!tt3 = Label45.Caption
nt!mo3 = Label46.Caption
nt!tt4 = Label32.Caption
nt!mo4 = Label38.Caption
nt!ncl = Label23.Caption
nt!img = App.Path & "\IMAGES\" & Text1.Text & ".jpg"
nt.Update
Next i
r = (19 - n)
For i = 1 To r
nt.AddNew
nt!ann = Interface.SBB1.Panels(1).Text
nt!sri = Text1.Text
nt!niv = Label4.Caption
nt!cla = Label8.Caption
nt!num = Label10.Caption
nt!nom = Label12.Caption
nt!nof = Label35.Caption
nt!mat = ""
nt!cmt = ""
nt!mmt = ""
nt!dv1 = ""
nt!dv2 = ""
nt!dv3 = ""
nt!dv4 = ""
nt!dv5 = ""
nt!dv6 = ""
nt!mdv = ""
nt!cdv = ""
nt!ex1 = ""
nt!cx1 = ""
nt!ex2 = ""
nt!cx2 = ""
nt!ex3 = ""
nt!cx3 = ""
nt!mym = ""
nt!tot = ""
nt!moy = Label40.Caption
nt!tto = ""
nt!tcf = ""
nt!men = Label34.Caption
nt!ran = Label42.Caption
nt!dat = Date
nt!Abs = ""
nt!obs = ""
nt!tt1 = Label36.Caption
nt!mo1 = Label33.Caption
nt!tt2 = Label43.Caption
nt!mo2 = Label44.Caption
nt!tt3 = Label45.Caption
nt!mo3 = Label46.Caption
nt!tt4 = Label32.Caption
nt!mo4 = Label38.Caption
nt!ncl = Label23.Caption
nt!img = App.Path & "\IMAGES\" & Text1.Text & ".jpg"
nt.Update
Next i
grd2.Visible = True
'MsgBox " „ Õ›Ÿ «·»Ì« « ", vbInformation

End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Else
Text1.BackColor = &H8080FF
End If
grd2.Visible = False
Call chargegrd2_tete
Picture2.Visible = False
Picture3.Visible = False
Label40.Caption = ""
Label4.Caption = ""
Label8.Caption = ""
Label10.Caption = ""
Label12.Caption = ""
Label35.Caption = ""
Label21.Caption = "0"
Label16.Caption = "0"
Label18.Caption = "0"
Label19.Caption = ""
Label25.Caption = ""
Label27.Caption = ""
Label36.Caption = ""
Label43.Caption = ""
Label45.Caption = ""
Label32.Caption = ""
Label33.Caption = ""
Label44.Caption = ""
Label46.Caption = ""
Label38.Caption = ""
Label34.Caption = ""
Label42.Caption = ""
Label4.BackColor = &HFF8080
Label8.BackColor = &HFF8080
Label10.BackColor = &HFF8080
Label12.BackColor = &HFF8080
grd2.Visible = True
End Sub

Private Sub Text1_Click()
Text1_Change
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Text1.Text <> "" Then
If KeyCode = 13 Then
Command7_Click
End If
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii

End Sub
Public Sub calcule_moyenne_lc()
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
Private Sub coff_dv_ex()
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
grd2.Text = cf!cof16
grd2.CellBackColor = &H80C0FF
grd2.Col = 14
grd2.Text = cf!cof17
grd2.CellBackColor = &H80C0FF
grd2.Col = 16
grd2.Text = cf!cof18
grd2.CellBackColor = &H80C0FF
Else
grd2.Row = i
grd2.Col = 10
grd2.Text = cf!cof0
grd2.CellBackColor = &H80C0FF
grd2.Col = 12
grd2.Text = cf!cof1
grd2.CellBackColor = &H80C0FF
grd2.Col = 14
grd2.Text = cf!cof2
grd2.CellBackColor = &H80C0FF
grd2.Col = 16
grd2.Text = cf!cof3
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
Dim i As Double
Dim n As Double
Dim tx1 As String
Dim tx2 As String
n = grd2.Rows
Call cont
Do While Not nt.EOF
If nt!sri = Text1.Text Or Val(nt!sri) = Val(Text1.Text) Then
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
Private Sub mention()
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
a = cf!cof4
b = cf!cof6
c = cf!cof8
d = cf!cof10
e = cf!cof12
f = cf!cof14
g = cf!cof15
ma1 = cf!tex9
MA2 = cf!tex12
MA3 = cf!tex15
MA4 = cf!tex18
MA5 = cf!tex19
MA6 = cf!tex20
mf1 = cf!tex21
mf2 = cf!tex22
mf3 = cf!tex23
mf4 = cf!tex24
mf5 = cf!tex25
mf6 = cf!tex26
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
Private Sub rang()
Dim j As Double
Dim i As Double
Dim tx As String
i = 1
grd1.Cols = 3
Call cont
grd1.Rows = rn.RecordCount + 3
Do While Not rn.EOF
If Etudiants.Label6.Caption = rn!cla Then
grd1.Row = i
grd1.Col = 0
grd1.Text = rn!cod
grd1.Col = 1
grd1.Text = rn!moy
grd1.Col = 2
grd1.Text = rn!ran
i = i + 1
End If
rn.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 2
For j = 1 To i - 1
grd1.Row = j
grd1.Col = 0
tx = grd1.Text
If Etudiants.Text2.Text = tx Or Val(Etudiants.Text2.Text) = Val(tx) Then
Label18.Caption = j
Exit Sub
End If
Next j

End Sub
Private Sub bulletin_primaire()
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



