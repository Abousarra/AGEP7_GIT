VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Caisse_PRT 
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
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   10440
      ScaleHeight     =   8745
      ScaleWidth      =   2385
      TabIndex        =   40
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
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
         TabIndex        =   41
         Top             =   8400
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid grd6 
         Height          =   8295
         Left            =   0
         TabIndex        =   42
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   -8640
      ScaleHeight     =   8175
      ScaleWidth      =   10215
      TabIndex        =   30
      Top             =   2760
      Width           =   10215
   End
   Begin VB.CommandButton Command8 
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
      TabIndex        =   27
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command9 
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
      Left            =   1080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   10440
      ScaleHeight     =   8745
      ScaleWidth      =   2385
      TabIndex        =   24
      Top             =   720
      Width           =   2415
      Begin ComctlLib.TreeView TreeView1 
         Height          =   8775
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   15478
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
      Left            =   6480
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   780
      Width           =   855
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
      Height          =   375
      Left            =   7440
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   780
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
      ItemData        =   "Caisse_PRT.frx":0000
      Left            =   4800
      List            =   "Caisse_PRT.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1440
      Width           =   2055
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
      Height          =   375
      Left            =   2520
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1440
      Width           =   1518
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "”Õ» "
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
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   735
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
      Height          =   345
      Left            =   1080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox Picture3 
      Height          =   5535
      Left            =   3000
      ScaleHeight     =   5475
      ScaleWidth      =   6555
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   6615
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   4200
         TabIndex        =   45
         Text            =   "Text6"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   4200
         TabIndex        =   44
         Text            =   "Text5"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   3720
         TabIndex        =   39
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text3 
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
         Left            =   1920
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   600
         Width           =   2055
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H000000FF&
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Text            =   "Text4"
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   255
         Left            =   480
         TabIndex        =   46
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   -360
         TabIndex        =   38
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„œÌ‰ »‹‹."
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
         TabIndex        =   37
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   0
         TabIndex        =   36
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»«·€ «·«Ìœ«⁄"
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
         Left            =   1440
         TabIndex        =   35
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
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
         Height          =   375
         Left            =   -120
         TabIndex        =   34
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»«·€ «·”Õ»"
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
         Left            =   1560
         TabIndex        =   33
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "0"
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ›«’Ì·"
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
         Left            =   1320
         TabIndex        =   29
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CheckBox Check13 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "⁄—÷ «·ﬂ·"
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
      Left            =   9120
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   345
      Left            =   8040
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16744576
      CalendarTitleBackColor=   16711680
      CalendarTrailingForeColor=   16744576
      Format          =   107872257
      CurrentDate     =   42638
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   6735
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11880
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
   Begin MSComCtl2.DTPicker DT2 
      Height          =   345
      Left            =   4440
      TabIndex        =   13
      Top             =   2040
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
      Format          =   107872257
      CurrentDate     =   42638
   End
   Begin MSComCtl2.DTPicker DT3 
      Height          =   345
      Left            =   1920
      TabIndex        =   14
      Top             =   2040
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
      Format          =   107872257
      CurrentDate     =   42638
   End
   Begin MSFlexGridLib.MSFlexGrid grd5 
      Height          =   615
      Left            =   10560
      TabIndex        =   43
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·≈”„"
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
      Left            =   5040
      TabIndex        =   23
      Top             =   780
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   10215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   2400
      TabIndex        =   22
      Top             =   780
      Width           =   3495
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
      Left            =   8640
      TabIndex        =   21
      Top             =   840
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Index           =   4
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   10215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„»·€"
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
      TabIndex        =   20
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”Õ»/ «Ìœ«⁄"
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
      Left            =   6600
      TabIndex        =   19
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«· «—ÌŒ"
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
      TabIndex        =   18
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   4680
      Y1              =   1320
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2400
      Y1              =   1320
      Y2              =   1920
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï  «—ÌŒ"
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
      Left            =   3120
      TabIndex        =   17
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "⁄—÷ ⁄„·Ì«  «·”Õ» Ê«·«Ìœ«⁄ „‰  «—ÌŒ"
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
      Left            =   5400
      TabIndex        =   16
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      Height          =   6975
      Index           =   3
      Left            =   120
      Top             =   2520
      Width           =   10215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "’‰œÊﬁ «·‘—ﬂ«¡"
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
      TabIndex        =   15
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "Caisse_PRT"
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
Dim rClientRect As RECT
Dim ReturnStyle As Long
ReturnStyle = GetWindowLong(TreeView1.hWnd, GWL_EXSTYLE)
SetWindowLong TreeView1.hWnd, GWL_EXSTYLE, ReturnStyle Or WS_EX_LAYOUTRTL
GetClientRect TreeView1.hWnd, rClientRect
InvalidateRect TreeView1.hWnd, rClientRect, True
End Sub
Private Sub couleur_treeview1()
Dim lngStyle As Long
Call SendMessage(TreeView1.hWnd, TVM_SETBKCOLOR, 0, ByVal RGB(250, 247, 13))    'Change the background 'color to red.
    ' Now reset the style so that the tree lines appear properly
    lngStyle = GetWindowLong(TreeView1.hWnd, GWL_STYLE)
    Call SetWindowLong(TreeView1.hWnd, GWL_STYLE, lngStyle - TVS_HASLINES)
    Call SetWindowLong(TreeView1.hWnd, GWL_STYLE, lngStyle)
TreeView1.Sorted = True
End Sub

Private Sub Check13_Click()
If Check13.Value = 0 Then
grd2.Visible = False
Call chargegrd_clear
grd2.Visible = True
Else
grd2.Visible = False
Call chargegrd2_T
grd2.Visible = True
End If
End Sub

Private Sub Combo1_Change()
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
Text2.SetFocus
Else
Combo1.BackColor = &H8080FF
End If

End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Command1_Click()
Label16.Caption = "rrrrrrrrr"
End Sub

Private Sub Command2_Click()
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
Call cont
Do While Not pr.EOF
If pr!sri = Text1.Text Or Val(pr!sri) = Val(Text1.Text) Then
If pr!act = "1" Then
Label1.Caption = pr!nom
Check13.Value = 0
grd2.Visible = False
Call chargegrd2
grd2.Visible = True
Picture1.Visible = False
Exit Sub
Else
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· ·‘—Ìﬂ  „ Õ–›Â", vbCritical + arabic
Exit Sub
End If
End If
pr.MoveNext
Loop
Call cont
Do While Not sr.EOF
If sr!sri = Text1.Text Or Val(sr!sri) = Val(Text1.Text) Then
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· ·Ì” —ﬁ„  ”·”·Ì ·‘—Ìﬂ Ê≈‰„« —ﬁ„  ”·”·Ì ·" + sr!eta, vbExclamation
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

Private Sub Command6_Click()
Picture4.Visible = False

End Sub

Private Sub Command7_Click()
Check13.Value = 0
grd2.Visible = False
Call chargegrd2
grd2.Visible = True
End Sub

Private Sub Command8_Click()
Text2.Text = ""
Text3.Text = ""
Text2.SetFocus
DT1.Value = Date
Label17.Caption = ""
Label11.Caption = ""
Label16.Caption = "0"
Check13.Value = 0
grd2.Visible = False
Call chargegrd2
grd2.Visible = True
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False
Call Operations

End Sub

Private Sub Command9_Click()
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As String
Dim x$
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
If Text1.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· «·—ﬁ„ «· ”·”·Ì À„ ⁄—÷ «·»Ì«‰« ", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
If Label1.Caption = "" Then
MsgBox "«·—Ã«¡ «·÷€ÿ ⁄·Ï “— ⁄—÷ √Ê «·÷€ÿ ⁄·Ï ENTER", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Or Combo1.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text2.BackColor = &H8080FF Then
Text2.SetFocus
End If
Exit Sub
End If
'** controle caisse
Call cont
a = eb!cca
b = Text2.Text
c = 0
If Label11.Caption = "”Õ»" Then
c = Label16.Caption
End If
If Label11.Caption = "«Ìœ«⁄" Then
c = Label16.Caption
c = c * (-1)
End If
If Combo1.Text = "”Õ»" Then
d = (a + c) - b
Else
d = (a + c) + b
End If
If d < 0 Then
MsgBox "—’Ìœ «·’‰œÊﬁ €Ì— ﬂ«› ·≈ „«„ «·⁄„·Ì…", vbExclamation
Exit Sub
End If
eb!cca = d
eb.Update
'*******
'**** archive de caisse ajou et modif
Adat = Date
Aheu = Time$
If Label17.Caption = "" Then
Atyp = "≈÷«›…"
Else
Atyp = " ⁄œÌ·"
End If
Adet = Combo1.Text + " „‰ ÿ—› «·‘—Ìﬂ ’«Õ» «·—ﬁ„ «· ”·”·Ì " + Text1.Text
Amon = Text2.Text
Acom = "”Ã· «·‘—ﬂ«¡"
Auti = directions.Label2.Caption
'****************************************
If Label17.Caption <> "" Then
Call cont
Do While Not cp.EOF
If Label17.Caption = cp!aut Then
cp!sri = Text1.Text
cp!nom = Label1.Caption
cp!dat = DT1.Value
cp!typ = Combo1.Text
cp!mon = Text2.Text
cp!det = Text3.Text
cp!heu = Time$
If cp!act = "2" Then
cp!act = "3"
End If
cp.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
cp.MoveNext
Loop
End If
cp.AddNew
cp!sri = Text1.Text
cp!nom = Label1.Caption
cp!dat = DT1.Value
cp!typ = Combo1.Text
cp!mon = Text2.Text
cp!det = Text3.Text
cp!heu = Time$
cp!act = "0"
cp!mtf = ""
cp.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
'Call chargegrd1
Call MakeTreeViewRTL
Call chargetreeview1
Call couleur_treeview1
DT1.Value = Date
DT2.Value = Date
DT3.Value = Date
Call Operations
End Sub
Private Sub chargegrd2()
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd2.Clear
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 2000
grd2.ColWidth(2) = 1400
grd2.ColWidth(3) = 2000
grd2.ColWidth(4) = 2700
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 800
grd2.ColWidth(7) = 800
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 3
grd2.ColAlignment(7) = 3
grd2.Row = 0
grd2.Col = 1
grd2.Text = "«· «—ÌŒ"
grd2.Col = 2
grd2.Text = "«·”«⁄…"
grd2.Col = 3
grd2.Text = "«·‰Ê⁄"
grd2.Col = 4
grd2.Text = "«·„»·€"
grd2.Col = 5
grd2.Text = "«· ›«’Ì·"
i = 1
dat1 = DT2.Value
dat2 = DT3.Value
P = 0
r = 0
s = 0
Call cont
grd2.Rows = cp.RecordCount + 3
Do While Not cp.EOF
If Text1.Text = cp!sri Or Val(Text1.Text) = Val(cp!sri) Then
dat3 = cp!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If cp!act <> "1" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = cp!aut
grd2.Col = 1
grd2.Text = cp!dat
If cp!act = "2" Then
grd2.CellBackColor = &HFF&
End If
grd2.Col = 2
grd2.Text = cp!heu
grd2.Col = 3
grd2.Text = cp!typ
End If
If cp!typ = "”Õ»" Then
a = cp!mon
P = P + a
Else
a = cp!mon
r = r + a
End If
If cp!act <> "1" Then
grd2.Col = 4
grd2.Text = cp!mon
grd2.Col = 5
grd2.Text = cp!det
grd2.Col = 6
grd2.Text = " ⁄œÌ·"
grd2.CellBackColor = &HFFFF&
grd2.Col = 7
grd2.Text = "Õ–›"
grd2.CellBackColor = &HFF&
i = i + 1
End If
End If
End If
cp.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 2
s = (P - r)
Label7.Caption = P
Label10.Caption = r
If s > 0 Then
Label12.Caption = "„œÌ‰ »‹ "
Label15.Caption = s
Else
Label12.Caption = "œ«∆‰ »‹ "
Label15.Caption = s * -1
End If
If s = 0 Then
Label15.Caption = "·« „œÌ‰ Ê·« œ«∆‰"
Label12.Caption = ""
End If
End Sub
Private Sub chargegrd2_T()
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd2.Clear
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 2000
grd2.ColWidth(2) = 1400
grd2.ColWidth(3) = 2000
grd2.ColWidth(4) = 2700
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 800
grd2.ColWidth(7) = 800
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 3
grd2.ColAlignment(7) = 3
grd2.Row = 0
grd2.Col = 1
grd2.Text = "«· «—ÌŒ"
grd2.Col = 2
grd2.Text = "«·”«⁄…"
grd2.Col = 3
grd2.Text = "«·‰Ê⁄"
grd2.Col = 4
grd2.Text = "«·„»·€"
grd2.Col = 5
grd2.Text = "«· ›«’Ì·"
i = 1
dat1 = DT2.Value
dat2 = DT3.Value
P = 0
r = 0
s = 0
Call cont
grd2.Rows = cp.RecordCount + 3
Do While Not cp.EOF
If Text1.Text = cp!sri Or Val(Text1.Text) = Val(cp!sri) Then
dat3 = cp!dat
If cp!act <> "1" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = cp!aut
grd2.Col = 1
grd2.Text = cp!dat
If cp!act = "2" Then
grd2.CellBackColor = &HFF&
End If
grd2.Col = 2
grd2.Text = cp!heu
grd2.Col = 3
grd2.Text = cp!typ
End If
If cp!typ = "”Õ»" Then
a = cp!mon
P = P + a
Else
a = cp!mon
r = r + a
End If
If cp!act <> "1" Then
grd2.Col = 4
grd2.Text = cp!mon
grd2.Col = 5
grd2.Text = cp!det
grd2.Col = 6
grd2.Text = " ⁄œÌ·"
grd2.CellBackColor = &HFFFF&
grd2.Col = 7
grd2.Text = "Õ–›"
grd2.CellBackColor = &HFF&
i = i + 1
End If
End If
cp.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 2
s = (P - r)
Label7.Caption = P
Label10.Caption = r
If s > 0 Then
Label12.Caption = "„œÌ‰ »‹ "
Label15.Caption = s
Else
Label12.Caption = "œ«∆‰ »‹ "
Label15.Caption = s * -1
End If
If s = 0 Then
Label15.Caption = "·« „œÌ‰ Ê·« œ«∆‰"
Label12.Caption = ""
End If
End Sub

Private Sub grd2_Click()
Dim i As Double
Dim j As Double
Dim a As Double
Dim b As Double
Dim tx1 As Double
i = grd2.Row
j = grd2.Col
If i > 0 Then
If j = 1 Then
grd2.Row = i
grd2.Col = 1
If grd2.CellBackColor = &HFF& Then
grd2.Row = i
grd2.Col = 0
tx1 = grd2.Text
Call cont
Do While Not cp.EOF
If cp!aut = tx1 Then
MsgBox cp!mtf
Exit Sub
End If
cp.MoveNext
Loop
End If
End If
If j = 6 Then
grd2.Row = i
grd2.Col = 0
Label17.Caption = grd2.Text
grd2.Col = 1
DT1.Value = grd2.Text
grd2.Col = 3
Combo1.Text = grd2.Text
Label11.Caption = grd2.Text
grd2.Col = 4
Text2.Text = grd2.Text
Label16.Caption = grd2.Text
grd2.Col = 5
Text3.Text = grd2.Text
End If
If j = 7 Then
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–Â «·⁄„·Ì…", vbInformation + vbYesNo + arabic, "AGEP7")
If g = vbYes Then
a = eb!cca
grd2.Row = i
grd2.Col = 0
Label17.Caption = grd2.Text
grd2.Col = 3
Label11.Caption = grd2.Text
grd2.Col = 4
b = grd2.Text
If Label11.Caption = "«Ìœ«⁄" Then
'*** controle caisse
If b > a Then
MsgBox "—’Ìœ «·’‰œÊﬁ ·« Ì”„Õ »« „«„ «·⁄„·Ì…... Ì—ÃÏ ÷Œ „»·€ ÃœÌœ ›Ì «·’‰œÊﬁ", vbExclamation
Label17.Caption = ""
Label11.Caption = ""
Exit Sub
End If
'***
b = -b
End If
a = a + b
eb!cca = a
eb.Update
Call cont
Do While Not cp.EOF
If Label17.Caption = cp!aut Then
'**** archive de caisse supp
Label18.Caption = cp!sri
If b < 0 Then
b = -b
End If
Adat = Date
Aheu = Time$
Atyp = "Õ–›"
Adet = Label11.Caption + " „‰ ÿ—› «·‘—Ìﬂ ’«Õ» «·—ﬁ„ «· ”·”·Ì " + Label18.Caption
Amon = b
Acom = "”Ã· «·‘—ﬂ«¡"
Auti = directions.Label2.Caption
'****************************************
cp.Delete
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
cp.MoveNext
Loop
End If
End If
End If

End Sub


Private Sub grd5_Click()
Dim i As Double
Dim j As Double
Dim k As Double
k = 0
i = grd5.Row
If i = 0 Then
j = 0
k = 3
Else
j = 2
k = 2
End If
If i = 0 Or i = 1 Then
If grd5.RowHeight(i) > 0 Then
Text6.Text = j
Text5.Text = k
grd6.Visible = False
Call chargegrd6
grd6.Visible = True
Picture4.Visible = True
End If
End If

End Sub

Private Sub grd6_Click()
Dim i As Double
Dim j As Double
i = grd6.Row
j = grd6.Col
If i > 0 Then
grd6.Row = i
grd6.Col = 0
DT2.Value = grd6.Text
DT3.Value = grd6.Text
grd6.Col = 1
Text1.Text = grd6.Text
Command2_Click
Command7_Click
Picture4.Visible = False

End If

End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Check13.Value = 0
Text2.Text = ""
Text3.Text = ""
DT1.Value = Date
Label17.Caption = ""
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False
Picture1.Visible = True
Else
Text1.BackColor = &H8080FF
End If

End Sub

Private Sub Text1_Click()
Text1_Change
End Sub

Private Sub Text2_Change()
If Len(Text2.Text) > 0 Then
Text2.BackColor = &HC000&
Else
Text2.BackColor = &H8080FF
End If

End Sub

Private Sub Text2_Click()
Text2_Change
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
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
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation + arabic
Command8_Click
Call archive_caisse
End If

End Sub
Private Sub chargegrd_clear()
grd2.Clear
grd2.Cols = 5
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1500
grd2.ColWidth(2) = 1500
grd2.ColWidth(3) = 2000
grd2.ColWidth(4) = 4500
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.Row = 0
grd2.Col = 1
grd2.Text = "«· «—ÌŒ"
grd2.Col = 2
grd2.Text = "«·‰Ê⁄"
grd2.Col = 3
grd2.Text = "«·„»·€"
grd2.Col = 4
grd2.Text = "«· ›«’Ì·"
End Sub
Private Sub chargetreeview1()
Dim id1 As String
Dim id2 As String
Dim i As Double
Dim n As Double
TreeView1.Nodes.Clear
'TreeView1.Nodes.Add , , "PR", "√”„«¡ «·‘—ﬂ«¡"
Call cont
Do While Not pr.EOF
If pr!act = "1" Then
id1 = pr!sri
id2 = "M" + id1
TreeView1.Nodes.Add , tvwChild, id2, pr!nom
End If
pr.MoveNext
Loop
End Sub
Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
Dim n As Double
Text4.Text = Node.Key
n = Len(Text4.Text)
If n > 2 Then
n = (n - 1)
vg = Mid$(Text4.Text, 2, n)
Text1.Text = vg
Command2_Click
End If
End Sub
Private Sub Operations()
Dim a As Double
Dim b As Double
grd5.Rows = 2
grd5.Cols = 2
grd5.ColWidth(0) = 1100
grd5.ColWidth(1) = 900
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 3
grd5.Col = 0
grd5.Row = 0
grd5.Text = "⁄„·Ì… ÃœÌœ…"
grd5.Row = 1
grd5.Text = "⁄„·Ì… „—›Ê÷…"
a = 0
b = 0
Call cont
Do While Not cp.EOF
If cp!act = "0" Or cp!act = "3" Then
a = a + 1
End If
If cp!act = "2" Then
b = b + 1
End If
cp.MoveNext
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
Dim i As Double
Dim tx As String
grd6.Clear
grd6.Rows = 1
grd6.Cols = 2
grd6.ColWidth(0) = 1150
grd6.ColWidth(1) = 1100
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.Row = 0
grd6.Col = 0
grd6.Text = " «—ÌŒ «·⁄„·Ì…"
grd6.Col = 1
grd6.Text = "«·—ﬁ„ «· ”·”·Ì"
If Text6.Text = "2" Then
tx = "„—›Ê÷…"
Else
tx = "ÃœÌœ…"
End If
i = 1
Call cont
grd6.Rows = cp.RecordCount + 3
Do While Not cp.EOF
If cp!act = Text6.Text Or cp!act = Text5.Text Then
grd6.Row = i
grd6.Col = 0
grd6.Text = cp!dat
grd6.Col = 1
grd6.Text = cp!sri
If Text6.Text = "0" Then
grd6.CellBackColor = &HFFFF&
Else
grd6.CellBackColor = &HFF&
End If
i = i + 1
End If
cp.MoveNext
Loop
grd6.Rows = i
End Sub
