VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Caisse_FNC 
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
      Left            =   10320
      ScaleHeight     =   8745
      ScaleWidth      =   2505
      TabIndex        =   43
      Top             =   720
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
         TabIndex        =   44
         Top             =   8400
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid grd6 
         Height          =   8295
         Left            =   0
         TabIndex        =   45
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
      Height          =   8055
      Left            =   -7680
      ScaleHeight     =   8055
      ScaleWidth      =   10215
      TabIndex        =   35
      Top             =   6000
      Width           =   10215
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
      TabIndex        =   34
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   1215
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
      TabIndex        =   33
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   10320
      ScaleHeight     =   8625
      ScaleWidth      =   2505
      TabIndex        =   29
      Top             =   720
      Width           =   2535
      Begin ComctlLib.TreeView TreeView1 
         Height          =   8775
         Left            =   0
         TabIndex        =   30
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
      Height          =   345
      Left            =   6480
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   12
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
      TabIndex        =   11
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
      ItemData        =   "Caisse_FNC.frx":0000
      Left            =   6360
      List            =   "Caisse_FNC.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
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
      Left            =   2520
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "”Õ» "
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
      TabIndex        =   8
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton Command7 
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
      Left            =   1080
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2040
      Width           =   735
   End
   Begin VB.PictureBox Picture3 
      Height          =   4095
      Left            =   3480
      ScaleHeight     =   4035
      ScaleWidth      =   4515
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   " ”œÌœ —« »"
         Height          =   495
         Left            =   1320
         TabIndex        =   50
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2880
         TabIndex        =   48
         Text            =   "Text6"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2760
         TabIndex        =   47
         Text            =   "Text5"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
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
         Left            =   360
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   360
         Width           =   1815
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H0000FF00&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Text            =   "Text4"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Label19"
         Height          =   255
         Left            =   600
         TabIndex        =   49
         Top             =   2520
         Width           =   1335
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
         Left            =   960
         TabIndex        =   42
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ «·Õ”«»"
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
         TabIndex        =   41
         Top             =   2160
         Width           =   1455
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
         Left            =   1200
         TabIndex        =   40
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„»«·€ «·„”ÕÊ»…"
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
         TabIndex        =   39
         Top             =   1800
         Width           =   1455
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
         Left            =   1800
         TabIndex        =   38
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„” Õﬁ« "
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
         Left            =   2880
         TabIndex        =   37
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ›«’Ì·"
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
         Left            =   1440
         TabIndex        =   32
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label20 
         Caption         =   " ”œÌœ —« » ‘Â— 7"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label22 
         Caption         =   "0"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
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
      Left            =   9000
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
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
      ForeColor       =   &H00000000&
      Height          =   375
      ItemData        =   "Caisse_FNC.frx":0024
      Left            =   5160
      List            =   "Caisse_FNC.frx":004C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   345
      Left            =   8160
      TabIndex        =   13
      Top             =   1440
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
      Format          =   99352577
      CurrentDate     =   42638
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grd2 
      Height          =   6375
      Left            =   240
      TabIndex        =   15
      Top             =   2520
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   11245
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
      Left            =   4320
      TabIndex        =   16
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
      Format          =   99352577
      CurrentDate     =   42638
   End
   Begin MSComCtl2.DTPicker DT3 
      Height          =   345
      Left            =   1920
      TabIndex        =   17
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
      Format          =   99352577
      CurrentDate     =   42638
   End
   Begin MSFlexGridLib.MSFlexGrid grd5 
      Height          =   615
      Left            =   10440
      TabIndex        =   46
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
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   8880
      Top             =   1920
      Width           =   1455
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
      TabIndex        =   28
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
      Left            =   2520
      TabIndex        =   27
      Top             =   780
      Width           =   3375
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
      TabIndex        =   26
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
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   25
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·Õ«·…"
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
      Left            =   6840
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   5040
      X2              =   5040
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
      TabIndex        =   22
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
      Left            =   5280
      TabIndex        =   21
      Top             =   2040
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      Height          =   7095
      Index           =   3
      Left            =   120
      Top             =   1920
      Width           =   10215
   End
   Begin VB.Shape Shape2 
      Height          =   375
      Left            =   120
      Top             =   9000
      Width           =   10215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "’‰œÊﬁ «·„ÊŸ›Ì‰"
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
      Left            =   5280
      TabIndex        =   20
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘Â—"
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
      Left            =   5760
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      Height          =   345
      Left            =   2520
      TabIndex        =   18
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "Caisse_FNC"
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
If Combo1.Text = " √ﬂÌœ —« »" Then
Label16.Visible = True
Combo2.Visible = True
Text2.Text = Label18.Caption
Text2.Visible = False
Else
Label16.Visible = False
Combo2.Visible = False
Text2.Visible = True
Text2.Text = ""
Text2.SetFocus
End If
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
'Text3.SetFocus
Else
Combo2.BackColor = &H8080FF
End If

End Sub

Private Sub Combo2_Click()
Combo2_Change
End Sub


Private Sub Command1_Click()
Call cont2
Do While Not fc5.EOF
Text1.Text = fc5!sri
Command2_Click
Call cont3
Do While Not pf3.EOF
If pf3!cas = " ”œÌœ —« »" And Label1.Caption = pf3!nom Then
DT1.Value = pf3!dat
Combo1.Text = " √ﬂÌœ —« »"
Combo2.Text = pf3!moi
Command9_Click
End If
pf3.MoveNext
Loop
fc5.MoveNext
Loop
MsgBox "OK...Reste", vbInformation
Call cont2
Do While Not fc5.EOF
Text1.Text = fc5!sri
Command2_Click
Call cont3
Do While Not pf3.EOF
If pf3!cas = "œ›⁄ „»·€" And Label1.Caption = pf3!nom Then
DT1.Value = pf3!dat
Combo1.Text = "”Õ» „»·€"
Text2.Text = pf3!mon
Command9_Click
End If
pf3.MoveNext
Loop
fc5.MoveNext
Loop
MsgBox "OK", vbInformation

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
Do While Not fc.EOF
If fc!sri = Text1.Text Or Val(fc!sri) = Val(Text1.Text) Then
If fc!act = "1" Then
Label1.Caption = fc!nom
Label18.Caption = fc!sal
Check13.Value = 0
grd2.Visible = False
Call chargegrd2
grd2.Visible = True
Picture1.Visible = False
Exit Sub
Else
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· ·„ÊŸ›  „ Õ–›Â", vbCritical + arabic
Exit Sub
End If
End If
fc.MoveNext
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

Private Sub Command4_Click()

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
Text3.Text = ""
If Text2.Visible = True Then
Text2.Text = ""
Text2.SetFocus
End If
DT1.Value = Date
Label17.Caption = ""
Label22.Caption = "0"
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
If Combo2.Text = "" And Combo2.Visible = True Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
Exit Sub
End If
b = 0
If Combo2.Visible = False Then
a = Text2.Text
b = Label22.Caption
c = Label15.Caption
d = ((c + b) - a)
If d < 0 Then
MsgBox "·« Ì„ﬂ‰ ”Õ» Â–« «·„»·€, ·√‰ —’Ìœ «·Õ”«» ·« Ì”„Õ »–·ﬂ Ì—ÃÏ  √ﬂÌœ —« » ‘Â— ÃœÌœ", vbExclamation
Exit Sub
End If
End If
If Combo2.Visible = True Then
a = Text2.Text
b = Label22.Caption
c = Label10.Caption
d = Label7.Caption
c = ((c - b) + a)
If d > c Then
MsgBox "·« Ì„ﬂ‰ «Ã—«¡ Â–Â «·⁄„·Ì…, ·√‰Â ›Ì Â–Â «·Õ«·… ” ﬂÊ‰ «·„»«·€ «·„”ÕÊ»… √ﬂ»— „‰ «·„” Õﬁ« , ÌÃ» √Ê·« Õ–› »⁄÷ «·„»«·€ «·„”ÕÊ»…", vbExclamation
Exit Sub
End If
End If
If Combo2.Visible = True And Label17.Caption = "" Then
Call cont
Do While Not cf.EOF
If Text1.Text = cf!sri Or Val(Text1.Text) = Val(cf!sri) Then
If cf!moi = Combo2.Text Then
MsgBox "”»ﬁ √‰  „  ”œÌœ —« » «·‘Â— " + Combo2.Text, vbCritical
Exit Sub
End If
End If
cf.MoveNext
Loop
End If
'** controle caisse
Call cont
a = eb!cca
If Combo2.Visible = True Then
b = 0
Else
b = Text2.Text
End If
If Label11.Caption = "”Õ» „»·€" Then
c = Label22.Caption
Else
c = 0
End If
d = (a + c) - b
If d < 0 Then
MsgBox "—’Ìœ «·’‰œÊﬁ €Ì— ﬂ«› ·≈ „«„ «·⁄„·Ì…", vbExclamation
Exit Sub
End If
eb!cca = d
eb.Update
'******* controle
'**** archive de caisse ajou et modif
Adat = Date
Aheu = Time$
If Label17.Caption = "" Then
Atyp = "≈÷«›…"
Else
Atyp = " ⁄œÌ·"
End If
Adet = Combo1.Text + " „‰ ÿ—› «·„ÊŸ› ’«Õ» «·—ﬁ„ «· ”·”·Ì " + Text1.Text
Amon = Text2.Text
Acom = "”Ã· «·„ÊŸ›Ì‰"
Auti = directions.Label2.Caption
'****************************************
If Label17.Caption <> "" Then
Call cont
Do While Not cf.EOF
If Label17.Caption = cf!aut Then
cf!sri = Text1.Text
cf!nom = Label1.Caption
cf!dat = DT1.Value
cf!typ = Combo1.Text
cf!mon = Text2.Text
cf!det = Text3.Text
cf!heu = Time$
cf!moi = Combo2.Text
If cf!act = "2" Then
cf!act = "3"
End If
cf.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
cf.MoveNext
Loop
End If
cf.AddNew
cf!sri = Text1.Text
cf!nom = Label1.Caption
cf!dat = DT1.Value
cf!typ = Combo1.Text
cf!mon = Text2.Text
cf!det = Text3.Text
cf!heu = Time$
cf!moi = Combo2.Text
cf!act = "0"
cf!mtf = ""
cf.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Call chargegrd2_T
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
grd2.Text = "‰Ê⁄Ì… «·⁄„·Ì…"
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
grd2.Rows = cf.RecordCount + 3
Do While Not cf.EOF
If Text1.Text = cf!sri Or Val(Text1.Text) = Val(cf!sri) Then
dat3 = cf!dat
If dat3 >= dat1 And dat3 <= dat2 Then
If cf!act <> "1" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = cf!aut
grd2.Col = 1
grd2.Text = cf!dat
If cf!act = "2" Then
grd2.CellBackColor = &HFF&
End If
grd2.Col = 2
grd2.Text = cf!heu
If cf!typ = "”Õ» „»·€" Then
grd2.Col = 3
grd2.Text = cf!typ
grd2.Col = 4
grd2.Text = cf!mon
Else
grd2.Col = 3
grd2.Text = cf!typ + " ‘Â— " + cf!moi
grd2.Col = 4
grd2.Text = cf!mon
grd2.CellBackColor = &H80000008
End If
grd2.Col = 5
grd2.Text = cf!det
grd2.Col = 6
grd2.Text = " ⁄œÌ·"
grd2.CellBackColor = &HFFFF&
grd2.Col = 7
grd2.Text = "Õ–›"
grd2.CellBackColor = &HFF&
i = i + 1
End If
End If
If cf!typ = "”Õ» „»·€" Then
a = cf!mon
P = P + a
Else
a = cf!mon
r = r + a
End If
End If
cf.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 2
s = (r - P)
Label7.Caption = P
Label10.Caption = r
Label12.Caption = "«·—’Ìœ «·«Ã„«·Ì"
Label15.Caption = s
If s <= 0 Then
Label15.ForeColor = &HFF&
Else
Label15.ForeColor = &HFF00&
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
grd2.Text = "‰Ê⁄Ì… «·⁄„·Ì…"
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
grd2.Rows = cf.RecordCount + 3
Do While Not cf.EOF
If Text1.Text = cf!sri Or Val(Text1.Text) = Val(cf!sri) Then
dat3 = cf!dat
'If dat3 >= dat1 And dat3 <= dat2 Then
If cf!act <> "1" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = cf!aut
grd2.Col = 1
grd2.Text = cf!dat
If cf!act = "2" Then
grd2.CellBackColor = &HFF&
End If
grd2.Col = 2
grd2.Text = cf!heu
If cf!typ = "”Õ» „»·€" Then
grd2.Col = 3
grd2.Text = cf!typ
grd2.Col = 4
grd2.Text = cf!mon
Else
grd2.Col = 3
grd2.Text = cf!typ + " ‘Â— " + cf!moi
grd2.Col = 4
grd2.Text = cf!mon
grd2.CellBackColor = &H80000008
End If
grd2.Col = 5
grd2.Text = cf!det
grd2.Col = 6
grd2.Text = " ⁄œÌ·"
grd2.CellBackColor = &HFFFF&
grd2.Col = 7
grd2.Text = "Õ–›"
grd2.CellBackColor = &HFF&
i = i + 1
End If
If cf!typ = "”Õ» „»·€" Then
a = cf!mon
P = P + a
Else
a = cf!mon
r = r + a
End If
End If
cf.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 2
s = (r - P)
Label7.Caption = P
Label10.Caption = r
Label12.Caption = "«·—’Ìœ «·«Ã„«·Ì"
Label15.Caption = s
If s <= 0 Then
Label15.ForeColor = &HFF&
Else
Label15.ForeColor = &HFF00&
End If
End Sub

Private Sub grd2_Click()
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim n As Double
Dim i As Double
Dim j As Double
Dim tx As String
Dim tx1 As String
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
Do While Not cf.EOF
If cf!aut = tx1 Then
MsgBox cf!mtf
Exit Sub
End If
cf.MoveNext
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
tx = grd2.Text
Label11.Caption = grd2.Text
grd2.Col = 4
Label22.Caption = grd2.Text
grd2.Col = 5
Text3.Text = grd2.Text
If tx = "”Õ» „»·€" Then
Combo1.Text = tx
Text2.Text = Label22.Caption
Else
Label20.Caption = tx
n = Len(Label20.Caption)
vg = Mid$(Label20.Caption, 1, 10)
Combo1.Text = vg
vg = Mid$(Label20.Caption, 16, n)
Combo2.Text = vg
End If
End If
If j = 7 Then
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–Â «·⁄„·Ì…", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
grd2.Row = i
grd2.Col = 0
Label17.Caption = grd2.Text
grd2.Col = 3
tx = grd2.Text
grd2.Col = 4
Label22.Caption = grd2.Text
If tx <> "”Õ» „»·€" Then
b = Label22.Caption
c = Label10.Caption
d = Label7.Caption
c = (c - b)
If d > c Then
MsgBox "·« Ì„ﬂ‰ «Ã—«¡ Â–Â «·⁄„·Ì…, ·√‰Â ›Ì Â–Â «·Õ«·… ” ﬂÊ‰ «·„»«·€ «·„”ÕÊ»… √ﬂ»— „‰ «·„” Õﬁ« , ÌÃ» √Ê·« Õ–› »⁄÷ «·„»«·€ «·„”ÕÊ»…", vbExclamation + arabic
Label17.Caption = ""
Label22.Caption = "0"
Exit Sub
End If
End If
If tx = "”Õ» „»·€" Then
c = eb!cca
b = Label22.Caption
c = (c + b)
eb!cca = c
eb.Update
End If
Call cont
Do While Not cf.EOF
If Label17.Caption = cf!aut Then
'**** archive de caisse supp
Label19.Caption = cf!sri
Adat = Date
Aheu = Time$
Atyp = "Õ–›"
Adet = tx + " ··„ÊŸ› ’«Õ» «·—ﬁ„ «· ”·”·Ì " + Label19.Caption
Amon = b
Acom = "”Ã· «·„ÊŸ›Ì‰"
Auti = directions.Label2.Caption
'****************************************
cf.Delete
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
cf.MoveNext
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
Label1.Caption = ""
Label18.Caption = "0"
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
'ProgressBar1.Value = ProgressBar1.Value + 8
'If ProgressBar1.Value > 90 Then
'MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation + arabic
Command8_Click
Call archive_caisse
'End If

End Sub
Private Sub chargegrd_clear()
grd2.Clear
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1200
grd2.ColWidth(2) = 1000
grd2.ColWidth(3) = 1800
grd2.ColWidth(4) = 1500
grd2.ColWidth(5) = 2400
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
grd2.Text = "‰Ê⁄Ì… «·⁄„·Ì…"
grd2.Col = 4
grd2.Text = "«·„»·€"
grd2.Col = 5
grd2.Text = "«· ›«’Ì·"
End Sub
Private Sub chargetreeview1()
Dim id1 As String
Dim id2 As String
Dim i As Double
Dim n As Double
TreeView1.Nodes.Clear
Call cont
Do While Not fc.EOF
If fc!act = "1" Then
id1 = fc!sri
id2 = "M" + id1
TreeView1.Nodes.Add , tvwChild, id2, fc!nom
End If
fc.MoveNext
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
Do While Not cf.EOF
If cf!typ = "”Õ» „»·€" Then
If cf!act = "0" Or cf!act = "3" Then
a = a + 1
End If
If cf!act = "2" Then
b = b + 1
End If
End If
cf.MoveNext
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
grd6.Rows = cf.RecordCount + 3
Do While Not cf.EOF
If cf!typ = "”Õ» „»·€" Then
If cf!act = Text6.Text Or cf!act = Text5.Text Then
grd6.Row = i
grd6.Col = 0
grd6.Text = cf!dat
grd6.Col = 1
grd6.Text = cf!sri
If Text6.Text = "0" Then
grd6.CellBackColor = &HFFFF&
Else
grd6.CellBackColor = &HFF&
End If
i = i + 1
End If
End If
cf.MoveNext
Loop
grd6.Rows = i
End Sub



