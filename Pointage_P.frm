VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Pointage_P 
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   8175
      Left            =   120
      ScaleHeight     =   8175
      ScaleWidth      =   10095
      TabIndex        =   8
      Top             =   1440
      Width           =   10095
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "Pointage_P.frx":0000
         Left            =   2040
         List            =   "Pointage_P.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox Combo7 
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
         Height          =   345
         ItemData        =   "Pointage_P.frx":002B
         Left            =   7680
         List            =   "Pointage_P.frx":0053
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   480
         Width           =   735
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
         Left            =   8880
         TabIndex        =   40
         Top             =   1080
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
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
         Height          =   345
         ItemData        =   "Pointage_P.frx":007E
         Left            =   2040
         List            =   "Pointage_P.frx":008B
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   120
         Width           =   1095
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
         TabIndex        =   24
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   1575
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
         Left            =   240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   345
         ItemData        =   "Pointage_P.frx":00A9
         Left            =   3960
         List            =   "Pointage_P.frx":00AB
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   120
         Width           =   975
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
         Left            =   7680
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   480
         Width           =   735
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
         Left            =   2040
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   480
         Width           =   2895
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
         Left            =   5640
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   480
         Width           =   1215
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
         Height          =   330
         Left            =   1080
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1080
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
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
         Height          =   345
         ItemData        =   "Pointage_P.frx":00AD
         Left            =   5640
         List            =   "Pointage_P.frx":00BD
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   120
         Width           =   1215
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
         Height          =   330
         Left            =   120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
      Begin VB.PictureBox Picture3 
         Height          =   3375
         Left            =   3000
         ScaleHeight     =   3315
         ScaleWidth      =   2715
         TabIndex        =   9
         Top             =   3120
         Visible         =   0   'False
         Width           =   2775
         Begin VB.CommandButton Command5 
            Caption         =   " ’ÕÌÕ Ê÷⁄Ì… ‰"
            Height          =   495
            Left            =   240
            TabIndex        =   47
            Top             =   2760
            Width           =   2295
         End
         Begin VB.CommandButton Command4 
            Caption         =   " ÕœÌœ «·„” ÊÌ« "
            Height          =   495
            Left            =   840
            TabIndex        =   46
            Top             =   960
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Õ÷Ê— «·√”« –…"
            Height          =   495
            Left            =   960
            TabIndex        =   45
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox Text11 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Text            =   "Text11"
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   50
            Left            =   0
            Top             =   0
         End
         Begin VB.Label Label10 
            Caption         =   "Label10"
            Height          =   255
            Left            =   1800
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label31 
            Caption         =   "0"
            Height          =   255
            Left            =   480
            TabIndex        =   15
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label30 
            Caption         =   "0"
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label Label17 
            Caption         =   "Label17"
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label26 
            Caption         =   "0"
            Height          =   255
            Left            =   480
            TabIndex        =   12
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "0"
            Height          =   255
            Left            =   480
            TabIndex        =   11
            Top             =   1200
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grd1 
         Height          =   6375
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   9855
         _ExtentX        =   17383
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
      Begin MSComCtl2.DTPicker DT1 
         Height          =   345
         HelpContextID   =   345
         Left            =   7680
         TabIndex        =   26
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16744576
         CalendarTitleBackColor=   16711680
         CalendarTrailingForeColor=   16744576
         Format          =   102825985
         CurrentDate     =   42638
      End
      Begin ComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   3480
         TabIndex        =   27
         Top             =   840
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSComCtl2.DTPicker DT2 
         Height          =   330
         Left            =   5160
         TabIndex        =   28
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16744576
         CalendarTitleBackColor=   16711680
         CalendarTrailingForeColor=   16744576
         Format          =   102825985
         CurrentDate     =   42638
      End
      Begin MSComCtl2.DTPicker DT3 
         Height          =   330
         Left            =   3000
         TabIndex        =   29
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16744576
         CalendarTitleBackColor=   16711680
         CalendarTrailingForeColor=   16744576
         Format          =   102825985
         CurrentDate     =   42638
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„” Õﬁ«  «·‘Â—"
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
         Index           =   1
         Left            =   8160
         TabIndex        =   42
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«· ﬁ«÷Ì"
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
         Left            =   3000
         TabIndex        =   39
         Top             =   120
         Width           =   855
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
         Left            =   4200
         TabIndex        =   37
         Top             =   480
         Width           =   1335
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
         Left            =   8640
         TabIndex        =   36
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ ”«⁄«  «· œ—Ì”"
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
         Left            =   8160
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label6 
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
         Left            =   6360
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label8 
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
         Left            =   4800
         TabIndex        =   33
         Top             =   120
         Width           =   735
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
         Left            =   5760
         TabIndex        =   32
         Top             =   480
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         Height          =   6615
         Index           =   3
         Left            =   0
         Top             =   1440
         Width           =   10095
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄—÷ ”Ã· «·Õ÷Ê— „‰  «—ÌŒ"
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
         Left            =   6000
         TabIndex        =   31
         Top             =   1080
         Width           =   2775
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
         Left            =   3960
         TabIndex        =   30
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Index           =   4
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   10095
      End
      Begin VB.Line Line3 
         X1              =   1920
         X2              =   1920
         Y1              =   0
         Y2              =   960
      End
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   10320
      ScaleHeight     =   8745
      ScaleWidth      =   2505
      TabIndex        =   6
      Top             =   840
      Width           =   2535
      Begin ComctlLib.TreeView TreeView1 
         Height          =   8775
         Left            =   0
         TabIndex        =   7
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
      Left            =   5400
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
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
      Height          =   345
      Left            =   6360
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   2535
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
      Left            =   3960
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   10095
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
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   4455
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
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”Ã· Õ÷Ê— «·√”« –…"
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
      Index           =   3
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Pointage_P"
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

Private Sub chargetreeview1()
Dim id1 As String
Dim id2 As String
Dim i As Double
Dim n As Double
TreeView1.Nodes.Clear
'TreeView1.Nodes.Add , , "PF", "√”„«¡ «·√”« –…"
Call cont
Do While Not pf.EOF
If pf!act = "1" Then
id1 = pf!sri
id2 = "P" + id1
TreeView1.Nodes.Add , tvwChild, id2, pf!nom
End If
pf.MoveNext
Loop
End Sub

Private Sub Combo1_Change()
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
Combo3.Clear
Combo3.AddItem "»«·”«⁄…"
Combo3.AddItem "»«·‘Â—"
Combo3.AddItem "»«·‰”»…"
Combo3.BackColor = &H8080FF
Else
Combo2.BackColor = &H8080FF
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
Else
Combo2.BackColor = &H8080FF
End If

End Sub

Private Sub Combo2_Click()
Combo2_Change
End Sub

Private Sub Combo3_Change()
If Len(Combo3.Text) > 0 Then
If Combo3.Text = "»«·”«⁄…" Then
Label10.Caption = "H"
ElseIf Combo3.Text = "»«·‘Â—" Then
Label10.Caption = "M"
ElseIf Combo3.Text = "»«·‰”»…" Then
Label10.Caption = "P"
End If
Combo3.BackColor = &HC000&
If Combo3.Text = "»«·”«⁄…" Then
Label5(1).Visible = False
Combo7.Visible = False
Label5(0).Visible = True
Text2.Visible = True
Text2.Text = ""
Text4.Enabled = True
Text4.Text = ""
Text4.BackColor = &H8080FF
Label9.Visible = True
ElseIf Combo3.Text = "»«·‘Â—" Then
Label5(0).Visible = False
Text2.Visible = False
Text2.Text = "1"
Label5(1).Visible = True
Combo7.Visible = True
Text4.Enabled = True
Text4.Text = ""
Text4.BackColor = &H8080FF
Label9.Visible = True
ElseIf Combo3.Text = "»«·‰”»…" Then
Label5(1).Visible = False
Combo7.Visible = False
Label5(0).Visible = True
Text2.Visible = True
Text2.Text = ""
Text4.Enabled = False
Text4.Text = "0"
Text4.BackColor = &H8000&
Label9.Visible = False
End If
Else
Combo3.BackColor = &H8080FF
End If

End Sub

Private Sub Combo3_Click()
Combo3_Change
End Sub

Private Sub Combo4_Change()
Label10.Caption = ""
If Combo4.Text = "»«·”«⁄…" Then
Label10.Caption = "H"
ElseIf Combo4.Text = "»«·‘Â—" Then
Label10.Caption = "M"
ElseIf Combo4.Text = "»«·‰”»…" Then
Label10.Caption = "P"
End If
Call chargegrd_clear

End Sub

Private Sub Combo4_Click()
Combo4_Change
End Sub

Private Sub Combo7_Change()
If Len(Combo7.Text) > 0 Then
Combo7.BackColor = &HC000&
Else
Combo7.BackColor = &H8080FF
End If

End Sub

Private Sub Combo7_Click()
Combo7_Change
End Sub

Private Sub Command1_Click()
Call cont3
Do While Not pp3.EOF
Call cont
'xe = eb!sri
'Call Series
Text1.Text = pp3!ser
Command2_Click
DT1.Value = pp3!dat
DT1.Month = pp3!mois
Combo2.Text = pp3!niv
Combo1.Text = pp3!cla
If pp3!cas = "m" Then
Combo3.Text = "»«·‘Â—"
Combo7.Text = pp3!mois
Text4.Text = pp3!mon
ElseIf pp3!cas = "p" Then
Combo3.Text = "»«·‰”»…"
Text2.Text = pp3!nbr
ElseIf pp3!cas = "h" Then
Combo3.Text = "»«·”«⁄…"
Text2.Text = pp3!nbr
Text4.Text = pp3!mon
End If
Command9_Click
pp3.MoveNext
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
Call cont
Do While Not pf.EOF
If pf!sri = Text1.Text Or Val(pf!sri) = Val(Text1.Text) Then
If pf!act = "1" Then
Label1.Caption = pf!nom
grd1.Visible = False
Combo4.Text = "»«·”«⁄…"
Call chargegrd1_T
grd1.Visible = True
Picture1.Visible = True
Exit Sub
Else
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· ·√” «–  „ Õ–›Â", vbCritical + arabic
Exit Sub
End If
End If
pf.MoveNext
Loop
Call cont
Do While Not sr.EOF
If sr!sri = Text1.Text Or Val(sr!sri) = Val(Text1.Text) Then
MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· ·Ì” —ﬁ„  ”·”·Ì ·√” «– Ê≈‰„« —ﬁ„  ”·”·Ì ·" + sr!eta, vbExclamation
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
Dim cl1 As String
Dim cl2 As String
Dim niv1 As String
Call cont
Call cont3
Do While Not pp3.EOF
cl1 = pp3!cla
cl.MoveFirst
Do While Not cl.EOF
cl2 = cl!cla
If cl1 = cl2 Then
niv1 = cl!niv
pp3!niv = niv1
pp3.Update
cl.MoveLast
End If
cl.MoveNext
Loop
pp3.MoveNext
Loop
MsgBox "OK", vbInformation

End Sub

Private Sub Command5_Click()
Call cont2
Do While Not pp2.EOF
If pp2!eta = "P" Then
Label17.Caption = pp2!aut
DT1.Value = pp2!dat
Combo2.Text = pp2!niv
Combo1.Text = pp2!cla
Combo3.Text = "»«·‰”»…"
Label26.Caption = pp2!nbh
Text2.Text = pp2!nbh
Text4.Text = "0"
Label27.Caption = pp2!mon
DT1.Enabled = False
Combo2.Enabled = False
Combo1.Enabled = False
Combo3.Enabled = False
Combo7.Enabled = False
Call Pourcentage
Command8_Click
End If
pp2.MoveNext
Loop
Call cont
Do While Not pp.EOF
If pp!eta = "P" Then
pp.Delete
End If
pp.MoveNext
Loop

Call cont3
Do While Not pp3.EOF
If pp3!cas = "p" Then
Call cont
'xe = eb!sri
'Call Series
Text1.Text = pp3!ser
Command2_Click
DT1.Value = pp3!dat
DT1.Month = pp3!mois
Combo2.Text = pp3!niv
Combo1.Text = pp3!cla
If pp3!cas = "m" Then
Combo3.Text = "»«·‘Â—"
Combo7.Text = pp3!mois
Text4.Text = pp3!mon
ElseIf pp3!cas = "p" Then
Combo3.Text = "»«·‰”»…"
Text2.Text = pp3!nbr
ElseIf pp3!cas = "h" Then
Combo3.Text = "»«·”«⁄…"
Text2.Text = pp3!nbr
Text4.Text = pp3!mon
End If
Command9_Click
End If
pp3.MoveNext
Loop
MsgBox "OK", vbInformation

End Sub

Private Sub Command7_Click()
grd1.Visible = False
Call chargegrd1_T
grd1.Visible = True

End Sub

Private Sub Command8_Click()
DT1.Enabled = True
Combo2.Enabled = True
Combo1.Enabled = True
Combo3.Enabled = True
Combo7.Enabled = True
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
DT1.Value = Date
Label17.Caption = ""
Label26.Caption = "0"
Label27.Caption = "0"
Combo4.Text = Combo3.Text
Combo3.Clear
Combo3.AddItem "»«·”«⁄…"
Combo3.AddItem "»«·‘Â—"
Combo3.AddItem "»«·‰”»…"
Combo3.BackColor = &H8080FF
grd1.Visible = False
Call chargegrd1_T
grd1.Visible = True
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False

End Sub

Private Sub Command9_Click()
Dim x$
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text4.Text = Trim(Text4.Text)
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
If Text2.Text = "" Or Text4.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo7.Text = "" And Combo7.Visible = True Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text2.BackColor = &H8080FF And Text2.Visible = True Then
Text2.SetFocus
ElseIf Text4.BackColor = &H8080FF And Text4.Enabled = True Then
Text4.SetFocus
End If
Exit Sub
End If
If Label17.Caption <> "" Then
Call cont
Do While Not pp.EOF
If Label17.Caption = pp!aut Then
pp!sri = Text1.Text
pp!nom = Label1.Caption
pp!dat = DT1.Value
pp!niv = Combo2.Text
pp!cla = Combo1.Text
pp!eta = Label10.Caption
pp!nbh = Text2.Text
pp!mon = Text4.Text
pp!det = Text3.Text
If Combo7.Visible = True Then
pp!moi = Combo7.Text
Else
pp!moi = Month(DT1.Value)
End If
pp.Update
Call Pourcentage
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
pp.MoveNext
Loop
End If
pp.AddNew
pp!sri = Text1.Text
pp!nom = Label1.Caption
pp!dat = DT1.Value
pp!niv = Combo2.Text
pp!cla = Combo1.Text
pp!eta = Label10.Caption
pp!nbh = Text2.Text
pp!mon = Text4.Text
pp!det = Text3.Text
pp!moi = Month(DT1.Value)
If Combo7.Visible = True Then
pp!moi = Combo7.Text
End If
pp.Update
Call Pourcentage
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
'DT1.Value = Date
'DT2.Value = Date
'DT3.Value = Date
Call MakeTreeViewRTL
Call chargetreeview1
Call couleur_treeview1
'Call chargegrd_clear

End Sub


Private Sub grd1_Click()
Dim i As Double
Dim j As Double
Dim au As Double
Dim a As Double
Dim b As Double
Dim tx As String
i = grd1.Row
j = grd1.Col
If i > 0 Then
If j = 9 Then
grd1.Row = i
grd1.Col = 0
Label17.Caption = grd1.Text
grd1.Col = 1
DT1.Value = grd1.Text
grd1.Col = 2
Combo2.Text = grd1.Text
grd1.Col = 3
Combo1.Text = grd1.Text
grd1.Col = 7
tx = grd1.Text
If tx = "H" Then
Combo3.Text = "»«·”«⁄…"
grd1.Col = 4
Text2.Text = grd1.Text
Label26.Caption = grd1.Text
ElseIf tx = "M" Then
Combo3.Text = "»«·‘Â—"
grd1.Col = 4
Combo7.Text = grd1.Text
Label26.Caption = grd1.Text
ElseIf tx = "P" Then
Combo3.Text = "»«·‰”»…"
grd1.Col = 4
Text2.Text = grd1.Text
Label26.Caption = grd1.Text
End If
grd1.Col = 5
Text4.Text = grd1.Text
Label27.Caption = grd1.Text
grd1.Col = 6
Text3.Text = grd1.Text
DT1.Enabled = False
Combo2.Enabled = False
Combo1.Enabled = False
Combo3.Enabled = False
Combo7.Enabled = False
End If
If j = 10 Then
g = MsgBox("Â·  —Ìœ Õﬁ« «·Õ–›ø", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
grd1.Row = i
grd1.Col = 0
Label17.Caption = grd1.Text
grd1.Col = 1
DT1.Value = grd1.Text
grd1.Col = 2
Combo2.Text = grd1.Text
grd1.Col = 3
Combo1.Text = grd1.Text
grd1.Col = 7
tx = grd1.Text
If tx = "H" Then
Combo3.Text = "»«·”«⁄…"
grd1.Col = 4
Label26.Caption = grd1.Text
ElseIf tx = "M" Then
Combo3.Text = "»«·‘Â—"
grd1.Col = 4
Combo7.Text = grd1.Text
Label26.Caption = grd1.Text
ElseIf tx = "P" Then
Combo3.Text = "»«·‰”»…"
grd1.Col = 4
Label26.Caption = grd1.Text
End If
Text2.Text = "0"
grd1.Col = 5
Text4.Text = "0"
Label27.Caption = grd1.Text
DT1.Enabled = False
Combo2.Enabled = False
Combo1.Enabled = False
Combo3.Enabled = False
Combo7.Enabled = False
Call cont
Do While Not pp.EOF
If Label17.Caption = pp!aut Then
pp.Delete
Call Pourcentage
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
pp.MoveNext
Loop
End If

End If
End If

End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Else
Text1.BackColor = &H8080FF
End If
Text4.Text = ""
Text2.Text = ""
DT1.Value = Date
Label17.Caption = ""
Label1.Caption = ""
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False
Picture1.Visible = False

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

Private Sub Text4_Change()
If Len(Text4.Text) > 0 Then
Text4.BackColor = &HC000&
Else
Text4.BackColor = &H8080FF
End If


End Sub

Private Sub Text4_Click()
Text4_Change
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
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
End If

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
Dim n As Double
Text11.Text = Node.Key
n = Len(Text11.Text)
If n > 2 Then
n = (n - 1)
vg = Mid$(Text11.Text, 2, n)
Text1.Text = vg
Command2_Click
End If

End Sub
Private Sub chargcombo1()
Combo1.Clear
Call cont
Do While Not cl.EOF
If Combo2.Text = cl!niv And cl!act = "1" Then
Combo1.AddItem cl!cla
End If
cl.MoveNext
Loop
End Sub
Private Sub chargegrd1_T()
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim i As Double
i = 1
dat1 = DT2.Value
dat2 = DT3.Value
Call cont
grd1.Rows = pp.RecordCount + 3
Do While Not pp.EOF
If Text1.Text = pp!sri Or Val(Text1.Text) = Val(pp!sri) Then
If pp!eta = Label10.Caption Or Combo4.Text = "«·ﬂ·" Then
dat3 = pp!dat
'If dat3 >= dat1 And dat3 <= dat2 Then
grd1.Row = i
grd1.Col = 0
grd1.Text = pp!aut
grd1.Col = 1
grd1.Text = pp!dat
grd1.Col = 2
grd1.Text = pp!niv
grd1.Col = 3
grd1.Text = pp!cla
grd1.Col = 4
If Combo4.Text = "»«·‘Â—" Then
grd1.Text = pp!moi
Else
grd1.Text = pp!nbh
End If
grd1.Col = 5
grd1.Text = pp!mon
If Combo4.Text = "»«·‰”»…" Then
grd1.CellBackColor = &H80000008
End If
grd1.Col = 6
grd1.Text = pp!det
grd1.Col = 7
grd1.Text = pp!eta
'grd1.Col = 8
'grd1.Text = pp!rtr
grd1.Col = 9
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 10
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
'End If
End If
End If
pp.MoveNext
Loop
grd1.Rows = i
End Sub
Private Sub chargegrd1_M()
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim i As Double
i = 1
dat1 = DT2.Value
dat2 = DT3.Value
Call cont
grd1.Rows = pp.RecordCount + 3
Do While Not pp.EOF
If Text1.Text = pp!sri Or Val(Text1.Text) = Val(pp!sri) Then
If pp!eta = Label10.Caption Or Combo4.Text = "«·ﬂ·" Then
dat3 = pp!dat
If dat3 >= dat1 And dat3 <= dat2 Then
grd1.Row = i
grd1.Col = 0
grd1.Text = pp!aut
grd1.Col = 1
grd1.Text = pp!dat
grd1.Col = 2
grd1.Text = pp!niv
grd1.Col = 3
grd1.Text = pp!cla
grd1.Col = 4
If Combo4.Text = "»«·‘Â—" Then
grd1.Text = pp!moi
Else
grd1.Text = pp!nbh
End If
grd1.Col = 5
grd1.Text = pp!mon
If Combo4.Text = "»«·‰”»…" Then
grd1.CellBackColor = &H80000008
End If
grd1.Col = 6
grd1.Text = pp!det
grd1.Col = 7
grd1.Text = pp!eta
grd1.Col = 8
grd1.Text = ""
grd1.Col = 9
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 10
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
End If
End If
End If
pp.MoveNext
Loop
grd1.Rows = i
End Sub
Private Sub chargegrd_clear()
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 1200
grd1.ColWidth(3) = 1200
grd1.ColWidth(4) = 1200
grd1.ColWidth(5) = 1200
grd1.ColWidth(6) = 2300
grd1.ColWidth(7) = 0
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 600
grd1.ColWidth(10) = 600
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 1
grd1.ColAlignment(9) = 3
grd1.ColAlignment(10) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«· «—ÌŒ"
grd1.Col = 2
grd1.Text = "«·„” ÊÏ"
grd1.Col = 3
grd1.Text = "«·ﬁ”„"
grd1.Col = 4
If Combo4.Text = "»«·‘Â—" Then
grd1.Text = "«·‘Â—"
Else
grd1.Text = "⁄œœ.”"
End If
grd1.Col = 5
grd1.Text = "«·„»·€"
grd1.Col = 6
grd1.Text = "«· ›«’Ì·"

End Sub
Private Sub Pourcentage()
Dim n1 As Double
Dim m1 As Double
Dim n2 As Double
Dim m2 As Double
Dim m3 As Double
Dim m4 As Double
Dim m As Double
n1 = Label26.Caption
m1 = Label27.Caption
m1 = (n1 * m1)
n2 = Text2.Text
m2 = Text4.Text
m2 = (n2 * m2)
m = Month(DT1.Value)
Call cont
Do While Not pc.EOF
If Combo1.Text = pc!cla And m = pc!moi Then
If Label10.Caption = "H" Or Label10.Caption = "M" Then
m3 = pc!pro
m4 = (m3 - m1 + m2)
pc!pro = m4
pc.Update
End If
If Label10.Caption = "P" Then
n3 = pc!nbr
n4 = (n3 - n1 + n2)
pc!nbr = n4
pc.Update
End If
Exit Sub
End If
pc.MoveNext
Loop
If Label10.Caption = "H" Or Label10.Caption = "M" Then
pc.AddNew
If Combo7.Visible = True Then
pc!moi = Combo7.Text
Else
pc!moi = Month(DT1.Value)
End If
pc!niv = Combo2.Text
pc!cla = Combo1.Text
pc!etu = "0"
pc!pro = m2
pc!nbr = "0"
pc.Update
End If
If Label10.Caption = "P" Then
pc.AddNew
pc!moi = Month(DT1.Value)
pc!niv = Combo2.Text
pc!cla = Combo1.Text
pc!etu = "0"
pc!pro = "0"
pc!nbr = n2
pc.Update
End If
End Sub

