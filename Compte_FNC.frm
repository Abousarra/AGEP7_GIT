VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Compte_FNC 
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
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8415
      ScaleWidth      =   9375
      TabIndex        =   75
      Top             =   1200
      Width           =   9375
   End
   Begin VB.CommandButton Command8 
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
      Left            =   2640
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   1320
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   10080
      ScaleHeight     =   2595
      ScaleWidth      =   2355
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   1440
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
         Format          =   124977153
         CurrentDate     =   42638
      End
      Begin VB.Label Label24 
         Caption         =   "Label24"
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label37 
         Caption         =   "Label37"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "30"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      Caption         =   "Õ›Ÿ «· €ÌÌ—"
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
      Left            =   480
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text6 
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
      Left            =   240
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox Text5 
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
      Left            =   240
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1920
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
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Õ”«» ‘Â—Ì"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "Õ”«» ”‰ÊÌ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5160
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
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
      ItemData        =   "Compte_FNC.frx":0000
      Left            =   6720
      List            =   "Compte_FNC.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1320
      Width           =   975
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
      TabIndex        =   3
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
      TabIndex        =   2
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
      TabIndex        =   1
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
      Left            =   8160
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   855
   End
   Begin ComctlLib.TreeView TreeView1 
      Height          =   8415
      Left            =   9600
      TabIndex        =   17
      Top             =   1200
      Width           =   3135
      _ExtentX        =   5530
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6255
      Left            =   240
      TabIndex        =   18
      Top             =   3240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   11033
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
      TabCaption(0)   =   "«·„»«·€ «·„” Õﬁ…"
      TabPicture(0)   =   "Compte_FNC.frx":0054
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Command9"
      Tab(0).Control(1)=   "grd1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "«·„»«·€ «·„”ÕÊ»…"
      TabPicture(1)   =   "Compte_FNC.frx":0070
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grd2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FF8080&
         BorderStyle     =   0  'None
         Height          =   6735
         Left            =   -74880
         ScaleHeight     =   6735
         ScaleWidth      =   9735
         TabIndex        =   21
         Top             =   360
         Width           =   9735
         Begin VB.ComboBox Combo7 
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
            ItemData        =   "Compte_FNC.frx":008C
            Left            =   4800
            List            =   "Compte_FNC.frx":00B4
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox Combo4 
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
            Height          =   405
            ItemData        =   "Compte_FNC.frx":00DF
            Left            =   7320
            List            =   "Compte_FNC.frx":00E9
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox Text10 
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
            Left            =   4800
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox Text9 
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
            Left            =   4800
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox Text8 
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
            Left            =   2280
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Text            =   "0"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Text7 
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
            Left            =   2280
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Text            =   "0"
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
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
            Height          =   375
            Left            =   2160
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   1800
            Width           =   975
         End
         Begin VB.ComboBox Combo3 
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
            Height          =   405
            ItemData        =   "Compte_FNC.frx":00FA
            Left            =   7320
            List            =   "Compte_FNC.frx":010A
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   720
            Width           =   1455
         End
         Begin VB.CommandButton Command7 
            Appearance      =   0  'Flat
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
            Height          =   375
            Left            =   120
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1800
            Width           =   1935
         End
         Begin VB.PictureBox Picture4 
            Height          =   2895
            Left            =   3000
            ScaleHeight     =   2835
            ScaleWidth      =   2715
            TabIndex        =   22
            Top             =   3120
            Visible         =   0   'False
            Width           =   2775
            Begin VB.Timer Timer2 
               Enabled         =   0   'False
               Interval        =   50
               Left            =   0
               Top             =   0
            End
            Begin VB.Label Label28 
               Caption         =   "0"
               Height          =   375
               Left            =   480
               TabIndex        =   24
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label12 
               Caption         =   "Label17"
               Height          =   255
               Left            =   480
               TabIndex        =   23
               Top             =   240
               Width           =   1335
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grd3 
            Height          =   4095
            Left            =   240
            TabIndex        =   34
            Top             =   2400
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   7223
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
         Begin MSComCtl2.DTPicker DT4 
            Height          =   375
            Left            =   7320
            TabIndex        =   35
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
            Format          =   124977153
            CurrentDate     =   42638
         End
         Begin ComctlLib.ProgressBar ProgressBar2 
            Height          =   375
            Left            =   2280
            TabIndex        =   36
            Top             =   1200
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   661
            _Version        =   327682
            Appearance      =   1
         End
         Begin MSComCtl2.DTPicker DT5 
            Height          =   375
            Left            =   5640
            TabIndex        =   37
            Top             =   1800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
            Format          =   124977153
            CurrentDate     =   42638
         End
         Begin MSComCtl2.DTPicker DT6 
            Height          =   375
            Left            =   3240
            TabIndex        =   38
            Top             =   1800
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
            Format          =   124977153
            CurrentDate     =   42638
         End
         Begin VB.Label Label11 
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
            Left            =   240
            TabIndex        =   50
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "≈·€«¡"
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
            Left            =   240
            TabIndex        =   49
            Top             =   1080
            Width           =   1815
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
            Index           =   1
            Left            =   5760
            TabIndex        =   48
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«· «—ÌŒ"
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
            Left            =   8160
            TabIndex        =   47
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„” Õﬁ«  «·‘Â—"
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
            Left            =   5280
            TabIndex        =   46
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label6 
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
            Index           =   1
            Left            =   8280
            TabIndex        =   45
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label30 
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
            Left            =   8280
            TabIndex        =   44
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label31 
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
            Left            =   5280
            TabIndex        =   43
            Top             =   720
            Width           =   1815
         End
         Begin VB.Shape Shape1 
            Height          =   4335
            Index           =   6
            Left            =   120
            Top             =   2280
            Width           =   9495
         End
         Begin VB.Image Image2 
            Height          =   645
            Index           =   3
            Left            =   240
            Picture         =   "Compte_FNC.frx":012C
            Top             =   240
            Width           =   1860
         End
         Begin VB.Image Image2 
            Height          =   645
            Index           =   2
            Left            =   240
            Picture         =   "Compte_FNC.frx":0D14
            Top             =   960
            Width           =   1860
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄·«Ê« "
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
            Left            =   3480
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„ √Œ—« "
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
            Left            =   3480
            TabIndex        =   41
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄—÷ ”Ã· «·Õ÷Ê— „‰  «—ÌŒ"
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
            Left            =   6840
            TabIndex        =   40
            Top             =   1800
            Width           =   2775
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "≈·Ï  «—ÌŒ"
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
            Left            =   3720
            TabIndex        =   39
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Shape Shape1 
            Height          =   1575
            Index           =   9
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   9495
         End
         Begin VB.Line Line6 
            X1              =   7200
            X2              =   7200
            Y1              =   120
            Y2              =   1680
         End
         Begin VB.Line Line5 
            X1              =   4680
            X2              =   4680
            Y1              =   120
            Y2              =   1680
         End
         Begin VB.Line Line4 
            X1              =   2160
            X2              =   2160
            Y1              =   120
            Y2              =   1680
         End
      End
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         Caption         =   "”Õ»"
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
         Left            =   -71160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Caption         =   "”Õ»"
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
         Left            =   3840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   5880
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grd1 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   51
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9763
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
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grd2 
         Height          =   5535
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9763
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
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Line Line2 
      X1              =   2520
      X2              =   9480
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " €ÌÌ— ﬂ·„… «·”—"
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
      Left            =   360
      TabIndex        =   73
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "≈⁄«œ…"
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
      TabIndex        =   72
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÃœÌœ…"
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
      TabIndex        =   71
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁœÌ„…"
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
      TabIndex        =   70
      Top             =   1560
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label26 
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
      TabIndex        =   69
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—’Ìœ «·⁄«„ ··Õ”«»"
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
      Left            =   4080
      TabIndex        =   68
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2760
      TabIndex        =   67
      Top             =   2640
      Width           =   1755
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   4
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      Height          =   6495
      Index           =   3
      Left            =   120
      Top             =   3120
      Width           =   9375
   End
   Begin VB.Label Label4 
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
      Index           =   0
      Left            =   4560
      TabIndex        =   66
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» «·„ÊŸ›Ì‰"
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
      TabIndex        =   65
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label3 
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
      Left            =   11040
      TabIndex        =   64
      Top             =   720
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   12615
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·≈”„"
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
      Left            =   8160
      TabIndex        =   63
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—« »"
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
      Left            =   7320
      TabIndex        =   62
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·ÊŸÌ›…"
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
      Index           =   0
      Left            =   8280
      TabIndex        =   61
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   6480
      TabIndex        =   60
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„»«·€ «·„” Õﬁ…"
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
      Left            =   4680
      TabIndex        =   59
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2640
      TabIndex        =   58
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   2
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label Label18 
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
      Left            =   2160
      TabIndex        =   57
      Top             =   720
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   600
      Y2              =   1080
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   6480
      TabIndex        =   56
      Top             =   1800
      Width           =   1995
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—’Ìœ"
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
      Left            =   7680
      TabIndex        =   55
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2640
      TabIndex        =   54
      Top             =   2160
      Width           =   1995
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF8080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   6120
      TabIndex        =   53
      Top             =   2640
      Width           =   2355
   End
End
Attribute VB_Name = "Compte_FNC"
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

Private Sub chargetreeview1()
On Error Resume Next
Dim id1 As String
Dim id2 As String
Dim i As Double
Dim n As Double
TreeView1.Nodes.Clear
TreeView1.Nodes.Add , , "FN", "√”„«¡ «·„ÊŸ›Ì‰"
Call cont
Do While Not fc.EOF
If fc!act = "1" Then
id1 = fc!sri
id2 = "F" + id1
TreeView1.Nodes.Add "FN", tvwChild, id2, fc!nom
End If
fc.MoveNext
Loop
End Sub

Private Sub Combo1_Change()
On Error Resume Next
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
Call tous_clear
Else
Combo1.BackColor = &H8080FF
End If
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
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
Do While Not fc.EOF
If Text1.Text = fc!sri Or Val(Text1.Text) = Val(fc!sri) Then
If fc!act = "1" Then
Label26.Caption = fc!nom
Label20.Caption = fc!sal
Label8.Caption = fc!fon
Label24.Caption = fc!mot
Option2.Value = True
Command8_Click
Text2.SetFocus
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


Private Sub Command2_Click()
On Error Resume Next
Text1.Text = Trim(Text1.Text)
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
If Text2.Text = Label24.Caption Then
Picture2.Visible = False
Else
MsgBox "ﬂ·„… «·”— «· Ì √œŒ· „ €Ì— ’ÕÌÕ…", vbExclamation + arabic
Text2.Text = ""
Text2.SetFocus
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Text4.Text = Trim(Text4.Text)
Text5.Text = Trim(Text5.Text)
Text6.Text = Trim(Text6.Text)
If Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text4.BackColor = &H8080FF Then
Text4.SetFocus
ElseIf Text5.BackColor = &H8080FF Then
Text5.SetFocus
ElseIf Text6.BackColor = &H8080FF Then
Text6.SetFocus
End If
Exit Sub
End If
If Text4.Text <> Label37.Caption Then
MsgBox "ﬂ·„… «·”— «·ﬁœÌ„… €Ì— ’ÕÌÕ…", vbCritical + arabic
Exit Sub
End If
If Text5.Text <> Text6.Text Then
MsgBox "ﬂ·„ « «·”— €Ì— „ ÿ«»ﬁ Ì‰", vbCritical + arabic
Exit Sub
End If
Call cont
Do While Not fc.EOF
If Text1.Text = fc!sri Or Val(Text1.Text) = Val(fc!sri) Then
fc!mot = Text5.Text
fc.Update
MsgBox " „ Õ›Ÿ «· €ÌÌ—", vbInformation
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Exit Sub
End If
fc.MoveNext
Loop
End Sub

Private Sub Command8_Click()
On Error Resume Next
grd1.Visible = False
grd2.Visible = False
If Option2.Value = True Then
Call chargegrd1_2_T
Else
Call chargegrd1_2_M
End If
grd1.Visible = True
grd2.Visible = True

End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0
Call MakeTreeViewRTL
Call chargetreeview1
Call couleur_treeview1
End Sub
Private Sub chargegrd1_2_T()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim m As Double
Dim e As Double
Dim se As Double
Dim P As Double
Dim sp As Double
Dim s As Double
grd1.Clear
grd1.Cols = 4
grd1.Rows = 1
grd1.ColWidth(0) = 1200
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 2000
grd1.ColWidth(3) = 4100
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.Row = 0
grd1.Col = 0
grd1.Text = "«· «—ÌŒ"
grd1.Col = 1
grd1.Text = "«·”«⁄…"
grd1.Col = 2
grd1.Text = "«·„»·€"
grd1.Col = 3
grd1.Text = "«· ›«’Ì·"
grd2.Clear
grd2.Cols = 4
grd2.Rows = 1
grd2.ColWidth(0) = 1200
grd2.ColWidth(1) = 1200
grd2.ColWidth(2) = 2000
grd2.ColWidth(3) = 4100
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.Row = 0
grd2.Col = 0
grd2.Text = "«· «—ÌŒ"
grd2.Col = 1
grd2.Text = "«·”«⁄…"
grd2.Col = 2
grd2.Text = "«·„»·€"
grd2.Col = 3
grd2.Text = "«· ›«’Ì·"
i = 1
j = 1
e = 0
se = 0
P = 0
sp = 0
s = 0
Call cont
grd1.Rows = cf.RecordCount + 3
grd2.Rows = cf.RecordCount + 3
Do While Not cf.EOF
If Text1.Text = cf!sri Or Val(Text1.Text) = Val(cf!sri) Then
e = 0
P = 0
If cf!typ = "”Õ» „»·€" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = cf!dat
grd2.Col = 1
grd2.Text = cf!heu
grd2.Col = 2
grd2.Text = cf!mon
grd2.Col = 3
grd2.Text = cf!det
e = cf!mon
se = se + e
i = i + 1
Else
grd1.Row = j
grd1.Col = 0
grd1.Text = cf!dat
grd1.Col = 1
grd1.Text = cf!heu
grd1.Col = 2
grd1.Text = cf!mon + " —« » ‘Â— " + cf!moi
grd1.Col = 3
grd1.Text = cf!det
P = cf!mon
sp = sp + P
j = j + 1
End If
End If
r = (sp - se)
cf.MoveNext
Loop
grd2.Rows = i
grd1.Rows = j
Label9.Caption = se
Label17.Caption = sp
Label22.Caption = r
Label15.Caption = r
End Sub
Private Sub chargegrd1_2_M()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim m As Double
Dim e As Double
Dim se As Double
Dim P As Double
Dim sp As Double
Dim s As Double
grd1.Clear
grd1.Cols = 4
grd1.Rows = 1
grd1.ColWidth(0) = 1200
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 2000
grd1.ColWidth(3) = 4100
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.Row = 0
grd1.Col = 0
grd1.Text = "«· «—ÌŒ"
grd1.Col = 1
grd1.Text = "«·”«⁄…"
grd1.Col = 2
grd1.Text = "«·„»·€"
grd1.Col = 3
grd1.Text = "«· ›«’Ì·"
grd2.Clear
grd2.Cols = 4
grd2.Rows = 1
grd2.ColWidth(0) = 1200
grd2.ColWidth(1) = 1200
grd2.ColWidth(2) = 2000
grd2.ColWidth(3) = 4100
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.Row = 0
grd2.Col = 0
grd2.Text = "«· «—ÌŒ"
grd2.Col = 1
grd2.Text = "«·”«⁄…"
grd2.Col = 2
grd2.Text = "«·„»·€"
grd2.Col = 3
grd2.Text = "«· ›«’Ì·"
i = 1
j = 1
e = 0
se = 0
P = 0
sp = 0
s = 0
Call cont
grd1.Rows = cf.RecordCount + 3
grd2.Rows = cf.RecordCount + 3
Do While Not cf.EOF
If Text1.Text = cf!sri Or Val(Text1.Text) = Val(cf!sri) Then
DT1.Value = cf!dat
m = DT1.Month
If m = Combo1.Text Then
e = 0
P = 0
If cf!typ = "”Õ» „»·€" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = cf!dat
grd2.Col = 1
grd2.Text = cf!heu
grd2.Col = 2
grd2.Text = cf!mon
grd2.Col = 3
grd2.Text = cf!det
e = cf!mon
se = se + e
i = i + 1
Else
grd1.Row = j
grd1.Col = 0
grd1.Text = cf!dat
grd1.Col = 1
grd1.Text = cf!heu
grd1.Col = 2
grd1.Text = cf!mon + " —« » ‘Â— " + cf!moi
grd1.Col = 3
grd1.Text = cf!det
P = cf!mon
sp = sp + P
j = j + 1
End If
End If
End If
r = (sp - se)
cf.MoveNext
Loop
grd2.Rows = i
grd1.Rows = j
Label9.Caption = se
Label17.Caption = sp
Label22.Caption = r
'Label15.Caption = r
End Sub

Private Sub Option1_Click()
On Error Resume Next
Combo1.Visible = True
Call tous_clear
End Sub

Private Sub Option2_Click()
On Error Resume Next
Combo1.Visible = False
Call tous_clear
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Picture2.Visible = True
Else
Text1.BackColor = &H8080FF
End If
Label37.Caption = ""
Label24.Caption = ""
Label26.Caption = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Text1_Click()
On Error Resume Next
Text1_Change
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

Private Sub Text6_Change()
On Error Resume Next
If Len(Text6.Text) > 0 Then
Text6.BackColor = &HC000&
Else
Text6.BackColor = &H8080FF
End If

End Sub

Private Sub Text6_Click()
On Error Resume Next
Text6_Change
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


Private Sub tous_clear()
On Error Resume Next
Label17.Caption = "0"
Label9.Caption = "0"
Label22.Caption = "0"
grd1.Clear
grd1.Cols = 4
grd1.Rows = 1
grd1.ColWidth(0) = 1500
grd1.ColWidth(1) = 1500
grd1.ColWidth(2) = 2000
grd1.ColWidth(3) = 5000
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.Row = 0
grd1.Col = 0
grd1.Text = "«· «—ÌŒ"
grd1.Col = 1
grd1.Text = "«·”«⁄…"
grd1.Col = 2
grd1.Text = "«·„»·€"
grd1.Col = 3
grd1.Text = "«· ›«’Ì·"
grd2.Clear
grd2.Cols = 4
grd2.Rows = 1
grd2.ColWidth(0) = 1500
grd2.ColWidth(1) = 1500
grd2.ColWidth(2) = 2000
grd2.ColWidth(3) = 5000
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.Row = 0
grd2.Col = 0
grd2.Text = "«· «—ÌŒ"
grd2.Col = 1
grd2.Text = "«·”«⁄…"
grd2.Col = 2
grd2.Text = "«·„»·€"
grd2.Col = 3
grd2.Text = "«· ›«’Ì·"

End Sub



