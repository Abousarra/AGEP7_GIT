VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Caisse_SLD 
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
      Left            =   6360
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   780
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "ÌÊ„Ì"
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
      Left            =   11880
      TabIndex        =   4
      Top             =   780
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H0000FFFF&
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
      ItemData        =   "Caisse_SLD.frx":0000
      Left            =   8640
      List            =   "Caisse_SLD.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "‘Â—Ì"
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
      Left            =   9360
      TabIndex        =   2
      Top             =   780
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "”‰ÊÌ"
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
      Left            =   7680
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabsPerRow      =   6
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
      TabCaption(0)   =   "”Ã· «·»‰ﬂ"
      TabPicture(0)   =   "Caisse_SLD.frx":0054
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture2"
      Tab(0).Control(1)=   "Command6"
      Tab(0).Control(2)=   "grd6"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "”Ã· «·„’—Ê›« "
      TabPicture(1)   =   "Caisse_SLD.frx":0070
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grd5"
      Tab(1).Control(1)=   "Command5"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "”Ã· «· ·«„Ì–"
      TabPicture(2)   =   "Caisse_SLD.frx":008C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command4"
      Tab(2).Control(1)=   "grd4"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "”Ã· «·√”« –…"
      TabPicture(3)   =   "Caisse_SLD.frx":00A8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command3"
      Tab(3).Control(1)=   "grd3"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "”Ã· «·„ÊŸ›Ì‰"
      TabPicture(4)   =   "Caisse_SLD.frx":00C4
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command2"
      Tab(4).Control(1)=   "grd2"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "”Ã· «·‘—ﬂ«¡"
      TabPicture(5)   =   "Caisse_SLD.frx":00E0
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "grd1"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Command1"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Picture1"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      Begin VB.PictureBox Picture2 
         Height          =   3015
         Left            =   -71640
         ScaleHeight     =   2955
         ScaleWidth      =   5475
         TabIndex        =   57
         Top             =   960
         Width           =   5535
         Begin MSFlexGridLib.MSFlexGrid grd10 
            Height          =   2775
            Left            =   240
            TabIndex        =   58
            Top             =   120
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   4895
            _Version        =   393216
            Appearance      =   0
         End
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         Height          =   345
         Left            =   -69360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5080
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         Height          =   345
         Left            =   -69360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   5080
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         Height          =   345
         Left            =   -69360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   5080
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         Height          =   345
         Left            =   -69360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   5080
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         Height          =   345
         Left            =   -69360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5080
         Width           =   1335
      End
      Begin VB.PictureBox Picture1 
         Height          =   2655
         Left            =   3600
         ScaleHeight     =   2595
         ScaleWidth      =   5115
         TabIndex        =   34
         Top             =   1320
         Visible         =   0   'False
         Width           =   5175
         Begin MSComCtl2.DTPicker DT4 
            Height          =   345
            Left            =   1680
            TabIndex        =   55
            Top             =   600
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
         Begin MSComCtl2.DTPicker DT3 
            Height          =   345
            Left            =   120
            TabIndex        =   54
            Top             =   600
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
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1440
            TabIndex        =   50
            Text            =   "Text1"
            Top             =   120
            Width           =   1815
         End
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   960
            TabIndex        =   35
            Text            =   "1"
            Top             =   1560
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DT2 
            Height          =   345
            Left            =   960
            TabIndex        =   36
            Top             =   1200
            Visible         =   0   'False
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
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   60
            Top             =   2160
            Width           =   3840
         End
         Begin VB.Label Label23 
            Caption         =   "Label23"
            Height          =   255
            Left            =   3360
            TabIndex        =   56
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Label22"
            Height          =   255
            Left            =   3360
            TabIndex        =   53
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label20 
            Caption         =   "Label20"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Label19"
            Height          =   375
            Left            =   3360
            TabIndex        =   51
            Top             =   120
            Width           =   1815
         End
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
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
         Height          =   345
         Left            =   5640
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   5080
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid grd1 
         Height          =   4695
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8281
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
      Begin MSFlexGridLib.MSFlexGrid grd2 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   39
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8281
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
      Begin MSFlexGridLib.MSFlexGrid grd3 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   41
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8281
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
      Begin MSFlexGridLib.MSFlexGrid grd4 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   43
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8281
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
      Begin MSFlexGridLib.MSFlexGrid grd5 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   45
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8281
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
      Begin MSFlexGridLib.MSFlexGrid grd6 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   47
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8281
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
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   345
      Left            =   10320
      TabIndex        =   5
      Top             =   780
      Visible         =   0   'False
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
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   59
      Top             =   780
      Width           =   3840
   End
   Begin VB.Label Label18 
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
      Left            =   6480
      TabIndex        =   49
      Top             =   2880
      Width           =   3000
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„”ÕÊ»«  «·»‰ﬂ"
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
      Left            =   10080
      TabIndex        =   48
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—’Ìœ ”«»ﬁ"
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
      Left            =   10080
      TabIndex        =   32
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label8 
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
      Left            =   6480
      TabIndex        =   31
      Top             =   3240
      Width           =   3000
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«Ìœ«⁄«  «·»‰ﬂ"
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
      Index           =   5
      Left            =   4560
      TabIndex        =   30
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   29
      Top             =   3240
      Width           =   3000
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   3
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„»«·€ «·Œ«—Ã…"
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
      Left            =   3720
      TabIndex        =   27
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„»«·€ «·œ«Œ·…"
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
      Left            =   10080
      TabIndex        =   26
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”Ã· «·’‰œÊﬁ"
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
      Left            =   4200
      TabIndex        =   25
      Top             =   0
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   2
      Left            =   7560
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      Height          =   2655
      Index           =   0
      Left            =   120
      Top             =   1320
      Width           =   12615
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   12720
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   6360
      X2              =   6360
      Y1              =   1320
      Y2              =   3960
   End
   Begin VB.Label Label1 
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
      Left            =   6480
      TabIndex        =   24
      Top             =   1800
      Width           =   3000
   End
   Begin VB.Label Label2 
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
      Left            =   6480
      TabIndex        =   23
      Top             =   2160
      Width           =   3000
   End
   Begin VB.Label Label3 
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
      Left            =   6480
      TabIndex        =   22
      Top             =   2520
      Width           =   3000
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   21
      Top             =   1800
      Width           =   3000
   End
   Begin VB.Label Label6 
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
      Left            =   240
      TabIndex        =   20
      Top             =   2160
      Width           =   3000
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
      Left            =   240
      TabIndex        =   19
      Top             =   2520
      Width           =   3000
   End
   Begin VB.Label Label9 
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
      Left            =   240
      TabIndex        =   18
      Top             =   2880
      Width           =   3000
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
      Left            =   6480
      TabIndex        =   17
      Top             =   3600
      Width           =   3000
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "≈Ìœ«⁄«  «·‘—ﬂ«¡"
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
      Left            =   9120
      TabIndex        =   16
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "≈Ìœ«⁄«  «·√”« –…"
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
      Left            =   10200
      TabIndex        =   15
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„œ›Ê⁄«  «· ·«„Ì–"
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
      Left            =   10080
      TabIndex        =   14
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„”ÕÊ»«  «·‘—ﬂ«¡"
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
      Left            =   4560
      TabIndex        =   13
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„”ÕÊ»«  «·√”« –…"
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
      Left            =   4560
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„”ÕÊ»«  «·„ÊŸ›Ì‰"
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
      Left            =   4560
      TabIndex        =   11
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„’—Ê›« "
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
      Left            =   4560
      TabIndex        =   10
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   12720
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label16 
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
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   3000
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «·Œ«—Ã"
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
      Left            =   4200
      TabIndex        =   8
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «·œ«Œ·"
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
      Left            =   10080
      TabIndex        =   7
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—’Ìœ «·’‰œÊﬁ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   6
      Top             =   780
      Width           =   2175
   End
End
Attribute VB_Name = "Caisse_SLD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pr1 As Double
Dim pf1 As Double
Dim et1 As Double
Dim bn1 As Double
Dim pr2 As Double
Dim pf2 As Double
Dim fn2 As Double
Dim bn2 As Double
Dim dp2 As Double
Dim tl1 As Double
Dim tl2 As Double
Private Sub Combo1_Change()
On Error Resume Next
Call grds_clear

End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Command7_Click()
On Error Resume Next
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  ⁄·Ï «·Ì„Ì‰", vbCritical
Exit Sub
End If
Command7.Enabled = False
grd1.Visible = False
grd2.Visible = False
grd3.Visible = False
grd4.Visible = False
grd5.Visible = False
grd6.Visible = False
Call grds_clear
If Option1.Value = True Then
Call chargegrd1_D
Call chargegrd2_D
Call chargegrd3_D
Call chargegrd4_D
Call chargegrd5_D
Call chargegrd6_D

End If
If Option2.Value = True Then
If Combo1.Text = "" Then
grd1.Visible = True
grd2.Visible = True
grd3.Visible = True
grd4.Visible = True
grd5.Visible = True
grd6.Visible = True
MsgBox "ﬁ„ » ÕœÌœ «·‘Â— „‰ Œ·«· «·ﬁ«∆„… «·„‰”œ·…", vbCritical
Command7.Enabled = True
Exit Sub
End If
Text1.Text = eb!ann
Label20.Caption = eb!moi
vg = Mid$(Text1.Text, 1, 4)
Label19.Caption = vg
vg = Mid$(Text1.Text, 6, 9)
Label22.Caption = vg
If Val(Combo1.Text) < Val(Label20.Caption) Then
Label23.Caption = Label22.Caption
Else
Label23.Caption = Label19.Caption
End If
DT3.Value = "01/" & Label20.Caption & "/" & Label19.Caption
DT4.Value = "01/" & Combo1.Text & "/" & Label23.Caption
Call chargegrd1_M
Call chargegrd2_M
Call chargegrd3_M
Call chargegrd4_M
Call chargegrd5_M
Call chargegrd6_M

End If
If Option3.Value = True Then
Call chargegrd1_T
Call chargegrd2_T
Call chargegrd3_T
Call chargegrd4_T
Call chargegrd5_T
Call chargegrd6_T
End If
Call solde_T
grd1.Visible = True
grd2.Visible = True
grd3.Visible = True
grd4.Visible = True
grd5.Visible = True
grd6.Visible = True
Command7.Enabled = True
End Sub

Private Sub DT1_Change()
On Error Resume Next
Call grds_clear

End Sub

Private Sub DT1_Click()
On Error Resume Next
DT1_Change
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0
Call grds_clear
Label24.Caption = eb!cca
DT1.Value = Date
End Sub


Private Sub Option1_Click()
On Error Resume Next
DT1.Visible = True
Combo1.Visible = False
Call grds_clear
End Sub

Private Sub Option2_Click()
On Error Resume Next
Call grds_clear
DT1.Visible = False
Combo1.Visible = True

End Sub

Private Sub Option3_Click()
On Error Resume Next
Call grds_clear
DT1.Visible = False
Combo1.Visible = False
End Sub
Private Sub chargegrd1_T()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd1.Clear
grd1.Cols = 7
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 1000
grd1.ColWidth(3) = 800
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 0
grd1.ColWidth(6) = 7000
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«· «—ÌŒ"
grd1.Col = 2
grd1.Text = "«·”«⁄…"
grd1.Col = 3
grd1.Text = "«·‰Ê⁄"
grd1.Col = 4
grd1.Text = "«·„»·€"
grd1.Col = 5
grd1.Text = "«· ›«’Ì·"
grd1.Col = 6
grd1.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
pr1 = 0
pr2 = 0

Call cont
grd1.Rows = cp.RecordCount + 3
Do While Not cp.EOF
If cp!act <> Combo2.Text Then
dat3 = cp!dat
'If dat3 >= dat1 And dat3 <= dat2 Then
grd1.Row = i
grd1.Col = 0
grd1.Text = cp!aut
grd1.Col = 1
grd1.Text = cp!dat
grd1.Col = 2
grd1.Text = cp!heu
grd1.Col = 3
grd1.Text = cp!typ
If cp!typ = "”Õ»" Then
a = cp!mon
P = P + a
Else
a = cp!mon
r = r + a
End If
grd1.Col = 4
grd1.Text = cp!mon
grd1.Col = 5
grd1.Text = cp!det
grd1.Col = 6
grd1.Text = cp!typ + " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cp!sri
i = i + 1
End If
cp.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 2
s = (P - r)
Label5.Caption = P
Label1.Caption = r
End Sub
Private Sub chargegrd1_D()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd1.Clear
grd1.Cols = 7
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 1000
grd1.ColWidth(3) = 800
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 0
grd1.ColWidth(6) = 7000
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«· «—ÌŒ"
grd1.Col = 2
grd1.Text = "«·”«⁄…"
grd1.Col = 3
grd1.Text = "«·‰Ê⁄"
grd1.Col = 4
grd1.Text = "«·„»·€"
grd1.Col = 5
grd1.Text = "«· ›«’Ì·"
grd1.Col = 6
grd1.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
a = 0
pr1 = 0
pr2 = 0
dat1 = DT1.Value
Call cont
grd1.Rows = cp.RecordCount + 3
Do While Not cp.EOF
dat2 = cp!dat
If cp!act = Combo2.Text Then
If dat2 < dat1 Then
If cp!typ = "”Õ»" Then
a = cp!mon
pr2 = pr2 + a
Else
a = cp!mon
pr1 = pr1 + a
End If
End If
End If
If cp!act <> Combo2.Text Then
If dat2 = dat1 Then
grd1.Row = i
grd1.Col = 0
grd1.Text = cp!aut
grd1.Col = 1
grd1.Text = cp!dat
grd1.Col = 2
grd1.Text = cp!heu
grd1.Col = 3
grd1.Text = cp!typ
If cp!typ = "”Õ»" Then
a = cp!mon
P = P + a
Else
a = cp!mon
r = r + a
End If
grd1.Col = 4
grd1.Text = cp!mon
grd1.Col = 5
grd1.Text = cp!det
grd1.Col = 6
grd1.Text = cp!typ + " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cp!sri
i = i + 1
End If
End If
cp.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 2
s = (P - r)
Label5.Caption = P
Label1.Caption = r
End Sub
Private Sub chargegrd1_M()
On Error Resume Next
Dim j1 As Double
Dim j2 As Double
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd1.Clear
grd1.Cols = 7
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 1000
grd1.ColWidth(3) = 800
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 0
grd1.ColWidth(6) = 7000
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«· «—ÌŒ"
grd1.Col = 2
grd1.Text = "«·”«⁄…"
grd1.Col = 3
grd1.Text = "«·‰Ê⁄"
grd1.Col = 4
grd1.Text = "«·„»·€"
grd1.Col = 5
grd1.Text = "«· ›«’Ì·"
grd1.Col = 6
grd1.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
pr1 = 0
pr2 = 0
j1 = Combo1.Text
Call cont
grd1.Rows = cp.RecordCount + 3
Do While Not cp.EOF
If cp!act <> Combo2.Text Then
dat1 = DT3.Value
dat2 = DT4.Value
dat3 = cp!dat
If dat3 >= dat1 And dat3 < dat2 Then
If cp!typ = "”Õ»" Then
a = cp!mon
pr2 = pr2 + a
Else
a = cp!mon
pr1 = pr1 + a
End If
End If
DT2.Value = cp!dat
j2 = DT2.Month
If j1 = j2 Then
grd1.Row = i
grd1.Col = 0
grd1.Text = cp!aut
grd1.Col = 1
grd1.Text = cp!dat
grd1.Col = 2
grd1.Text = cp!heu
grd1.Col = 3
grd1.Text = cp!typ
If cp!typ = "”Õ»" Then
a = cp!mon
P = P + a
Else
a = cp!mon
r = r + a
End If
grd1.Col = 4
grd1.Text = cp!mon
grd1.Col = 5
grd1.Text = cp!det
grd1.Col = 6
grd1.Text = cp!typ + " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cp!sri
i = i + 1
End If
End If
cp.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 2
s = (P - r)
Label5.Caption = P
Label1.Caption = r
End Sub
Private Sub chargegrd2_T()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd2.Clear
grd2.Cols = 7
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1200
grd2.ColWidth(2) = 1000
grd2.ColWidth(3) = 1000
grd2.ColWidth(4) = 1800
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 7000
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
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
grd2.Col = 6
grd2.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
fn2 = 0
Call cont
grd2.Rows = cf.RecordCount + 3
Do While Not cf.EOF
If cf!act <> Combo2.Text Then
dat3 = cf!dat
If cf!typ = "”Õ» „»·€" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = cf!aut
grd2.Col = 1
grd2.Text = cf!dat
grd2.Col = 2
grd2.Text = cf!heu
grd2.Col = 3
grd2.Text = cf!typ
a = cf!mon
P = P + a
grd2.Col = 4
grd2.Text = cf!mon
grd2.Col = 5
grd2.Text = cf!det
grd2.Col = 6
grd2.Text = cf!typ + " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cf!sri
i = i + 1
End If
End If
cf.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 2
s = (P - r)
Label7.Caption = P
End Sub
Private Sub chargegrd2_D()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd2.Clear
grd2.Cols = 7
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1200
grd2.ColWidth(2) = 1000
grd2.ColWidth(3) = 1000
grd2.ColWidth(4) = 1800
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 7000
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
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
grd2.Col = 6
grd2.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
fn2 = 0
dat1 = DT1.Value
Call cont
grd2.Rows = cf.RecordCount + 3
Do While Not cf.EOF
dat2 = cf!dat
If cf!act = Combo2.Text Then
If dat2 < dat1 Then
If cf!typ = "”Õ» „»·€" Then
a = cf!mon
fn2 = fn2 + a
End If
End If
End If
If cf!act <> Combo2.Text Then
If dat2 = dat1 Then
If cf!typ = "”Õ» „»·€" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = cf!aut
grd2.Col = 1
grd2.Text = cf!dat
grd2.Col = 2
grd2.Text = cf!heu
grd2.Col = 3
grd2.Text = cf!typ
a = cf!mon
P = P + a
grd2.Col = 4
grd2.Text = cf!mon
grd2.Col = 5
grd2.Text = cf!det
grd2.Col = 6
grd2.Text = cf!typ + " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cf!sri
i = i + 1
End If
End If
End If
cf.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 2
s = (P - r)
Label7.Caption = P
End Sub
Private Sub chargegrd2_M()
On Error Resume Next
Dim i As Double
Dim j1 As Double
Dim j2 As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd2.Clear
grd2.Cols = 7
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1200
grd2.ColWidth(2) = 1000
grd2.ColWidth(3) = 1000
grd2.ColWidth(4) = 1800
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 7000
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
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
grd2.Col = 6
grd2.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
fn2 = 0
j1 = Combo1.Text
Call cont
grd2.Rows = cf.RecordCount + 3
Do While Not cf.EOF
If cf!act <> Combo2.Text Then
dat1 = DT3.Value
dat2 = DT4.Value
dat3 = cf!dat
If dat3 >= dat1 And dat3 < dat2 Then
If cf!typ = "”Õ» „»·€" Then
a = cf!mon
fn2 = fn2 + a
End If
End If
DT2.Value = cf!dat
j2 = DT2.Month
If j1 = j2 Then
If cf!typ = "”Õ» „»·€" Then
grd2.Row = i
grd2.Col = 0
grd2.Text = cf!aut
grd2.Col = 1
grd2.Text = cf!dat
grd2.Col = 2
grd2.Text = cf!heu
grd2.Col = 3
grd2.Text = cf!typ
a = cf!mon
P = P + a
grd2.Col = 4
grd2.Text = cf!mon
grd2.Col = 5
grd2.Text = cf!det
grd2.Col = 6
grd2.Text = cf!typ + " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cf!sri
i = i + 1
End If
End If
End If
cf.MoveNext
Loop
grd2.Rows = i
grd2.Col = 1
grd2.Sort = 2
s = (P - r)
Label7.Caption = P
End Sub

Private Sub chargegrd3_T()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd3.Clear
grd3.Cols = 7
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1200
grd3.ColWidth(2) = 1000
grd3.ColWidth(3) = 800
grd3.ColWidth(4) = 2000
grd3.ColWidth(5) = 0
grd3.ColWidth(6) = 7000
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.ColAlignment(5) = 1
grd3.ColAlignment(6) = 1
grd3.Row = 0
grd3.Col = 1
grd3.Text = "«· «—ÌŒ"
grd3.Col = 2
grd3.Text = "«·”«⁄…"
grd3.Col = 3
grd3.Text = "«·‰Ê⁄"
grd3.Col = 4
grd3.Text = "«·„»·€"
grd3.Col = 5
grd3.Text = "«· ›«’Ì·"
grd3.Col = 6
grd3.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
pf1 = 0
pf2 = 0
Call cont
grd3.Rows = cs.RecordCount + 3
Do While Not cs.EOF
If cs!act <> Combo2.Text Then
dat3 = cs!dat
'If dat3 >= dat1 And dat3 <= dat2 Then
grd3.Row = i
grd3.Col = 0
grd3.Text = cs!aut
grd3.Col = 1
grd3.Text = cs!dat
grd3.Col = 2
grd3.Text = cs!heu
grd3.Col = 3
grd3.Text = cs!typ
If cs!typ = "”Õ»" Then
a = cs!mon
P = P + a
Else
a = cs!mon
r = r + a
End If
grd3.Col = 4
grd3.Text = cs!mon
grd3.Col = 5
grd3.Text = cs!det
grd3.Col = 6
grd3.Text = cs!typ + " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cs!sri
i = i + 1
End If
cs.MoveNext
Loop
grd3.Rows = i
grd3.Col = 1
grd3.Sort = 2
s = (P - r)
Label6.Caption = P
Label2.Caption = r
End Sub
Private Sub chargegrd3_D()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd3.Clear
grd3.Cols = 7
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1200
grd3.ColWidth(2) = 1000
grd3.ColWidth(3) = 800
grd3.ColWidth(4) = 2000
grd3.ColWidth(5) = 0
grd3.ColWidth(6) = 7000
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.ColAlignment(5) = 1
grd3.ColAlignment(6) = 1
grd3.Row = 0
grd3.Col = 1
grd3.Text = "«· «—ÌŒ"
grd3.Col = 2
grd3.Text = "«·”«⁄…"
grd3.Col = 3
grd3.Text = "«·‰Ê⁄"
grd3.Col = 4
grd3.Text = "«·„»·€"
grd3.Col = 5
grd3.Text = "«· ›«’Ì·"
grd3.Col = 6
grd3.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
pf1 = 0
pf2 = 0
dat1 = DT1.Value
Call cont
grd3.Rows = cs.RecordCount + 3
Do While Not cs.EOF
dat2 = cs!dat
If cs!act = Combo2.Text Then
If dat2 < dat1 Then
If cs!typ = "”Õ»" Then
a = cs!mon
pf2 = pf2 + a
Else
a = cs!mon
pf1 = pf1 + a
End If
End If
End If
If cs!act <> Combo2.Text Then
If dat2 = dat1 Then
grd3.Row = i
grd3.Col = 0
grd3.Text = cs!aut
grd3.Col = 1
grd3.Text = cs!dat
grd3.Col = 2
grd3.Text = cs!heu
grd3.Col = 3
grd3.Text = cs!typ
If cs!typ = "”Õ»" Then
a = cs!mon
P = P + a
Else
a = cs!mon
r = r + a
End If
grd3.Col = 4
grd3.Text = cs!mon
grd3.Col = 5
grd3.Text = cs!det
grd3.Col = 6
grd3.Text = cs!typ + " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cs!sri
i = i + 1
End If
End If
cs.MoveNext
Loop
grd3.Rows = i
grd3.Col = 1
grd3.Sort = 2
s = (P - r)
Label6.Caption = P
Label2.Caption = r
End Sub
Private Sub chargegrd3_M()
On Error Resume Next
Dim i As Double
Dim j1 As Double
Dim j2 As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd3.Clear
grd3.Cols = 7
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1200
grd3.ColWidth(2) = 1000
grd3.ColWidth(3) = 800
grd3.ColWidth(4) = 2000
grd3.ColWidth(5) = 0
grd3.ColWidth(6) = 7000
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.ColAlignment(5) = 1
grd3.ColAlignment(6) = 1
grd3.Row = 0
grd3.Col = 1
grd3.Text = "«· «—ÌŒ"
grd3.Col = 2
grd3.Text = "«·”«⁄…"
grd3.Col = 3
grd3.Text = "«·‰Ê⁄"
grd3.Col = 4
grd3.Text = "«·„»·€"
grd3.Col = 5
grd3.Text = "«· ›«’Ì·"
grd3.Col = 6
grd3.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
pf1 = 0
pf2 = 0
j1 = Combo1.Text
Call cont
grd3.Rows = cs.RecordCount + 3
Do While Not cs.EOF
If cs!act <> Combo2.Text Then
dat1 = DT3.Value
dat2 = DT4.Value
dat3 = cs!dat
If dat3 >= dat1 And dat3 < dat2 Then
If cs!typ = "”Õ»" Then
a = cs!mon
pf2 = pf2 + a
Else
a = cs!mon
pf1 = pf1 + a
End If
End If
DT2.Value = cs!dat
j2 = DT2.Month
If j1 = j2 Then
grd3.Row = i
grd3.Col = 0
grd3.Text = cs!aut
grd3.Col = 1
grd3.Text = cs!dat
grd3.Col = 2
grd3.Text = cs!heu
grd3.Col = 3
grd3.Text = cs!typ
If cs!typ = "”Õ»" Then
a = cs!mon
P = P + a
Else
a = cs!mon
r = r + a
End If
grd3.Col = 4
grd3.Text = cs!mon
grd3.Col = 5
grd3.Text = cs!det
grd3.Col = 6
grd3.Text = cs!typ + " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cs!sri
i = i + 1
End If
End If
cs.MoveNext
Loop
grd3.Rows = i
grd3.Col = 1
grd3.Sort = 2
s = (P - r)
Label6.Caption = P
Label2.Caption = r
End Sub
Private Sub chargegrd4_T()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim b As Double
Dim c As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd4.Clear
grd4.Cols = 8
grd4.Rows = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1200
grd4.ColWidth(2) = 1000
grd4.ColWidth(3) = 1200
grd4.ColWidth(4) = 1200
grd4.ColWidth(5) = 1200
grd4.ColWidth(6) = 1200
grd4.ColWidth(7) = 5000
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.ColAlignment(5) = 1
grd4.ColAlignment(6) = 1
grd4.ColAlignment(7) = 1
grd4.Row = 0
grd4.Col = 1
grd4.Text = "«· «—ÌŒ"
grd4.Col = 2
grd4.Text = "—”Ê„"
grd4.Col = 3
grd4.Text = "«·Ê’·"
grd4.Col = 4
grd4.Text = "«·„” Õﬁ"
grd4.Col = 5
grd4.Text = "«·„œ›Ê⁄"
grd4.Col = 6
grd4.Text = "«·»«ﬁÌ"
grd4.Col = 7
grd4.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
et1 = 0
Call cont
grd4.Rows = ct.RecordCount + 3
Do While Not ct.EOF
If ct!act <> Combo2.Text Then
dat3 = ct!dat
'If dat3 >= dat1 And dat3 <= dat2 Then
If ct!rcu <> "0" Then
grd4.Row = i
grd4.Col = 0
grd4.Text = ct!aut
grd4.Col = 1
grd4.Text = ct!dat
grd4.Col = 2
grd4.Text = ct!mois
grd4.Col = 3
grd4.Text = ct!rec
a = ct!tpy
P = P + a
b = ct!tpy
c = ct!trs
c = (b + c)
grd4.Col = 4
grd4.Text = c
grd4.Col = 5
grd4.Text = ct!tpy
grd4.Col = 6
grd4.Text = ct!trs
grd4.Col = 7
grd4.Text = "  „ œ›⁄Â „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + ct!sri
i = i + 1
End If
End If
ct.MoveNext
Loop
grd4.Rows = i
grd4.Col = 1
grd4.Sort = 2
Label3.Caption = P
End Sub
Private Sub chargegrd4_D()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim b As Double
Dim c As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd4.Clear
grd4.Cols = 8
grd4.Rows = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1200
grd4.ColWidth(2) = 1000
grd4.ColWidth(3) = 1200
grd4.ColWidth(4) = 1200
grd4.ColWidth(5) = 1200
grd4.ColWidth(6) = 1200
grd4.ColWidth(7) = 5000
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.ColAlignment(5) = 1
grd4.ColAlignment(6) = 1
grd4.ColAlignment(7) = 1
grd4.Row = 0
grd4.Col = 1
grd4.Text = "«· «—ÌŒ"
grd4.Col = 2
grd4.Text = "—”Ê„"
grd4.Col = 3
grd4.Text = "«·Ê’·"
grd4.Col = 4
grd4.Text = "«·„” Õﬁ"
grd4.Col = 5
grd4.Text = "«·„œ›Ê⁄"
grd4.Col = 6
grd4.Text = "«·»«ﬁÌ"
grd4.Col = 7
grd4.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
et1 = 0
dat1 = DT1.Value
Call cont
grd4.Rows = ct.RecordCount + 3
Do While Not ct.EOF
dat2 = ct!dat
If ct!act = Combo2.Text Then
If dat2 < dat1 Then
a = ct!tpy
et1 = et1 + a
End If
End If
If ct!act <> Combo2.Text Then
If dat1 = dat2 Then
If ct!rcu <> "0" Then
grd4.Row = i
grd4.Col = 0
grd4.Text = ct!aut
grd4.Col = 1
grd4.Text = ct!dat
grd4.Col = 2
grd4.Text = ct!mois
grd4.Col = 3
grd4.Text = ct!rec
a = ct!tpy
P = P + a
b = ct!tpy
c = ct!trs
c = (b + c)
grd4.Col = 4
grd4.Text = c
grd4.Col = 5
grd4.Text = ct!tpy
grd4.Col = 6
grd4.Text = ct!trs
grd4.Col = 7
grd4.Text = "  „ œ›⁄Â „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + ct!sri
i = i + 1
End If
End If
End If
ct.MoveNext
Loop
grd4.Rows = i
grd4.Col = 1
grd4.Sort = 2
Label3.Caption = P
End Sub
Private Sub chargegrd4_M()
On Error Resume Next
Dim i As Double
Dim j1 As Double
Dim j2 As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim b As Double
Dim c As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd4.Clear
grd4.Cols = 8
grd4.Rows = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1200
grd4.ColWidth(2) = 1000
grd4.ColWidth(3) = 1200
grd4.ColWidth(4) = 1200
grd4.ColWidth(5) = 1200
grd4.ColWidth(6) = 1200
grd4.ColWidth(7) = 5000
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.ColAlignment(5) = 1
grd4.ColAlignment(6) = 1
grd4.ColAlignment(7) = 1
grd4.Row = 0
grd4.Col = 1
grd4.Text = "«· «—ÌŒ"
grd4.Col = 2
grd4.Text = "—”Ê„"
grd4.Col = 3
grd4.Text = "«·Ê’·"
grd4.Col = 4
grd4.Text = "«·„” Õﬁ"
grd4.Col = 5
grd4.Text = "«·„œ›Ê⁄"
grd4.Col = 6
grd4.Text = "«·»«ﬁÌ"
grd4.Col = 7
grd4.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
et1 = 0
j1 = Combo1.Text
Call cont
grd4.Rows = ct.RecordCount + 3
Do While Not ct.EOF
If ct!act <> Combo2.Text Then
dat1 = DT3.Value
dat2 = DT4.Value
dat3 = ct!dat
If dat3 >= dat1 And dat3 < dat2 Then
a = ct!tpy
et1 = et1 + a
End If
DT2.Value = ct!dat
j2 = DT2.Month
If j1 = j2 Then
If ct!rcu <> "0" Then
grd4.Row = i
grd4.Col = 0
grd4.Text = ct!aut
grd4.Col = 1
grd4.Text = ct!dat
grd4.Col = 2
grd4.Text = ct!mois
grd4.Col = 3
grd4.Text = ct!rec
a = ct!tpy
P = P + a
b = ct!tpy
c = ct!trs
c = (b + c)
grd4.Col = 4
grd4.Text = c
grd4.Col = 5
grd4.Text = ct!tpy
grd4.Col = 6
grd4.Text = ct!trs
grd4.Col = 7
grd4.Text = "  „ œ›⁄Â „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + ct!sri
i = i + 1
End If
End If
End If
ct.MoveNext
Loop
grd4.Rows = i
grd4.Col = 1
grd4.Sort = 2
Label3.Caption = P
End Sub
Private Sub chargegrd5_T()
On Error Resume Next
Dim a As Double
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim d As Double
Dim sd As Double
grd5.Clear
grd5.Cols = 5
grd5.Rows = 1
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 1500
grd5.ColWidth(2) = 1500
grd5.ColWidth(3) = 2000
grd5.ColWidth(4) = 7000
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 1
grd5.ColAlignment(2) = 1
grd5.ColAlignment(3) = 1
grd5.ColAlignment(4) = 1
grd5.Row = 0
grd5.Col = 1
grd5.Text = "«· «—ÌŒ"
grd5.Col = 2
grd5.Text = "«·”«⁄…"
grd5.Col = 3
grd5.Text = "«·„»·€"
grd5.Col = 4
grd5.Text = " ›«’Ì·"
i = 1
a = 0
sd = 0
dp2 = 0
Call cont
grd5.Rows = dp.RecordCount + 3
Do While Not dp.EOF
If dp!act <> Combo2.Text Then
dat3 = dp!dat
grd5.Row = i
grd5.Col = 0
grd5.Text = dp!aut
grd5.Col = 1
grd5.Text = dp!dat
grd5.Col = 2
grd5.Text = dp!heu
grd5.Col = 3
grd5.Text = dp!mon
d = dp!mon
sd = sd + d
grd5.Col = 4
grd5.Text = dp!det
i = i + 1
End If
dp.MoveNext
Loop
grd5.Rows = i
Label9.Caption = sd
End Sub
Private Sub chargegrd5_D()
On Error Resume Next
Dim a As Double
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim d As Double
Dim sd As Double
grd5.Clear
grd5.Cols = 5
grd5.Rows = 1
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 1500
grd5.ColWidth(2) = 1500
grd5.ColWidth(3) = 2000
grd5.ColWidth(4) = 7000
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 1
grd5.ColAlignment(2) = 1
grd5.ColAlignment(3) = 1
grd5.ColAlignment(4) = 1
grd5.Row = 0
grd5.Col = 1
grd5.Text = "«· «—ÌŒ"
grd5.Col = 2
grd5.Text = "«·”«⁄…"
grd5.Col = 3
grd5.Text = "«·„»·€"
grd5.Col = 4
grd5.Text = " ›«’Ì·"
i = 1
a = 0
sd = 0
dp2 = 0
dat1 = DT1.Value
Call cont
grd5.Rows = dp.RecordCount + 3
Do While Not dp.EOF
dat2 = dp!dat
If dp!act = Combo2.Text Then
If dat2 < dat1 Then
a = dp!mon
dp2 = dp2 + a
End If
End If
If dp!act <> Combo2.Text Then
If dat1 = dat2 Then
grd5.Row = i
grd5.Col = 0
grd5.Text = dp!aut
grd5.Col = 1
grd5.Text = dp!dat
grd5.Col = 2
grd5.Text = dp!heu
grd5.Col = 3
grd5.Text = dp!mon
d = dp!mon
sd = sd + d
grd5.Col = 4
grd5.Text = dp!det
i = i + 1
End If
End If
dp.MoveNext
Loop
grd5.Rows = i
Label9.Caption = sd
End Sub
Private Sub chargegrd5_M()
On Error Resume Next
Dim a As Double
Dim i As Double
Dim j1 As Double
Dim j2 As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim d As Double
Dim sd As Double
grd5.Clear
grd5.Cols = 5
grd5.Rows = 1
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 1500
grd5.ColWidth(2) = 1500
grd5.ColWidth(3) = 2000
grd5.ColWidth(4) = 7000
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 1
grd5.ColAlignment(2) = 1
grd5.ColAlignment(3) = 1
grd5.ColAlignment(4) = 1
grd5.Row = 0
grd5.Col = 1
grd5.Text = "«· «—ÌŒ"
grd5.Col = 2
grd5.Text = "«·”«⁄…"
grd5.Col = 3
grd5.Text = "«·„»·€"
grd5.Col = 4
grd5.Text = " ›«’Ì·"
i = 1
a = 0
sd = 0
dp2 = 0
j1 = Combo1.Text
Call cont
grd5.Rows = dp.RecordCount + 3
Do While Not dp.EOF
If dp!act <> Combo2.Text Then
dat1 = DT3.Value
dat2 = DT4.Value
dat3 = dp!dat
If dat3 >= dat1 And dat3 < dat2 Then
a = dp!mon
dp2 = dp2 + a
End If
DT2.Value = dp!dat
j2 = DT2.Month
If j1 = j2 Then
grd5.Row = i
grd5.Col = 0
grd5.Text = dp!aut
grd5.Col = 1
grd5.Text = dp!dat
grd5.Col = 2
grd5.Text = dp!heu
grd5.Col = 3
grd5.Text = dp!mon
d = dp!mon
sd = sd + d
grd5.Col = 4
grd5.Text = dp!det
i = i + 1
End If
End If
dp.MoveNext
Loop
grd5.Rows = i
Label9.Caption = sd
End Sub
Private Sub chargegrd6_T()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd6.Clear
grd6.Cols = 7
grd6.Rows = 1
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 1200
grd6.ColWidth(2) = 1000
grd6.ColWidth(3) = 800
grd6.ColWidth(4) = 2000
grd6.ColWidth(5) = 0
grd6.ColWidth(6) = 7000
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.ColAlignment(4) = 1
grd6.ColAlignment(5) = 1
grd6.ColAlignment(6) = 1
grd6.Row = 0
grd6.Col = 1
grd6.Text = "«· «—ÌŒ"
grd6.Col = 2
grd6.Text = "«·”«⁄…"
grd6.Col = 3
grd6.Text = "«·‰Ê⁄"
grd6.Col = 4
grd6.Text = "«·„»·€"
grd6.Col = 5
grd6.Text = "«· ›«’Ì·"
grd6.Col = 6
grd6.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
bn1 = 0
bn2 = 0
Call cont
grd6.Rows = bn.RecordCount + 3
Do While Not bn.EOF
If bn!act <> Combo2.Text Then
dat3 = bn!dat
'If dat3 >= dat1 And dat3 <= dat2 Then
grd6.Row = i
grd6.Col = 0
grd6.Text = bn!aut
grd6.Col = 1
grd6.Text = bn!dat
grd6.Col = 2
grd6.Text = bn!heu
grd6.Col = 3
grd6.Text = bn!typ
If bn!typ = "”Õ»" Then
a = bn!mon
P = P + a
Else
a = bn!mon
r = r + a
End If
grd6.Col = 4
grd6.Text = bn!mon
grd6.Col = 5
grd6.Text = bn!det
grd6.Col = 6
grd6.Text = bn!det
i = i + 1
End If
bn.MoveNext
Loop
grd6.Rows = i
grd6.Col = 1
grd6.Sort = 2
s = (P - r)
Label8.Caption = P
Label4.Caption = r
End Sub
Private Sub chargegrd6_D()
On Error Resume Next
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd6.Clear
grd6.Cols = 7
grd6.Rows = 1
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 1200
grd6.ColWidth(2) = 1000
grd6.ColWidth(3) = 800
grd6.ColWidth(4) = 2000
grd6.ColWidth(5) = 0
grd6.ColWidth(6) = 7000
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.ColAlignment(4) = 1
grd6.ColAlignment(5) = 1
grd6.ColAlignment(6) = 1
grd6.Row = 0
grd6.Col = 1
grd6.Text = "«· «—ÌŒ"
grd6.Col = 2
grd6.Text = "«·”«⁄…"
grd6.Col = 3
grd6.Text = "«·‰Ê⁄"
grd6.Col = 4
grd6.Text = "«·„»·€"
grd6.Col = 5
grd6.Text = "«· ›«’Ì·"
grd6.Col = 6
grd6.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
bn1 = 0
bn2 = 0
dat1 = DT1.Value
Call cont
grd6.Rows = bn.RecordCount + 3
Do While Not bn.EOF
dat2 = bn!dat
If bn!act = Combo2.Text Then
If dat2 < dat1 Then
If bn!typ = "”Õ»" Then
a = bn!mon
bn1 = bn1 + a
Else
a = bn!mon
bn2 = bn2 + a
End If
End If
End If
If bn!act <> Combo2.Text Then
If dat2 = dat1 Then
grd6.Row = i
grd6.Col = 0
grd6.Text = bn!aut
grd6.Col = 1
grd6.Text = bn!dat
grd6.Col = 2
grd6.Text = bn!heu
grd6.Col = 3
grd6.Text = bn!typ
If bn!typ = "”Õ»" Then
a = bn!mon
P = P + a
Else
a = bn!mon
r = r + a
End If
grd6.Col = 4
grd6.Text = bn!mon
grd6.Col = 5
grd6.Text = bn!det
grd6.Col = 6
grd6.Text = bn!det
i = i + 1
End If
End If
bn.MoveNext
Loop
grd6.Rows = i
grd6.Col = 1
grd6.Sort = 2
s = (P - r)
Label8.Caption = P
Label4.Caption = r
End Sub
Private Sub chargegrd6_M()
On Error Resume Next
Dim i As Double
Dim j1 As Double
Dim j2 As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd6.Clear
grd6.Cols = 7
grd6.Rows = 1
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 1200
grd6.ColWidth(2) = 1000
grd6.ColWidth(3) = 800
grd6.ColWidth(4) = 2000
grd6.ColWidth(5) = 0
grd6.ColWidth(6) = 7000
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.ColAlignment(4) = 1
grd6.ColAlignment(5) = 1
grd6.ColAlignment(6) = 1
grd6.Row = 0
grd6.Col = 1
grd6.Text = "«· «—ÌŒ"
grd6.Col = 2
grd6.Text = "«·”«⁄…"
grd6.Col = 3
grd6.Text = "«·‰Ê⁄"
grd6.Col = 4
grd6.Text = "«·„»·€"
grd6.Col = 5
grd6.Text = "«· ›«’Ì·"
grd6.Col = 6
grd6.Text = " ›«’Ì·"
i = 1
P = 0
r = 0
s = 0
bn1 = 0
bn2 = 0
j1 = Combo1.Text
Call cont
grd6.Rows = bn.RecordCount + 3
Do While Not bn.EOF
If bn!act <> Combo2.Text Then
dat1 = DT3.Value
dat2 = DT4.Value
dat3 = bn!dat
If dat3 >= dat1 And dat3 < dat2 Then
If bn!typ = "”Õ»" Then
a = bn!mon
bn1 = bn1 + a
Else
a = bn!mon
bn2 = bn2 + a
End If
End If
DT2.Value = bn!dat
j2 = DT2.Month
If j1 = j2 Then
grd6.Row = i
grd6.Col = 0
grd6.Text = bn!aut
grd6.Col = 1
grd6.Text = bn!dat
grd6.Col = 2
grd6.Text = bn!heu
grd6.Col = 3
grd6.Text = bn!typ
If bn!typ = "”Õ»" Then
a = bn!mon
P = P + a
Else
a = bn!mon
r = r + a
End If
grd6.Col = 4
grd6.Text = bn!mon
grd6.Col = 5
grd6.Text = bn!det
grd6.Col = 6
grd6.Text = bn!det
i = i + 1
End If
End If
bn.MoveNext
Loop
grd6.Rows = i
grd6.Col = 1
grd6.Sort = 2
s = (P - r)
Label8.Caption = P
Label4.Caption = r
End Sub

Private Sub solde_T()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim f As Double
Dim g As Double
Dim h As Double
Dim i As Double
Dim j As Double
Dim s1 As Double
Dim s2 As Double
tl1 = (pr1 + pf1 + et1 + bn1)
tl2 = (pr2 + pf2 + fn2 + bn2 + dp2)
Label18.Caption = (tl1 - tl2)
Call chargegrd10
a = Label1.Caption
b = Label2.Caption
c = Label3.Caption
d = Label8.Caption
j = Label18.Caption
s1 = a + b + c + d + j
Label10.Caption = s1
e = Label5.Caption
f = Label6.Caption
g = Label7.Caption
h = Label4.Caption
i = Label9.Caption
s2 = e + f + g + h + i
Label16.Caption = s2
Label17.Caption = (s1 - s2)
End Sub
Private Sub grds_clear()
On Error Resume Next
grd1.Clear
grd1.Cols = 7
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 1000
grd1.ColWidth(3) = 800
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 0
grd1.ColWidth(6) = 7000
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«· «—ÌŒ"
grd1.Col = 2
grd1.Text = "«·”«⁄…"
grd1.Col = 3
grd1.Text = "«·‰Ê⁄"
grd1.Col = 4
grd1.Text = "«·„»·€"
grd1.Col = 5
grd1.Text = "«· ›«’Ì·"
grd1.Col = 6
grd1.Text = " ›«’Ì·"
grd2.Clear
grd2.Cols = 7
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1200
grd2.ColWidth(2) = 1000
grd2.ColWidth(3) = 1000
grd2.ColWidth(4) = 1800
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 7000
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
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
grd2.Col = 6
grd2.Text = " ›«’Ì·"
grd3.Clear
grd3.Cols = 7
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1200
grd3.ColWidth(2) = 1000
grd3.ColWidth(3) = 800
grd3.ColWidth(4) = 2000
grd3.ColWidth(5) = 0
grd3.ColWidth(6) = 7000
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.ColAlignment(5) = 1
grd3.ColAlignment(6) = 1
grd3.Row = 0
grd3.Col = 1
grd3.Text = "«· «—ÌŒ"
grd3.Col = 2
grd3.Text = "«·”«⁄…"
grd3.Col = 3
grd3.Text = "«·‰Ê⁄"
grd3.Col = 4
grd3.Text = "«·„»·€"
grd3.Col = 5
grd3.Text = "«· ›«’Ì·"
grd3.Col = 6
grd3.Text = " ›«’Ì·"
grd4.Clear
grd4.Cols = 8
grd4.Rows = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1200
grd4.ColWidth(2) = 1000
grd4.ColWidth(3) = 1200
grd4.ColWidth(4) = 1200
grd4.ColWidth(5) = 1200
grd4.ColWidth(6) = 1200
grd4.ColWidth(7) = 5000
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.ColAlignment(5) = 1
grd4.ColAlignment(6) = 1
grd4.ColAlignment(7) = 1
grd4.Row = 0
grd4.Col = 1
grd4.Text = "«· «—ÌŒ"
grd4.Col = 2
grd4.Text = "«·‘Â—"
grd4.Col = 3
grd4.Text = "«·Ê’·"
grd4.Col = 4
grd4.Text = "«·„” Õﬁ"
grd4.Col = 5
grd4.Text = "«·„œ›Ê⁄"
grd4.Col = 6
grd4.Text = "«·»«ﬁÌ"
grd4.Col = 7
grd4.Text = " ›«’Ì·"
grd5.Clear
grd5.Cols = 5
grd5.Rows = 1
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 1500
grd5.ColWidth(2) = 1500
grd5.ColWidth(3) = 2000
grd5.ColWidth(4) = 7000
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 1
grd5.ColAlignment(2) = 1
grd5.ColAlignment(3) = 1
grd5.ColAlignment(4) = 1
grd5.Row = 0
grd5.Col = 1
grd5.Text = "«· «—ÌŒ"
grd5.Col = 2
grd5.Text = "«·”«⁄…"
grd5.Col = 3
grd5.Text = "«·„»·€"
grd5.Col = 4
grd5.Text = " ›«’Ì·"
grd6.Clear
grd6.Cols = 7
grd6.Rows = 1
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 1200
grd6.ColWidth(2) = 1000
grd6.ColWidth(3) = 800
grd6.ColWidth(4) = 2000
grd6.ColWidth(5) = 0
grd6.ColWidth(6) = 7000
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.ColAlignment(4) = 1
grd6.ColAlignment(5) = 1
grd6.ColAlignment(6) = 1
grd6.Row = 0
grd6.Col = 1
grd6.Text = "«· «—ÌŒ"
grd6.Col = 2
grd6.Text = "«·”«⁄…"
grd6.Col = 3
grd6.Text = "«·‰Ê⁄"
grd6.Col = 4
grd6.Text = "«·„»·€"
grd6.Col = 5
grd6.Text = "«· ›«’Ì·"
grd6.Col = 6
grd6.Text = " ›«’Ì·"

Label1.Caption = "0"
Label2.Caption = "0"
Label4.Caption = "0"
Label3.Caption = "0"
Label10.Caption = "0"
Label5.Caption = "0"
Label6.Caption = "0"
Label8.Caption = "0"
Label18.Caption = "0"
Label7.Caption = "0"
Label9.Caption = "0"
Label16.Caption = "0"
Label17.Caption = "0"
End Sub
Private Sub chargegrd10()
On Error Resume Next
grd10.Clear
grd10.Cols = 2
grd10.Rows = 6
grd10.ColWidth(0) = 1500
grd10.ColWidth(1) = 1500
grd10.ColAlignment(0) = 1
grd10.ColAlignment(1) = 1
grd10.Col = 0
grd10.Row = 0
grd10.Text = pr1
grd10.Row = 1
grd10.Text = pf1
grd10.Row = 2
grd10.Text = et1
grd10.Row = 3
grd10.Text = "" '(tl1 - tl2)
grd10.Row = 4
grd10.Text = bn1
grd10.Row = 5
grd10.Text = tl1
grd10.Col = 1
grd10.Row = 0
grd10.Text = pr2
grd10.Row = 1
grd10.Text = pf2
grd10.Row = 2
grd10.Text = fn2
grd10.Row = 3
grd10.Text = dp2
grd10.Row = 4
grd10.Text = bn2
grd10.Row = 5
grd10.Text = tl2
End Sub

