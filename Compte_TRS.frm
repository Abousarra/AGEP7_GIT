VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Compte_TRS 
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
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   120
      ScaleHeight     =   2775
      ScaleWidth      =   12735
      TabIndex        =   22
      Top             =   6960
      Width           =   12735
      Begin VB.Label Label19 
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
         Left            =   6360
         TabIndex        =   48
         Top             =   2280
         Width           =   3000
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
         Index           =   6
         Left            =   9960
         TabIndex        =   47
         Top             =   2280
         Width           =   2535
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
         Index           =   3
         Left            =   4080
         TabIndex        =   46
         Top             =   2280
         Width           =   2055
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
         Left            =   120
         TabIndex        =   45
         Top             =   2280
         Width           =   3000
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   0
         X2              =   12600
         Y1              =   2280
         Y2              =   2280
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
         Index           =   2
         Left            =   4440
         TabIndex        =   44
         Top             =   1560
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
         Left            =   4440
         TabIndex        =   43
         Top             =   1200
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
         Left            =   4440
         TabIndex        =   42
         Top             =   840
         Width           =   1695
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
         Left            =   4440
         TabIndex        =   41
         Top             =   480
         Width           =   1695
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
         Left            =   9960
         TabIndex        =   40
         Top             =   1200
         Width           =   2535
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
         Left            =   10080
         TabIndex        =   39
         Top             =   840
         Width           =   2415
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
         Left            =   9000
         TabIndex        =   38
         Top             =   480
         Width           =   3495
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
         Left            =   120
         TabIndex        =   37
         Top             =   1560
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
         Left            =   120
         TabIndex        =   36
         Top             =   1200
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
         Left            =   120
         TabIndex        =   35
         Top             =   840
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
         Left            =   120
         TabIndex        =   34
         Top             =   480
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
         Left            =   6360
         TabIndex        =   33
         Top             =   1200
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
         Left            =   6360
         TabIndex        =   32
         Top             =   840
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
         Left            =   6360
         TabIndex        =   31
         Top             =   480
         Width           =   3000
      End
      Begin VB.Line Line2 
         X1              =   6240
         X2              =   6240
         Y1              =   0
         Y2              =   2640
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   0
         X2              =   12600
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Shape Shape1 
         Height          =   2655
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   12615
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
         Index           =   1
         Left            =   9960
         TabIndex        =   30
         Top             =   0
         Width           =   2535
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
         Left            =   3600
         TabIndex        =   29
         Top             =   0
         Width           =   2535
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
         Left            =   120
         TabIndex        =   28
         Top             =   1920
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
         Index           =   0
         Left            =   4440
         TabIndex        =   27
         Top             =   1920
         Width           =   1695
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
         Left            =   6360
         TabIndex        =   26
         Top             =   1920
         Width           =   3000
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
         Left            =   9960
         TabIndex        =   25
         Top             =   1560
         Width           =   2535
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
         Left            =   9960
         TabIndex        =   24
         Top             =   1920
         Width           =   2535
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
         Left            =   6360
         TabIndex        =   23
         Top             =   1560
         Width           =   3000
      End
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H0000FFFF&
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
      ItemData        =   "Compte_TRS.frx":0000
      Left            =   9120
      List            =   "Compte_TRS.frx":0014
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   2055
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
      Left            =   3840
      TabIndex        =   4
      Top             =   840
      Width           =   855
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
      Left            =   5520
      TabIndex        =   3
      Top             =   840
      Width           =   855
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
      ItemData        =   "Compte_TRS.frx":0060
      Left            =   4800
      List            =   "Compte_TRS.frx":0089
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   615
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
      Left            =   8040
      TabIndex        =   1
      Top             =   840
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
      Left            =   240
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   5
      Top             =   1440
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
      TabPicture(0)   =   "Compte_TRS.frx":00B4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grd6"
      Tab(0).Control(1)=   "Command12"
      Tab(0).Control(2)=   "Command13"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "”Ã· «·„’—Ê›« "
      TabPicture(1)   =   "Compte_TRS.frx":00D0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grd5"
      Tab(1).Control(1)=   "Command10"
      Tab(1).Control(2)=   "Command11"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "”Ã· «· ·«„Ì–"
      TabPicture(2)   =   "Compte_TRS.frx":00EC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command4"
      Tab(2).Control(1)=   "Command3"
      Tab(2).Control(2)=   "grd4"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "”Ã· «·√”« –…"
      TabPicture(3)   =   "Compte_TRS.frx":0108
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grd3"
      Tab(3).Control(1)=   "Command8"
      Tab(3).Control(2)=   "Command9"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "”Ã· «·„ÊŸ›Ì‰"
      TabPicture(4)   =   "Compte_TRS.frx":0124
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grd2"
      Tab(4).Control(1)=   "Command5"
      Tab(4).Control(2)=   "Command6"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "”Ã· «·‘—ﬂ«¡"
      TabPicture(5)   =   "Compte_TRS.frx":0140
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "grd1"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Picture1"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Command1"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Command2"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      Begin VB.CommandButton Command13 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   " √ﬂÌœ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   -68640
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "—›÷ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   -72000
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   " √ﬂÌœ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   -68640
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command10 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "—›÷ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   -72000
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   " √ﬂÌœ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   -68640
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "—›÷ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   -72000
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   " √ﬂÌœ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   -68640
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "—›÷ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   -72000
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   " √ﬂÌœ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   -68640
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "—›÷ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   -72000
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "—›÷ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         Left            =   3000
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   5160
         Width           =   3255
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         Caption         =   " √ﬂÌœ ﬂ«›… «·⁄„·Ì«  «·Ÿ«Â—… √⁄·«Â"
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
         MaskColor       =   &H0080FFFF&
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   5160
         Width           =   3255
      End
      Begin VB.PictureBox Picture1 
         Height          =   975
         Left            =   720
         ScaleHeight     =   915
         ScaleWidth      =   7275
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   7335
         Begin MSComCtl2.DTPicker DT3 
            Height          =   345
            Left            =   0
            TabIndex        =   51
            Top             =   0
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
         Begin VB.ComboBox Combo2 
            Height          =   315
            Left            =   4080
            TabIndex        =   12
            Text            =   "0"
            Top             =   120
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker DT2 
            Height          =   345
            Left            =   2520
            TabIndex        =   13
            Top             =   120
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
         Begin VB.Label Label4 
            BackColor       =   &H00000000&
            Caption         =   "Label4"
            Height          =   255
            Left            =   5640
            TabIndex        =   16
            Top             =   120
            Width           =   1455
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
            Left            =   0
            TabIndex        =   15
            Top             =   480
            Width           =   3840
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
            Left            =   3600
            TabIndex        =   14
            Top             =   480
            Width           =   2175
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grd1 
         Height          =   4815
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8493
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
         Height          =   4815
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8493
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
         Height          =   4815
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8493
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
         Height          =   4815
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8493
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
         Height          =   4815
         Left            =   -74880
         TabIndex        =   20
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8493
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
         Height          =   4815
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8493
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
      Left            =   6480
      TabIndex        =   6
      Top             =   840
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
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄Ì… «·⁄„·Ì« "
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
      Left            =   11160
      TabIndex        =   9
      Top             =   840
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   4
      Left            =   9000
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   2
      Left            =   3720
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   5175
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» «·’‰œÊﬁ"
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
      Left            =   4800
      TabIndex        =   7
      Top             =   0
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   3
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   3495
   End
End
Attribute VB_Name = "Compte_TRS"
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

Private Sub Combo3_Change()
On Error Resume Next
If Combo3.Text = "⁄„·Ì«  ÃœÌœ…" Then
Combo2.Text = "0"
Label4.Caption = "ÃœÌœ"
Label4.ForeColor = &HFFFF&
ElseIf Combo3.Text = "⁄„·Ì«  „ƒﬂœ…" Then
Combo2.Text = "1"
Label4.Caption = "„ƒﬂœ"
Label4.ForeColor = &HFF0000
ElseIf Combo3.Text = "⁄„·Ì«  „—›Ê÷…" Then
Combo2.Text = "2"
Label4.Caption = "„—›Ê÷"
Label4.ForeColor = &HFF&
ElseIf Combo3.Text = "⁄„·Ì«  „⁄œ·…" Then
Combo2.Text = "3"
Label4.Caption = "„⁄œ·"
Label4.ForeColor = &HFFFF&
ElseIf Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
Combo2.Text = "6"
Label4.Caption = ""
Label4.ForeColor = &H0&
End If
Call grds_clear

End Sub

Private Sub Combo3_Click()
On Error Resume Next
Combo3_Change
End Sub


Private Sub Combo1_Change()
On Error Resume Next
Call grds_clear

End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim m1 As Double
Dim m2 As Double
Dim j As Double
Dim tx As String
If Combo3.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄Ì… «·⁄„·Ì« ", vbCritical
Exit Sub
End If
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  √⁄·«Â", vbCritical
Exit Sub
End If
Command1.Enabled = False
j = 0
Call cont
Do While Not cp.EOF
dat1 = DT1.Value
dat2 = cp!dat
DT3.Value = cp!dat
m2 = DT3.Month
'***** date
If Option1.Value = True Then
If dat1 = dat2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If cp!act = "0" Or cp!act = "3" Then
cp!act = "1"
cp!mtf = ""
cp.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If cp!act = "2" Then
cp!act = "1"
cp!mtf = ""
cp.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end date
'***** jour
If Option2.Value = True Then
If Combo1.Text = "" Then
MsgBox "ﬁ„ » ÕœÌœ «·‘Â— „‰ Œ·«· «·ﬁ«∆„… «·„‰”œ·…", vbCritical
Command1.Enabled = True
Exit Sub
End If
m1 = Combo1.Text
If m1 = m2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If cp!act = "0" Or cp!act = "3" Then
cp!act = "1"
cp!mtf = ""
cp.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If cp!act = "2" Then
cp!act = "1"
cp!mtf = ""
cp.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end jour
'***** annee
If Option3.Value = True Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If cp!act = "0" Or cp!act = "3" Then
cp!act = "1"
cp!mtf = ""
cp.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If cp!act = "2" Then
cp!act = "1"
cp!mtf = ""
cp.Update
j = j + 1
End If
End If
'***** end mervoud
End If
'***** end annee
cp.MoveNext
Loop
Combo3.Text = "⁄„·Ì«  „ƒﬂœ…"
Command7_Click
tx = j
MsgBox " „  √ﬂÌœ " + tx + " ⁄„·Ì…", vbInformation
Command1.Enabled = True

End Sub

Private Sub Command11_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim m1 As Double
Dim m2 As Double
Dim j As Double
Dim tx As String
If Combo3.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄Ì… «·⁄„·Ì« ", vbCritical
Exit Sub
End If
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  √⁄·«Â", vbCritical
Exit Sub
End If
Command11.Enabled = False
j = 0
Call cont
Do While Not dp.EOF
dat1 = DT1.Value
dat2 = dp!dat
DT3.Value = dp!dat
m2 = DT3.Month
'***** date
If Option1.Value = True Then
If dat1 = dat2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If dp!act = "0" Or dp!act = "3" Then
dp!act = "1"
dp!mtf = ""
dp.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If dp!act = "2" Then
dp!act = "1"
dp!mtf = ""
dp.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end date
'***** jour
If Option2.Value = True Then
If Combo1.Text = "" Then
MsgBox "ﬁ„ » ÕœÌœ «·‘Â— „‰ Œ·«· «·ﬁ«∆„… «·„‰”œ·…", vbCritical
Command11.Enabled = True
Exit Sub
End If
m1 = Combo1.Text
If m1 = m2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If dp!act = "0" Or dp!act = "3" Then
dp!act = "1"
dp!mtf = ""
dp.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If dp!act = "2" Then
dp!act = "1"
dp!mtf = ""
dp.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end jour
'***** annee
If Option3.Value = True Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If dp!act = "0" Or dp!act = "3" Then
dp!act = "1"
dp!mtf = ""
dp.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If dp!act = "2" Then
dp!act = "1"
dp!mtf = ""
dp.Update
j = j + 1
End If
End If
'***** end mervoud
End If
'***** end annee
dp.MoveNext
Loop
Combo3.Text = "⁄„·Ì«  „ƒﬂœ…"
Command7_Click
tx = j
MsgBox " „  √ﬂÌœ " + tx + " ⁄„·Ì…", vbInformation
Command11.Enabled = True

End Sub

Private Sub Command13_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim m1 As Double
Dim m2 As Double
Dim j As Double
Dim tx As String
If Combo3.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄Ì… «·⁄„·Ì« ", vbCritical
Exit Sub
End If
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  √⁄·«Â", vbCritical
Exit Sub
End If
Command13.Enabled = False
j = 0
Call cont
Do While Not bn.EOF
dat1 = DT1.Value
dat2 = bn!dat
DT3.Value = bn!dat
m2 = DT3.Month
'***** date
If Option1.Value = True Then
If dat1 = dat2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If bn!act = "0" Or bn!act = "3" Then
bn!act = "1"
bn!mtf = ""
bn.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If bn!act = "2" Then
bn!act = "1"
bn!mtf = ""
bn.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end date
'***** jour
If Option2.Value = True Then
If Combo1.Text = "" Then
MsgBox "ﬁ„ » ÕœÌœ «·‘Â— „‰ Œ·«· «·ﬁ«∆„… «·„‰”œ·…", vbCritical
Command13.Enabled = True
Exit Sub
End If
m1 = Combo1.Text
If m1 = m2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If bn!act = "0" Or bn!act = "3" Then
bn!act = "1"
bn!mtf = ""
bn.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If bn!act = "2" Then
bn!act = "1"
bn!mtf = ""
bn.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end jour
'***** annee
If Option3.Value = True Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If bn!act = "0" Or bn!act = "3" Then
bn!act = "1"
bn!mtf = ""
bn.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If bn!act = "2" Then
bn!act = "1"
bn!mtf = ""
bn.Update
j = j + 1
End If
End If
'***** end mervoud
End If
'***** end annee
bn.MoveNext
Loop
Combo3.Text = "⁄„·Ì«  „ƒﬂœ…"
Command7_Click
tx = j
MsgBox " „  √ﬂÌœ " + tx + " ⁄„·Ì…", vbInformation
Command13.Enabled = True

End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim m1 As Double
Dim m2 As Double
Dim j As Double
Dim tx As String
If Combo3.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄Ì… «·⁄„·Ì« ", vbCritical
Exit Sub
End If
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  √⁄·«Â", vbCritical
Exit Sub
End If
Command2.Enabled = False
j = 0
Call cont
Do While Not cp.EOF
dat1 = DT1.Value
dat2 = cp!dat
DT3.Value = cp!dat
m2 = DT3.Month
'***** date
If Option1.Value = True Then
If dat1 = dat2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Or Label4.Caption = "„ƒﬂœ" Then
If cp!act = "0" Or cp!act = "1" Or cp!act = "3" Then
cp!act = "2"
cp!mtf = " „ —›÷ Â–Â «·⁄„·Ì… „⁄ Ã„·… „‰ «·⁄„·Ì«  «·√Œ—Ï œ›⁄… Ê«Õœ… œÊ‰  »ÌÌ‰ ”»» «·—›÷"
cp.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
End If
End If
'***** end date
'***** jour
If Option2.Value = True Then
If Combo1.Text = "" Then
MsgBox "ﬁ„ » ÕœÌœ «·‘Â— „‰ Œ·«· «·ﬁ«∆„… «·„‰”œ·…", vbCritical
Command2.Enabled = True
Exit Sub
End If
m1 = Combo1.Text
If m1 = m2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Or Label4.Caption = "„ƒﬂœ" Then
If cp!act = "0" Or cp!act = "1" Or cp!act = "3" Then
cp!act = "2"
cp!mtf = " „ —›÷ Â–Â «·⁄„·Ì… „⁄ Ã„·… „‰ «·⁄„·Ì«  «·√Œ—Ï œ›⁄… Ê«Õœ… œÊ‰  »ÌÌ‰ ”»» «·—›÷"
cp.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
End If
End If
'***** end jour
'***** annee
If Option3.Value = True Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Or Label4.Caption = "„ƒﬂœ" Then
If cp!act = "0" Or cp!act = "1" Or cp!act = "3" Then
cp!act = "2"
cp!mtf = " „ —›÷ Â–Â «·⁄„·Ì… „⁄ Ã„·… „‰ «·⁄„·Ì«  «·√Œ—Ï œ›⁄… Ê«Õœ… œÊ‰  »ÌÌ‰ ”»» «·—›÷"
cp.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
End If
'***** end annee
cp.MoveNext
Loop
tx = j
Combo3.Text = "⁄„·Ì«  „—›Ê÷…"
Command7_Click
MsgBox " „  √ﬂÌœ " + tx + " ⁄„·Ì…", vbInformation
Command2.Enabled = True

End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim m1 As Double
Dim m2 As Double
Dim j As Double
Dim tx As String
If Combo3.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄Ì… «·⁄„·Ì« ", vbCritical
Exit Sub
End If
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  √⁄·«Â", vbCritical
Exit Sub
End If
Command2.Enabled = False
j = 0
Call cont
Do While Not ct.EOF
dat1 = DT1.Value
dat2 = ct!dat
DT3.Value = ct!dat
m2 = DT3.Month
'***** date
If Option1.Value = True Then
If dat1 = dat2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Or Label4.Caption = "„ƒﬂœ" Then
If ct!act = "0" Or ct!act = "1" Or ct!act = "3" Then
ct!act = "2"
ct!mtf = " „ —›÷ Â–Â «·⁄„·Ì… „⁄ Ã„·… „‰ «·⁄„·Ì«  «·√Œ—Ï œ›⁄… Ê«Õœ… œÊ‰  »ÌÌ‰ ”»» «·—›÷"
ct.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
End If
End If
'***** end date
'***** jour
If Option2.Value = True Then
If Combo1.Text = "" Then
MsgBox "ﬁ„ » ÕœÌœ «·‘Â— „‰ Œ·«· «·ﬁ«∆„… «·„‰”œ·…", vbCritical
Command2.Enabled = True
Exit Sub
End If
m1 = Combo1.Text
If m1 = m2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Or Label4.Caption = "„ƒﬂœ" Then
If ct!act = "0" Or ct!act = "1" Or ct!act = "3" Then
ct!act = "2"
ct!mtf = " „ —›÷ Â–Â «·⁄„·Ì… „⁄ Ã„·… „‰ «·⁄„·Ì«  «·√Œ—Ï œ›⁄… Ê«Õœ… œÊ‰  »ÌÌ‰ ”»» «·—›÷"
ct.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
End If
End If
'***** end jour
'***** annee
If Option3.Value = True Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Or Label4.Caption = "„ƒﬂœ" Then
If ct!act = "0" Or ct!act = "1" Or ct!act = "3" Then
ct!act = "2"
ct!mtf = " „ —›÷ Â–Â «·⁄„·Ì… „⁄ Ã„·… „‰ «·⁄„·Ì«  «·√Œ—Ï œ›⁄… Ê«Õœ… œÊ‰  »ÌÌ‰ ”»» «·—›÷"
ct.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
End If
'***** end annee
ct.MoveNext
Loop
tx = j
Combo3.Text = "⁄„·Ì«  „—›Ê÷…"
Command7_Click
MsgBox " „  √ﬂÌœ " + tx + " ⁄„·Ì…", vbInformation
Command2.Enabled = True
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim m1 As Double
Dim m2 As Double
Dim j As Double
Dim tx As String
If Combo3.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄Ì… «·⁄„·Ì« ", vbCritical
Exit Sub
End If
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  √⁄·«Â", vbCritical
Exit Sub
End If
Command4.Enabled = False
j = 0
Call cont
Do While Not ct.EOF
dat1 = DT1.Value
dat2 = ct!dat
DT3.Value = ct!dat
m2 = DT3.Month
'***** date
If Option1.Value = True Then
If dat1 = dat2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If ct!act = "0" Or ct!act = "3" Then
ct!act = "1"
ct!mtf = ""
ct.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If ct!act = "2" Then
ct!act = "1"
ct!mtf = ""
ct.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end date
'***** jour
If Option2.Value = True Then
If Combo1.Text = "" Then
MsgBox "ﬁ„ » ÕœÌœ «·‘Â— „‰ Œ·«· «·ﬁ«∆„… «·„‰”œ·…", vbCritical
Command4.Enabled = True
Exit Sub
End If
m1 = Combo1.Text
If m1 = m2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If ct!act = "0" Or ct!act = "3" Then
ct!act = "1"
ct!mtf = ""
ct.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If ct!act = "2" Then
ct!act = "1"
ct!mtf = ""
ct.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end jour
'***** annee
If Option3.Value = True Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If ct!act = "0" Or ct!act = "3" Then
ct!act = "1"
ct!mtf = ""
ct.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If ct!act = "2" Then
ct!act = "1"
ct!mtf = ""
ct.Update
j = j + 1
End If
End If
'***** end mervoud
End If
'***** end annee
ct.MoveNext
Loop
tx = j
Combo3.Text = "⁄„·Ì«  „ƒﬂœ…"
Command7_Click
MsgBox " „  √ﬂÌœ " + tx + " ⁄„·Ì…", vbInformation
Command4.Enabled = True

End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim m1 As Double
Dim m2 As Double
Dim j As Double
Dim tx As String
If Combo3.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄Ì… «·⁄„·Ì« ", vbCritical
Exit Sub
End If
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  √⁄·«Â", vbCritical
Exit Sub
End If
Command6.Enabled = False
j = 0
Call cont
Do While Not cf.EOF
dat1 = DT1.Value
dat2 = cf!dat
DT3.Value = cf!dat
m2 = DT3.Month
'***** date
If Option1.Value = True Then
If dat1 = dat2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If cf!act = "0" Or cf!act = "3" Then
cf!act = "1"
cf!mtf = ""
cf.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If cf!act = "2" Then
cf!act = "1"
cf!mtf = ""
cf.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end date
'***** jour
If Option2.Value = True Then
If Combo1.Text = "" Then
MsgBox "ﬁ„ » ÕœÌœ «·‘Â— „‰ Œ·«· «·ﬁ«∆„… «·„‰”œ·…", vbCritical
Command6.Enabled = True
Exit Sub
End If
m1 = Combo1.Text
If m1 = m2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If cf!act = "0" Or cf!act = "3" Then
cf!act = "1"
cf!mtf = ""
cf.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If cf!act = "2" Then
cf!act = "1"
cf!mtf = ""
cf.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end jour
'***** annee
If Option3.Value = True Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If cf!act = "0" Or cf!act = "3" Then
cf!act = "1"
cf!mtf = ""
cf.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If cf!act = "2" Then
cf!act = "1"
cf!mtf = ""
cf.Update
j = j + 1
End If
End If
'***** end mervoud
End If
'***** end annee
cf.MoveNext
Loop
Combo3.Text = "⁄„·Ì«  „ƒﬂœ…"
Command7_Click
tx = j
MsgBox " „  √ﬂÌœ " + tx + " ⁄„·Ì…", vbInformation
Command6.Enabled = True

End Sub

Private Sub Command7_Click()
On Error Resume Next
If Combo3.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄Ì… «·⁄„·Ì« ", vbCritical
Exit Sub
End If
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

Private Sub Command9_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim m1 As Double
Dim m2 As Double
Dim j As Double
Dim tx As String
If Combo3.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ ‰Ê⁄Ì… «·⁄„·Ì« ", vbCritical
Exit Sub
End If
If Option1.Value = False And Option2.Value = False And Option3.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  √⁄·«Â", vbCritical
Exit Sub
End If
Command9.Enabled = False
j = 0
Call cont
Do While Not cs.EOF
dat1 = DT1.Value
dat2 = cs!dat
DT3.Value = cs!dat
m2 = DT3.Month
'***** date
If Option1.Value = True Then
If dat1 = dat2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If cs!act = "0" Or cs!act = "3" Then
cs!act = "1"
cs!mtf = ""
cs.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If cs!act = "2" Then
cs!act = "1"
cs!mtf = ""
cs.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end date
'***** jour
If Option2.Value = True Then
If Combo1.Text = "" Then
MsgBox "ﬁ„ » ÕœÌœ «·‘Â— „‰ Œ·«· «·ﬁ«∆„… «·„‰”œ·…", vbCritical
Command9.Enabled = True
Exit Sub
End If
m1 = Combo1.Text
If m1 = m2 Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If cs!act = "0" Or cs!act = "3" Then
cs!act = "1"
cs!mtf = ""
cs.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If cs!act = "2" Then
cs!act = "1"
cs!mtf = ""
cs.Update
j = j + 1
End If
End If
'***** end mervoud
End If
End If
'***** end jour
'***** annee
If Option3.Value = True Then
'***** jedid or moaadel
If Label4.Caption = "ÃœÌœ" Or Label4.Caption = "„⁄œ·" Then
If cs!act = "0" Or cs!act = "3" Then
cs!act = "1"
cs!mtf = ""
cs.Update
j = j + 1
End If
End If
'***** end jedid or moaadel
'***** mervoud
If Label4.Caption = "„—›Ê÷" Then
If cs!act = "2" Then
cs!act = "1"
cs!mtf = ""
cs.Update
j = j + 1
End If
End If
'***** end mervoud
End If
'***** end annee
cs.MoveNext
Loop
Combo3.Text = "⁄„·Ì«  „ƒﬂœ…"
Command7_Click
tx = j
MsgBox " „  √ﬂÌœ " + tx + " ⁄„·Ì…", vbInformation
Command9.Enabled = True

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
DT1.Value = Date
End Sub


Private Sub grd1_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Integer
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim tx6 As String
Dim tx5 As String
Dim g As String
g = ""
i = grd1.Row
j = grd1.Col
If i > 0 Then
If j = 7 Then
grd1.Row = i
grd1.Col = 0
tx1 = grd1.Text
grd1.Col = 1
tx5 = grd1.Text
grd1.Col = 3
tx3 = grd1.Text
grd1.Col = 4
tx4 = grd1.Text
grd1.Col = 6
tx6 = grd1.Text
grd1.Col = 7
tx2 = grd1.Text
If tx2 = "ÃœÌœ" Or tx2 = "„⁄œ·" Then
k = 1
grd1.Row = i
grd1.Col = 7
grd1.Text = "„ƒﬂœ"
grd1.CellBackColor = &HFF0000
Call cont
Do While Not cp.EOF
If cp!aut = tx1 Then
cp!act = k
cp!mtf = g
cp.Update
Exit Sub
End If
cp.MoveNext
Loop
ElseIf tx2 = "„ƒﬂœ" Then
g = InputBox("√‰ „ ⁄·Ï Ê‘ﬂ —›÷ ⁄„·Ì… " + tx3 + " „»·€ " + tx4 + " " + tx6 + " √Ã—Ì  » «—ÌŒ " + tx5 + " ", "«·—Ã«¡ ﬂ «»… ”»» «·—›÷ ≈–« ﬂ‰ „  Ê«›ﬁÊ‰ ⁄·ÌÂ")
If g = Cancel Then
Exit Sub
End If
k = 2
grd1.Row = i
grd1.Col = 7
grd1.Text = "„—›Ê÷"
grd1.CellBackColor = &HFF&
Call cont
Do While Not cp.EOF
If cp!aut = tx1 Then
cp!act = k
cp!mtf = g
cp.Update
Exit Sub
End If
cp.MoveNext
Loop
End If
End If
End If
End Sub

Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Integer
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim tx6 As String
Dim tx5 As String
Dim g As String
g = ""
i = grd2.Row
j = grd2.Col
If i > 0 Then
If j = 7 Then
grd2.Row = i
grd2.Col = 0
tx1 = grd2.Text
grd2.Col = 1
tx5 = grd2.Text
grd2.Col = 3
tx3 = grd2.Text
grd2.Col = 4
tx4 = grd2.Text
grd2.Col = 6
tx6 = grd2.Text
grd2.Col = 7
tx2 = grd2.Text
If tx2 = "ÃœÌœ" Or tx2 = "„⁄œ·" Then
k = 1
grd2.Row = i
grd2.Col = 7
grd2.Text = "„ƒﬂœ"
grd2.CellBackColor = &HFF0000
Call cont
Do While Not cf.EOF
If cf!aut = tx1 Then
cf!act = k
cf!mtf = g
cf.Update
Exit Sub
End If
cf.MoveNext
Loop
ElseIf tx2 = "„ƒﬂœ" Then
g = InputBox("√‰ „ ⁄·Ï Ê‘ﬂ —›÷ ⁄„·Ì… " + tx3 + " „»·€ " + tx4 + " " + tx6 + " √Ã—Ì  » «—ÌŒ " + tx5 + " ", "«·—Ã«¡ ﬂ «»… ”»» «·—›÷ ≈–« ﬂ‰ „  Ê«›ﬁÊ‰ ⁄·ÌÂ")
If g = Cancel Then
Exit Sub
End If
k = 2
grd2.Row = i
grd2.Col = 7
grd2.Text = "„—›Ê÷"
grd2.CellBackColor = &HFF&
Call cont
Do While Not cf.EOF
If cf!aut = tx1 Then
cf!act = k
cf!mtf = g
cf.Update
Exit Sub
End If
cf.MoveNext
Loop
End If
End If
End If

End Sub

Private Sub grd3_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Integer
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim tx6 As String
Dim tx5 As String
Dim g As String
g = ""
i = grd3.Row
j = grd3.Col
If i > 0 Then
If j = 7 Then
grd3.Row = i
grd3.Col = 0
tx1 = grd3.Text
grd3.Col = 1
tx5 = grd3.Text
grd3.Col = 3
tx3 = grd3.Text
grd3.Col = 4
tx4 = grd3.Text
grd3.Col = 6
tx6 = grd3.Text
grd3.Col = 7
tx2 = grd3.Text
If tx2 = "ÃœÌœ" Or tx2 = "„⁄œ·" Then
k = 1
grd3.Row = i
grd3.Col = 7
grd3.Text = "„ƒﬂœ"
grd3.CellBackColor = &HFF0000
Call cont
Do While Not cs.EOF
If cs!aut = tx1 Then
cs!act = k
cs!mtf = g
cs.Update
Exit Sub
End If
cs.MoveNext
Loop
ElseIf tx2 = "„ƒﬂœ" Then
g = InputBox("√‰ „ ⁄·Ï Ê‘ﬂ —›÷ ⁄„·Ì… " + tx3 + " „»·€ " + tx4 + " " + tx6 + " √Ã—Ì  » «—ÌŒ " + tx5 + " ", "«·—Ã«¡ ﬂ «»… ”»» «·—›÷ ≈–« ﬂ‰ „  Ê«›ﬁÊ‰ ⁄·ÌÂ")
If g = Cancel Then
Exit Sub
End If
k = 2
grd3.Row = i
grd3.Col = 7
grd3.Text = "„—›Ê÷"
grd3.CellBackColor = &HFF&
Call cont
Do While Not cs.EOF
If cs!aut = tx1 Then
cs!act = k
cs!mtf = g
cs.Update
Exit Sub
End If
cs.MoveNext
Loop
End If
End If
End If

End Sub

Private Sub grd4_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Integer
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim tx6 As String
Dim tx5 As String
Dim tx7 As String
Dim tx8 As String
Dim g As String
g = ""
i = grd4.Row
j = grd4.Col
If i > 0 Then
If j = 8 Then
grd4.Row = i
grd4.Col = 0
tx1 = grd4.Text
grd4.Col = 1
tx5 = grd4.Text
grd4.Col = 2
tx7 = grd4.Text
grd4.Col = 3
tx3 = grd4.Text
grd4.Col = 4
tx4 = grd4.Text
grd4.Col = 5
tx8 = grd4.Text
grd4.Col = 6
tx6 = grd4.Text
grd4.Col = 8
tx2 = grd4.Text
If tx2 = "ÃœÌœ" Or tx2 = "„⁄œ·" Then
k = 1
grd4.Row = i
grd4.Col = 8
grd4.Text = "„ƒﬂœ"
grd4.CellBackColor = &HFF0000
Call cont
Do While Not ct.EOF
If ct!aut = tx1 Then
ct!act = k
ct!mtf = g
ct.Update
Exit Sub
End If
ct.MoveNext
Loop
ElseIf tx2 = "„ƒﬂœ" Then
g = InputBox("√‰ „ ⁄·Ï Ê‘ﬂ —›÷ ⁄„·Ì… œ›⁄ —”Ê„ " + tx7 + " »„»·€ ﬁœ—Â " + tx8 + " √Ã—Ì  » «—ÌŒ " + tx5 + " ", "«·—Ã«¡ ﬂ «»… ”»» «·—›÷ ≈–« ﬂ‰ „  Ê«›ﬁÊ‰ ⁄·ÌÂ")
If g = Cancel Then
Exit Sub
End If
k = 2
grd4.Row = i
grd4.Col = 8
grd4.Text = "„—›Ê÷"
grd4.CellBackColor = &HFF&
Call cont
Do While Not ct.EOF
If ct!aut = tx1 Then
ct!act = k
ct!mtf = g
ct.Update
Exit Sub
End If
ct.MoveNext
Loop
End If
End If
End If

End Sub

Private Sub grd5_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Integer
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim tx6 As String
Dim tx5 As String
Dim g As String
g = ""
i = grd5.Row
j = grd5.Col
If i > 0 Then
If j = 5 Then
grd5.Row = i
grd5.Col = 0
tx1 = grd5.Text
grd5.Col = 1
tx5 = grd5.Text
grd5.Col = 3
tx3 = grd5.Text
grd5.Col = 4
tx4 = grd5.Text
grd5.Col = 5
tx2 = grd5.Text
If tx2 = "ÃœÌœ" Or tx2 = "„⁄œ·" Then
k = 1
grd5.Row = i
grd5.Col = 5
grd5.Text = "„ƒﬂœ"
grd5.CellBackColor = &HFF0000
Call cont
Do While Not dp.EOF
If dp!aut = tx1 Then
dp!act = k
dp!mtf = g
dp.Update
Exit Sub
End If
dp.MoveNext
Loop
ElseIf tx2 = "„ƒﬂœ" Then
g = InputBox("√‰ „ ⁄·Ï Ê‘ﬂ —›÷ ⁄„·Ì… " + tx3 + " „»·€ " + tx4 + " " + tx6 + " √Ã—Ì  » «—ÌŒ " + tx5 + " ", "«·—Ã«¡ ﬂ «»… ”»» «·—›÷ ≈–« ﬂ‰ „  Ê«›ﬁÊ‰ ⁄·ÌÂ")
If g = Cancel Then
Exit Sub
End If
k = 2
grd5.Row = i
grd5.Col = 5
grd5.Text = "„—›Ê÷"
grd5.CellBackColor = &HFF&
Call cont
Do While Not dp.EOF
If dp!aut = tx1 Then
dp!act = k
dp!mtf = g
dp.Update
Exit Sub
End If
dp.MoveNext
Loop
End If
End If
End If

End Sub

Private Sub grd6_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim k As Integer
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim tx6 As String
Dim tx5 As String
Dim g As String
g = ""
i = grd6.Row
j = grd6.Col
If i > 0 Then
If j = 7 Then
grd6.Row = i
grd6.Col = 0
tx1 = grd6.Text
grd6.Col = 1
tx5 = grd6.Text
grd6.Col = 3
tx3 = grd6.Text
grd6.Col = 4
tx4 = grd6.Text
grd6.Col = 6
tx6 = grd6.Text
grd6.Col = 7
tx2 = grd6.Text
If tx2 = "ÃœÌœ" Or tx2 = "„⁄œ·" Then
k = 1
grd6.Row = i
grd6.Col = 7
grd6.Text = "„ƒﬂœ"
grd6.CellBackColor = &HFF0000
Call cont
Do While Not bn.EOF
If bn!aut = tx1 Then
bn!act = k
bn!mtf = g
bn.Update
Exit Sub
End If
bn.MoveNext
Loop
ElseIf tx2 = "„ƒﬂœ" Then
g = InputBox("√‰ „ ⁄·Ï Ê‘ﬂ —›÷ ⁄„·Ì… " + tx3 + " „»·€ " + tx4 + " " + tx6 + " √Ã—Ì  » «—ÌŒ " + tx5 + " ", "«·—Ã«¡ ﬂ «»… ”»» «·—›÷ ≈–« ﬂ‰ „  Ê«›ﬁÊ‰ ⁄·ÌÂ")
If g = Cancel Then
Exit Sub
End If
k = 2
grd6.Row = i
grd6.Col = 7
grd6.Text = "„—›Ê÷"
grd6.CellBackColor = &HFF&
Call cont
Do While Not bn.EOF
If bn!aut = tx1 Then
bn!act = k
bn!mtf = g
bn.Update
Exit Sub
End If
bn.MoveNext
Loop
End If
End If
End If

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
grd1.Cols = 8
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 1000
grd1.ColWidth(3) = 800
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 0
grd1.ColWidth(6) = 5800
grd1.ColWidth(7) = 1200
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 3
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
If cp!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat3 = cp!dat
If cp!typ = "”Õ»" Then
a = cp!mon
pr2 = pr2 + a
Else
a = cp!mon
pr1 = pr1 + a
End If
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
grd1.Text = " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cp!sri
grd1.Col = 7
grd1.Text = Label4.Caption
grd1.CellBackColor = Label4.ForeColor
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
grd1.Cols = 8
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 1000
grd1.ColWidth(3) = 800
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 0
grd1.ColWidth(6) = 5800
grd1.ColWidth(7) = 1200
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 3
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
dat1 = DT1.Value
Call cont
grd1.Rows = cp.RecordCount + 3
Do While Not cp.EOF
If cp!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat2 = cp!dat
If dat2 = dat1 Then
If cp!typ = "”Õ»" Then
a = cp!mon
pr2 = pr2 + a
Else
a = cp!mon
pr1 = pr1 + a
End If
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
grd1.Text = " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cp!sri
grd1.Col = 7
grd1.Text = Label4.Caption
grd1.CellBackColor = Label4.ForeColor
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
grd1.Cols = 8
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1200
grd1.ColWidth(2) = 1000
grd1.ColWidth(3) = 800
grd1.ColWidth(4) = 2000
grd1.ColWidth(5) = 0
grd1.ColWidth(6) = 5800
grd1.ColWidth(7) = 1200
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 3
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
If cp!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
DT2.Value = cp!dat
j2 = DT2.Month
If j1 = j2 Then
If cp!typ = "”Õ»" Then
a = cp!mon
pr2 = pr2 + a
Else
a = cp!mon
pr1 = pr1 + a
End If
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
grd1.Text = " „‰ ÿ—› ’«Õ» «·—ﬁ„ «· ”·”·Ì: " + cp!sri
grd1.Col = 7
grd1.Text = Label4.Caption
grd1.CellBackColor = Label4.ForeColor
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
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1200
grd2.ColWidth(2) = 1000
grd2.ColWidth(3) = 1000
grd2.ColWidth(4) = 1800
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 5800
grd2.ColWidth(7) = 1200
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
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
If cf!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat3 = cf!dat
If cf!typ = "”Õ» „»·€" Then
a = cf!mon
fn2 = fn2 + a
End If
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
grd2.Col = 7
grd2.Text = Label4.Caption
grd2.CellBackColor = Label4.ForeColor
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
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1200
grd2.ColWidth(2) = 1000
grd2.ColWidth(3) = 1000
grd2.ColWidth(4) = 1800
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 5800
grd2.ColWidth(7) = 1200
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
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
If cf!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat2 = cf!dat
If dat2 = dat1 Then
If cf!typ = "”Õ» „»·€" Then
a = cf!mon
fn2 = fn2 + a
End If
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
grd2.Col = 7
grd2.Text = Label4.Caption
grd2.CellBackColor = Label4.ForeColor
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
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1200
grd2.ColWidth(2) = 1000
grd2.ColWidth(3) = 1000
grd2.ColWidth(4) = 1800
grd2.ColWidth(5) = 0
grd2.ColWidth(6) = 5800
grd2.ColWidth(7) = 1200
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
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
If cf!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
DT2.Value = cf!dat
j2 = DT2.Month
If j1 = j2 Then
If cf!typ = "”Õ» „»·€" Then
a = cf!mon
fn2 = fn2 + a
End If
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
grd2.Col = 7
grd2.Text = Label4.Caption
grd2.CellBackColor = Label4.ForeColor
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
grd3.Cols = 8
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1200
grd3.ColWidth(2) = 1000
grd3.ColWidth(3) = 800
grd3.ColWidth(4) = 2000
grd3.ColWidth(5) = 0
grd3.ColWidth(6) = 5800
grd3.ColWidth(7) = 1200
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.ColAlignment(5) = 1
grd3.ColAlignment(6) = 1
grd3.ColAlignment(7) = 3
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
If cs!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat3 = cs!dat
'If dat3 >= dat1 And dat3 <= dat2 Then
If cs!typ = "”Õ»" Then
a = cs!mon
pf2 = pf2 + a
Else
a = cs!mon
pf1 = pf1 + a
End If
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
grd3.Col = 7
grd3.Text = Label4.Caption
grd3.CellBackColor = Label4.ForeColor
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
grd3.Cols = 8
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1200
grd3.ColWidth(2) = 1000
grd3.ColWidth(3) = 800
grd3.ColWidth(4) = 2000
grd3.ColWidth(5) = 0
grd3.ColWidth(6) = 5800
grd3.ColWidth(7) = 1200
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.ColAlignment(5) = 1
grd3.ColAlignment(6) = 1
grd3.ColAlignment(7) = 3
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
If cs!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat2 = cs!dat
If dat2 = dat1 Then
If cs!typ = "”Õ»" Then
a = cs!mon
pf2 = pf2 + a
Else
a = cs!mon
pf1 = pf1 + a
End If
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
grd3.Col = 7
grd3.Text = Label4.Caption
grd3.CellBackColor = Label4.ForeColor
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
grd3.Cols = 8
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1200
grd3.ColWidth(2) = 1000
grd3.ColWidth(3) = 800
grd3.ColWidth(4) = 2000
grd3.ColWidth(5) = 0
grd3.ColWidth(6) = 5800
grd3.ColWidth(7) = 1200
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.ColAlignment(5) = 1
grd3.ColAlignment(6) = 1
grd3.ColAlignment(7) = 3
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
If cs!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
DT2.Value = cs!dat
j2 = DT2.Month
If j1 = j2 Then
If cs!typ = "”Õ»" Then
a = cs!mon
pf2 = pf2 + a
Else
a = cs!mon
pf1 = pf1 + a
End If
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
grd3.Col = 7
grd3.Text = Label4.Caption
grd3.CellBackColor = Label4.ForeColor
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
grd4.Cols = 9
grd4.Rows = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1200
grd4.ColWidth(2) = 1000
grd4.ColWidth(3) = 1000
grd4.ColWidth(4) = 1000
grd4.ColWidth(5) = 1000
grd4.ColWidth(6) = 1000
grd4.ColWidth(7) = 4800
grd4.ColWidth(8) = 1000
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.ColAlignment(5) = 1
grd4.ColAlignment(6) = 1
grd4.ColAlignment(7) = 1
grd4.ColAlignment(8) = 3
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
If ct!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat3 = ct!dat
'If dat3 >= dat1 And dat3 <= dat2 Then
If ct!rcu <> "0" Then
a = ct!tpy
et1 = et1 + a
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
grd4.Col = 8
grd4.Text = Label4.Caption
grd4.CellBackColor = Label4.ForeColor
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
grd4.Cols = 9
grd4.Rows = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1200
grd4.ColWidth(2) = 1000
grd4.ColWidth(3) = 1000
grd4.ColWidth(4) = 1000
grd4.ColWidth(5) = 1000
grd4.ColWidth(6) = 1000
grd4.ColWidth(7) = 4800
grd4.ColWidth(8) = 1000
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.ColAlignment(5) = 1
grd4.ColAlignment(6) = 1
grd4.ColAlignment(7) = 1
grd4.ColAlignment(8) = 3
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
If ct!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat2 = ct!dat
If dat1 = dat2 Then
If ct!rcu <> "0" Then
a = ct!tpy
et1 = et1 + a
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
grd4.Col = 8
grd4.Text = Label4.Caption
grd4.CellBackColor = Label4.ForeColor
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
grd4.Cols = 9
grd4.Rows = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1200
grd4.ColWidth(2) = 1000
grd4.ColWidth(3) = 1000
grd4.ColWidth(4) = 1000
grd4.ColWidth(5) = 1000
grd4.ColWidth(6) = 1000
grd4.ColWidth(7) = 4800
grd4.ColWidth(8) = 1000
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.ColAlignment(5) = 1
grd4.ColAlignment(6) = 1
grd4.ColAlignment(7) = 1
grd4.ColAlignment(8) = 3
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
If ct!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
DT2.Value = ct!dat
j2 = DT2.Month
If j1 = j2 Then
If ct!rcu <> "0" Then
a = ct!tpy
et1 = et1 + a
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
grd4.Col = 8
grd4.Text = Label4.Caption
grd4.CellBackColor = Label4.ForeColor
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
grd5.Cols = 6
grd5.Rows = 1
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 1500
grd5.ColWidth(2) = 1500
grd5.ColWidth(3) = 2000
grd5.ColWidth(4) = 5800
grd5.ColWidth(5) = 1200
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 1
grd5.ColAlignment(2) = 1
grd5.ColAlignment(3) = 1
grd5.ColAlignment(4) = 1
grd5.ColAlignment(5) = 3
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
If dp!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat3 = dp!dat
a = dp!mon
dp2 = dp2 + a
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
grd5.Col = 5
grd5.Text = Label4.Caption
grd5.CellBackColor = Label4.ForeColor
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
grd5.Cols = 6
grd5.Rows = 1
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 1500
grd5.ColWidth(2) = 1500
grd5.ColWidth(3) = 2000
grd5.ColWidth(4) = 5800
grd5.ColWidth(5) = 1200
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 1
grd5.ColAlignment(2) = 1
grd5.ColAlignment(3) = 1
grd5.ColAlignment(4) = 1
grd5.ColAlignment(5) = 3
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
If dp!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat2 = dp!dat
If dat1 = dat2 Then
a = dp!mon
dp2 = dp2 + a
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
grd5.Col = 5
grd5.Text = Label4.Caption
grd5.CellBackColor = Label4.ForeColor
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
grd5.Cols = 6
grd5.Rows = 1
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 1500
grd5.ColWidth(2) = 1500
grd5.ColWidth(3) = 2000
grd5.ColWidth(4) = 5800
grd5.ColWidth(5) = 1200
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 1
grd5.ColAlignment(2) = 1
grd5.ColAlignment(3) = 1
grd5.ColAlignment(4) = 1
grd5.ColAlignment(5) = 3
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
If dp!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
DT2.Value = dp!dat
j2 = DT2.Month
If j1 = j2 Then
a = dp!mon
dp2 = dp2 + a
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
grd5.Col = 5
grd5.Text = Label4.Caption
grd5.CellBackColor = Label4.ForeColor
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
grd6.Cols = 8
grd6.Rows = 1
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 1200
grd6.ColWidth(2) = 1000
grd6.ColWidth(3) = 800
grd6.ColWidth(4) = 2000
grd6.ColWidth(5) = 0
grd6.ColWidth(6) = 5800
grd6.ColWidth(7) = 1200
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.ColAlignment(4) = 1
grd6.ColAlignment(5) = 1
grd6.ColAlignment(6) = 1
grd6.ColAlignment(7) = 3
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
If bn!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat3 = bn!dat
'If dat3 >= dat1 And dat3 <= dat2 Then
If bn!typ = "”Õ»" Then
a = bn!mon
bn1 = bn1 + a
Else
a = bn!mon
bn2 = bn2 + a
End If
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
grd6.Col = 7
grd6.Text = Label4.Caption
grd6.CellBackColor = Label4.ForeColor
i = i + 1
End If
bn.MoveNext
Loop
grd6.Rows = i
grd6.Col = 1
grd6.Sort = 2
s = (P - r)
Label5.Caption = P
Label1.Caption = r
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
grd6.Cols = 8
grd6.Rows = 1
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 1200
grd6.ColWidth(2) = 1000
grd6.ColWidth(3) = 800
grd6.ColWidth(4) = 2000
grd6.ColWidth(5) = 0
grd6.ColWidth(6) = 5800
grd6.ColWidth(7) = 1200
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.ColAlignment(4) = 1
grd6.ColAlignment(5) = 1
grd6.ColAlignment(6) = 1
grd6.ColAlignment(7) = 3
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
If bn!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
dat2 = bn!dat
If dat2 = dat1 Then
If bn!typ = "”Õ»" Then
a = bn!mon
bn1 = bn1 + a
Else
a = bn!mon
bn2 = bn2 + a
End If
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
grd6.Col = 7
grd6.Text = Label4.Caption
grd6.CellBackColor = Label4.ForeColor
i = i + 1
End If
End If
bn.MoveNext
Loop
grd6.Rows = i
grd6.Col = 1
grd6.Sort = 2
s = (P - r)
Label5.Caption = P
Label1.Caption = r
End Sub
Private Sub chargegrd6_M()
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
grd6.Clear
grd6.Cols = 8
grd6.Rows = 1
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 1200
grd6.ColWidth(2) = 1000
grd6.ColWidth(3) = 800
grd6.ColWidth(4) = 2000
grd6.ColWidth(5) = 0
grd6.ColWidth(6) = 5800
grd6.ColWidth(7) = 1200
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.ColAlignment(4) = 1
grd6.ColAlignment(5) = 1
grd6.ColAlignment(6) = 1
grd6.ColAlignment(7) = 3
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
If bn!act = Combo2.Text Or Combo3.Text = "Ã„Ì⁄ «·⁄„·Ì« " Then
DT2.Value = bn!dat
j2 = DT2.Month
If j1 = j2 Then
If bn!typ = "”Õ»" Then
a = bn!mon
bn1 = bn1 + a
Else
a = bn!mon
bn2 = bn2 + a
End If
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
grd6.Col = 7
grd6.Text = Label4.Caption
grd6.CellBackColor = Label4.ForeColor
i = i + 1
End If
End If
bn.MoveNext
Loop
grd6.Rows = i
grd6.Col = 1
grd6.Sort = 2
s = (P - r)
Label5.Caption = P
Label1.Caption = r
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
Dim s1 As Double
Dim s2 As Double
tl1 = (pr1 + pf1 + et1 + bn1)
tl2 = (pr2 + pf2 + fn2 + bn2 + dp2)
Label2.Caption = pr1
Label3.Caption = pf1
Label5.Caption = et1
Label8.Caption = bn1

Label6.Caption = pr2
Label7.Caption = pf2
Label9.Caption = fn2
Label10.Caption = dp2
Label1.Caption = bn2

Label19.Caption = tl1
Label16.Caption = tl2

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
grd6.Cols = 8
grd6.Rows = 1
grd6.ColWidth(0) = 0
grd6.ColWidth(1) = 1200
grd6.ColWidth(2) = 1000
grd6.ColWidth(3) = 800
grd6.ColWidth(4) = 2000
grd6.ColWidth(5) = 0
grd6.ColWidth(6) = 5800
grd6.ColWidth(7) = 1200
grd6.ColAlignment(0) = 1
grd6.ColAlignment(1) = 1
grd6.ColAlignment(2) = 1
grd6.ColAlignment(3) = 1
grd6.ColAlignment(4) = 1
grd6.ColAlignment(5) = 1
grd6.ColAlignment(6) = 1
grd6.ColAlignment(7) = 3
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
Label3.Caption = "0"
Label10.Caption = "0"
Label5.Caption = "0"
Label6.Caption = "0"
Label7.Caption = "0"
Label9.Caption = "0"
Label16.Caption = "0"
Label17.Caption = "0"
Label8.Caption = "0"
Label19.Caption = "0"
End Sub


