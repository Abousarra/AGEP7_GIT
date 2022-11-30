VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Compte_PRF 
   AutoRedraw      =   -1  'True
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
      TabIndex        =   58
      Top             =   1320
      Width           =   9375
   End
   Begin VB.TextBox Text1 
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
      Height          =   330
      Left            =   9120
      ScrollBars      =   2  'Vertical
      TabIndex        =   55
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
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
      Left            =   3960
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   1440
      Width           =   1215
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
      Top             =   720
      Width           =   855
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
      TabIndex        =   18
      Top             =   1440
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
      Left            =   5400
      TabIndex        =   17
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "Compte_PRF.frx":0000
      Left            =   6960
      List            =   "Compte_PRF.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1440
      Width           =   855
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
      Left            =   600
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text6 
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
      Height          =   330
      Left            =   240
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text5 
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
      Height          =   330
      Left            =   240
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text4 
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
      Height          =   330
      Left            =   240
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   240
      TabIndex        =   4
      Top             =   3960
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9763
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
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
      TabCaption(0)   =   "«·«Ìœ«⁄"
      TabPicture(0)   =   "Compte_PRF.frx":0054
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grd5"
      Tab(0).Control(1)=   "Command9"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "«·”Õ»"
      TabPicture(1)   =   "Compte_PRF.frx":0070
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grd4"
      Tab(1).Control(1)=   "Command8"
      Tab(1).Control(2)=   "Picture1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "„” Õﬁ«  «·‰”»…"
      TabPicture(2)   =   "Compte_PRF.frx":008C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grd3"
      Tab(2).Control(1)=   "Command6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "„” Õﬁ«  «·‘Â—"
      TabPicture(3)   =   "Compte_PRF.frx":00A8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grd2"
      Tab(3).Control(1)=   "Command5"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "„” Õﬁ«  «·”«⁄…"
      TabPicture(4)   =   "Compte_PRF.frx":00C4
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "grd1"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Command4"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).ControlCount=   2
      Begin VB.PictureBox Picture1 
         Height          =   1575
         Left            =   -70800
         ScaleHeight     =   1515
         ScaleWidth      =   3555
         TabIndex        =   10
         Top             =   1680
         Visible         =   0   'False
         Width           =   3615
         Begin MSComCtl2.DTPicker DT1 
            Height          =   375
            Left            =   1920
            TabIndex        =   57
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Format          =   124977153
            CurrentDate     =   43029
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   360
            TabIndex        =   11
            Text            =   "Text3"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label27 
            Caption         =   "Label27"
            Height          =   255
            Left            =   2400
            TabIndex        =   56
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄·«Ê« "
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
            TabIndex        =   52
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label5 
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
            Left            =   0
            TabIndex        =   51
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ √Œ—« "
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
            Left            =   1080
            TabIndex        =   50
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label11 
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
            Left            =   120
            TabIndex        =   49
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "30"
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   120
            Width           =   1335
         End
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
         Height          =   330
         Left            =   3840
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5160
         Width           =   1455
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
         Height          =   330
         Left            =   -71160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5160
         Width           =   1455
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
         Height          =   330
         Left            =   -71160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5160
         Width           =   1455
      End
      Begin VB.CommandButton Command8 
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
         Height          =   330
         Left            =   -71160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5160
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
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
         Height          =   330
         Left            =   -71160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5160
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grd3 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8493
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
      Begin MSFlexGridLib.MSFlexGrid grd4 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8493
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
      Begin MSFlexGridLib.MSFlexGrid grd5 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8493
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
         Height          =   4815
         Left            =   -74880
         TabIndex        =   53
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8493
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
      Begin MSFlexGridLib.MSFlexGrid grd1 
         Height          =   4815
         Left            =   120
         TabIndex        =   54
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8493
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
   Begin ComctlLib.TreeView TreeView1 
      Height          =   8295
      Left            =   9600
      TabIndex        =   22
      Top             =   1320
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   14631
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
   Begin VB.Line Line2 
      X1              =   2520
      X2              =   9480
      Y1              =   3120
      Y2              =   3120
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
      TabIndex        =   47
      Top             =   720
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   615
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
      Left            =   6720
      TabIndex        =   46
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„” Õﬁ«  «·”«⁄…"
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
      Left            =   7920
      TabIndex        =   45
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label6 
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
      Left            =   7920
      TabIndex        =   44
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label8 
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
      Left            =   6360
      TabIndex        =   43
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„” Õﬁ«  «·‰”»…"
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
      Left            =   7920
      TabIndex        =   42
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label10 
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
      Left            =   6360
      TabIndex        =   41
      Top             =   2640
      Width           =   1395
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «·„” Õﬁ« "
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
      Left            =   7680
      TabIndex        =   40
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label13 
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
      Left            =   6360
      TabIndex        =   39
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «·”Õ»"
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
      Left            =   4560
      TabIndex        =   38
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label15 
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
      Left            =   3120
      TabIndex        =   37
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—’Ìœ «·‰Â«∆Ì"
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
      Left            =   4440
      TabIndex        =   36
      Top             =   3240
      Width           =   1455
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
      Left            =   2760
      TabIndex        =   35
      Top             =   3240
      Width           =   1695
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
      TabIndex        =   34
      Top             =   720
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   600
      Y2              =   1200
   End
   Begin VB.Label Label20 
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
      Left            =   6360
      TabIndex        =   33
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» «·√”« –…"
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
      Left            =   4680
      TabIndex        =   32
      Top             =   0
      Width           =   3735
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
      Left            =   1920
      TabIndex        =   31
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label22 
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   30
      Top             =   1440
      Width           =   1875
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   5
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   3
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   4
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   6975
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
      TabIndex        =   29
      Top             =   1800
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
      TabIndex        =   28
      Top             =   2880
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
      TabIndex        =   27
      Top             =   2520
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
      TabIndex        =   26
      Top             =   2160
      Width           =   855
   End
   Begin VB.Shape Shape1 
      Height          =   5775
      Index           =   2
      Left            =   120
      Top             =   3840
      Width           =   9375
   End
   Begin VB.Label Label24 
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
      TabIndex        =   25
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label Label25 
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
      Left            =   3120
      TabIndex        =   24
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «·«Ìœ«⁄"
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
      Left            =   4440
      TabIndex        =   23
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "Compte_PRF"
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
TreeView1.Nodes.Add , , "PF", "√”„«¡ «·√”« –…"
Call cont
Label19.Caption = eb!pcp
Do While Not pf.EOF
If pf!act = "1" Then
id1 = pf!sri
id2 = "P" + id1
TreeView1.Nodes.Add "PF", tvwChild, id2, pf!nom
End If
pf.MoveNext
Loop
End Sub

Private Sub Combo1_Change()
On Error Resume Next
Call chargegrd_clear

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
Do While Not pf.EOF
If Text1.Text = pf!sri Or Val(Text1.Text) = Val(pf!sri) Then
If pf!act = "1" Then
Label24.Caption = pf!nom
Label27.Caption = pf!mot
Call chargegrd_clear
Option2.Value = True
Command10_Click
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

Private Sub Command10_Click()
On Error Resume Next
Dim e As Double
If Option1.Value = False And Option2.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ √Õœ «·ŒÌ«—«  ⁄·Ï «·Ì„Ì‰", vbCritical
Exit Sub
End If
If Option1.Value = True Then
If Combo1.Text = "" Then
MsgBox "Ì—ÃÏ  ÕœÌœ «·‘Â—", vbCritical
Exit Sub
End If
End If
Command10.Enabled = False
grd1.Visible = False
grd2.Visible = False
grd3.Visible = False
grd4.Visible = False
grd5.Visible = False
If Option2.Value = True Then
Call chargegrd3_T
Call chargegrd1_2_res3_T
Call chargegrd4_5_T
Label22.Caption = Label17.Caption
e = Label22.Caption
If e < 0 Then
Label22.ForeColor = &HFF&
Else
Label22.ForeColor = &HFF00&
End If
Else
Call chargegrd3_M
Call chargegrd1_2_res3_M
Call chargegrd4_5_M
End If
grd1.Visible = True
grd2.Visible = True
grd3.Visible = True
grd4.Visible = True
grd5.Visible = True
Command10.Enabled = True

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
If Text2.Text = Label27.Caption Then
Picture2.Visible = False
Else
MsgBox "ﬂ·„… «·”— «· Ì √œŒ· „ €Ì— ’ÕÌÕ…", vbExclamation + arabic
Text2.Text = ""
Text2.SetFocus
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0
Call MakeTreeViewRTL
Call chargetreeview1
Call couleur_treeview1
End Sub
Private Sub chargegrd3_T()
On Error Resume Next
Dim i As Double
Dim e As Double
Dim P As Double
Dim r As Double
Dim t As Double
Dim c As Double
Dim h As Double
i = 1
Call cont
grd3.Rows = pc.RecordCount + 3
Do While Not pc.EOF
If Val(pc!nbr) > 0 Then
grd3.RowHeight(i) = 0
e = 0
P = 0
r = 0
t = 0
m = 0
c = Label19.Caption
grd3.Row = i
grd3.Col = 0
grd3.Text = pc!moi
grd3.Col = 1
grd3.Text = pc!cla
grd3.Col = 2
grd3.Text = pc!etu
e = pc!etu
grd3.Col = 3
grd3.Text = pc!pro
P = pc!pro
r = (e - P)
grd3.Col = 4
grd3.Text = r
t = (r * c / 100)
MyNumber = Round(t, 0)
t = MyNumber
grd3.Col = 5
grd3.Text = c
grd3.Col = 6
grd3.Text = t
grd3.Col = 7
grd3.Text = pc!nbr
h = pc!nbr
If h > 0 Then
t = t / h
MyNumber = Round(t, 0)
t = MyNumber
grd3.Col = 8
grd3.Text = t
Else
grd3.Col = 8
grd3.Text = "0"
End If
grd3.Col = 9
grd3.Text = "0"
i = i + 1
End If
pc.MoveNext
Loop
grd3.Rows = i
'grd3.Col = 1
'grd3.Sort = 2
End Sub
Private Sub chargegrd3_M()
On Error Resume Next
Dim i As Double
Dim e As Double
Dim P As Double
Dim r As Double
Dim t As Double
Dim c As Double
Dim h As Double
i = 1
Call cont
grd3.Rows = pc.RecordCount + 3
Do While Not pc.EOF
If pc!moi = Combo1.Text Then
If Val(pc!nbr) > 0 Then
grd3.RowHeight(i) = 0
e = 0
P = 0
r = 0
t = 0
m = 0
c = Label19.Caption
grd3.Row = i
grd3.Col = 0
grd3.Text = pc!moi
grd3.Col = 1
grd3.Text = pc!cla
grd3.Col = 2
grd3.Text = pc!etu
e = pc!etu
grd3.Col = 3
grd3.Text = pc!pro
P = pc!pro
r = (e - P)
grd3.Col = 4
grd3.Text = r
t = (r * c / 100)
MyNumber = Round(t, 0)
t = MyNumber
grd3.Col = 5
grd3.Text = c
grd3.Col = 6
grd3.Text = t
grd3.Col = 7
grd3.Text = pc!nbr
h = pc!nbr
If h > 0 Then
t = t / h
MyNumber = Round(t, 0)
t = MyNumber
grd3.Col = 8
grd3.Text = t
Else
grd3.Col = 8
grd3.Text = "0"
End If
grd3.Col = 9
grd3.Text = "0"
i = i + 1
End If
End If
pc.MoveNext
Loop
grd3.Rows = i
'grd3.Col = 1
'grd3.Sort = 2
End Sub
Private Sub chargegrd4_5_T()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim e As Double
Dim s As Double
Dim se As Double
Dim ss As Double
i = 1
j = 1
se = 0
ss = 0
Call cont
grd5.Rows = cs.RecordCount + 3
grd4.Rows = cs.RecordCount + 3
Do While Not cs.EOF
If cs!sri = Text1.Text Or Val(cs!sri) = Val(Text1.Text) Then
If cs!typ = "”Õ»" Then
e = cs!mon
se = se + e
grd4.Row = i
grd4.Col = 0
grd4.Text = cs!dat
grd4.Col = 1
grd4.Text = cs!heu
grd4.Col = 2
grd4.Text = cs!mon
grd4.Col = 3
grd4.Text = cs!det
i = i + 1
Else
s = cs!mon
ss = ss + s
grd5.Row = j
grd5.Col = 0
grd5.Text = cs!dat
grd5.Col = 1
grd5.Text = cs!heu
grd5.Col = 2
grd5.Text = cs!mon
grd5.Col = 3
grd5.Text = cs!det
j = j + 1
End If
End If
cs.MoveNext
Loop
grd4.Rows = i
grd5.Rows = j
Label15.Caption = se
Label25.Caption = ss
s = Label13.Caption
e = (s + ss) - se
Label17.Caption = e
End Sub
Private Sub chargegrd4_5_M()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim e As Double
Dim s As Double
Dim se As Double
Dim ss As Double
Dim q As Double
i = 1
j = 1
se = 0
ss = 0
Call cont
grd5.Rows = cs.RecordCount + 3
grd4.Rows = cs.RecordCount + 3
Do While Not cs.EOF
DT1.Value = cs!dat
q = DT1.Month
If q = Combo1.Text Then
If cs!sri = Text1.Text Or Val(cs!sri) = Val(Text1.Text) Then
If cs!typ = "”Õ»" Then
e = cs!mon
se = se + e
grd4.Row = i
grd4.Col = 0
grd4.Text = cs!dat
grd4.Col = 1
grd4.Text = cs!heu
grd4.Col = 2
grd4.Text = cs!mon
grd4.Col = 3
grd4.Text = cs!det
i = i + 1
Else
s = cs!mon
ss = ss + s
grd5.Row = j
grd5.Col = 0
grd5.Text = cs!dat
grd5.Col = 1
grd5.Text = cs!heu
grd5.Col = 2
grd5.Text = cs!mon
grd5.Col = 3
grd5.Text = cs!det
j = j + 1
End If
End If
End If
cs.MoveNext
Loop
grd4.Rows = i
grd5.Rows = j
Label15.Caption = se
Label25.Caption = ss
s = Label13.Caption
e = (s + ss) - se
Label17.Caption = e
End Sub

Private Sub Option1_Click()
On Error Resume Next
Combo1.Visible = True
Call chargegrd_clear

End Sub

Private Sub Option2_Click()
On Error Resume Next
Combo1.Visible = False
Call chargegrd_clear
End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Else
Text1.BackColor = &H8080FF
End If
Label24.Caption = ""
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
Private Sub chargegrd_clear()
On Error Resume Next
Label20.Caption = "0"
Label8.Caption = "0"
Label10.Caption = "0"
Label13.Caption = "0"
Label15.Caption = "0"
Label25.Caption = "0"
Label17.Caption = "0"
grd1.Clear
grd1.Cols = 7
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 2000
grd1.ColWidth(2) = 0
grd1.ColWidth(3) = 2100
grd1.ColWidth(4) = 2200
grd1.ColWidth(5) = 2200
grd1.ColWidth(6) = 0
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
grd1.Text = "«·ﬁ”„"
grd1.Col = 4
grd1.Text = "«·”«⁄« "
grd1.Col = 5
grd1.Text = "«·„»·€"
grd1.Col = 6
grd1.Text = "«· ›«’Ì·"
grd2.Clear
grd2.Cols = 7
grd2.Rows = 1
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 2000
grd2.ColWidth(2) = 0
grd2.ColWidth(3) = 2100
grd2.ColWidth(4) = 2200
grd2.ColWidth(5) = 2200
grd2.ColWidth(6) = 0
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
grd2.Text = "«·ﬁ”„"
grd2.Col = 4
grd2.Text = "«·‘Â—"
grd2.Col = 5
grd2.Text = "«·„»·€"
grd2.Col = 6
grd2.Text = "«· ›«’Ì·"
grd3.Clear
grd3.Cols = 11
grd3.Rows = 1
grd3.ColWidth(0) = 600
grd3.ColWidth(1) = 600
grd3.ColWidth(2) = 900
grd3.ColWidth(3) = 900
grd3.ColWidth(4) = 900
grd3.ColWidth(5) = 800
grd3.ColWidth(6) = 900
grd3.ColWidth(7) = 500
grd3.ColWidth(8) = 900
grd3.ColWidth(9) = 600
grd3.ColWidth(10) = 900
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.ColAlignment(5) = 1
grd3.ColAlignment(6) = 1
grd3.ColAlignment(7) = 1
grd3.ColAlignment(8) = 1
grd3.ColAlignment(9) = 1
grd3.ColAlignment(10) = 1
grd3.Row = 0
grd3.Col = 0
grd3.Text = "«·‘Â—"
grd3.Col = 1
grd3.Text = "«·ﬁ”„"
grd3.Col = 2
grd3.Text = "„.«· ·«„Ì–"
grd3.Col = 3
grd3.Text = "„.«·√”« –…"
grd3.Col = 4
grd3.Text = "«·»«ﬁÌ"
grd3.Col = 5
grd3.Text = "‰.«·√”« –…"
grd3.Col = 6
grd3.Text = "’.«·√”« –…"
grd3.Col = 7
grd3.Text = "”.ﬁ"
grd3.Col = 8
grd3.Text = "‰.«·”«⁄…"
grd3.Col = 9
grd3.Text = "”.√”"
grd3.Col = 10
grd3.Text = "’.«·√” «–"
grd4.Clear
grd4.Cols = 4
grd4.Rows = 1
grd4.ColWidth(0) = 2000
grd4.ColWidth(1) = 2100
grd4.ColWidth(2) = 2200
grd4.ColWidth(3) = 2200
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.Row = 0
grd4.Col = 0
grd4.Text = "«· «—ÌŒ"
grd4.Col = 1
grd4.Text = "«·”«⁄…"
grd4.Col = 2
grd4.Text = "«·„»·€"
grd4.Col = 3
grd4.Text = "«· ›«’Ì·"
grd5.Clear
grd5.Cols = 4
grd5.Rows = 1
grd5.ColWidth(0) = 2000
grd5.ColWidth(1) = 2100
grd5.ColWidth(2) = 2200
grd5.ColWidth(3) = 2200
grd5.ColAlignment(0) = 1
grd5.ColAlignment(1) = 1
grd5.ColAlignment(2) = 1
grd5.ColAlignment(3) = 1
grd5.Row = 0
grd5.Col = 0
grd5.Text = "«· «—ÌŒ"
grd5.Col = 1
grd5.Text = "«·”«⁄…"
grd5.Col = 2
grd5.Text = "«·„»·€"
grd5.Col = 3
grd5.Text = "«· ›«’Ì·"
End Sub


Private Sub chargegrd1_2_res3_T()
On Error Resume Next
Dim n As Double
Dim i As Double
Dim m1 As Double
Dim m2 As Double
Dim m3 As Double
Dim cl1 As String
Dim cl2 As String
Dim n1 As Double
Dim n2 As Double
Dim n3 As Double
Dim i1 As Double
Dim i2 As Double
Dim i3 As Double
Dim h As Double
Dim k As Double
Dim sh As Double
Dim m As Double
Dim sm As Double
Dim s As Double
Dim sp As Double
Dim tsp As Double
i1 = 1
i2 = 1
i3 = 1
sh = 0
sm = 0
sp = 0
tsp = 0
Call cont
grd1.Rows = pp.RecordCount + 3
grd2.Rows = pp.RecordCount + 3
Do While Not pp.EOF
If Text1.Text = pp!sri Or Val(Text1.Text) = Val(pp!sri) Then
If pp!eta = "H" Then
grd1.Row = i1
grd1.Col = 0
grd1.Text = pp!aut
grd1.Col = 1
grd1.Text = pp!dat
grd1.Col = 2
grd1.Text = ""
grd1.Col = 3
grd1.Text = pp!cla
grd1.Col = 4
grd1.Text = pp!nbh
grd1.Col = 5
grd1.Text = pp!mon
grd1.Col = 6
grd1.Text = pp!det
k = pp!nbh
h = pp!mon
h = (k * h)
sh = sh + h
i1 = i1 + 1
End If
If pp!eta = "M" Then
grd2.Row = i2
grd2.Col = 0
grd2.Text = pp!aut
grd2.Col = 1
grd2.Text = pp!dat
grd2.Col = 2
grd2.Text = ""
grd2.Col = 3
grd2.Text = pp!cla
grd2.Col = 4
grd2.Text = pp!moi
grd2.Col = 5
grd2.Text = pp!mon
grd2.Col = 6
grd2.Text = pp!det
m = pp!mon
sm = sm + m
i2 = i2 + 1
End If
If pp!eta = "P" Then
m1 = pp!moi
cl1 = pp!cla
n1 = pp!nbh
n = grd3.Rows
For i = 1 To n - 1
s = 0
grd3.Row = i
grd3.Col = 0
m2 = grd3.Text
grd3.Col = 1
cl2 = grd3.Text
grd3.Col = 8
s = grd3.Text
grd3.Col = 9
n2 = grd3.Text
If m1 = m2 And cl1 = cl2 Then
sp = (s * n1)
tsp = tsp + sp
n3 = (n1 + n2)
grd3.Row = i
grd3.Col = 9
grd3.Text = n3
s = (n3 * s)
grd3.Col = 10
grd3.Text = s
grd3.RowHeight(i) = 270
End If
Next i
End If
End If
pp.MoveNext
Loop
grd1.Rows = i1
grd2.Rows = i2
Label20.Caption = sh
Label8.Caption = sm
Label10.Caption = tsp
Label13.Caption = (sh + sm + tsp)
End Sub
Private Sub chargegrd1_2_res3_M()
On Error Resume Next
Dim n As Double
Dim i As Double
Dim m1 As Double
Dim m2 As Double
Dim m3 As Double
Dim cl1 As String
Dim cl2 As String
Dim n1 As Double
Dim n2 As Double
Dim n3 As Double
Dim i1 As Double
Dim i2 As Double
Dim i3 As Double
Dim h As Double
Dim k As Double
Dim sh As Double
Dim m As Double
Dim sm As Double
Dim s As Double
Dim sp As Double
Dim tsp As Double
Dim q As Double
i1 = 1
i2 = 1
i3 = 1
sh = 0
sm = 0
sp = 0
tsp = 0
Call cont
grd1.Rows = pp.RecordCount + 3
grd2.Rows = pp.RecordCount + 3
Do While Not pp.EOF
If Text1.Text = pp!sri Or Val(Text1.Text) = Val(pp!sri) Then
If pp!eta = "H" Then
DT1.Value = pp!dat
q = DT1.Month
If q = Combo1.Text Then
grd1.Row = i1
grd1.Col = 0
grd1.Text = pp!aut
grd1.Col = 1
grd1.Text = pp!dat
grd1.Col = 2
grd1.Text = ""
grd1.Col = 3
grd1.Text = pp!cla
grd1.Col = 4
grd1.Text = pp!nbh
grd1.Col = 5
grd1.Text = pp!mon
grd1.Col = 6
grd1.Text = pp!det
k = pp!nbh
h = pp!mon
h = (k * h)
sh = sh + h
i1 = i1 + 1
End If
End If
If pp!eta = "M" Then
q = pp!moi
If q = Combo1.Text Then
grd2.Row = i2
grd2.Col = 0
grd2.Text = pp!aut
grd2.Col = 1
grd2.Text = pp!dat
grd2.Col = 2
grd2.Text = ""
grd2.Col = 3
grd2.Text = pp!cla
grd2.Col = 4
grd2.Text = pp!moi
grd2.Col = 5
grd2.Text = pp!mon
grd2.Col = 6
grd2.Text = pp!det
m = pp!mon
sm = sm + m
i2 = i2 + 1
End If
End If
If pp!eta = "P" Then
m1 = pp!moi
cl1 = pp!cla
n1 = pp!nbh
n = grd3.Rows
For i = 1 To n - 1
s = 0
grd3.Row = i
grd3.Col = 0
m2 = grd3.Text
grd3.Col = 1
cl2 = grd3.Text
grd3.Col = 8
s = grd3.Text
grd3.Col = 9
n2 = grd3.Text
If m1 = m2 And cl1 = cl2 Then
sp = (s * n1)
tsp = tsp + sp
n3 = (n1 + n2)
grd3.Row = i
grd3.Col = 9
grd3.Text = n3
s = (n3 * s)
grd3.Col = 10
grd3.Text = s
grd3.RowHeight(i) = 270
End If
Next i
End If
End If
pp.MoveNext
Loop
grd1.Rows = i1
grd2.Rows = i2
Label20.Caption = sh
Label8.Caption = sm
Label10.Caption = tsp
Label13.Caption = (sh + sm + tsp)
End Sub


