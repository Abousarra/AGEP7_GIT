VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Compte_PRT 
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
      Height          =   8295
      Left            =   120
      ScaleHeight     =   8295
      ScaleWidth      =   9375
      TabIndex        =   63
      Top             =   1200
      Width           =   9375
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
      IMEMode         =   3  'DISABLE
      Left            =   9120
      ScrollBars      =   2  'Vertical
      TabIndex        =   58
      Top             =   720
      Width           =   2175
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
      Left            =   3960
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   1320
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
      TabIndex        =   24
      Top             =   720
      Width           =   855
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
      TabIndex        =   23
      Top             =   720
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
      TabIndex        =   22
      Top             =   720
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
      ItemData        =   "Compte_PRT.frx":0000
      Left            =   6960
      List            =   "Compte_PRT.frx":0029
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1320
      Width           =   855
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
      TabIndex        =   20
      Top             =   1320
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
      TabIndex        =   19
      Top             =   1320
      Width           =   1455
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
      Left            =   360
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   2040
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
      Left            =   360
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
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
      Left            =   360
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2760
      Width           =   1335
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
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   10080
      ScaleHeight     =   2595
      ScaleWidth      =   2355
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Text            =   "Text3"
         Top             =   1440
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   375
         Left            =   0
         TabIndex        =   2
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
         Format          =   33161217
         CurrentDate     =   42638
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
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄ —”Ê„ «· ”ÃÌ·"
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
         Left            =   1560
         TabIndex        =   60
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label Label11 
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
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label Label19 
         Caption         =   "30"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label36 
         Caption         =   "Label36"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label37 
         Caption         =   "Label37"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   1575
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   240
      TabIndex        =   10
      Top             =   3720
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   16744576
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "«·„»«·€ «·„Êœ⁄…"
      TabPicture(0)   =   "Compte_PRT.frx":0054
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grd5"
      Tab(0).Control(1)=   "Command4"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "«·„»«·€ «·„”ÕÊ»…"
      TabPicture(1)   =   "Compte_PRT.frx":0070
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command5"
      Tab(1).Control(1)=   "grd4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   " ›«’Ì· «·‰›ﬁ« "
      TabPicture(2)   =   "Compte_PRT.frx":008C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command6"
      Tab(2).Control(1)=   "grd3"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   " ›«’Ì· «·Ê«—œ« "
      TabPicture(3)   =   "Compte_PRT.frx":00A8
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "grd2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Command7"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).ControlCount=   2
      Begin VB.CommandButton Command4 
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
         TabIndex        =   14
         Top             =   5040
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
         Left            =   -71160
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
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
         TabIndex        =   12
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
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
         TabIndex        =   11
         Top             =   5040
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grd2 
         Height          =   4695
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
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
            Size            =   9.75
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
         TabIndex        =   16
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
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
            Size            =   9.75
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
         TabIndex        =   17
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
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
            Size            =   9.75
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
         TabIndex        =   18
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
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
      TabIndex        =   25
      Top             =   1200
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
      X1              =   5640
      X2              =   5640
      Y1              =   1680
      Y2              =   3600
   End
   Begin VB.Label Label10 
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
      Left            =   5760
      TabIndex        =   62
      Top             =   2520
      Width           =   1395
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
      Left            =   5760
      TabIndex        =   56
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Line Line1 
      X1              =   3840
      X2              =   3840
      Y1              =   600
      Y2              =   1080
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
      TabIndex        =   55
      Top             =   720
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   2
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   5655
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
      Left            =   5760
      TabIndex        =   54
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·‰ ÌÃ…"
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
      TabIndex        =   53
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„’—Ê›« "
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
      Left            =   7920
      TabIndex        =   52
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label13 
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
      Left            =   5760
      TabIndex        =   51
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ ‰’Ì» «·„ƒ””…"
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
      Left            =   7560
      TabIndex        =   50
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   5760
      TabIndex        =   49
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ „” Õﬁ«  ««·„ÊŸ›Ì‰"
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
      Left            =   7080
      TabIndex        =   48
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„Ã„Ê⁄ «·—”Ê„ «·‘Â—Ì…"
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
      TabIndex        =   47
      Top             =   1800
      Width           =   2055
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
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   1
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   12615
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
      TabIndex        =   45
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ”«» «·‘—ﬂ«¡"
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
      Left            =   4320
      TabIndex        =   44
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰”»… «·‘—«ﬂ…"
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
      Left            =   4200
      TabIndex        =   43
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Index           =   3
      Left            =   4080
      TabIndex        =   42
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   2640
      TabIndex        =   41
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label Label22 
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
      Left            =   2640
      TabIndex        =   40
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      Height          =   5895
      Index           =   3
      Left            =   120
      Top             =   3600
      Width           =   9375
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·‰’Ì»"
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
      Left            =   4320
      TabIndex        =   39
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label27 
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
      TabIndex        =   38
      Top             =   2520
      Width           =   1395
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„»«·€ «·„”ÕÊ»…"
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
      TabIndex        =   37
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label29 
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
      TabIndex        =   36
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„»«·€ «·„Êœ⁄…"
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
      TabIndex        =   35
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—’Ìœ"
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
      TabIndex        =   34
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label31 
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
      TabIndex        =   33
      Top             =   3240
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   4
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Index           =   5
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   3735
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
      Left            =   240
      TabIndex        =   32
      Top             =   1320
      Width           =   1755
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
      Left            =   1680
      TabIndex        =   31
      Top             =   1320
      Width           =   2055
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
      TabIndex        =   30
      Top             =   720
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   2415
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
      Left            =   1440
      TabIndex        =   29
      Top             =   2040
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
      Left            =   1440
      TabIndex        =   28
      Top             =   2400
      Width           =   855
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
      Left            =   1440
      TabIndex        =   27
      Top             =   2760
      Width           =   855
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
      TabIndex        =   26
      Top             =   1680
      Width           =   1935
   End
End
Attribute VB_Name = "Compte_PRT"
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
TreeView1.Nodes.Add , , "PR", "√”„«¡ «·‘—ﬂ«¡"
Call cont
Label19.Caption = eb!pce
Label36.Caption = eb!moi
Do While Not pr.EOF
If pr!act = "1" Then
id1 = pr!sri
id2 = "R" + id1
TreeView1.Nodes.Add "PR", tvwChild, id2, pr!nom
End If
pr.MoveNext
Loop
End Sub

Private Sub Combo1_Change()
Call tous_clear

End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Command1_Click()
Text1.Text = Trim(Text1.Text)
If Text1.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· «·—ﬁ„ «· ”·”·Ì ", vbCritical + arabic
Text1.SetFocus
Exit Sub
End If
Call cont
Do While Not pr.EOF
If Text1.Text = pr!sri Or Val(Text1.Text) = Val(pr!sri) Then
If pr!act = "1" Then
Label26.Caption = pr!nom
Label2.Caption = pr!prc
Label37.Caption = pr!mot
Call tous_clear
Option2.Value = True
Command8_Click
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


Private Sub Command2_Click()
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
If Text2.Text = Label37.Caption Then
Picture2.Visible = False
Else
MsgBox "ﬂ·„… «·”— «· Ì √œŒ· „ €Ì— ’ÕÌÕ…", vbExclamation + arabic
Text2.Text = ""
Text2.SetFocus
End If
End Sub

Private Sub Command3_Click()
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
Do While Not pr.EOF
If Text1.Text = pr!sri Or Val(Text1.Text) = Val(pr!sri) Then
pr!mot = Text5.Text
pr.Update
MsgBox " „ Õ›Ÿ «· €ÌÌ—", vbInformation
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Exit Sub
End If
pr.MoveNext
Loop
End Sub

Private Sub Command8_Click()
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
Command8.Enabled = False
grd2.Visible = False
grd3.Visible = False
grd4.Visible = False
grd5.Visible = False
If Option2.Value = True Then
Call chargegrd2_T
Call chargegrd3_T
Call chargegrd4_5_T
Call Sold_PRT
Label15.Caption = Label31.Caption
Else
Call chargegrd2_M
Call chargegrd3_M
Call chargegrd4_5_M
Call Sold_PRT
End If
grd2.Visible = True
grd3.Visible = True
grd4.Visible = True
grd5.Visible = True
Command8.Enabled = True

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Call MakeTreeViewRTL
Call chargetreeview1
Call couleur_treeview1
Label19.Caption = eb!pce
Label36.Caption = eb!moi
End Sub
Private Sub Option1_Click()
Combo1.Visible = True
Call tous_clear
End Sub

Private Sub Option2_Click()
Combo1.Visible = False
Call tous_clear
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Picture2.Visible = True
Else
Text1.BackColor = &H8080FF
End If
Label37.Caption = ""
Label26.Caption = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
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

Private Sub Text5_Change()
If Len(Text5.Text) > 0 Then
Text5.BackColor = &HC000&
Else
Text5.BackColor = &H8080FF
End If

End Sub

Private Sub Text5_Click()
Text5_Change
End Sub

Private Sub Text6_Change()
If Len(Text6.Text) > 0 Then
Text6.BackColor = &HC000&
Else
Text6.BackColor = &H8080FF
End If

End Sub

Private Sub Text6_Click()
Text6_Change
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As ComctlLib.Node)
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
Private Sub chargegrd4_5_T()
Dim i As Double
Dim j As Double
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd4.Clear
grd4.Cols = 5
grd4.Rows = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1600
grd4.ColWidth(2) = 1500
grd4.ColWidth(3) = 1500
grd4.ColWidth(4) = 3900
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.Row = 0
grd4.Col = 1
grd4.Text = "«· «—ÌŒ"
grd4.Col = 2
grd4.Text = "«·”«⁄…"
grd4.Col = 3
grd4.Text = "«·„»·€"
grd4.Col = 4
grd4.Text = "«· ›«’Ì·"
grd5.Clear
grd5.Cols = 5
grd5.Rows = 1
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 1600
grd5.ColWidth(2) = 1500
grd5.ColWidth(3) = 1500
grd5.ColWidth(4) = 3900
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
grd5.Text = "«· ›«’Ì·"
i = 1
j = 1
P = 0
r = 0
s = 0
Call cont
grd4.Rows = cp.RecordCount + 3
grd5.Rows = cp.RecordCount + 3
Do While Not cp.EOF
If Text1.Text = cp!sri Or Val(Text1.Text) = Val(cp!sri) Then
If cp!typ = "”Õ»" Then
grd4.Row = j
grd4.Col = 0
grd4.Text = cp!aut
grd4.Col = 1
grd4.Text = cp!dat
grd4.Col = 2
grd4.Text = cp!heu
grd4.Col = 3
grd4.Text = cp!mon
grd4.Col = 4
grd4.Text = cp!det
a = cp!mon
P = P + a
j = j + 1
Else
grd5.Row = i
grd5.Col = 0
grd5.Text = cp!aut
grd5.Col = 1
grd5.Text = cp!dat
grd5.Col = 2
grd5.Text = cp!heu
grd5.Col = 3
grd5.Text = cp!mon
grd5.Col = 4
grd5.Text = cp!det
a = cp!mon
r = r + a
i = i + 1
End If
End If
cp.MoveNext
Loop
grd4.Rows = j
grd4.Col = 1
grd4.Sort = 1
grd5.Rows = i
grd5.Col = 1
grd5.Sort = 1
s = (P - r)
Label27.Caption = P
Label29.Caption = r
End Sub
Private Sub chargegrd4_5_M()
Dim i As Double
Dim j As Double
Dim a As Double
Dim P As Double
Dim r As Double
Dim s As Double
grd4.Clear
grd4.Cols = 5
grd4.Rows = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1600
grd4.ColWidth(2) = 1500
grd4.ColWidth(3) = 1500
grd4.ColWidth(4) = 3900
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.Row = 0
grd4.Col = 1
grd4.Text = "«· «—ÌŒ"
grd4.Col = 2
grd4.Text = "«·”«⁄…"
grd4.Col = 3
grd4.Text = "«·„»·€"
grd4.Col = 4
grd4.Text = "«· ›«’Ì·"
grd5.Clear
grd5.Cols = 5
grd5.Rows = 1
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 1600
grd5.ColWidth(2) = 1500
grd5.ColWidth(3) = 1500
grd5.ColWidth(4) = 3900
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
grd5.Text = "«· ›«’Ì·"
i = 1
j = 1
P = 0
r = 0
s = 0
Call cont
grd4.Rows = cp.RecordCount + 3
grd5.Rows = cp.RecordCount + 3
Do While Not cp.EOF
If Text1.Text = cp!sri Or Val(Text1.Text) = Val(cp!sri) Then
DT1.Value = cp!dat
If DT1.Month = Combo1.Text Then
If cp!typ = "”Õ»" Then
grd4.Row = j
grd4.Col = 0
grd4.Text = cp!aut
grd4.Col = 1
grd4.Text = cp!dat
grd4.Col = 2
grd4.Text = cp!heu
grd4.Col = 3
grd4.Text = cp!mon
grd4.Col = 4
grd4.Text = cp!det
a = cp!mon
P = P + a
j = j + 1
Else
grd5.Row = i
grd5.Col = 0
grd5.Text = cp!aut
grd5.Col = 1
grd5.Text = cp!dat
grd5.Col = 2
grd5.Text = cp!heu
grd5.Col = 3
grd5.Text = cp!mon
grd5.Col = 4
grd5.Text = cp!det
a = cp!mon
r = r + a
i = i + 1
End If
End If
End If
cp.MoveNext
Loop
grd4.Rows = j
grd4.Col = 1
grd4.Sort = 1
grd5.Rows = i
grd5.Col = 1
grd5.Sort = 1
s = (P - r)
Label27.Caption = P
Label29.Caption = r
End Sub
Private Sub tous_clear()
Label27.Caption = "0"
Label29.Caption = "0"
Label20.Caption = "0"
Label8.Caption = "0"
Label5.Caption = "0"
Label11.Caption = "0"
Label17.Caption = "0"
Label22.Caption = "0"
Label31.Caption = "0"
Label13.Caption = "0"
Label10.Caption = "0"
grd4.Clear
grd4.Cols = 5
grd4.Rows = 1
grd4.ColWidth(0) = 0
grd4.ColWidth(1) = 1600
grd4.ColWidth(2) = 1500
grd4.ColWidth(3) = 1500
grd4.ColWidth(4) = 3900
grd4.ColAlignment(0) = 1
grd4.ColAlignment(1) = 1
grd4.ColAlignment(2) = 1
grd4.ColAlignment(3) = 1
grd4.ColAlignment(4) = 1
grd4.Row = 0
grd4.Col = 1
grd4.Text = "«· «—ÌŒ"
grd4.Col = 2
grd4.Text = "«·”«⁄…"
grd4.Col = 3
grd4.Text = "«·„»·€"
grd4.Col = 4
grd4.Text = "«· ›«’Ì·"
grd5.Clear
grd5.Cols = 5
grd5.Rows = 1
grd5.ColWidth(0) = 0
grd5.ColWidth(1) = 1600
grd5.ColWidth(2) = 1500
grd5.ColWidth(3) = 1500
grd5.ColWidth(4) = 3900
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
grd5.Text = "«· ›«’Ì·"
grd2.Clear
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 600
grd2.ColWidth(1) = 1100
grd2.ColWidth(2) = 1100
grd2.ColWidth(3) = 1250
grd2.ColWidth(4) = 1100
grd2.ColWidth(5) = 1100
grd2.ColWidth(6) = 1100
grd2.ColWidth(7) = 1250
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
grd2.Row = 0
grd2.Col = 0
grd2.Text = "«·‘Â—"
grd2.Col = 1
grd2.Text = "«·ﬁ”„"
grd2.Col = 2
grd2.Text = "„” Õﬁ«  «· ·«„Ì–"
grd2.Col = 3
grd2.Text = "„” Õﬁ«  «·√”« –…"
grd2.Col = 4
grd2.Text = "«·»«ﬁÌ"
grd2.Col = 5
grd2.Text = "‰”»… «·„ƒ””…"
grd2.Col = 6
grd2.Text = "‰’Ì» «·„ƒ””…"
grd3.Clear
grd3.Cols = 5
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1500
grd3.ColWidth(2) = 1600
grd3.ColWidth(3) = 2000
grd3.ColWidth(4) = 3500
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.Row = 0
grd3.Col = 1
grd3.Text = "«· «—ÌŒ"
grd3.Col = 2
grd3.Text = "«·”«⁄…"
grd3.Col = 3
grd3.Text = "«·„»·€"
grd3.Col = 4
grd3.Text = "«· ›«’Ì·"

End Sub
Private Sub chargegrd3_T()
Dim i As Double
Dim dat1 As Date
Dim dat2 As Date
Dim dat3 As Date
Dim d As Double
Dim sd As Double
grd3.Clear
grd3.Cols = 5
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1500
grd3.ColWidth(2) = 1600
grd3.ColWidth(3) = 1500
grd3.ColWidth(4) = 4000
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.Row = 0
grd3.Col = 1
grd3.Text = "«· «—ÌŒ"
grd3.Col = 2
grd3.Text = "«·”«⁄…"
grd3.Col = 3
grd3.Text = "«·„»·€"
grd3.Col = 4
grd3.Text = "«· ›«’Ì·"
i = 1
a = 0
sd = 0
Call cont
grd3.Rows = dp.RecordCount + 3
Do While Not dp.EOF
dat3 = dp!dat
grd3.Row = i
grd3.Col = 0
grd3.Text = dp!aut
grd3.Col = 1
grd3.Text = dp!dat
grd3.Col = 2
grd3.Text = dp!heu
grd3.Col = 3
grd3.Text = dp!mon
d = dp!mon
sd = sd + d
grd3.Col = 4
grd3.Text = dp!det
i = i + 1
dp.MoveNext
Loop
Label13.Caption = sd
'******
a = 0
sd = 0
Call cont
grd3.Rows = cf.RecordCount + 3
Do While Not cf.EOF
DT1.Value = cf!dat
m = DT1.Month
If cf!typ = "”Õ» „»·€" Then
grd3.Row = i
grd3.Col = 0
grd3.Text = cf!aut
grd3.Col = 1
grd3.Text = cf!dat
grd3.Col = 2
grd3.Text = cf!heu
grd3.Col = 3
grd3.Text = cf!mon
d = cf!mon
sd = sd + d
grd3.Col = 4
grd3.Text = "„‰ ÿ—› «·„ÊŸ› ’«Õ» «·—ﬁ„ «· ”·”·Ì " + cf!sri
i = i + 1
End If
cf.MoveNext
Loop
grd3.Rows = i
Label10.Caption = sd
End Sub
Private Sub chargegrd3_M()
Dim i As Double
Dim m As Double
Dim d As Double
Dim sd As Double
grd3.Clear
grd3.Cols = 5
grd3.Rows = 1
grd3.ColWidth(0) = 0
grd3.ColWidth(1) = 1500
grd3.ColWidth(2) = 1600
grd3.ColWidth(3) = 1500
grd3.ColWidth(4) = 4000
grd3.ColAlignment(0) = 1
grd3.ColAlignment(1) = 1
grd3.ColAlignment(2) = 1
grd3.ColAlignment(3) = 1
grd3.ColAlignment(4) = 1
grd3.Row = 0
grd3.Col = 1
grd3.Text = "«· «—ÌŒ"
grd3.Col = 2
grd3.Text = "«·”«⁄…"
grd3.Col = 3
grd3.Text = "«·„»·€"
grd3.Col = 4
grd3.Text = "«· ›«’Ì·"
i = 1
a = 0
sd = 0
Call cont
grd3.Rows = dp.RecordCount + 3
Do While Not dp.EOF
DT1.Value = dp!dat
m = DT1.Month
If m = Combo1.Text Then
grd3.Row = i
grd3.Col = 0
grd3.Text = dp!aut
grd3.Col = 1
grd3.Text = dp!dat
grd3.Col = 2
grd3.Text = dp!heu
grd3.Col = 3
grd3.Text = dp!mon
d = dp!mon
sd = sd + d
grd3.Col = 4
grd3.Text = dp!det
i = i + 1
End If
dp.MoveNext
Loop
Label13.Caption = sd
'******
a = 0
sd = 0
Call cont
grd3.Rows = cf.RecordCount + 3
Do While Not cf.EOF
DT1.Value = cf!dat
m = DT1.Month
If m = Combo1.Text And cf!typ = "”Õ» „»·€" Then
grd3.Row = i
grd3.Col = 0
grd3.Text = cf!aut
grd3.Col = 1
grd3.Text = cf!dat
grd3.Col = 2
grd3.Text = cf!heu
grd3.Col = 3
grd3.Text = cf!mon
d = cf!mon
sd = sd + d
grd3.Col = 4
grd3.Text = "„‰ ÿ—› «·„ÊŸ› ’«Õ» «·—ﬁ„ «· ”·”·Ì " + cf!sri
i = i + 1
End If
cf.MoveNext
Loop
grd3.Rows = i
Label10.Caption = sd
End Sub

Private Sub chargegrd2_M()
Dim i As Double
Dim j As Double
Dim m As Double
Dim k As Double
Dim e As Double
Dim se As Double
Dim P As Double
Dim sp As Double
Dim r As Double
Dim c As Double
Dim l As Double
Dim sl As Double
Dim s As Double
grd2.Clear
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 600
grd2.ColWidth(1) = 1100
grd2.ColWidth(2) = 1100
grd2.ColWidth(3) = 1250
grd2.ColWidth(4) = 1100
grd2.ColWidth(5) = 1100
grd2.ColWidth(6) = 1100
grd2.ColWidth(7) = 1250
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
grd2.ColAlignment(7) = 1
grd2.Row = 0
grd2.Col = 0
grd2.Text = "«·‘Â—"
grd2.Col = 1
grd2.Text = "«·ﬁ”„"
grd2.Col = 2
grd2.Text = "«Ì—«œ«  «· ·«„Ì–"
grd2.Col = 3
grd2.Text = "„” Õﬁ«  √”« –… ”"
grd2.Col = 4
grd2.Text = "«·»«ﬁÌ"
grd2.Col = 5
grd2.Text = "‰”»… «·„ƒ””…"
grd2.Col = 6
grd2.Text = "‰’Ì» «·„ƒ””…"
grd2.Col = 7
grd2.Text = "‰. Â–« «·‘—Ìﬂ"
i = 1
e = 0
se = 0
P = 0
sp = 0
r = 0
c = 0
l = 0
sl = 0
s = 0
k = Label2.Caption
Call cont
grd2.Rows = pc.RecordCount + 3
Do While Not pc.EOF
m = pc!moi
If m = Combo1.Text And Combo1.Enabled = True Or Label36.Caption = Combo1.Text And m = "0" And Combo1.Enabled = True Then
j = pc!nbr
e = 0
P = 0
r = 0
t = 0
grd2.Row = i
grd2.Col = 0
If pc!moi = "0" Then
grd2.Text = "—. "
Else
grd2.Text = pc!moi
End If
grd2.Col = 1
grd2.Text = pc!cla
grd2.Col = 2
grd2.Text = pc!etu
e = pc!etu
se = se + e
grd2.Col = 3
grd2.Text = pc!pro
P = pc!pro
sp = sp + P
r = (e - P)
grd2.Col = 4
grd2.Text = r
If j > 0 Then
c = Label19.Caption
Else
c = 100
End If
grd2.Col = 5
grd2.Text = c
l = (r * c / 100)
MyNumber = Round(l, 0)
l = MyNumber
sl = sl + l
grd2.Col = 6
grd2.Text = l
s = (l * k / 100)
MyNumber = Round(s, 0)
s = MyNumber
grd2.Col = 7
grd2.Text = s
i = i + 1
End If
pc.MoveNext
Loop
grd2.Rows = i
Label20.Caption = se
Label8.Caption = sp
Label5.Caption = sl
End Sub
Private Sub chargegrd2_T()
Dim i As Double
Dim j As Double
Dim m As Double
Dim k As Double
Dim e As Double
Dim se As Double
Dim P As Double
Dim sp As Double
Dim r As Double
Dim c As Double
Dim l As Double
Dim sl As Double
Dim s As Double
grd2.Clear
grd2.Cols = 8
grd2.Rows = 1
grd2.ColWidth(0) = 600
grd2.ColWidth(1) = 1100
grd2.ColWidth(2) = 1100
grd2.ColWidth(3) = 1250
grd2.ColWidth(4) = 1100
grd2.ColWidth(5) = 1100
grd2.ColWidth(6) = 1100
grd2.ColWidth(7) = 1250
grd2.ColAlignment(0) = 1
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.ColAlignment(4) = 1
grd2.ColAlignment(5) = 1
grd2.ColAlignment(6) = 1
grd2.ColAlignment(7) = 1
grd2.Row = 0
grd2.Col = 0
grd2.Text = "«·‘Â—"
grd2.Col = 1
grd2.Text = "«·ﬁ”„"
grd2.Col = 2
grd2.Text = "«Ì—«œ«  «· ·«„Ì–"
grd2.Col = 3
grd2.Text = "„” Õﬁ«  √”« –… ”"
grd2.Col = 4
grd2.Text = "«·»«ﬁÌ"
grd2.Col = 5
grd2.Text = "‰”»… «·„ƒ””…"
grd2.Col = 6
grd2.Text = "‰’Ì» «·„ƒ””…"
grd2.Col = 7
grd2.Text = "‰.Â–« «·‘—Ìﬂ"
i = 1
e = 0
se = 0
P = 0
sp = 0
r = 0
c = 0
l = 0
sl = 0
s = 0
k = Label2.Caption
Call cont
grd2.Rows = pc.RecordCount + 3
Do While Not pc.EOF
m = pc!moi
j = pc!nbr
e = 0
P = 0
r = 0
t = 0
grd2.Row = i
grd2.Col = 0
If pc!moi = "0" Then
grd2.Text = "—. "
Else
grd2.Text = pc!moi
End If
grd2.Col = 1
grd2.Text = pc!cla
grd2.Col = 2
grd2.Text = pc!etu
e = pc!etu
se = se + e
grd2.Col = 3
grd2.Text = pc!pro
P = pc!pro
sp = sp + P
r = (e - P)
grd2.Col = 4
grd2.Text = r
If j > 0 Then
c = Label19.Caption
Else
c = 100
End If
grd2.Col = 5
grd2.Text = c
l = (r * c / 100)
MyNumber = Round(l, 0)
l = MyNumber
sl = sl + l
grd2.Col = 6
grd2.Text = l
s = (l * k / 100)
MyNumber = Round(s, 0)
s = MyNumber
grd2.Col = 7
grd2.Text = s
i = i + 1
pc.MoveNext
Loop
grd2.Rows = i
Label20.Caption = se
Label8.Caption = sp
Label5.Caption = sl
End Sub
Private Sub Sold_PRT()
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim f As Double
Dim g As Double
Dim h As Double
Dim s As Double
a = Label5.Caption
b = Label13.Caption
s = Label10.Caption
c = (a - b - s)
Label17.Caption = c
d = Label2.Caption
e = (c * d / 100)
MyNumber = Round(e, 0)
e = MyNumber
Label22.Caption = e
f = Label27.Caption
g = Label29.Caption
h = (e + g - f)
Label31.Caption = h
End Sub
