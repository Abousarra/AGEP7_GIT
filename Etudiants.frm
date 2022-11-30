VERSION 5.00
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Etudiants 
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
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   360
      ScaleHeight     =   2385
      ScaleWidth      =   5385
      TabIndex        =   46
      Top             =   2040
      Visible         =   0   'False
      Width           =   5415
      Begin MSFlexGridLib.MSFlexGrid grd2 
         Height          =   2175
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   3836
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
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   6
         Height          =   2415
         Left            =   0
         Top             =   0
         Width           =   5415
      End
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
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
      IMEMode         =   3  'DISABLE
      Left            =   240
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Top             =   2680
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "”Õ»"
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
      Left            =   5520
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   9000
      UseMaskColor    =   -1  'True
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   3480
      ScaleHeight     =   3195
      ScaleWidth      =   5835
      TabIndex        =   33
      Top             =   3360
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton Command8 
         Caption         =   " ÕœÌœ «·„” ÊÌ« "
         Height          =   495
         Left            =   4440
         TabIndex        =   41
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   " Œ“Ì‰ «· ·«„Ì–"
         Height          =   495
         Left            =   4440
         TabIndex        =   40
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         Caption         =   "√Œ– ’Ê—…"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   720
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   1800
         Top             =   120
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   255
         Left            =   2880
         TabIndex        =   44
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Label10"
         Height          =   255
         Left            =   2880
         TabIndex        =   39
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "Label18"
         Height          =   255
         Left            =   4080
         TabIndex        =   36
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "Label17"
         Height          =   375
         Left            =   2520
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label16 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command4 
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
      Left            =   8160
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
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
      Left            =   10920
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2160
      UseMaskColor    =   -1  'True
      Width           =   1575
   End
   Begin VB.TextBox Text7 
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
      IMEMode         =   3  'DISABLE
      Left            =   5880
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   1680
      Width           =   1335
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
      Left            =   8160
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   1440
      Width           =   2895
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
      ItemData        =   "Etudiants.frx":0000
      Left            =   5880
      List            =   "Etudiants.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1200
      Width           =   1335
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
      ItemData        =   "Etudiants.frx":001B
      Left            =   5880
      List            =   "Etudiants.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
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
      IMEMode         =   3  'DISABLE
      Left            =   3000
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Text3 
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
      Left            =   3000
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "„”Õ «·’Ê—…"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      MaskColor       =   &H00FF0000&
      Picture         =   "Etudiants.frx":004D
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "«—›«ﬁ ’Ê—…"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      MaskColor       =   &H00FF0000&
      Picture         =   "Etudiants.frx":03AC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
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
      Height          =   285
      Left            =   9480
      TabIndex        =   7
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "«”„ «· ·„Ì–"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8160
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "—ﬁ„ «·‰œ«¡"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "«·—ﬁ„ «·Êÿ‰Ì"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   4
      Top             =   2760
      Width           =   1335
   End
   Begin VB.OptionButton Option5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "—ﬁ„ Â« › «·ÊﬂÌ·"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.OptionButton Option6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "√‰ÀÏ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5880
      TabIndex        =   2
      Top             =   2160
      Width           =   615
   End
   Begin VB.OptionButton Option7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "–ﬂ—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   1
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox Text4 
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
      Left            =   8160
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1800
      Width           =   2895
   End
   Begin BARCODEXLib.BarcodeX BX1 
      Height          =   615
      Left            =   8160
      Top             =   840
      Width           =   2895
      _Version        =   65536
      _ExtentX        =   5106
      _ExtentY        =   1085
      _StockProps     =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0000001"
      BarcodeType     =   6
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   5775
      Left            =   240
      TabIndex        =   15
      Top             =   3120
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10186
      _Version        =   393216
      BackColor       =   32768
      BackColorFixed  =   32768
      BackColorBkg    =   32768
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
      Left            =   3000
      TabIndex        =   16
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
   Begin VB.Label Label11 
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
      Left            =   2040
      TabIndex        =   42
      Top             =   2680
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      Height          =   400
      Left            =   120
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·Ã‰”"
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
      TabIndex        =   30
      Top             =   2160
      Width           =   855
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   240
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1455
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
      Left            =   6720
      TabIndex        =   28
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰«  «· ·«„Ì–"
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
      TabIndex        =   27
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·«”„ »«·⁄—»Ì…"
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
      Left            =   10920
      TabIndex        =   26
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Left            =   11040
      TabIndex        =   25
      Top             =   840
      Width           =   1455
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
      Left            =   6720
      TabIndex        =   24
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·—ﬁ„ «·Êÿ‰Ì"
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
      TabIndex        =   23
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
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
      TabIndex        =   22
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Index           =   9
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   12615
   End
   Begin VB.Shape Shape1 
      Height          =   6855
      Index           =   0
      Left            =   120
      Top             =   2640
      Width           =   12615
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·‰œ«¡"
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
      TabIndex        =   21
      Top             =   1680
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   8040
      X2              =   8040
      Y1              =   720
      Y2              =   2520
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   5760
      Y1              =   720
      Y2              =   2520
   End
   Begin VB.Line Line3 
      X1              =   2880
      X2              =   2880
      Y1              =   720
      Y2              =   2520
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Â« › «·ÊﬂÌ·"
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
      TabIndex        =   20
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«· — Ì» ÌﬂÊ‰ Õ”»"
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
      Left            =   10440
      TabIndex        =   19
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·«”„ »«·›—‰”Ì…"
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
      Left            =   10920
      TabIndex        =   18
      Top             =   1800
      Width           =   1575
   End
End
Attribute VB_Name = "Etudiants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Dim data As New Access.Application
Dim seri As String
Dim basa As String

Private Sub chargcombo1()
On Error Resume Next
Combo1.Clear
Call cont
Do While Not cl.EOF
If Combo2.Text = cl!niv And cl!act = "1" Then
Combo1.AddItem cl!cla
End If
cl.MoveNext
Loop
End Sub

Private Sub Combo1_Change()
On Error Resume Next
If Len(Combo1.Text) > 0 Then
Combo1.BackColor = &HC000&
grd1.Visible = False
Call chargegrd1
grd1.Visible = True
Call cont
Do While Not cl.EOF
If Combo1.Text = cl!cla Then
Label17.Caption = cl!aut
Exit Sub
End If
cl.MoveNext
Loop
Else
Combo1.BackColor = &H8080FF
End If
End Sub

Private Sub Combo1_Click()
On Error Resume Next
Combo1_Change
End Sub

Private Sub Combo2_Change()
On Error Resume Next
If Len(Combo2.Text) > 0 Then
Combo2.BackColor = &HC000&
Call chargcombo1
Combo1.BackColor = &H8080FF
Call chargegrd1_clear
Else
Combo2.BackColor = &H8080FF
End If

End Sub

Private Sub Combo2_Click()
On Error Resume Next
Combo2_Change
End Sub

Private Sub Command1_Click()
On Error GoTo P
PicFile = ""
    CommonDialog1.CancelError = True
    CommonDialog1.Filter = "Picture JPG |*.jpg|Picture Gif |*.Gif|Picture Bmp |*.Bmp|Picture Icon |*.ICO|All Picture |*.*"
    CommonDialog1.DialogTitle = "Picture"
    CommonDialog1.ShowOpen
    PicFile = CommonDialog1.FileName 'lien d'image
    Image1.Picture = LoadPicture(PicFile) 'Afficher l'image
Exit Sub
P:
MsgBox "Êﬁ⁄ Œÿ√ ›Ì  Õ„Ì· «·ÊÀÌﬁ… , «·—Ã«¡ «⁄«œ… «·„Õ«Ê·…", vbExclamation
PicFile = ""
   Image1.Picture = LoadPicture(PicFile) 'Afficher l'image
End Sub

Private Sub Command2_Click()
On Error Resume Next
PicFile = ""
Image1.Picture = LoadPicture(PicFile) 'Afficher l'image
fName = ""

End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim x$
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
Text7.Text = Trim(Text7.Text)
If Text1.Text = "" Or Text7.Text = "" Or Text3.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text1.BackColor = &H8080FF Then
Text1.SetFocus
ElseIf Text7.BackColor = &H8080FF Then
Text7.SetFocus
ElseIf Text3.BackColor = &H8080FF Then
Text3.SetFocus
End If
Exit Sub
End If
If Option7.Value = False And Option6.Value = False Then
MsgBox "ÌÃ»  ÕœÌœ «·Ã‰”", vbCritical
Exit Sub
End If
aab = Text4.Text
Text4.Text = StrConv(aab, vbProperCase)
Call cont
Do While Not et.EOF
If Combo1.Text = et!cla And Val(Text7.Text) = Val(et!num) And Label16.Caption <> et!aut Then
MsgBox "€Ì— „„ﬂ‰... ·ﬁœ  „ ÕÃ“  —ﬁ„ «·‰œ«¡ Â–« ”«»ﬁ«", vbCritical
Exit Sub
End If
et.MoveNext
Loop
Label13.Caption = Interface.SBB1.Panels(1).Text
If Label16.Caption <> "" Then
Call cont
Do While Not et.EOF
If Label16.Caption = et!aut Then
et!sri = BX1.Caption
et!nom = Text1.Text
et!nfr = Text4.Text
et!niv = Combo2.Text
et!cla = Combo1.Text
et!num = Text7.Text
If Option7.Value = True Then
et!sex = Option7.Caption
End If
If Option6.Value = True Then
et!sex = Option6.Caption
End If
et!nni = Text2.Text
et!dat = DT1.Value
et!tel = Text3.Text
et!img = App.Path & "\" & Label13.Caption & "\IMAGES\" & Combo1.Text & "\" & BX1.Caption & ".jpg"
et!ncl = Label17.Caption
et!eta = Label18.Caption
et.Update
If PicFile = "" Then
PicFile = App.Path & "\Pardefaut.jpg"
Image1.Picture = LoadPicture(PicFile)
End If
Call enregistrerphotto
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
et.MoveNext
Loop
End If
et.AddNew
et!sri = BX1.Caption
et!nom = Text1.Text
et!nfr = Text4.Text
et!niv = Combo2.Text
et!cla = Combo1.Text
et!num = Text7.Text
If Option7.Value = True Then
et!sex = Option7.Caption
End If
If Option6.Value = True Then
et!sex = Option6.Caption
End If
et!nni = Text2.Text
et!dat = DT1.Value
et!tel = Text3.Text
et!img = App.Path & "\" & Label13.Caption & "\IMAGES\" & Combo1.Text & "\" & BX1.Caption & ".jpg"
et!ncl = Label17.Caption
et!eta = Label18.Caption
et!act = "1"
et.Update
eb!sri = Val(eb!sri) + 1
eb.Update
sr.AddNew
sr!sri = BX1.Caption
sr!eta = " ·„Ì–"
sr.Update
If PicFile = "" Then
PicFile = App.Path & "\Pardefaut.jpg"
Image1.Picture = LoadPicture(PicFile)
End If
Call enregistrerphotto
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Command4_Click()
On Error Resume Next
Text2.Text = ""
Text3.Text = ""
Text7.Text = ""
Text4.Text = ""
Text1.Text = ""
Text1.SetFocus
PicFile = ""
Image1.Picture = LoadPicture(PicFile) 'Afficher l'image
fName = ""
DT1.Value = Date
Label16.Caption = ""
Label13.Caption = ""
Label10.Caption = "A"
'Call num_etudiants
grd1.Visible = False
grd1.Clear
grd1.Rows = 1
Call chargegrd1
grd1.Visible = True
xe = eb!sri
Call Series
BX1.Caption = xs
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False

End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim i As Double
Dim n As Double
Call cont
Do While Not ie.EOF
ie.Delete
ie.MoveNext
Loop
n = grd1.Rows
For i = 1 To n - 1
ie.AddNew
grd1.Row = i
grd1.Col = 1
ie!sri = grd1.Text
grd1.Col = 1
ie!ser = grd1.Text
grd1.Col = 2
ie!nom = "  " + grd1.Text
grd1.Col = 3
ie!cla = Combo1.Text
grd1.Col = 4
ie!num = grd1.Text
grd1.Col = 5
ie!sex = grd1.Text
grd1.Col = 6
ie!nni = grd1.Text
grd1.Col = 7
ie!dat = grd1.Text
grd1.Col = 8
ie!tel = grd1.Text
ie!sim = App.Path & "\Tete_Long2266.jpg"
ie.Update
Next i
Call cont
'basa = Interface.SBB1.Panels(1).Text
data.OpenCurrentDatabase App.Path & "\" & Interface.SBB1.Panels(1).Text & ".mdb", False, "7346804"
'data.DoCmd.ApplyFilter cla
data.DoCmd.Maximize
'data.DoCmd.OpenReport "Bulletin", acViewPreview, , "num =" & a, acWindowNormal, OpenArgs
data.DoCmd.OpenReport "List_Etudiants", acViewPreview, , , acWindowNormal, OpenArgs
'data.DoCmd.RunMacro "Opnreport", 2
'data.DCount
    'data.CloseCurrentDatabase
data.Visible = True
Set data = Nothing


End Sub

Private Sub Command6_Click()
On Error Resume Next
Call cont3
Do While Not et3.EOF
Call cont
'xe = eb!sri
'Call Series
BX1.Caption = et3!ser
Text1.Text = et3!nom
Combo2.Text = et3!niv
Combo1.Text = et3!cla
Text2.Text = et3!adr
DT1.Value = et3!dat
Text3.Text = et3!tel
If et3!sex = "–ﬂ—" Then
Option7.Value = True
Else
Option6.Value = True
End If
Command3_Click
et3.MoveNext
Loop
MsgBox "OK", vbInformation

End Sub

Private Sub Command7_Click()
On Error Resume Next
PicFile = ""
Image1.Picture = LoadPicture(PicFile)
mCapHwnd = capCreateCaptureWindow("AlsahemCapture", 0, 0, 0, 320, 240, Me.hWnd, 0)

DoEvents: SendMessage mCapHwnd, CONNECT, 0, 0
        SendMessage mCapHwnd, GetObject, 0, 0

        SendMessage mCapHwnd, COPY, 0, 0

        Image1.Picture = Clipboard.GetData
        PicFile = Clipboard.GetData
Clipboard.Clear
        
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0 '
SavePicture Image1.Picture, App.Path & "\image.jpg"
PicFile = App.Path & "\image.jpg"

End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim cl1 As String
Dim cl2 As String
Dim niv1 As String
Call cont
Call cont3
Do While Not et3.EOF
cl1 = et3!cla
cl.MoveFirst
Do While Not cl.EOF
cl2 = cl!cla
If cl1 = cl2 Then
niv1 = cl!niv
et3!niv = niv1
et3.Update
cl.MoveLast
End If
cl.MoveNext
Loop
et3.MoveNext
Loop
MsgBox "OK", vbInformation
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = 0
Me.Top = 0
Call cont
xe = eb!sri
Label18.Caption = eb!eta
Call Series
BX1.Caption = xs
Call chargegrd1_clear
End Sub
Private Sub grd1_Click()
On Error Resume Next
Dim sx As String
Dim i As Double
Dim j As Double
Dim au As Double
Dim a As Double
Dim b As Double
Dim x$
i = grd1.Row
j = grd1.Col
If i > 0 Then
If j = 9 Then
Label13.Caption = Interface.SBB1.Panels(1).Text
Label10.Caption = Combo1.Text
grd1.Row = i
grd1.Col = 0
Label16.Caption = grd1.Text
grd1.Col = 1
BX1.Caption = grd1.Text
grd1.Col = 2
Text1.Text = grd1.Text
grd1.Col = 3
Text4.Text = grd1.Text
grd1.Col = 4
Text7.Text = grd1.Text
grd1.Col = 5
sx = grd1.Text
If sx = "–ﬂ—" Then
Option7.Value = True
End If
If sx = "√‰ÀÏ" Then
Option6.Value = True
End If
grd1.Col = 6
Text2.Text = grd1.Text
grd1.Col = 7
DT1.Value = grd1.Text
grd1.Col = 8
Text3.Text = grd1.Text
PicFile = ""
Image1.Picture = LoadPicture(PicFile)
x$ = ""
x$ = dir$(App.Path & "\" & Label13.Caption & "\IMAGES\" & Combo1.Text & "\" & BX1.Caption & ".jpg")
If x$ <> "" Then
PicFile = App.Path & "\" & Label13.Caption & "\IMAGES\" & Combo1.Text & "\" & BX1.Caption & ".jpg"
Image1.Picture = LoadPicture(PicFile)
End If
End If
If j = 10 Then
grd1.Row = i
grd1.Col = 0
Label16.Caption = grd1.Text
grd1.Row = i
grd1.Col = 1
seri = grd1.Text
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–« «· ·„Ì–", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Call cont
Do While Not et.EOF
If Label16.Caption = et!aut Then
et!act = "0"
et.Update
'Call supression_series
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
et.MoveNext
Loop
Else
Label16.Caption = ""
End If
End If
End If

End Sub

Private Sub grd2_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
i = grd2.Row
j = grd2.Col
If i > 0 Then
grd2.Row = i
grd2.Col = 3
Text3.Text = grd2.Text
End If
End Sub



Private Sub Option1_Click()
On Error Resume Next
grd1.Col = 1
grd1.Sort = 1

End Sub

Private Sub Option2_Click()
On Error Resume Next
grd1.Col = 2
grd1.Sort = 1

End Sub

Private Sub Option3_Click()
On Error Resume Next
grd1.Col = 4
grd1.Sort = 1

End Sub

Private Sub Option4_Click()
On Error Resume Next
grd1.Col = 6
grd1.Sort = 1

End Sub

Private Sub Option5_Click()
On Error Resume Next
grd1.Col = 8
grd1.Sort = 1

End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Else
Text1.BackColor = &H8080FF
End If
ActivateKeyboardLayout 67175425, 167175425

End Sub

Private Sub Text1_Click()
On Error Resume Next
Text1_Change
End Sub

Private Sub Text3_Change()
On Error Resume Next
grd2.Clear
grd2.Rows = 1
grd2.Cols = 4
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1300
grd2.ColWidth(2) = 2400
grd2.ColWidth(3) = 1200
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.Row = 0
grd2.Col = 1
grd2.Text = "—.  ··ÊﬂÌ·"
grd2.Col = 2
grd2.Text = "«”„ «·ÊﬂÌ·"
grd2.Col = 3
grd2.Text = "«·Â« ›"
If Len(Text3.Text) > 0 Then
Text3.BackColor = &HC000&
Call recheche
Else
Text3.BackColor = &H8080FF
End If

End Sub

Private Sub Text3_Click()
On Error Resume Next
Text3_Change
End Sub

Private Sub Text3_GotFocus()
On Error Resume Next
Picture2.Visible = True

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii <> 8 Then
If KeyAscii = 46 Then
KeyAscii = 0
End If
If KeyAscii < 46 Or KeyAscii > 57 Or KeyAscii = 47 Then
KeyAscii = 0
End If
End If

End Sub


Private Sub Text3_LostFocus()
On Error Resume Next
Picture2.Visible = False
End Sub

Private Sub Text4_Change()
On Error Resume Next
ActivateKeyboardLayout 67896332, 67896332

End Sub

Private Sub Text4_Click()
On Error Resume Next
Text4_Change
End Sub

Private Sub Text5_Change()
On Error Resume Next
Call chargegrd1_clear
End Sub

Private Sub Text5_Click()
On Error Resume Next
Text5_Change
End Sub

Private Sub Text5_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If Text5.Text <> "" Then
If KeyCode = 13 Then
Call cont
Do While Not et.EOF
If et!sri = Text5.Text Or Val(et!sri) = Val(Text5.Text) Then
grd1.Visible = False
Label12.Caption = et!cla
Combo2.Text = et!niv
Combo1.Text = Label12.Caption
grd1.Visible = False
Call chargegrd1_sri
grd1.Visible = True
Exit Sub
End If
et.MoveNext
Loop
End If
End If

End Sub

Private Sub Text7_Change()
On Error Resume Next
If Len(Text7.Text) > 0 Then
Text7.BackColor = &HC000&
Else
Text7.BackColor = &H8080FF
End If

End Sub

Private Sub Text7_Click()
On Error Resume Next
Text7_Change
End Sub
Public Sub recheche()
On Error Resume Next
Dim i As Double
grd2.Clear
grd2.Visible = False
grd2.Rows = 1
grd2.Cols = 4
grd2.ColWidth(0) = 0
grd2.ColWidth(1) = 1300
grd2.ColWidth(2) = 2400
grd2.ColWidth(3) = 1200
grd2.ColAlignment(1) = 1
grd2.ColAlignment(2) = 1
grd2.ColAlignment(3) = 1
grd2.Row = 0
grd2.Col = 1
grd2.Text = "—.  ··ÊﬂÌ·"
grd2.Col = 2
grd2.Text = "«”„ «·ÊﬂÌ·"
grd2.Col = 3
grd2.Text = "«·Â« ›"
 ' **  **  **  ** chargemant des donnÈes **  **  **  **
Call cont
 i = 1
'***** recherch par tel
If Text3.Text <> "" Then
n = cr.RecordCount
grd2.Rows = n + 2
cr.Filter = "[tel]" & "Like '*" & Text3 & "*'" 'entre
Do While Not cr.EOF
grd2.Row = i
grd2.Col = 1
grd2.Text = cr!sri
grd2.Col = 2
grd2.Text = cr!nom
grd2.Col = 3
grd2.Text = cr!tel
i = i + 1
cr.MoveNext
Loop
grd2.Rows = i
End If
grd2.Visible = True
'****************************************
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
'MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation + arabic
Command4_Click
End If

End Sub
Private Sub chargegrd1()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As String
Dim j As Double
Dim i As Double
Dim P As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1400
grd1.ColWidth(2) = 2250
grd1.ColWidth(3) = 2250
grd1.ColWidth(4) = 700
grd1.ColWidth(5) = 700
grd1.ColWidth(6) = 1300
grd1.ColWidth(7) = 1200
grd1.ColWidth(8) = 1100
grd1.ColWidth(9) = 600
grd1.ColWidth(10) = 600
grd1.ColAlignment(0) = 1
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
grd1.Text = "«·—ﬁ„ «· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«”„ «· ·„Ì–"
grd1.Col = 3
grd1.Text = "Nom d'etudiant"
grd1.Col = 4
grd1.Text = "«·‰œ«¡"
grd1.Col = 5
grd1.Text = "«·Ã‰”"
grd1.Col = 6
grd1.Text = "«·—ﬁ„ «·Êÿ‰Ì"
grd1.Col = 7
grd1.Text = " «—ÌŒ «· ”ÃÌ·"
grd1.Col = 8
grd1.Text = "Â« › «·ÊﬂÌ·"
i = 1
b = 0
c = 0
Call cont
grd1.Rows = et.RecordCount + 3
Do While Not et.EOF
If Combo1.Text = et!cla Then
If c = 0 Then
b = b + 1
a = et!num
If a <> b Then
d = b
If b < 10 Then
d = "00" + d
ElseIf b < 100 And b >= 10 Then
d = "0" + d
Else
d = d
End If
Text7.Text = d
c = 1
End If
End If
If et!act = "1" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = et!aut
grd1.Col = 1
grd1.Text = et!sri
grd1.Col = 2
grd1.Text = et!nom
grd1.Col = 3
grd1.Text = et!nfr
grd1.Col = 4
grd1.Text = et!num
grd1.Col = 5
grd1.Text = et!sex
grd1.Col = 6
grd1.Text = et!nni
grd1.Col = 7
grd1.Text = et!dat
grd1.Col = 8
grd1.Text = et!tel
grd1.Col = 9
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 10
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
End If
End If
et.MoveNext
Loop
If c = 0 Then
b = b + 1
d = b
If b < 10 Then
d = "00" + d
ElseIf b < 100 And b >= 10 Then
d = "0" + d
Else
d = d
End If
Text7.Text = d
End If
grd1.Rows = i
grd1.Col = 4
grd1.Sort = 2
End Sub
Private Sub chargegrd1_clear()
On Error Resume Next
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1400
grd1.ColWidth(2) = 4500
grd1.ColWidth(3) = 0
grd1.ColWidth(4) = 700
grd1.ColWidth(5) = 700
grd1.ColWidth(6) = 1300
grd1.ColWidth(7) = 1200
grd1.ColWidth(8) = 1100
grd1.ColWidth(9) = 600
grd1.ColWidth(10) = 600
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 1
grd1.ColAlignment(9) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«·—ﬁ„ «· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«”„ «· ·„Ì–"
grd1.Col = 3
grd1.Text = "«·ﬁ”„"
grd1.Col = 4
grd1.Text = "«·‰œ«¡"
grd1.Col = 5
grd1.Text = "«·Ã‰”"
grd1.Col = 6
grd1.Text = "«·—ﬁ„ «·Êÿ‰Ì"
grd1.Col = 7
grd1.Text = " «—ÌŒ «· ”ÃÌ·"
grd1.Col = 8
grd1.Text = "Â« › «·ÊﬂÌ·"
End Sub
Private Sub enregistrerphotto()
On Error Resume Next
Dim Security As SECURITY_ATTRIBUTES
Dim x$
Dim y$
y$ = dir$(App.Path & "\" & Label13.Caption & "\IMAGES\" & Combo1.Text & "\" & BX1.Caption & ".jpg")
If y$ <> "" Then
Kill App.Path & "\" & Label13.Caption & "\IMAGES\" & Combo1.Text & "\" & BX1.Caption & ".jpg"
End If
x$ = ""
x$ = dir$(App.Path & "\" & Label13.Caption)
If x$ = "" Then
'Create a directory dossier images
Ret& = CreateDirectory(App.Path & "\" & Label13.Caption, Security)
End If
x$ = ""
x$ = dir$(App.Path & "\" & Label13.Caption & "\IMAGES\")
If x$ = "" Then
'Create a directory dossier images
Ret& = CreateDirectory(App.Path & "\" & Label13.Caption & "\IMAGES\", Security)
End If
x$ = dir$(App.Path & "\" & Label13.Caption & "\IMAGES\" & Combo1.Text)
If x$ = "" Then
'Create a directory dossier images
Ret& = CreateDirectory(App.Path & "\" & Label13.Caption & "\IMAGES\" & Combo1.Text, Security)
End If
SavePicture Image1.Picture, App.Path & "\" & Label13.Caption & "\IMAGES\" & Combo1.Text & "\" & BX1.Caption & ".jpg"
End Sub

Private Sub supression_series()
On Error Resume Next
Call cont
Do While Not sr.EOF
If seri = sr!sri Then
sr.Delete
If sr.RecordCount > 3 Then
sr.MoveLast
End If
End If
sr.MoveNext
Loop
End Sub
Private Sub chargegrd1_sri()
On Error Resume Next
Dim a As Double
Dim b As Double
Dim c As Double
Dim d As String
Dim j As Double
Dim i As Double
Dim P As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1400
grd1.ColWidth(2) = 2250
grd1.ColWidth(3) = 2250
grd1.ColWidth(4) = 700
grd1.ColWidth(5) = 700
grd1.ColWidth(6) = 1300
grd1.ColWidth(7) = 1200
grd1.ColWidth(8) = 1100
grd1.ColWidth(9) = 600
grd1.ColWidth(10) = 600
grd1.ColAlignment(0) = 1
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
grd1.Text = "«·—ﬁ„ «· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«”„ «· ·„Ì–"
grd1.Col = 3
grd1.Text = "Nom d'etudiant"
grd1.Col = 4
grd1.Text = "«·‰œ«¡"
grd1.Col = 5
grd1.Text = "«·Ã‰”"
grd1.Col = 6
grd1.Text = "«·—ﬁ„ «·Êÿ‰Ì"
grd1.Col = 7
grd1.Text = " «—ÌŒ «· ”ÃÌ·"
grd1.Col = 8
grd1.Text = "Â« › «·ÊﬂÌ·"
i = 1
b = 0
c = 0
Call cont
grd1.Rows = et.RecordCount + 3
Do While Not et.EOF
If et!sri = Text5.Text Or Val(et!sri) = Val(Text5.Text) Then
If et!act = "1" Then
grd1.Row = i
grd1.Col = 0
grd1.Text = et!aut
grd1.Col = 1
grd1.Text = et!sri
grd1.Col = 2
grd1.Text = et!nom
grd1.Col = 3
grd1.Text = et!nfr
grd1.Col = 4
grd1.Text = et!num
grd1.Col = 5
grd1.Text = et!sex
grd1.Col = 6
grd1.Text = et!nni
grd1.Col = 7
grd1.Text = et!dat
grd1.Col = 8
grd1.Text = et!tel
grd1.Col = 9
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 10
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
End If
End If
et.MoveNext
Loop
grd1.Rows = i
grd1.Col = 4
grd1.Sort = 2
End Sub

