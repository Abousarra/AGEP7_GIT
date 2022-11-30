VERSION 5.00
Object = "{8E515444-86DF-11D3-A630-444553540001}#1.0#0"; "barcodex.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Pointage_E 
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
   Begin VB.CommandButton Command13 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "€·ﬁ »«» «·Õ÷Ê— ·Â–« «· «—ÌŒ"
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
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "«—‘Ì› «·Õ÷Ê—"
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
      Left            =   10440
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   9000
      Width           =   1695
   End
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   3240
      ScaleHeight     =   3795
      ScaleWidth      =   5475
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton Command11 
         Caption         =   "Command11"
         Height          =   255
         Left            =   4080
         TabIndex        =   44
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Command10"
         Height          =   255
         Left            =   4080
         TabIndex        =   43
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         Height          =   375
         Left            =   3960
         TabIndex        =   42
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   4080
         TabIndex        =   41
         Top             =   1920
         Width           =   1335
      End
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   3000
         TabIndex        =   40
         Top             =   120
         Width           =   2175
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   1920
         TabIndex        =   39
         Top             =   1920
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DT1 
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
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
         Format          =   124977153
         CurrentDate     =   41924
      End
      Begin MSComCtl2.DTPicker DT3 
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
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
         Format          =   124977153
         CurrentDate     =   41924
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "Label12"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄—÷ » «—ÌŒ"
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
         Left            =   1320
         TabIndex        =   17
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "»Ì«‰«  «· ·„Ì–"
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
      Left            =   10440
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8640
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   240
      ScaleHeight     =   1545
      ScaleWidth      =   9345
      TabIndex        =   12
      Top             =   8040
      Width           =   9375
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "⁄—÷ Õ’’ Â–« «·ﬁ”„ ·Ã„Ì⁄ «· Ê«—ÌŒ"
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
         Left            =   3240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "⁄—÷ Õ÷Ê— Â–« «·—ﬁ„ ·Ã„Ì⁄ «· Ê«—ÌŒ"
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
         Left            =   120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
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
         Left            =   3240
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "⁄—÷ Õ’’ Ã„Ì⁄ «·√ﬁ”«„ ·Â–« «· «—ÌŒ"
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
         Left            =   3240
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "⁄—÷ Õ÷Ê— Ã„Ì⁄ «·√ﬁ”«„ ·Â–« «· «—ÌŒ"
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
         Left            =   120
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   2895
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H0000C000&
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
         ItemData        =   "Pointage_E.frx":0000
         Left            =   6360
         List            =   "Pointage_E.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "⁄—÷ Õ÷Ê— Â–« «·ﬁ”„ ·Â–« «· «—ÌŒ"
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
         Left            =   6360
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1080
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DT2 
         Height          =   300
         Left            =   6360
         TabIndex        =   18
         Top             =   120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
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
         Format          =   124977153
         CurrentDate     =   41924
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         Height          =   495
         Left            =   3120
         Top             =   480
         Width           =   6135
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FFFF&
         Height          =   1455
         Left            =   6240
         Top             =   0
         Width           =   3015
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00FF00FF&
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   9255
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         Height          =   495
         Left            =   0
         Top             =   960
         Width           =   6255
      End
      Begin VB.Label Label14 
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
         Left            =   4680
         TabIndex        =   35
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label13 
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
         Left            =   7920
         TabIndex        =   31
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         Left            =   8280
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.OptionButton Option6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "«·Õ’… «·”«œ”… (18:00) ‹"
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
      Left            =   10200
      TabIndex        =   5
      Top             =   3240
      Width           =   2415
   End
   Begin VB.OptionButton Option5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "«·Õ’… «·Œ«„”… (16:00) ‹"
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
      Left            =   10200
      TabIndex        =   4
      Top             =   2880
      Width           =   2415
   End
   Begin VB.OptionButton Option4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "«·Õ’… «·—«»⁄… (14:00) ‹ "
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
      Left            =   10320
      TabIndex        =   3
      Top             =   2520
      Width           =   2295
   End
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "«·Õ’… «·À«·À… (12:00)‹"
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
      Left            =   10440
      TabIndex        =   2
      Top             =   2160
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "«·Õ’… «·À«‰Ì… (10:00)‹"
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
      Left            =   10440
      TabIndex        =   1
      Top             =   1800
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "«·Õ’… «·√Ê·Ï (8:00)‹"
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
      Left            =   10560
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   " ”ÃÌ· «·‰œ«¡"
      Default         =   -1  'True
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
      Left            =   10560
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
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
      Left            =   9960
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4200
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   6615
      Left            =   240
      TabIndex        =   11
      Top             =   1200
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11668
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
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BARCODEXLib.BarcodeX BX1 
      Height          =   615
      Left            =   9960
      Top             =   7920
      Width           =   2655
      _Version        =   65536
      _ExtentX        =   4683
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
      Left            =   3000
      TabIndex        =   29
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„ €Ì»Ê‰"
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
      Left            =   4080
      TabIndex        =   28
      Top             =   840
      Width           =   975
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
      Left            =   5160
      TabIndex        =   27
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·Õ«÷—Ê‰"
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
      Top             =   840
      Width           =   975
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
      Left            =   7440
      TabIndex        =   25
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«·„”Ã·Ê‰"
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
      Left            =   8640
      TabIndex        =   24
      Top             =   840
      Width           =   975
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   2655
      Left            =   9960
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   7215
      Index           =   0
      Left            =   120
      Top             =   720
      Width           =   9615
   End
   Begin VB.Line Line1 
      X1              =   9840
      X2              =   12720
      Y1              =   3720
      Y2              =   3720
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
      TabIndex        =   10
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   8775
      Index           =   9
      Left            =   9840
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Õ÷Ê— «· ·«„Ì–"
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
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   12615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ÷Ê—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   10920
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Pointage_E"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Type FILETIME
dwLowDateTime As Long
dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
dwFileAttributes As Long
ftCreationTime As FILETIME
ftLastAccessTime As FILETIME
ftLastWriteTime As FILETIME
nFileSizeHigh As Long
nFileSizeLow As Long
dwReserved0 As Long
dwReserved1 As Long
cFileName As String * MAX_PATH
cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal _
lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal _
hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long


Private Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Private Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
Private Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Dim data As New Access.Application
Private Function NumFiles(sPath As String) As Long
On Error Resume Next
Dim f As WIN32_FIND_DATA
Dim hFile As Long
NumFiles = 0
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
sPath = sPath & "*.*"
hFile = FindFirstFile(sPath, f)
If hFile = INVALID_HANDLE_VALUE Then Exit Function
If (f.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then NumFiles = 1
Do While FindNextFile(hFile, f)
If (f.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then NumFiles = _
NumFiles + 1
Loop
FindClose (hFile)
End Function
Function StripPath(t$) As String
On Error Resume Next
On Error Resume Next
Dim x%, ct%
StripPath$ = t$
x% = InStr(t$, "\")
Do While x%
ct% = x%
x% = InStr(ct% + 1, t$, "\")
Loop
If ct% > 0 Then StripPath$ = Mid$(t$, ct% + 1)
End Function

Private Sub Check1_Click()
On Error Resume Next
grd1.Visible = False
Call chargegrd1
grd1.Visible = True
If Check1.Value = 1 Then
Picture1.Visible = False
Else
Picture1.Visible = True
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
Dim noms As String
Dim sers As String
Dim clas As String
Dim nums As String
Dim nivs As String
sers = BX1.Caption
Call cont
Do While Not et.EOF
If et!sri = sers Then
nivs = et!niv
clas = et!cla
nums = et!num
noms = et!nom
MsgBox "«·«”„: " + noms & vbCr & _
" «·„” ÊÏ: " + nivs & vbCr & _
" «·ﬁ”„: " + clas & vbCr & _
" —ﬁ„ «·‰œ«¡ " + nums, vbInformation + arabic
Exit Sub
End If
et.MoveNext
Loop
End Sub



Private Sub Command10_Click()
On Error Resume Next
Dim n1 As Double
Dim n2 As Double
Dim k1 As Double
Dim k2 As Double
Dim s As Double
n1 = 0
n2 = 0
s = 0
    Open App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Combo1.Text & ".txt" For Input As #1
    Open App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Text2.Text & ".txt" For Input As #2
    
    Do While Not EOF(1)
        s = s + 1
        Line Input #1, x
        If s > 2 And s Mod 2 <> 0 Then
        Label15.Caption = x
        vg = Mid$(Label15.Caption, 13, 1)
        k1 = vg
        n1 = n1 + k1
        End If
    Loop
    s = 0
    Do While Not EOF(2)
        s = s + 1
        Line Input #2, x
        If s > 2 And s Mod 2 <> 0 Then
        Label15.Caption = x
        vg = Mid$(Label15.Caption, 13, 1)
        k2 = vg
        n2 = n2 + k2
        End If
    Loop
     Close #1
     Close #2
     MsgBox n1          'nombre de cours classe
     MsgBox n2          'nombre de cours etudiant
     MsgBox (n1 - n2)   'nombre de cours absence
'Count:
 '       ss = ss + 1
  '      Line Input #1, x
   '     If ss > 2 And ss Mod 2 <> 0 Then
    '    Label15.Caption = x
     '   vg = Mid$(Label15.Caption, 13, 1)
      '  k = vg
      '  n = n + k
      '  End If
      '  If EOF(1) Then
            'MsgBox ss
       '     MsgBox n
        '    Close #1
         '   Exit Sub
        'Else
         '   GoTo Count:
        'End If
   ' Close

End Sub

Private Sub Command11_Click()
On Error Resume Next
Open App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Combo1.Text & ".txt" For Input As #1
Do While Not EOF(1)
myChar = Input(2, #1) 'one char a line
WholeWord = WholeWord & myChar
Loop
Close #1
MsgBox WholeWord
End Sub

Private Sub Command12_Click()
On Error Resume Next
DT2.Value = Date
For i = 1 To 500
DT2.Value = DT2.Value + 1
Text1.Text = "004522263"
Command7_Click
DT3.Value = DT2.Value - 1
Call supression_pointage
Next i
MsgBox "ok", vbInformation
End Sub

Private Sub Command13_Click()
On Error Resume Next
g = MsgBox("Â·  —Ìœ Õﬁ« €·ﬁ »«» «·Õ÷Ê— ·Â–« «·ÌÊ„ø", vbInformation + vbYesNo + arabic, "AGEP7")
If g = vbYes Then
Call cont
n = pe.RecordCount
If n > 0 Then
DT3.Value = pe!dat
Call supression_pointage
Call chargegrd1_clear
MsgBox " „ €·ﬁ »«» «·Õ÷Ê— ·Â–« «·ÌÊ„ »‰Ã«Õ", vbInformation
Else
MsgBox "·«ÌÊÃœ ”Ã· Õ÷Ê—", vbCritical
End If
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim x$
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & BX1.Caption & ".txt")
If x$ = "" Then
MsgBox "«·„·› «·„ÿ·Ê» €Ì— „ÊÃÊœ", vbExclamation
Exit Sub
End If
Shell "notepad.exe" & " " & App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & BX1.Caption & ".txt", vbNormalFocus
End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim d As String
d = NumFiles(App.Path & "\IMAGES\1AF")
MsgBox d
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim x$
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & DT2.Day & DT2.Month & DT2.Year & Combo1.Text & ".txt")
If x$ = "" Then
MsgBox "«·„·› «·„ÿ·Ê» €Ì— „ÊÃÊœ", vbExclamation
Exit Sub
End If
Shell "notepad.exe" & " " & App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & DT2.Day & DT2.Month & DT2.Year & Combo1.Text & ".txt", vbNormalFocus
End Sub

Private Sub Command5_Click()
On Error Resume Next
Dim x$
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\N" & DT2.Day & DT2.Month & DT2.Year & ".txt")
If x$ = "" Then
MsgBox "«·„·› «·„ÿ·Ê» €Ì— „ÊÃÊœ", vbExclamation
Exit Sub
End If
Shell "notepad.exe" & " " & App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\N" & DT2.Day & DT2.Month & DT2.Year & ".txt", vbNormalFocus

End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim x$
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\C" & DT2.Day & DT2.Month & DT2.Year & ".txt")
If x$ = "" Then
MsgBox "«·„·› «·„ÿ·Ê» €Ì— „ÊÃÊœ", vbExclamation
Exit Sub
End If
Shell "notepad.exe" & " " & App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\C" & DT2.Day & DT2.Month & DT2.Year & ".txt", vbNormalFocus
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim dat1 As Date
Dim dat2 As Date
Dim j As Double
Dim tx As String
Dim x$
Text1.Text = Trim(Text1.Text)
If Text1.Text = "" Then
MsgBox "«·—Ã«¡ «œŒ«· «·—ﬁ„ «· ”·”·Ì ·· ·„Ì–", vbCritical + arabic
Exit Sub
End If
dat1 = Date
Call cont
Do While Not pd.EOF
dat2 = pd!dat
tx = pd!dat
If dat1 = dat2 Then
MsgBox "⁄›Ê«... ·« Ì„ﬂ‰  ”ÃÌ· «·Õ÷Ê— ·Â–« «· «—ÌŒ " + tx + " ≈– ”»ﬁ √‰  „ €·ﬁ «·Õ÷Ê— ·Â", vbCritical
Text1.Text = ""
Text1.SetFocus
Exit Sub
End If
pd.MoveNext
Loop
'*** verif n s
vtx1 = Text1.Text
Call verif_n_serie
Text1.Text = vtx2
'*** end verif n s
PicFile = ""
BX1.Caption = ""
Image1.Picture = LoadPicture(PicFile)
j = 0
Call cont
Do While Not cl.EOF
tx = cl!cla
x$ = ""
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & tx & "\" & Text1.Text & ".jpg")
If x$ <> "" Then
Label12.Caption = tx
j = 1
PicFile = App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & tx & "\" & Text1.Text & ".jpg"
Image1.Picture = LoadPicture(PicFile)
BX1.Caption = Text1.Text
sndPlaySound App.Path & "\bsmOUI.wav", 1
cl.MoveLast
End If
cl.MoveNext
Loop
If j = 0 Then
sndPlaySound App.Path & "\bsmNON.wav", 1
'MsgBox "«·—ﬁ„ «· ”·”·Ì «·„œŒ· €Ì— „”Ã·  Õ Â √Ì  ·„Ì–", vbExclamation
Text1.Text = ""
Text1.SetFocus
Exit Sub
End If
Call cont
Do While Not pe.EOF
If Text1.Text = pe!sri Then
pe!dat = DT1.Value
If Option1.Value = True Then
pe!Pmt = Time$
End If
If Option2.Value = True Then
pe!pc1 = Time$
End If
If Option3.Value = True Then
pe!pc2 = Time$
End If
If Option4.Value = True Then
pe!pam = Time$
End If
If Option5.Value = True Then
pe!psr = Time$
End If
If Option6.Value = True Then
pe!dep = Time$
End If
pe.Update
grd1.Visible = False
Call chargegrd1
grd1.Visible = True
Text1.Text = ""
Text1.SetFocus
Exit Sub
End If
pe.MoveNext
Loop
pe.AddNew
pe!sri = Text1.Text
pe!dat = DT1.Value
pe!cla = Label12.Caption
pe!Pmt = ""
pe!pc1 = ""
pe!pc2 = ""
pe!pam = ""
pe!psr = ""
pe!dep = ""
If Option1.Value = True Then
pe!Pmt = Time$
End If
If Option2.Value = True Then
pe!pc1 = Time$
End If
If Option3.Value = True Then
pe!pc2 = Time$
End If
If Option4.Value = True Then
pe!pam = Time$
End If
If Option5.Value = True Then
pe!psr = Time$
End If
If Option6.Value = True Then
pe!dep = Time$
End If
pe.Update
grd1.Visible = False
Call chargegrd1
grd1.Visible = True
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim x$
'*** verif n s
vtx1 = Text2.Text
Call verif_n_serie
Text2.Text = vtx2
'*** end verif n s
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Text2.Text & ".txt")
If x$ = "" Then
MsgBox "«·„·› «·„ÿ·Ê» €Ì— „ÊÃÊœ", vbExclamation
Exit Sub
End If
Shell "notepad.exe" & " " & App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Text2.Text & ".txt", vbNormalFocus

End Sub

Private Sub Command9_Click()
On Error Resume Next
Dim x$
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Combo1.Text & ".txt")
If x$ = "" Then
MsgBox "«·„·› «·„ÿ·Ê» €Ì— „ÊÃÊœ", vbExclamation
Exit Sub
End If
Shell "notepad.exe" & " " & App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & Combo1.Text & ".txt", vbNormalFocus

End Sub

Private Sub DT1_Change()
On Error Resume Next
grd1.Visible = False
Call chargegrd1
grd1.Visible = True

End Sub

Private Sub DT1_Click()
On Error Resume Next
DT1_Change
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim n As Double
Dim dat1 As Date
Dim dat2 As Date
Me.Left = 0
Me.Top = 0
DT1.Value = Date
DT2.Value = Date
dat1 = DT1.Value
Call inscriptions
Call chargcombo1
Call cont
n = pe.RecordCount
If n > 0 Then
DT3.Value = pe!dat
dat2 = pe!dat
If dat1 <> dat2 Then
Call supression_pointage
Call chargegrd1_clear
Exit Sub
End If
Call chargegrd1
Else
Call chargegrd1_clear
End If
End Sub
Private Sub chargegrd1()
On Error Resume Next
Dim n As Double
Dim m As Double
Dim k As Double
Dim a As Double
Dim b As Double
Dim j As Double
Dim i As Double
Dim P As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Cols = 12
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1250
grd1.ColWidth(2) = 1200
grd1.ColWidth(3) = 1250
grd1.ColWidth(4) = 1250
grd1.ColWidth(5) = 1200
grd1.ColWidth(6) = 1200
grd1.ColWidth(7) = 1200
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 0
grd1.ColWidth(11) = 0
grd1.ColWidth(10) = 0
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 3
grd1.ColAlignment(10) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«·—ﬁ„ «· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«·’»«ÕÌ"
grd1.Col = 3
grd1.Text = "«·«” —«Õ… «·√Ê·Ï"
grd1.Col = 4
grd1.Text = "«·«” —«Õ… «·À«‰Ì…"
grd1.Col = 5
grd1.Text = "»⁄œ «·ŸÂÌ—…"
grd1.Col = 6
grd1.Text = "«·„”«∆Ì"
grd1.Col = 7
grd1.Text = "«·«‰’—«›"
i = 1
Call cont
grd1.Rows = pe.RecordCount + 3
Do While Not pe.EOF
grd1.Row = i
grd1.Col = 0
grd1.Text = pe!aut
grd1.Col = 1
grd1.Text = pe!sri
grd1.Col = 2
grd1.Text = pe!Pmt
grd1.Col = 3
grd1.Text = pe!pc1
grd1.Col = 4
grd1.Text = pe!pc2
grd1.Col = 5
grd1.Text = pe!pam
grd1.Col = 6
grd1.Text = pe!psr
grd1.Col = 7
grd1.Text = pe!dep
grd1.Col = 8
grd1.Col = 9
grd1.Text = ""
grd1.Col = 11
grd1.Text = pe!cla
i = i + 1
pe.MoveNext
Loop
grd1.Rows = i
grd1.Col = 11
grd1.Sort = 1
n = Label5.Caption
i = (i - 1)
m = (n - i)
Label8.Caption = i
Label10.Caption = m

End Sub

Private Sub grd1_Click()
On Error Resume Next
Dim i As Double
Dim j As Double
Dim a As Double
Dim s As String
i = grd1.Row
j = grd1.Col
If j = 3 Then
s = i
MsgBox s
End If
If j = 10 Then
grd1.Row = i
grd1.Col = 0
a = grd1.Text
grd1.Row = i
grd1.Col = 1
s = grd1.Text
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–›  Õ÷Ê— «· ·„Ì– ’«Õ» «·—ﬁ„ «· ”·”·Ì " + s, vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Call cont
Do While Not pe.EOF
If a = pe!aut Then
pe.Delete
grd1.Visible = False
Call chargegrd1
grd1.Visible = True
Exit Sub
End If
pe.MoveNext
Loop
End If
End If
End Sub

Private Sub Option1_Click()
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub Option2_Click()
On Error Resume Next
Text1.SetFocus

End Sub

Private Sub Option3_Click()
On Error Resume Next
Text1.SetFocus

End Sub

Private Sub Option4_Click()
On Error Resume Next
Text1.SetFocus

End Sub

Private Sub Option5_Click()
On Error Resume Next
Text1.SetFocus

End Sub

Private Sub Option6_Click()
On Error Resume Next
Text1.SetFocus

End Sub

Private Sub Text1_Change()
On Error Resume Next
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
Else
Text1.BackColor = &H8080FF
End If

End Sub

Private Sub Text1_Click()
On Error Resume Next
Text1_Change
End Sub
Private Sub chargegrd1_clear()
On Error Resume Next
grd1.Clear
grd1.Cols = 11
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 1250
grd1.ColWidth(2) = 1200
grd1.ColWidth(3) = 1250
grd1.ColWidth(4) = 1250
grd1.ColWidth(5) = 1200
grd1.ColWidth(6) = 1200
grd1.ColWidth(7) = 1200
grd1.ColWidth(8) = 0
grd1.ColWidth(9) = 0
grd1.ColWidth(10) = 0
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.ColAlignment(4) = 1
grd1.ColAlignment(5) = 1
grd1.ColAlignment(6) = 1
grd1.ColAlignment(7) = 1
grd1.ColAlignment(8) = 3
grd1.ColAlignment(10) = 3
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«·—ﬁ„ «· ”·”·Ì"
grd1.Col = 2
grd1.Text = "«·’»«ÕÌ"
grd1.Col = 3
grd1.Text = "«·«” —«Õ… «·√Ê·Ï"
grd1.Col = 4
grd1.Text = "«·«” —«Õ… «·À«‰Ì…"
grd1.Col = 5
grd1.Text = "»⁄œ «·ŸÂÌ—…"
grd1.Col = 6
grd1.Text = "«·„”«∆Ì"
grd1.Col = 7
grd1.Text = "«·«‰’—«›"

End Sub
Private Sub supression_pointageo()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim tx0 As String
Dim tx1 As String
Dim tx2 As String
Dim tx3 As String
Dim tx4 As String
Dim tx5 As String
Dim tx6 As String
Dim tx7 As String
Dim tx8 As String
Dim tx9 As String
Dim j As Double
Dim k As Double
Dim Security As SECURITY_ATTRIBUTES
Dim x$
x$ = ""
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\")
If x$ = "" Then
'Create a directory dossier images
Ret& = CreateDirectory(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES", Security)
Ret& = CreateDirectory(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\DATES", Security)
Ret& = CreateDirectory(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\SERIES", Security)
Ret& = CreateDirectory(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\CLASSES", Security)
End If
grd1.Visible = False
Call chargegrd1
grd1.Visible = True
tx0 = DT3.Value
n = grd1.Rows
j = 0
Open (App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\DATES\" & DT3.Day & DT3.Month & DT3.Year & ".txt") For Append As #1
Print #1, " SERIALE " + "||" + "MATIN   " + "||" + "PAUSE1  " + "||" + "PAUSE2  " + "||" + "MIDI    " + "||" + "SOIRE   " + "||" + "DEPART  "
Print #1, "---------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
For i = 1 To n - 1
grd1.Row = i
grd1.Col = 1
tx1 = grd1.Text
grd1.Col = 2
tx2 = grd1.Text
grd1.Col = 3
tx3 = grd1.Text
grd1.Col = 4
tx4 = grd1.Text
grd1.Col = 5
tx5 = grd1.Text
grd1.Col = 6
tx6 = grd1.Text
grd1.Col = 7
tx7 = grd1.Text
grd1.Col = 11
tx8 = grd1.Text
If i > 1 And tx8 <> tx9 Then
k = 0
k = NumFiles(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & tx9)
pl.AddNew
pl!dat = DT3.Value
pl!cla = tx9
pl!nbi = k
pl!nbp = j
pl!nba = (k - j)
pl.Update
j = 0
End If
grd1.Row = i
grd1.Col = 11
tx9 = grd1.Text
j = j + 1
If tx1 = "" Then
tx1 = "........"
End If
If tx2 = "" Then
tx2 = "........"
End If
If tx3 = "" Then
tx3 = "........"
End If
If tx4 = "" Then
tx4 = "........"
End If
If tx5 = "" Then
tx5 = "........"
End If
If tx6 = "" Then
tx6 = "........"
End If
If tx7 = "" Then
tx7 = "........"
End If
Print #1, tx1 + "||" + tx2 + "||" + tx3 + "||" + tx4 + "||" + tx5 + "||" + tx6 + "||" + tx7
Print #1, "---------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\SERIES\" & tx1 & ".txt")
If x$ = "" Then
Open (App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\SERIES\" & tx1 & ".txt") For Append As #2
Print #2, "   DATE   " + "||" + "MATIN   " + "||" + "PAUSE1  " + "||" + "PAUSE2  " + "||" + "MIDI    " + "||" + "SOIRE   " + "||" + "DEPART  "
Print #2, "---------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
Close #2
End If
Open (App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\SERIES\" & tx1 & ".txt") For Append As #2
Print #2, tx0 + "||" + tx2 + "||" + tx3 + "||" + tx4 + "||" + tx5 + "||" + tx6 + "||" + tx7
Print #2, "---------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
Close #2
Next i
Close #1
If i > 0 Then
k = 0
k = NumFiles(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & tx9)
pl.AddNew
pl!dat = DT3.Value
pl!cla = tx9
pl!nbi = k
pl!nbp = j
pl!nba = (k - j)
pl.Update
End If
Call cont
Do While Not pe.EOF
pe.Delete
pe.MoveNext
Loop
End Sub
Private Sub supression_pointage()
On Error Resume Next
Dim i As Double
Dim n As Double
Dim dt0 As String
Dim sr0 As String
Dim cl0 As String
Dim cl1 As String
Dim cr1 As String
Dim cr2 As String
Dim cr3 As String
Dim cr4 As String
Dim cr5 As String
Dim cr6 As String
Dim ndc As Double
Dim snd As String
Dim nsc As Double
Dim sns As String
Dim c1 As Double
Dim c2 As Double
Dim c3 As Double
Dim c4 As Double
Dim c5 As Double
Dim c6 As Double
Dim sc1 As Double
Dim sc2 As Double
Dim sc3 As Double
Dim sc4 As Double
Dim sc5 As Double
Dim sc6 As Double
Dim sc7 As Double
Dim j As Double
Dim nb As Double
Dim sb As String
Dim sc As String
Dim c7 As String
Dim c8 As String
Dim c9 As String
Dim c10 As String
Dim c11 As String
Dim c12 As String
Dim Security As SECURITY_ATTRIBUTES
Dim x$
x$ = ""
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text)
If x$ = "" Then
'Create a directory dossier POINTAGES
Ret& = CreateDirectory(App.Path & "\" & Interface.SBB1.Panels(1).Text, Security)
End If
x$ = ""
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\")
If x$ = "" Then
'Create a directory dossier POINTAGES
Ret& = CreateDirectory(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES", Security)
End If
grd1.Visible = False
Call chargegrd1
grd1.Visible = True
dt0 = DT3.Value
n = grd1.Rows
c1 = 0
c2 = 0
c3 = 0
c4 = 0
c5 = 0
c6 = 0
sc1 = 0
sc2 = 0
sc3 = 0
sc4 = 0
sc5 = 0
sc6 = 0
nb = 0
j = 0
For i = 1 To n - 1
grd1.Row = i
grd1.Col = 1
sr0 = grd1.Text
grd1.Col = 2
cr1 = grd1.Text
grd1.Col = 3
cr2 = grd1.Text
grd1.Col = 4
cr3 = grd1.Text
grd1.Col = 5
cr4 = grd1.Text
grd1.Col = 6
cr5 = grd1.Text
grd1.Col = 7
cr6 = grd1.Text
grd1.Col = 11
cl0 = grd1.Text
'1_1**** fichiers un classe ‡ toutes dates
If i > 1 Then
If cl0 <> cl1 Then
sb = nb
sc7 = sc1 + sc2 + sc3 + sc4 + sc5 + sc6
sc = sc7
c7 = c1
c8 = c2
c9 = c3
c10 = c4
c11 = c5
c12 = c6
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & cl1 & ".txt")
Open (App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & cl1 & ".txt") For Append As #1
If x$ = "" Then
Print #1, "   DATE   " + "||" + "P" + "||"
Print #1, "----------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
End If
Print #1, dt0 + "||" + sb + "||  " + c7 + "  ||  " + c8 + "  ||  " + c9 + "  ||  " + c10 + "  ||  " + c11 + "  ||  " + c12
Print #1, "----------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
Close #1
End If
End If
'end 1_1**** fichiers un classe ‡ toutes dates
'4_1**** fichiers un date ‡ toutes classes
If i > 1 Then
If cl0 <> cl1 Then
sc7 = sc1 + sc2 + sc3 + sc4 + sc5 + sc6
sb = nb
c7 = c1
c8 = c2
c9 = c3
c10 = c4
c11 = c5
c12 = c6
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\C" & DT3.Day & DT3.Month & DT3.Year & ".txt")
Open (App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\C" & DT3.Day & DT3.Month & DT3.Year & ".txt") For Append As #4
If x$ = "" Then
Print #4, "  CLASSE  " + "||" + "P" + "||"
Print #4, "----------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
End If
Print #4, cl1 + "||" + sb + "||  " + c7 + "  ||  " + c8 + "  ||  " + c9 + "  ||  " + c10 + "  ||  " + c11 + "  ||  " + c12
Print #4, "----------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
Close #4
c1 = 0
c2 = 0
c3 = 0
c4 = 0
c5 = 0
c6 = 0
sc1 = 0
sc2 = 0
sc3 = 0
sc4 = 0
sc5 = 0
sc6 = 0
nb = 0
End If
End If
'end 4_1**** fichiers un classe ‡ toutes dates
ndc = 0
nsc = 0
If cr1 = "" Then
cr1 = "........"
ndc = ndc + 1
nsc = nsc + 1
Else
c1 = c1 + 1
sc1 = 1
j = j + 1
End If
If cr2 = "" Then
cr2 = "........"
nsc = nsc + 1
ndc = ndc + 1
Else
c2 = c2 + 1
sc2 = 1
j = j + 1
End If
If cr3 = "" Then
cr3 = "........"
ndc = ndc + 1
nsc = nsc + 1
Else
c3 = c3 + 1
sc3 = 1
j = j + 1
End If
If cr4 = "" Then
cr4 = "........"
ndc = ndc + 1
nsc = nsc + 1
Else
c4 = c4 + 1
sc4 = 1
j = j + 1
End If
If cr5 = "" Then
cr5 = "........"
ndc = ndc + 1
nsc = nsc + 1
Else
c5 = c5 + 1
sc5 = 1
j = j + 1
End If
If cr6 = "" Then
cr6 = "........"
ndc = ndc + 1
nsc = nsc + 1
Else
c6 = c6 + 1
sc6 = 1
j = j + 1
End If
If nb < j Then
nb = j
End If
j = 0
'2 **** fichiers un date ‡ toutes series
ndc = 6 - ndc
snd = ndc
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\N" & DT3.Day & DT3.Month & DT3.Year & ".txt")
Open (App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\N" & DT3.Day & DT3.Month & DT3.Year & ".txt") For Append As #2
If x$ = "" Then
'**** tete fichier date ‡ toutes series
Print #2, "N∞ SÈrie " + "||" + "P" + "||" + "1erCours" + "||" + "2emCours" + "||" + "3emCours" + "||" + "4emCours" + "||" + "5emCours" + "||" + "6emCours"
Print #2, "---------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
'**** data fichier date ‡ toutes series
Print #2, sr0 + "||" + snd + "||" + cr1 + "||" + cr2 + "||" + cr3 + "||" + cr4 + "||" + cr5 + "||" + cr6
Print #2, "---------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
Else
'**** data fichier date ‡ toutes series
Print #2, sr0 + "||" + snd + "||" + cr1 + "||" + cr2 + "||" + cr3 + "||" + cr4 + "||" + cr5 + "||" + cr6
Print #2, "---------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
End If
Close #2
'end fichiers un date ‡ toutes series
'3 **** fichiers un serie ‡ toutes dates
nsc = 6 - nsc
sns = nsc
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & sr0 & ".txt")
Open (App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & sr0 & ".txt") For Append As #3
If x$ = "" Then
Print #3, "   DATE   " + "||" + "P" + "||" + "1erCours" + "||" + "2emCours" + "||" + "3emCours" + "||" + "4emCours" + "||" + "5emCours" + "||" + "6emCours"
Print #3, "----------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
End If
Print #3, dt0 + "||" + sns + "||" + cr1 + "||" + cr2 + "||" + cr3 + "||" + cr4 + "||" + cr5 + "||" + cr6
Print #3, "----------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
Close #3
'end fichiers un serie ‡ toutes dates
grd1.Row = i
grd1.Col = 11
cl1 = grd1.Text
'5**** fichiers un classe ‡ un dates
'ndc = 6 - ndc
snd = ndc
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & DT3.Day & DT3.Month & DT3.Year & cl0 & ".txt")
Open (App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & DT3.Day & DT3.Month & DT3.Year & cl0 & ".txt") For Append As #5
If x$ = "" Then
'**** tete fichier un classe ‡ un dates
Print #5, "N∞ SÈrie " + "||" + "P" + "||" + "1erCours" + "||" + "2emCours" + "||" + "3emCours" + "||" + "4emCours" + "||" + "5emCours" + "||" + "6emCours"
Print #5, "---------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
'**** data fichier un classe ‡ un dates
Print #5, sr0 + "||" + snd + "||" + cr1 + "||" + cr2 + "||" + cr3 + "||" + cr4 + "||" + cr5 + "||" + cr6
Print #5, "---------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
Else
'**** data fichier un classe ‡ un dates
Print #5, sr0 + "||" + snd + "||" + cr1 + "||" + cr2 + "||" + cr3 + "||" + cr4 + "||" + cr5 + "||" + cr6
Print #5, "---------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
End If
Close #5
'end fichiers un classe ‡ un dates
Next i
'1_2**** fichiers un classe ‡ toutes dates
sb = nb
sc7 = sc1 + sc2 + sc3 + sc4 + sc5 + sc6
sc = sc7
c7 = c1
c8 = c2
c9 = c3
c10 = c4
c11 = c5
c12 = c6
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & cl1 & ".txt")
Open (App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\" & cl1 & ".txt") For Append As #1
If x$ = "" Then
Print #1, "   DATE   " + "||" + "P" + "||"
Print #1, "----------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
End If
Print #1, dt0 + "||" + sb + "||  " + c7 + "  ||  " + c8 + "  ||  " + c9 + "  ||  " + c10 + "  ||  " + c11 + "  ||  " + c12
Print #1, "----------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
Close #1
'end 2**** fichiers un classe ‡ toutes dates
'4_2**** fichiers un date ‡ toutes classes
sc7 = sc1 + sc2 + sc3 + sc4 + sc5 + sc6
sb = nb
c7 = c1
c8 = c2
c9 = c3
c10 = c4
c11 = c5
c12 = c6
x$ = dir$(App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\C" & DT3.Day & DT3.Month & DT3.Year & ".txt")
Open (App.Path & "\" & Interface.SBB1.Panels(1).Text & "\POINTAGES\C" & DT3.Day & DT3.Month & DT3.Year & ".txt") For Append As #4
If x$ = "" Then
Print #4, "  CLASSE  " + "||" + "P" + "||"
Print #4, "----------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
End If
Print #4, cl1 + "||" + sb + "||  " + c7 + "  ||  " + c8 + "  ||  " + c9 + "  ||  " + c10 + "  ||  " + c11 + "  ||  " + c12
Print #4, "----------" + "--" + "-" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------" + "--" + "--------"
Close #4
'end 2**** fichiers un date ‡ toutes classes
Call cont
Do While Not pe.EOF
pe.Delete
pe.MoveNext
Loop
pd.AddNew
pd!dat = DT3.Value
pd.Update
End Sub
Private Sub chargcombo1()
On Error Resume Next
Combo1.Clear
Call cont
Do While Not cl.EOF
Combo1.AddItem cl!cla
cl.MoveNext
Loop
End Sub

Private Sub inscriptions()
On Error Resume Next
Dim tx As String
Dim n As Double
Dim i As Double
Dim k As Double
n = 0
Dir1.Path = App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\"
For i = 0 To Dir1.ListCount - 1
tx = StripPath(Dir1.List(i))
File1.Path = App.Path & "\" & Interface.SBB1.Panels(1).Text & "\IMAGES\" & tx
k = File1.ListCount
n = n + k
Next i
Label5.Caption = n
End Sub
