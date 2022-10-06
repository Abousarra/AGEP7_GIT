VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Utilisateurs 
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
   Begin VB.CheckBox tjr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "—ﬂ‰ «·Êﬂ·«¡"
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
      Left            =   9000
      TabIndex        =   63
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CheckBox tca 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "«—‘Ì› «·’‰œÊﬁ"
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
      Left            =   6600
      TabIndex        =   60
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CheckBox tbn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Õ”«» «·»‰ﬂ"
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
      Left            =   6840
      TabIndex        =   59
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CheckBox cbn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "”Ã· «·»‰ﬂ"
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
      Left            =   9120
      TabIndex        =   58
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "≈€·«ﬁ"
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
      Left            =   120
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "ﬂ· «·’·«ÕÌ« "
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
      Left            =   11160
      TabIndex        =   50
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CheckBox spn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "«·√—‘›…"
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
      Left            =   11400
      TabIndex        =   40
      Top             =   7680
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   1200
      ScaleHeight     =   4635
      ScaleWidth      =   3315
      TabIndex        =   38
      Top             =   3480
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CheckBox arc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "≈œ«—… «·≈—‘Ì›"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   65
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CheckBox buk 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "«·‰”Œ «·«Õ Ì«ÿÌ"
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
         Left            =   1560
         TabIndex        =   64
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox afn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "› Õ ”‰… ÃœÌœ… "
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
         Left            =   840
         TabIndex        =   62
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CheckBox rpn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "«” —Ã«⁄ »Ì«‰« "
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
         Left            =   720
         TabIndex        =   61
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CheckBox tcn 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "«·„—ﬂ“ «·„«·Ì "
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
         Left            =   960
         TabIndex        =   57
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CheckBox tli 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "ﬁ«∆„… «·œŒ·"
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
         Left            =   1080
         TabIndex        =   56
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CheckBox tbu 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         Caption         =   "„Ì“«‰ «·„—«Ã⁄…"
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
         Left            =   840
         TabIndex        =   55
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   0
         Top             =   0
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CheckBox bil 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "»ÿ«ﬁ«  «·œŒÊ·"
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
      Left            =   10800
      TabIndex        =   37
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CheckBox trc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Õ”«» «·’‰œÊﬁ"
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
      Left            =   6600
      TabIndex        =   36
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CheckBox tpr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Õ”«» «·‘—ﬂ«¡"
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
      Left            =   6600
      TabIndex        =   35
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CheckBox tfn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Õ”«» «·„ÊŸ›Ì‰"
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
      Left            =   6480
      TabIndex        =   34
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CheckBox tpf 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Õ”«» «·√”« –…"
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
      Left            =   6600
      TabIndex        =   33
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CheckBox tet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Õ”«» «· ·«„Ì–"
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
      Left            =   6720
      TabIndex        =   32
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CheckBox tdp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Õ”«» «·‰›ﬁ« "
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
      Left            =   6720
      TabIndex        =   31
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CheckBox tcl 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Õ”«» «·√ﬁ”«„"
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
      Left            =   6720
      TabIndex        =   30
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CheckBox com 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "≈œ«—… «·„Õ«”»…"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   29
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CheckBox cpr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "”Ã· «·‘—ﬂ«¡"
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
      Left            =   8880
      TabIndex        =   28
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CheckBox cfn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "”Ã· «·„ÊŸ›Ì‰"
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
      Left            =   8760
      TabIndex        =   27
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CheckBox cpf 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "”Ã· «·√”« –…"
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
      Left            =   8880
      TabIndex        =   26
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CheckBox cet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "”Ã· «· ·«„Ì–"
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
      Left            =   8880
      TabIndex        =   25
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CheckBox cdp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "”Ã· «·‰›ﬁ« "
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
      Left            =   8880
      TabIndex        =   24
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CheckBox cca 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "”Ã· «·’‰œÊﬁ"
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
      Left            =   8760
      TabIndex        =   23
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CheckBox cai 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "≈œ«—… «·’‰œÊﬁ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9000
      TabIndex        =   22
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CheckBox ppr 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Õ÷Ê— «·√”« –…"
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
      Left            =   8760
      TabIndex        =   21
      Top             =   4440
      Width           =   1455
   End
   Begin VB.CheckBox rch 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "»ÕÀ ⁄«„"
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
      Left            =   11280
      TabIndex        =   20
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CheckBox pet 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Õ÷Ê— «· ·«„Ì–"
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
      Left            =   8760
      TabIndex        =   19
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CheckBox note 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "«·‰ «∆Ã Ê«·ﬂ‘Ê›"
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
      Left            =   6600
      TabIndex        =   18
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CheckBox cnt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "≈œ«—… «·—ﬁ«»…"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   9120
      TabIndex        =   17
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CheckBox mat 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "«·„Ê«œ Ê«·÷Ê«—»"
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
      Left            =   6480
      TabIndex        =   16
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CheckBox emp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "Ãœ«Ê· «·“„‰"
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
      Left            =   8880
      TabIndex        =   15
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CheckBox les 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "≈œ«—… «·œ—Ê”"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6840
      TabIndex        =   14
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CheckBox drp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "«·≈œ«—… «·⁄«„…"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   11280
      TabIndex        =   13
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CheckBox etu 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "»Ì«‰«  «· ·«„Ì–"
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
      Left            =   10920
      TabIndex        =   12
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CheckBox agn 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "»Ì«‰«  «·Êﬂ·«¡"
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
      Left            =   10920
      TabIndex        =   11
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CheckBox cla 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "»Ì«‰«  «·√ﬁ”«„"
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
      Left            =   10920
      TabIndex        =   10
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CheckBox prf 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "»Ì«‰«  «·√”« –…"
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
      Left            =   10920
      TabIndex        =   9
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CheckBox fnc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "»Ì«‰«  «·„ÊŸ›Ì‰"
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
      Left            =   10800
      TabIndex        =   8
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CheckBox prt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "»Ì«‰«  «·‘—ﬂ«¡"
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
      Left            =   10920
      TabIndex        =   7
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CheckBox uti 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "»Ì«‰«  «·„” Œœ„Ì‰"
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
      Left            =   10680
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CheckBox dir 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      Caption         =   "»Ì«‰«  «·„ƒ””…"
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
      Left            =   10800
      TabIndex        =   5
      Top             =   4080
      Width           =   1575
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
      Left            =   6480
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   960
      Width           =   4935
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
      IMEMode         =   3  'DISABLE
      Left            =   9720
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text3 
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
      IMEMode         =   3  'DISABLE
      Left            =   6480
      PasswordChar    =   "*"
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Left            =   9600
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
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
      Left            =   6480
      MaskColor       =   &H00FF0000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      UseMaskColor    =   -1  'True
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid grd1 
      Height          =   7335
      Left            =   240
      TabIndex        =   41
      Top             =   1320
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   12938
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
      Left            =   6480
      TabIndex        =   42
      Top             =   1680
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Index           =   2
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ„« ·« Ì„ﬂ‰ Õ–›Â »√Ì Õ«· „‰ «·√ÕÊ«·"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   53
      Top             =   9120
      Width           =   4935
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”ÊÏ ﬂ·„… «·”— ·Â ›ﬁÿ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   8760
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "admin "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   51
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„·«ÕŸ… :·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰«  «·„” Œœ„ "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   49
      Top             =   8760
      Width           =   3135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬁ«∆„… «·„” Œœ„Ì‰"
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
      TabIndex        =   48
      Top             =   840
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      Height          =   8295
      Index           =   10
      Left            =   120
      Top             =   1200
      Width           =   6015
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Index           =   0
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   6495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "’·«ÕÌ«  «·„” Œœ„"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   47
      Top             =   2640
      Width           =   6495
   End
   Begin VB.Shape Shape1 
      Height          =   3735
      Index           =   6
      Left            =   6240
      Shape           =   4  'Rounded Rectangle
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   3015
      Index           =   5
      Left            =   8400
      Shape           =   4  'Rounded Rectangle
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Index           =   3
      Left            =   8400
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   4455
      Index           =   1
      Left            =   10560
      Shape           =   4  'Rounded Rectangle
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«”„ «·„” Œœ„"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   11400
      TabIndex        =   46
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ·„… «·”—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   11400
      TabIndex        =   45
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "≈⁄«œ… ﬂ·„… «·”—"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   8040
      TabIndex        =   44
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰«  «·„” Œœ„Ì‰"
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
      TabIndex        =   43
      Top             =   120
      Width           =   12615
   End
End
Attribute VB_Name = "Utilisateurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Double

Private Sub arc_Click()
If arc.Value = 0 Then
'afn.Value = 0
 'spn.Value = 0
 'rpn.Value = 0
 'buk.Value = 0
Else
'afn.Value = 1
 'spn.Value = 1
' rpn.Value = 1
 'buk.Value = 1
End If
End Sub

Private Sub cai_Click()
If cai.Value = 0 Then
cpr.Value = 0
 cpf.Value = 0
 cfn.Value = 0
 cet.Value = 0
 cdp.Value = 0
 cbn.Value = 0
 cca.Value = 0
Else
cpr.Value = 1
 cpf.Value = 1
 cfn.Value = 1
 cet.Value = 1
 cdp.Value = 1
 cbn.Value = 1
 cca.Value = 1
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
drp.Value = 0
 les.Value = 0
 cnt.Value = 0
 cai.Value = 0
 com.Value = 0
 'arc.Value = 0
Else
drp.Value = 1
 les.Value = 1
 cnt.Value = 1
 cai.Value = 1
 com.Value = 1
 'arc.Value = 1

End If
End Sub

Private Sub cnt_Click()
If cnt.Value = 0 Then
pet.Value = 0
 ppr.Value = 0
 emp.Value = 0
tjr.Value = 0
 Else
pet.Value = 1
 ppr.Value = 1
 emp.Value = 1
tjr.Value = 1
 End If
End Sub

Private Sub com_Click()
If com.Value = 0 Then
tpr.Value = 0
 tfn.Value = 0
 tpf.Value = 0
 tet.Value = 0
 tdp.Value = 0
 tbn.Value = 0
 tcl.Value = 0
 trc.Value = 0
 tca.Value = 0
 'tjr.Value = 0
 'tbu.Value = 0
 'tli.Value = 0
 'tcn.Value = 0
Else
tpr.Value = 1
 tfn.Value = 1
 tpf.Value = 1
 tet.Value = 1
 tdp.Value = 1
 tbn.Value = 1
 tcl.Value = 1
 trc.Value = 1
 tca.Value = 1
 'tjr.Value = 1
 'tbu.Value = 1
 'tli.Value = 1
 'tcn.Value = 1

End If
End Sub

Private Sub Command1_Click()
Text1.Text = Trim(Text1.Text)
Text2.Text = Trim(Text2.Text)
Text3.Text = Trim(Text3.Text)
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox "«·—Ã«¡ „·¡ Ã„Ì⁄ «·ÕﬁÊ· «·„·Ê‰… »«··Ê‰ «·√Õ„—", vbCritical + arabic
If Text1.BackColor = &H8080FF Then
Text1.SetFocus
ElseIf Text2.BackColor = &H8080FF Then
Text2.SetFocus
ElseIf Text3.BackColor = &H8080FF Then
Text3.SetFocus
End If
Exit Sub
End If
If Text2.Text <> Text3.Text Then
MsgBox "ﬂ·„ « «·”— €Ì— „ ÿ«»ﬁ Ì‰", vbCritical + arabic
Text3.Text = ""
Text3.SetFocus
Exit Sub
End If
Call verifier_Checks
If x = 0 Then
MsgBox "·«Ì„ﬂ‰ ≈÷«›… Â–« «·„” Œœ„ ·√‰Â ·« Ì„·ﬂ √Ì ’·«ÕÌ…", vbExclamation
Exit Sub
End If
Call cont
Do While Not ut.EOF
If ut!aut <> Label7.Caption And Text1.Text = ut!nom Then
MsgBox "·ﬁœ  „ ÕÃ“ Â–« «·«”„ ”«»ﬁ«", vbCritical
Exit Sub
End If
ut.MoveNext
Loop
If Label7.Caption <> "" Then
Call cont
Do While Not ut.EOF
If ut!aut = Label7.Caption Then
If ut!adm = "1" Then
MsgBox " Â–« «·„” Œœ„ ·« Ì„ﬂ‰  ⁄œÌ· »Ì«‰« Â ”ÊÏ ﬂ·„… «·”— ›ﬁÿ", vbExclamation
ut!mot = Text2.Text
ut.Update
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
ut!mot = Text2.Text
ut!nom = Text1.Text
ut.Update
Call Ses_Options
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
ut.MoveNext
Loop
End If
ut.AddNew
ut!mot = Text2.Text
ut!nom = Text1.Text
ut!adm = "0"
ut.Update
Call Ses_Options
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus
Text2.Text = ""
Text3.Text = ""
Label7.Caption = ""
Check1.Value = 1
Check1.Value = 0
grd1.Visible = False
grd1.Clear
grd1.Rows = 1
Call chargegrd1
grd1.Visible = True
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = False

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub drp_Click()
If drp.Value = 0 Then
dir.Value = 0
 uti.Value = 0
 prt.Value = 0
 fnc.Value = 0
 prf.Value = 0
 cla.Value = 0
 agn.Value = 0
 etu.Value = 0
 rch.Value = 0
 bil.Value = 0
 spn.Value = 0
Else
dir.Value = 1
 uti.Value = 1
 prt.Value = 1
 fnc.Value = 1
 prf.Value = 1
 cla.Value = 1
 agn.Value = 1
 etu.Value = 1
 rch.Value = 1
 bil.Value = 1
 spn.Value = 1
End If
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Call chargegrd1
End Sub
Private Sub verifier_Checks()
x = 0
x = x + dir.Value
x = x + uti.Value
x = x + prt.Value
x = x + fnc.Value
x = x + prf.Value
x = x + cla.Value
x = x + agn.Value
x = x + etu.Value
x = x + rch.Value
x = x + bil.Value
x = x + mat.Value
x = x + note.Value
x = x + pet.Value
x = x + ppr.Value
x = x + emp.Value
x = x + cpr.Value
x = x + cpf.Value
x = x + cfn.Value
x = x + cet.Value
x = x + cdp.Value
x = x + cbn.Value
x = x + cca.Value
x = x + tpr.Value
x = x + tfn.Value
x = x + tpf.Value
x = x + tet.Value
x = x + tdp.Value
x = x + tbn.Value
x = x + tcl.Value
x = x + trc.Value
x = x + tca.Value
x = x + tjr.Value
x = x + tbu.Value
x = x + tli.Value
x = x + tcn.Value
x = x + afn.Value
x = x + spn.Value
x = x + rpn.Value
x = x + buk.Value
End Sub
Private Sub Ses_Options()
Dim i As Double
Dim n As Double
Call cont
Do While Not ou.EOF
If Text1.Text = ou!nom Then
ou.Delete
End If
ou.MoveNext
Loop
For i = 1 To 34
ou.AddNew
ou!nom = Text1.Text
'**** reception
If i = 1 Then
ou!dir = "1"
ou!div = dir.Caption
ou!frm = dir.Name
ou!act = dir.Value
ElseIf i = 2 Then
ou!dir = "1"
ou!div = uti.Caption
ou!frm = uti.Name
ou!act = uti.Value
ElseIf i = 3 Then
ou!dir = "1"
ou!div = prt.Caption
ou!frm = prt.Name
ou!act = prt.Value
ElseIf i = 4 Then
ou!dir = "1"
ou!div = fnc.Caption
ou!frm = fnc.Name
ou!act = fnc.Value
ElseIf i = 5 Then
ou!dir = "1"
ou!div = prf.Caption
ou!frm = prf.Name
ou!act = prf.Value
ElseIf i = 6 Then
ou!dir = "1"
ou!div = cla.Caption
ou!frm = cla.Name
ou!act = cla.Value
ElseIf i = 7 Then
ou!dir = "1"
ou!div = agn.Caption
ou!frm = agn.Name
ou!act = agn.Value
ElseIf i = 8 Then
ou!dir = "1"
ou!div = etu.Caption
ou!frm = etu.Name
ou!act = etu.Value
ElseIf i = 9 Then
ou!dir = "1"
ou!div = bil.Caption
ou!frm = bil.Name
ou!act = bil.Value
ElseIf i = 10 Then
ou!dir = "1"
ou!div = rch.Caption
ou!frm = rch.Name
ou!act = rch.Value
ElseIf i = 11 Then
ou!dir = "1"
ou!div = spn.Caption
ou!frm = spn.Name
ou!act = spn.Value
'**** reception
'**** control
ElseIf i = 12 Then
ou!dir = "2"
ou!div = pet.Caption
ou!frm = pet.Name
ou!act = pet.Value
ElseIf i = 13 Then
ou!dir = "2"
ou!div = ppr.Caption
ou!frm = ppr.Name
ou!act = ppr.Value
ElseIf i = 14 Then
ou!dir = "2"
ou!div = emp.Caption
ou!frm = emp.Name
ou!act = emp.Value
ElseIf i = 15 Then
ou!dir = "2"
ou!div = tjr.Caption
ou!frm = tjr.Name
ou!act = tjr.Value
'**** control
'**** lessons
ElseIf i = 16 Then
ou!dir = "3"
ou!div = mat.Caption
ou!frm = mat.Name
ou!act = mat.Value
ElseIf i = 17 Then
ou!dir = "3"
ou!div = note.Caption
ou!frm = note.Name
ou!act = note.Value
'**** lessons
'**** caisse
ElseIf i = 18 Then
ou!dir = "4"
ou!div = cpr.Caption
ou!frm = cpr.Name
ou!act = cpr.Value
ElseIf i = 19 Then
ou!dir = "4"
ou!div = cfn.Caption
ou!frm = cfn.Name
ou!act = cfn.Value
ElseIf i = 20 Then
ou!dir = "4"
ou!div = cpf.Caption
ou!frm = cpf.Name
ou!act = cpf.Value
ElseIf i = 21 Then
ou!dir = "4"
ou!div = cet.Caption
ou!frm = cet.Name
ou!act = cet.Value
ElseIf i = 22 Then
ou!dir = "4"
ou!div = cdp.Caption
ou!frm = cdp.Name
ou!act = cdp.Value
ElseIf i = 23 Then
ou!dir = "4"
ou!div = cbn.Caption
ou!frm = cbn.Name
ou!act = cbn.Value
ElseIf i = 24 Then
ou!dir = "4"
ou!div = cca.Caption
ou!frm = cca.Name
ou!act = cca.Value
'**** caisse
'**** comptabilitÈ
ElseIf i = 25 Then
ou!dir = "5"
ou!div = tpr.Caption
ou!frm = tpr.Name
ou!act = tpr.Value
ElseIf i = 26 Then
ou!dir = "5"
ou!div = tfn.Caption
ou!frm = tfn.Name
ou!act = tfn.Value
ElseIf i = 27 Then
ou!dir = "5"
ou!div = tpf.Caption
ou!frm = tpf.Name
ou!act = tpf.Value
ElseIf i = 28 Then
ou!dir = "5"
ou!div = tet.Caption
ou!frm = tet.Name
ou!act = tet.Value
ElseIf i = 29 Then
ou!dir = "5"
ou!div = tdp.Caption
ou!frm = tdp.Name
ou!act = tdp.Value
ElseIf i = 30 Then
ou!dir = "5"
ou!div = tbn.Caption
ou!frm = tbn.Name
ou!act = tbn.Value
ElseIf i = 31 Then
ou!dir = "5"
ou!div = trc.Caption
ou!frm = trc.Name
ou!act = trc.Value
ElseIf i = 32 Then
ou!dir = "5"
ou!div = tca.Caption
ou!frm = tca.Name
ou!act = tca.Value
ElseIf i = 33 Then
ou!dir = "5"
ou!div = tcl.Caption
ou!frm = tcl.Name
ou!act = tcl.Value
End If
Next i
End Sub

Private Sub grd1_Click()
Dim i As Double
Dim j As Double
Dim au As Double
Dim k As Double
i = grd1.Row
j = grd1.Col
If i > 0 Then
If j = 2 Then
grd1.Row = i
grd1.Col = 0
au = grd1.Text
Call cont
Do While Not ut.EOF
If ut!aut = au Then
Label7.Caption = ut!aut
Text1.Text = ut!nom
Text2.Text = ut!mot
Text3.Text = ut!mot
Check1.Value = 1
Check1.Value = 0
Call rec_options
Exit Sub
End If
ut.MoveNext
Loop
End If
If j = 3 Then
grd1.Row = i
grd1.Col = 0
au = grd1.Text
g = MsgBox("Â·  —Ìœ Õﬁ« Õ–› Â–« «·„” Œœ„", vbInformation + vbYesNo + arabic, "AGEP6")
If g = vbYes Then
Call cont
Do While Not ut.EOF
If au = ut!aut And ut!adm <> "1" Then
ut.Delete
ProgressBar1.Value = 0
ProgressBar1.Visible = True
Timer1.Enabled = True
Exit Sub
End If
ut.MoveNext
Loop
MsgBox "⁄›Ê«..·« Ì„ﬂ‰ Õ–› „” Œœ„ admin »√Ì Õ«· „‰ «·√ÕÊ«·", vbExclamation + arabic
End If
End If
End If

End Sub

Private Sub les_Click()
If les.Value = 0 Then
mat.Value = 0
 note.Value = 0
Else
mat.Value = 1
 note.Value = 1
End If
End Sub

Private Sub Text1_Change()
If Len(Text1.Text) > 0 Then
Text1.BackColor = &HC000&
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

Private Sub Text3_Change()
If Len(Text3.Text) > 0 Then
Text3.BackColor = &HC000&
Else
Text3.BackColor = &H8080FF
End If

End Sub

Private Sub Text3_Click()
Text3_Change
End Sub

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 8
If ProgressBar1.Value > 90 Then
MsgBox " „  «·⁄„·Ì… »‰Ã«Õ", vbInformation + arabic
Command2_Click
End If

End Sub
Private Sub chargegrd1()
Dim j As Double
Dim i As Double
Dim P As Double
Dim sm As String
Dim m1 As String
grd1.Clear
grd1.Cols = 4
grd1.Rows = 1
grd1.ColWidth(0) = 0
grd1.ColWidth(1) = 4000
grd1.ColWidth(2) = 800
grd1.ColWidth(3) = 700
grd1.ColAlignment(0) = 1
grd1.ColAlignment(1) = 1
grd1.ColAlignment(2) = 1
grd1.ColAlignment(3) = 1
grd1.Row = 0
grd1.Col = 1
grd1.Text = "«·„” Œœ„"
i = 1
Call cont
grd1.Rows = ut.RecordCount + 3
Do While Not ut.EOF
grd1.Row = i
grd1.Col = 0
grd1.Text = ut!aut
grd1.Col = 1
grd1.Text = ut!nom
grd1.Col = 2
grd1.Text = " ⁄œÌ·"
grd1.CellBackColor = &HFFFF&
grd1.Col = 3
grd1.Text = "Õ–›"
grd1.CellBackColor = &HC0&
i = i + 1
ut.MoveNext
Loop
grd1.Rows = i
grd1.Col = 1
grd1.Sort = 1
End Sub
Private Sub rec_options()
Call cont
Do While Not ou.EOF
If Text1.Text = ou!nom Then
If ou!frm = "dir" Then
dir.Value = ou!act
End If
If ou!frm = "uti" Then
uti.Value = ou!act
End If
If ou!frm = "prt" Then
prt.Value = ou!act
End If
If ou!frm = "fnc" Then
fnc.Value = ou!act
End If
If ou!frm = "prf" Then
prf.Value = ou!act
End If
If ou!frm = "cla" Then
cla.Value = ou!act
End If
If ou!frm = "agn" Then
agn.Value = ou!act
End If
If ou!frm = "etu" Then
etu.Value = ou!act
End If
If ou!frm = "rch" Then
rch.Value = ou!act
End If
If ou!frm = "bil" Then
bil.Value = ou!act
End If
If ou!frm = "mat" Then
mat.Value = ou!act
End If
If ou!frm = "note" Then
note.Value = ou!act
End If
If ou!frm = "pet" Then
pet.Value = ou!act
End If
If ou!frm = "ppr" Then
ppr.Value = ou!act
End If
If ou!frm = "emp" Then
emp.Value = ou!act
End If
If ou!frm = "cpr" Then
cpr.Value = ou!act
End If
If ou!frm = "cpf" Then
cpf.Value = ou!act
End If
If ou!frm = "cfn" Then
cfn.Value = ou!act
End If
If ou!frm = "cet" Then
cet.Value = ou!act
End If
If ou!frm = "cdp" Then
cdp.Value = ou!act
End If
If ou!frm = "cbn" Then
cbn.Value = ou!act
End If
If ou!frm = "cca" Then
cca.Value = ou!act
End If
If ou!frm = "tpr" Then
tpr.Value = ou!act
End If
If ou!frm = "tfn" Then
tfn.Value = ou!act
End If
If ou!frm = "tpf" Then
tpf.Value = ou!act
End If
If ou!frm = "tet" Then
tet.Value = ou!act
End If
If ou!frm = "tet" Then
tet.Value = ou!act
End If
If ou!frm = "tdp" Then
tdp.Value = ou!act
End If
If ou!frm = "tbn" Then
tbn.Value = ou!act
End If
If ou!frm = "tcl" Then
tcl.Value = ou!act
End If
If ou!frm = "trc" Then
trc.Value = ou!act
End If
If ou!frm = "tca" Then
tca.Value = ou!act
End If
If ou!frm = "tjr" Then
tjr.Value = ou!act
End If
If ou!frm = "tbu" Then
tbu.Value = ou!act
End If
If ou!frm = "tli" Then
tli.Value = ou!act
End If
If ou!frm = "tcn" Then
tcn.Value = ou!act
End If
If ou!frm = "afn" Then
afn.Value = ou!act
End If
If ou!frm = "spn" Then
spn.Value = ou!act
End If
If ou!frm = "rpn" Then
rpn.Value = ou!act
End If
If ou!frm = "buk" Then
buk.Value = ou!act
End If
End If
ou.MoveNext
Loop
End Sub
