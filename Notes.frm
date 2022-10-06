VERSION 5.00
Begin VB.Form Notes 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12885
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   480
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton Option3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Caption         =   "ﬂ‘› œ—Ã«  "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   9480
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.OptionButton Option2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Caption         =   "ﬁ”„ ﬁ”„"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Caption         =   " ”ÃÌ· ‰ «∆Ã"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   285
      Left            =   11040
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008000&
      BorderWidth     =   5
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   12855
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ”ÃÌ· ‰ «∆Ã"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   375
      Left            =   11160
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Notes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub

Private Sub Option1_Click()
Unload Notes_B
Unload Notes_C
Notes_E.Show
End Sub

Private Sub Option2_Click()
Unload Notes_B
Unload Notes_E
Notes_C.Show
End Sub

Private Sub Option3_Click()
Unload Notes_C
Unload Notes_E
Notes_B.Show

End Sub
