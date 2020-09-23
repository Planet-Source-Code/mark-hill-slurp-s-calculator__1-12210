VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculator"
   ClientHeight    =   4635
   ClientLeft      =   5640
   ClientTop       =   4380
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6270
   Begin VB.CommandButton Command16 
      Caption         =   "Off"
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   1080
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   1620
      ItemData        =   "Form1.frx":0000
      Left            =   4080
      List            =   "Form1.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton cmdEquals 
      Caption         =   "="
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   3480
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      Caption         =   "÷"
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      Caption         =   "x"
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      Caption         =   "-"
      Height          =   495
      Left            =   3000
      TabIndex        =   13
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      Caption         =   "+"
      Height          =   495
      Left            =   3000
      TabIndex        =   12
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   2040
      TabIndex        =   8
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'www.slurpcandy.co.uk/welcome.htm
'Slurp's Calculator v1.0
'21st October 00
'Author: Mark A. Hill

Private Sub cmdEquals_Click()
Call Equals
End Sub

Private Sub Command1_Click()
Call One
End Sub

Private Sub Command10_Click()
Call zero
End Sub

Private Sub Command11_Click()
Call clear
End Sub

Private Sub Command12_Click()
Call Add
End Sub

Private Sub Command13_Click()
Call Subtract
End Sub

Private Sub Command14_Click()
Call multiply
End Sub

Private Sub Command15_Click()
Call Divide
End Sub

Private Sub Command16_Click()
Call EndIt
End Sub

Private Sub Command2_Click()
Call two
End Sub

Private Sub Command3_Click()
Call three
End Sub

Private Sub Command4_Click()
Call four
End Sub

Private Sub Command5_Click()
Call five
End Sub

Private Sub Command6_Click()
Call six
End Sub

Private Sub Command7_Click()
Call seven
End Sub

Private Sub Command8_Click()
Call eight
End Sub

Private Sub Command9_Click()
Call nine
End Sub

Private Sub Form_Load()
Form1.Label1.RightToLeft = True
End Sub

