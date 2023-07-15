VERSION 5.00
Begin VB.Form reg 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton EXIT 
      Caption         =   "EXIT"
      Height          =   315
      Left            =   1560
      TabIndex        =   17
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton RESET 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton SUBMIT 
      Caption         =   "Submit"
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox txtn 
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtc 
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   6
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtu 
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtp 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txta 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtco 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtr 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UserName"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Roll Nome"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTRATION FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "reg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EXIT_Click()
answer = MsgBox("Do you want to quit?", vbExclamation + vbYesNo, "Warning")
If answer = vbYes Then
Me.Hide
Else
End If
End Sub
Private Sub Form_Load()
connectdb
End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub
