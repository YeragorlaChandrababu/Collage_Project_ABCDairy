VERSION 5.00
Begin VB.Form dlogw 
   Caption         =   "DEALER LOGIN WINDOW"
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   Picture         =   "dlogw.frx":0000
   ScaleHeight     =   4890
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2055
      Left            =   2880
      TabIndex        =   5
      Top             =   2160
      Width           =   6015
      Begin VB.Label lblu 
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "YOUR USER NAME :"
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   1560
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVIDING HEALTH TO INDIA"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   3840
      TabIndex        =   4
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6960
      TabIndex        =   3
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DAIRY DAY "
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   585
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Label date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8400
      TabIndex        =   1
      Top             =   240
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATE :  "
      Height          =   255
      Index           =   0
      Left            =   7680
      TabIndex        =   0
      Top             =   240
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   120
      Picture         =   "dlogw.frx":37223
      Top             =   120
      Width           =   3345
   End
End
Attribute VB_Name = "dlogw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
date = DateTime.Now
lblu = dlogin.txtu.Text
End Sub

