VERSION 5.00
Begin VB.Form cpw 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CHANGE YOUR PASSWOR"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "cpw.frx":0000
   ScaleHeight     =   4650
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame clogin 
      BackColor       =   &H00FFFFC0&
      Caption         =   "CHANGE YOUR PASSWORD...."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   2535
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   5535
      Begin VB.TextBox txtcn 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtn 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton login 
         BackColor       =   &H0080FF80&
         Caption         =   "CONFIRM"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton back 
         BackColor       =   &H000000FF&
         Caption         =   "GO BACK"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIRM NEW PASSWORD:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   2580
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "NEW PASSWORD :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   2880
         TabIndex        =   12
         Top             =   600
         Width           =   105
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "USER NAME :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   225
         Index           =   0
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   1170
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights are Reserved @ ABC DAIRY "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   15
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   1095
      Index           =   1
      Left            =   -120
      Picture         =   "cpw.frx":797B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ABC DAIRY"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   270
      Left            =   2385
      TabIndex        =   11
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVIDING HEALTH TO INDIA"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   10
      Top             =   960
      Width           =   2625
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   720
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WARNING :"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Index           =   1
      Left            =   1320
      TabIndex        =   8
      Top             =   3960
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DATE :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   660
   End
   Begin VB.Label date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   105
   End
End
Attribute VB_Name = "cpw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SSQL As String
Private Sub BACK_Click()
cpw.Hide
End Sub
Private Sub Form_Load()
lblu = clogw.lblu.Caption
date = DateTime.Now
connectdb
rs.Close
End Sub
Private Sub login_Click()
If (txtn.Text <> txtcn.Text) Then
    MsgBox "Password Mismatching!"
Else
If txtn.Text = "" Then
MsgBox "Password can't be Blank!", vbCritical, "WARNING"
Else
    SSQL = "select * from Companydb Where Uname='" & lblu.Caption & "'"
rs.Open SSQL, con, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs!Password = txtn.Text
rs.Update
MsgBox "Password Updated Successfully", vbInformation, "Success"
txtn = ""
txtcn = ""
Me.Hide
Else
MsgBox "Failed to Updated", vbInformation, "Failed"
End If
rs.Close
End If
End If
End Sub

