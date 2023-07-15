VERSION 5.00
Begin VB.Form adw 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADD DEALER"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "adw.frx":0000
   ScaleHeight     =   6285
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "                                                           ADD NEW DEALER DETAILS                                          "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   3855
      Left            =   1800
      TabIndex        =   16
      Top             =   1800
      Width           =   7335
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000B&
         Caption         =   "* Chech Dealer USER NAME Availability."
         DownPicture     =   "adw.frx":D542
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3240
         MaskColor       =   &H00C0FFFF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   3375
      End
      Begin VB.CommandButton BACK 
         BackColor       =   &H000000FF&
         Caption         =   "BACK"
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
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton RESET 
         BackColor       =   &H00FF8080&
         Caption         =   "RESET"
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
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton SUBMIT 
         BackColor       =   &H00FF8080&
         Caption         =   "SUBMIT"
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
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txta 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3720
         TabIndex        =   6
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtc 
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
         Left            =   3720
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
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
         Left            =   3720
         TabIndex        =   4
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtp 
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
         Left            =   3720
         TabIndex        =   3
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtu 
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
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "PASSWORD :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ADDRESS  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTACT NO  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "NAME  :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEALER USER NAME :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1440
         TabIndex        =   17
         Top             =   360
         Width           =   1950
      End
   End
   Begin VB.Image Image1 
      Height          =   3975
      Index           =   1
      Left            =   -600
      Picture         =   "adw.frx":1558E
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   4440
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ABC DAIRY"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   585
      Index           =   1
      Left            =   4695
      TabIndex        =   24
      Top             =   2760
      Width           =   2865
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
      Index           =   1
      Left            =   7800
      TabIndex        =   23
      Top             =   2760
      Width           =   495
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
      Index           =   1
      Left            =   4680
      TabIndex        =   22
      Top             =   3240
      Width           =   3615
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
      Left            =   4440
      TabIndex        =   15
      Top             =   5880
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights are Reserved @ ABC DAIRY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Index           =   1
      Left            =   5640
      TabIndex        =   14
      Top             =   5880
      Width           =   3315
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6600
      TabIndex        =   13
      Top             =   240
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   7320
      TabIndex        =   12
      Top             =   240
      Width           =   105
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
      ForeColor       =   &H00FF0000&
      Height          =   270
      Index           =   0
      Left            =   3480
      TabIndex        =   11
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Index           =   0
      Left            =   6840
      TabIndex        =   10
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ABC DAIRY"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   585
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   960
      Width           =   2865
   End
End
Attribute VB_Name = "adw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BACK_Click()
adw.Hide
End Sub
Private Sub Command1_Click()
Set rs = con.Execute("select * from Dealerdb where uname='" + txtu.Text + "' ")
If (Not rs.EOF) Then
MsgBox "Try Another USER NAME", vbInformation, "USER NAME"
txtu = ""
txtu.SetFocus
Else
MsgBox "You Can Proceed With Your USER NAME", vbInformation, "USER NAME VERIFICATION"
End If
End Sub
Private Sub Form_Load()
date = DateTime.Now
Call connectdb
End Sub




Private Sub RESET_Click()
txtu = ""
txtp = ""
txtn = ""
txta = ""
txtc = ""
txtu.SetFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub
Private Sub SUBMIT_Click()
If txtu = "" Then
MsgBox "Enter User Name.", vbInformation
txtu.SetFocus
Exit Sub
End If

If txtn.Text = "" Then
MsgBox "Enter Dealer Name.", vbInformation
txtn.SetFocus
Exit Sub
End If

If txtp.Text = "" Then
MsgBox "Enter Password.", vbInformation
txtp.SetFocus
Exit Sub
End If

If txtc.Text = "" Then
MsgBox "Enter Contact Number.", vbInformation
txtc.SetFocus
Exit Sub
End If

If txta.Text = "" Then
MsgBox "Enter Address.", vbInformation
txta.SetFocus
Exit Sub
End If

con.Execute ("insert into Dealerdb values('" + txtu.Text + "','" + txtn.Text + "','" + txtc.Text + "', '" + txta.Text + "','" + txtp.Text + "')")
MsgBox ("Dealer Details Saved Succsessfully..!")
txtu = ""
txtp = ""
txtn = ""
txta = ""
txtc = ""
End Sub
Private Sub txtu_Change()
If KeyAscii = 13 Then
txtu.Text = UCase(txtnumber.Text)
txtu.SetFocus
End If
End Sub
