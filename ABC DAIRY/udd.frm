VERSION 5.00
Begin VB.Form udud 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "UPDATE OR DELETE DEALER DETAILS"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "udd.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "                                                     UPDATE DEALER DETAILS                                                   "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3735
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   7335
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF00&
         Caption         =   "DLELTE"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   " * Get Dealer Details "
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
         Left            =   3360
         TabIndex        =   2
         Top             =   600
         Width           =   2055
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
         Left            =   3120
         TabIndex        =   1
         ToolTipText     =   "ENTER FORMER ID TO EDIT & Click on the Below Button."
         Top             =   240
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
         Left            =   3120
         TabIndex        =   3
         Top             =   960
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
         Left            =   3120
         TabIndex        =   4
         Top             =   1320
         Width           =   2655
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
         Left            =   3120
         TabIndex        =   5
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CommandButton SUBMIT 
         BackColor       =   &H00FFFF00&
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
         Left            =   720
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3120
         Width           =   1215
      End
      Begin VB.CommandButton RESET 
         BackColor       =   &H00FFFF00&
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
         Left            =   3720
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3120
         Width           =   1215
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
         Left            =   5160
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
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
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   1950
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         Height          =   225
         Left            =   2040
         TabIndex        =   11
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   225
         Left            =   1320
         TabIndex        =   10
         Top             =   1320
         Width           =   1410
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Height          =   225
         Left            =   1680
         TabIndex        =   9
         Top             =   1680
         Width           =   1035
      End
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Note:"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   240
      TabIndex        =   22
      Top             =   5160
      Width           =   705
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "On Deleting Dealer Details All the Former Details and Dealer Milk Production Details Will Deleted Which are Assigned to Him."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   5520
      Width           =   10125
   End
   Begin VB.Image Image1 
      Height          =   3975
      Index           =   1
      Left            =   -480
      Picture         =   "udd.frx":C966
      Stretch         =   -1  'True
      Top             =   -1320
      Width           =   4200
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
      Left            =   6960
      TabIndex        =   19
      Top             =   5880
      Width           =   3315
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
      Left            =   5760
      TabIndex        =   18
      Top             =   5880
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   6840
      TabIndex        =   17
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   7560
      TabIndex        =   16
      Top             =   120
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
      ForeColor       =   &H00FF00FF&
      Height          =   270
      Left            =   3960
      TabIndex        =   15
      Top             =   1320
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
      ForeColor       =   &H8000000D&
      Height          =   330
      Left            =   6960
      TabIndex        =   14
      Top             =   840
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
      ForeColor       =   &H0000FFFF&
      Height          =   585
      Left            =   3975
      TabIndex        =   13
      Top             =   840
      Width           =   2865
   End
End
Attribute VB_Name = "udud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SSQL As String
Private Sub BACK_Click()
udud.Hide
End Sub
Private Sub Command1_Click()
Set rs = con.Execute("select * from Dealerdb where uname='" + txtu.Text + "'")
If (Not rs.EOF) Then
txtn.Text = rs("NAME")
txtc.Text = rs("CONTACT")
txta.Text = rs("ADDRESS")
rs.Close
Else
MsgBox "USER NAME You Entered Not Valid", vbInformation, "Warning"
txtu = ""
txtn = ""
txtc = ""
txta = ""
End If
End Sub
Private Sub Command2_Click()
answer = MsgBox("Do you want to 'Delete' ?", vbExclamation + vbYesNo, "Warning")
If answer = vbYes Then
rs.Open "delete from dealerdb where uname='" & txtu.Text & "'"
rs.Open "delete from formerdb where assdid='" & txtu.Text & "'"
rs.Open "delete from milkdb where dealeruname='" & txtu.Text & "'"
Else
End If
MsgBox "Data Deleted Successfully..!", vbInformation, "Warning"
txtu = ""
txtn = ""
txta = ""
txtc = ""
End Sub

Private Sub Form_Load()
date = DateTime.Now
connectdb
Set rs = con.Execute("select * from Dealerdb")
End Sub
Private Sub RESET_Click()
txtn = ""
txtc = ""
txta = ""
End Sub
Private Sub SUBMIT_Click()
If txtn.Text = "" Then
MsgBox "Please Enter the Stop .", vbInformation
txtn.SetFocus
Exit Sub
End If

If txtc.Text = "" Then
MsgBox "Please Enter the biginning stop .", vbInformation
txtc.SetFocus
Exit Sub
End If

If txta.Text = "" Then
MsgBox "Please Enter the Ending stop .", vbInformation
txta.SetFocus
Exit Sub
End If

SSQL = "select * from Dealerdb Where Uname='" & txtu.Text & "'"
rs.Open SSQL, con, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs("name") = txtn.Text
rs("contact") = txtc.Text
rs("address") = txta.Text
rs.Update
MsgBox "Details Updated Successfully", vbInformation, "Success"
Else
MsgBox "Failed to Updated", vbInformation, "Failed"
End If
rs.Close

txtu = ""
txtn = ""
txta = ""
txtc = ""
End Sub
