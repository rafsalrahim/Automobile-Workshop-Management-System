VERSION 5.00
Begin VB.Form frm_chpswrd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   5535
   ClientLeft      =   6075
   ClientTop       =   3480
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6405
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox txt_cpassword 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   9
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1440
         TabIndex        =   8
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox txt_new 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txt_old 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txt_name 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Conform Pasword"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   11
         Top             =   3600
         Width           =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   6
         Top             =   3000
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   2280
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Name "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   1560
         Width           =   1110
      End
      Begin VB.Label Label1 
         Caption         =   "CHANGE PASSWORD"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frm_chpswrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Dim psw As String
Dim uname As String
Dim utype As String
Public Sub Test()
psw = Trim(txt_old.Text)
uname = Trim(txt_name.Text)
STRSQL = "update tbl_user set password ='" & txt_new.Text & "' where usernam='" & uname & "' and password='" & psw & "'"
Set RS = adocn.Execute(STRSQL)
MsgBox " Password has been changed "
End Sub
Public Function fnValidation()
    Dim ok As Boolean
    If Trim(txt_old.Text) = "" Or txt_old.Text <> unam Then
        ok = False
    Else
    If Trim(txt_new.Text) = "" Then
        ok = False
    Else
    If Trim(txt_cpassword.Text) = "" Then
        ok = False
    Else
    If txt_new.Text <> txt_cpassword.Text Then
        ok = False
    Else
    ok = True
    End If
    End If
    End If
    End If
    fnValidation = ok
End Function

Private Sub Command1_Click()
If fnValidation Then
Test
clear
Unload Me
MDIForm1.Show
Else
  MsgBox "Invalid Password or Username....."
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
txt_name.Text = pw
End Sub
Public Sub clear()
txt_old.Text = ""
txt_new.Text = ""
txt_cpassword.Text = ""
End Sub





Private Sub txt_cpassword_KeyPress(KeyAscii As Integer)
If Len(txt_cpassword.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_cpassword.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_new_KeyPress(KeyAscii As Integer)
If Len(txt_new.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_new.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_old_KeyPress(KeyAscii As Integer)
If Len(txt_old.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_old.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub
