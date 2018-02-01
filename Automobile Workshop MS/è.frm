VERSION 5.00
Begin VB.Form frm_login 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5805
   ClientLeft      =   6180
   ClientTop       =   2580
   ClientWidth     =   7785
   FillColor       =   &H00C0C0FF&
   BeginProperty Font 
      Name            =   "Cooper Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "è.frx":0000
   ScaleHeight     =   5805
   ScaleWidth      =   7785
   Begin VB.Timer Timer1 
      Left            =   6360
      Top             =   3960
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4680
      Picture         =   "è.frx":5F1C
      ScaleHeight     =   2145
      ScaleWidth      =   2865
      TabIndex        =   12
      Top             =   1440
      Width           =   2895
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4680
      Picture         =   "è.frx":9E90
      ScaleHeight     =   2145
      ScaleWidth      =   2865
      TabIndex        =   11
      Top             =   1440
      Width           =   2895
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4680
      Picture         =   "è.frx":DF53
      ScaleHeight     =   2145
      ScaleWidth      =   2865
      TabIndex        =   10
      Top             =   1440
      Width           =   2895
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4680
      Picture         =   "è.frx":11E44
      ScaleHeight     =   2145
      ScaleWidth      =   2865
      TabIndex        =   9
      Top             =   1440
      Width           =   2895
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4680
      Picture         =   "è.frx":140D6
      ScaleHeight     =   2145
      ScaleWidth      =   2865
      TabIndex        =   8
      Top             =   1440
      Width           =   2895
   End
   Begin VB.PictureBox Picture2 
      Height          =   2175
      Left            =   4680
      Picture         =   "è.frx":160C4
      ScaleHeight     =   2115
      ScaleWidth      =   2835
      TabIndex        =   7
      Top             =   1440
      Width           =   2895
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4680
      Picture         =   "è.frx":1806C
      ScaleHeight     =   2145
      ScaleWidth      =   2865
      TabIndex        =   6
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton cmd_login 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txt_uname 
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin VB.TextBox txt_pword 
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
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   720
      TabIndex        =   5
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
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
      Left            =   720
      TabIndex        =   4
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BackStyle       =   0  'Transparent
      Caption         =   " LOGIN"
      BeginProperty Font 
         Name            =   "Cooper Black"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Dim STRSQL As String
Dim RS As ADODB.Recordset
'Dim pw1 As String
'Dim unam1 As String
Dim utype As String
Public Sub login()
pw = Trim(txt_uname.Text)
unam = Trim(txt_pword.Text)
If ((pw = "ADMIN") And (unam = "admin")) Then
MDIForm1.mnu_service.Visible = False
MDIForm1.mnu_tester.Visible = False
MDIForm1.mnu_machanic.Visible = False
frm_splash.Show
Unload Me
Else
STRSQL = "select * from tbl_user where usernam='" & pw & "' and password='" & unam & "'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
mech_id = RS!u_id
utype = RS!usertype
If (utype = "Machanic") Then
MDIForm1.mnu_service.Visible = False
MDIForm1.mnu_admin.Visible = False
MDIForm1.mnu_tester.Visible = False
MDIForm1.mnu_uprate.Visible = False
MDIForm1.mnu_report.Visible = False
frm_splash.Show
Unload Me
Else
If (utype = "Tester") Then
MDIForm1.mnu_uprate.Visible = False
MDIForm1.mnu_admin.Visible = False
MDIForm1.mnu_machanic.Visible = False
MDIForm1.mnu_report.Visible = False
frm_splash.Show
Unload Me
End If
End If
Else
MsgBox "invallied user", vbOKOnly + vbInformation, "Warning"
clear
End If
End If
'End If
End Sub

Private Sub cmd_login_Click()
login

End Sub
Private Sub Form_Load()
Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Timer1.Enabled = True
Timer1.Interval = 200
End Sub

Private Sub Timer1_Timer()
If Picture1.Visible = True Then
Picture1.Visible = False
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
ElseIf Picture2.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = True
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
ElseIf Picture3.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
ElseIf Picture4.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = True
Picture6.Visible = False
Picture7.Visible = False
ElseIf Picture5.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = True
Picture7.Visible = False
ElseIf Picture6.Visible = True Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = True
ElseIf Picture7.Visible = True Then
Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
End If
End Sub
Public Sub clear()
txt_uname.Text = ""
txt_pword.Text = ""
End Sub
