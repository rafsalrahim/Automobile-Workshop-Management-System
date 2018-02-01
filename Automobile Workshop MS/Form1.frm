VERSION 5.00
Begin VB.Form frm_registration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registration"
   ClientHeight    =   8115
   ClientLeft      =   7155
   ClientTop       =   1965
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fram_regi 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.ComboBox comb_comp 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Form1.frx":0000
         Left            =   2280
         List            =   "Form1.frx":0013
         TabIndex        =   28
         Text            =   "------SELECT ONE------"
         Top             =   4920
         Width           =   2295
      End
      Begin VB.ComboBox comb_vmodel 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2280
         TabIndex        =   27
         Text            =   "------SELECT ONE------"
         Top             =   5520
         Width           =   2295
      End
      Begin VB.TextBox txt_name 
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
         Left            =   2280
         TabIndex        =   9
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txt_add 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2760
         Width           =   2295
      End
      Begin VB.TextBox txt_phno 
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
         Left            =   2280
         TabIndex        =   7
         Top             =   3720
         Width           =   2295
      End
      Begin VB.TextBox txt_email 
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
         Left            =   2280
         TabIndex        =   6
         Top             =   4320
         Width           =   2295
      End
      Begin VB.TextBox txt_rno 
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
         Left            =   2280
         TabIndex        =   5
         Top             =   6120
         Width           =   2295
      End
      Begin VB.TextBox txt_cheno 
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
         Left            =   2280
         TabIndex        =   4
         Top             =   6720
         Width           =   2295
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   7320
         Width           =   1095
      End
      Begin VB.CommandButton cmd_submit 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   7320
         Width           =   1095
      End
      Begin VB.TextBox txt_vhid 
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
         Height          =   405
         Left            =   2280
         TabIndex        =   1
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label lbl_vhmod 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1800
         TabIndex        =   30
         Top             =   5520
         Width           =   150
      End
      Begin VB.Label lbl_comp 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   1440
         TabIndex        =   29
         Top             =   4920
         Width           =   150
      End
      Begin VB.Label lbl_chase 
         AutoSize        =   -1  'True
         Caption         =   "* This field is required"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   3000
         TabIndex        =   26
         Top             =   6480
         Width           =   1530
      End
      Begin VB.Label lbl_regno 
         AutoSize        =   -1  'True
         Caption         =   "* This field is required"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   3000
         TabIndex        =   25
         Top             =   5880
         Width           =   1530
      End
      Begin VB.Label lbl_addr 
         AutoSize        =   -1  'True
         Caption         =   "* This field is required"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   3000
         TabIndex        =   24
         Top             =   2520
         Width           =   1530
      End
      Begin VB.Label lbl_name 
         AutoSize        =   -1  'True
         Caption         =   "* This field is required"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   3000
         TabIndex        =   23
         Top             =   1920
         Width           =   1530
      End
      Begin VB.Label lbl_email 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   3120
         TabIndex        =   22
         Top             =   4080
         Width           =   45
      End
      Begin VB.Label lbl_mob 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   3120
         TabIndex        =   21
         Top             =   3480
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Name"
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
         Left            =   360
         TabIndex        =   20
         Top             =   2160
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Address"
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
         Left            =   360
         TabIndex        =   19
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ph.Number"
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
         Left            =   360
         TabIndex        =   18
         Top             =   3720
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Email"
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
         Left            =   360
         TabIndex        =   17
         Top             =   4320
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Register No."
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
         Left            =   360
         TabIndex        =   16
         Top             =   6120
         Width           =   1155
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Chase no."
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
         Left            =   360
         TabIndex        =   15
         Top             =   6720
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "  REGISTRATION"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         TabIndex        =   14
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lbl_ftime 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " <date>"
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
         Height          =   285
         Left            =   3720
         TabIndex        =   13
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle Model"
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
         Left            =   360
         TabIndex        =   12
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle ID"
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
         Left            =   360
         TabIndex        =   11
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Company"
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
         Left            =   360
         TabIndex        =   10
         Top             =   4920
         Width           =   885
      End
   End
End
Attribute VB_Name = "frm_registration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim STRSQL As String
Dim RS1 As New ADODB.Recordset
Dim STRSQL1 As String
Dim cmpdate As Date
Dim AUNT As Integer
Dim s As String
Public Sub cleardata()
txt_vhid.Text = " "
txt_name.Text = " "
txt_add.Text = " "
txt_phno.Text = " "
txt_email.Text = " "
comb_comp.Text = " "
comb_vmodel.Text = " "
txt_rno.Text = " "
txt_cheno = " "
End Sub

Public Sub cmd_cancel_Click()
cleardata
clearlabel
End Sub

Public Sub insert()
STRSQL = "insert into tbl_vehicleregistration(name,address,mob_no,email,vh_compnam,vh_model, " _
         & " regi_no,chai_no,date,status) values ('" & txt_name.Text & "','" & txt_add.Text & "','" & txt_phno.Text & "' ," _
         & " '" & txt_email.Text & "','" & comb_comp.List(comb_comp.ListIndex) & "','" & comb_vmodel.List(comb_vmodel.ListIndex) & "','" & txt_rno.Text & "'," _
         & " '" & txt_cheno.Text & "','" & cmpdate & "','REGISTERED')"

Set RS = adocn.Execute(STRSQL)
clearlabel
End Sub
Public Sub cmd_submit_Click()
If fnValidation And fnEmailValidation And fnMobileValidation Then
insert
J_ID = txt_vhid.Text
frm_jobsheet.Show
'MsgBox "Registration successfull"
cleardata
Unload Me
MDIForm1.Show
Else
MsgBox "registration unsuccessfull"
End If
End Sub



Private Sub comb_comp_Click()
fillmodel
End Sub

Private Sub comb_comp_KeyPress(KeyAscii As Integer)
lbl_comp.Visible = True
KeyAscii = 0
End Sub

Private Sub comb_vmodel_Change()
MsgBox "Should Select One"
End Sub

Public Sub Form_Load()
lbl_ftime.Caption = DateValue(Now)
s = DateValue(Now)
cmpdate = CDate(s)
fillvehicleid
fillmodel
clearlabel
End Sub
Public Sub fillvehicleid()
STRSQL = "select * from tbl_vehicleregistration"
Set RS = adocn.Execute(STRSQL)
If RS.BOF = True And RS.EOF = True Then
txt_vhid = 1
Else
While Not RS.EOF
AUNT = RS.Fields(0)
RS.MoveNext
Wend
txt_vhid.Text = AUNT + 1
End If
End Sub
Public Function fnValidation()
Dim ok As Boolean
If (Trim(txt_name.Text) = "") Then
   lbl_name.Visible = True
ok = False
Else
 If (Trim(txt_add.Text) = "") Then
  lbl_addr.Visible = True
    ok = False
    Else
    If (comb_comp.Text = "------SELECT ONE------") Then
     lbl_comp.Visible = True
    ok = False
    Else
    If (comb_vmodel.Text = "------SELECT ONE------") Then
     lbl_vhmod.Visible = True
     ok = False
     Else
     If (Trim(txt_rno.Text) = "") Then
       lbl_regno.Visible = True
     ok = False
     Else
     If (Trim(txt_cheno.Text) = "") Then
        lbl_chase.Visible = True
     ok = False
     Else
    ok = True
    End If
    End If
    End If
    End If
    End If
    End If
    'End If
    'End If
    fnValidation = ok
End Function

Public Function fnEmailValidation()
    Dim Email As String
    Dim ok As Boolean
    Email = txt_email.Text
    LCase (Email)
    If (Email Like "*@*.com" Or Email Like "*@*.co.in") Then
        ok = True
        lbl_email.Visible = False
    Else
        ok = False
        lbl_email.Caption = "* Invalid Email Id"
        lbl_email.Visible = True
    End If
    fnEmailValidation = ok
End Function

Public Function fnMobileValidation()
        Dim Mobile As String
        Mobile = txt_phno.Text
        Dim ok As Boolean
        If (IsNumeric(Mobile) And Len(Mobile) = "10") Then
            ok = True
            lbl_mob.Visible = False
        Else
            ok = False
            lbl_mob.Caption = "* Invalid Mobile Number"
            lbl_mob.Visible = True
        End If
        fnMobileValidation = ok
End Function



Private Sub txt_add_Change()
 If Trim(txt_add.Text) = "" Then
        lbl_addr.Visible = True
    Else
        lbl_addr.Visible = False
    End If
End Sub

Private Sub txt_add_KeyPress(KeyAscii As Integer)
If Len(txt_add.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_add.Text = ""
End If
End Sub

Private Sub txt_cheno_Change()
If Trim(txt_cheno.Text) = "" Then
        lbl_chase.Visible = True
    Else
        lbl_chase.Visible = False
    End If
End Sub

Private Sub txt_cheno_KeyPress(KeyAscii As Integer)
If Len(txt_cheno.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_cheno.Text = ""
End If
End Sub

Private Sub txt_email_KeyPress(KeyAscii As Integer)
If Len(txt_email.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_email.Text = ""
End If
End Sub

Private Sub txt_email_Lostfocus()
fnEmailValidation
End Sub
Private Sub txt_email_change()
clearlabel
End Sub

Private Sub txt_name_Change()
 If Trim(txt_name.Text) = "" Then
        lbl_name.Visible = True
    Else
        lbl_name.Visible = False
    End If
End Sub

Private Sub txt_name_KeyPress(KeyAscii As Integer)
If Len(txt_name.Text) = 30 Then
MsgBox "Limit Crossed", vhinformation
txt_name.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_phno_KeyPress(KeyAscii As Integer)
If Len(txt_add.Text) = 11 Then
MsgBox "Limit Crossed", vhinformation
txt_add.Text = ""
End If
If KeyAscii > 48 And KeyAscii > 58 And KeyAscii <> 8 And KeyAscii <> 13 Then
MsgBox "Only integer is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_phno_lOSTFocus()
fnMobileValidation
End Sub
Private Sub txt_phno_change()
clearlabel
End Sub

Private Sub txt_rno_Change()
 If Trim(txt_rno.Text) = "" Then
        lbl_regno.Visible = True
    Else
        lbl_regno.Visible = False
    End If
End Sub

Public Sub clearlabel()
    lbl_name.Visible = False
    lbl_addr.Visible = False
    lbl_mob.Visible = False
    lbl_email.Visible = False
   lbl_comp.Visible = False
    lbl_regno.Visible = False
    lbl_chase.Visible = False
    lbl_vhmod.Visible = False
End Sub
Public Sub fillmodel()
STRSQL1 = "select * from tbl_wheelrate where vh_company='" & comb_comp.List(comb_comp.ListIndex) & "'"
Set RS1 = adocn.Execute(STRSQL1)
Do While Not RS1.EOF
        comb_vmodel.AddItem RS1!service
        RS1.MoveNext
    Loop
End Sub


Private Sub txt_rno_KeyPress(KeyAscii As Integer)
If Len(txt_rno.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_rno.Text = ""
End If
End Sub
