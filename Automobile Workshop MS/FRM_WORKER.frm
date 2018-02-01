VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRM_WORKER 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Worker"
   ClientHeight    =   8145
   ClientLeft      =   2970
   ClientTop       =   1650
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   12210
   Begin VB.Frame Frame1 
      Height          =   7815
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.CommandButton cmd_delete 
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   28
         Top             =   6840
         Width           =   1215
      End
      Begin VB.CommandButton cmd_update 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   27
         Top             =   6840
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid gridadduser 
         Height          =   4335
         Left            =   5520
         TabIndex        =   26
         Top             =   1320
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   7646
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   2760
         TabIndex        =   17
         Top             =   5760
         Width           =   2895
         Begin VB.OptionButton Option2 
            Caption         =   "Machanic"
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
            Left            =   1440
            TabIndex        =   19
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Tester"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox txt_cont 
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
         Left            =   2760
         TabIndex        =   15
         Top             =   3960
         Width           =   2415
      End
      Begin VB.CommandButton cmd_submit 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   14
         Top             =   6840
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "CANCAL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   13
         Top             =   6840
         Width           =   1335
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
         Left            =   2760
         TabIndex        =   11
         Top             =   5160
         Width           =   2415
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
         Left            =   2760
         TabIndex        =   9
         Top             =   4560
         Width           =   2415
      End
      Begin VB.TextBox txt_waddr 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2760
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txt_wname 
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
         Left            =   2760
         TabIndex        =   5
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox txt_uid 
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
         Left            =   2760
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lbl_utype 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4200
         TabIndex        =   25
         Top             =   5520
         Width           =   45
      End
      Begin VB.Label lbl_pword 
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
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   3600
         TabIndex        =   24
         Top             =   4920
         Width           =   1530
      End
      Begin VB.Label lbl_unam 
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
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   3600
         TabIndex        =   23
         Top             =   4320
         Width           =   1530
      End
      Begin VB.Label lbl_cont 
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
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   3600
         TabIndex        =   22
         Top             =   3720
         Width           =   1530
      End
      Begin VB.Label lbl_waddr 
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
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   3600
         TabIndex        =   21
         Top             =   2520
         Width           =   1530
      End
      Begin VB.Label lbl_wnam 
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
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   3600
         TabIndex        =   20
         Top             =   1800
         Width           =   1530
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CONTACT"
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
         Left            =   480
         TabIndex        =   16
         Top             =   4080
         Width           =   1125
      End
      Begin VB.Label Label7 
         Caption         =   "USER TYPE"
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
         Left            =   480
         TabIndex        =   12
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "PASSWORD"
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
         Left            =   480
         TabIndex        =   10
         Top             =   5280
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "USERNAME"
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
         Left            =   480
         TabIndex        =   8
         Top             =   4680
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ADDRESS"
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
         Left            =   480
         TabIndex        =   6
         Top             =   2880
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "WORKER NAME"
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
         TabIndex        =   4
         Top             =   2160
         Width           =   1770
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User ID"
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
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "WORKERS DATA"
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
         Left            =   3960
         TabIndex        =   1
         Top             =   480
         Width           =   2370
      End
   End
End
Attribute VB_Name = "FRM_WORKER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim STRSQL As String
Dim AUNT As Integer
Dim ty As String
Dim SLNO As Integer
Dim Test As String

Private Sub cmd_cancel_Click()
cleardata
clrlabel
End Sub


Private Sub cmd_delete_Click()
If fnValidation Then
STRSQL = "delete from tbl_user where u_id = '" & gridadduser.TextMatrix(gridadduser.RowSel, 0) & "'"
    Set RS = adocn.Execute(STRSQL)
    cmd_delete.Enabled = False
    cmd_update.Enabled = False
    cleardata
    clrlabel
    'subClearlabel
    cmd_submit.Enabled = True
    cmd_update.Enabled = False
    cmd_delete.Enabled = False
    cmd_cancel.Enabled = True
    MsgBox "User deleted . . ."
    'OptM.Value = True
    subaddtogrid
   Unload Me
   MDIForm1.Show
   Else
   MsgBox "Field Required"
   End If
    
End Sub

Private Sub cmd_submit_Click()
If fnValidation And fnMobileValidation Then
fillworker
insert
subaddtogrid
Unload Me
MDIForm1.Show
Else
MsgBox "registration unsuccessfull"
Form_Load
 End If
End Sub
Public Sub fillworker()
STRSQL = "select * from tbl_user"
Set RS = adocn.Execute(STRSQL)
If RS.BOF = True And RS.EOF = True Then
txt_uid.Text = 1
Else
While Not RS.EOF
AUTN = RS.Fields(0)
RS.MoveNext
Wend
txt_uid.Text = AUTN + 1
End If
End Sub
Public Sub insert()
If Option1.Value = True Then
ty = "Tester"
Else
ty = "Machanic"
End If
If check Then
 lbl_val1.Caption = "USERNAME ALREADY EXIST"
 txt_uname.Text = ""
 Else
STRSQL = "insert into tbl_user(u_name,u_add,u_cont,usernam,password,usertype) values ('" & txt_wname.Text & "','" & txt_waddr & "','" & txt_cont & "','" & txt_uname.Text & "','" & txt_pword.Text & "','" & ty & "')"
       Set RS = adocn.Execute(STRSQL)
       MsgBox "registration successfull"
       subsetgrid
 Form_Load
       End If
End Sub
Public Function fnValidation()
Dim ok, ok2, ok3, ok4, ok5, ok6 As Boolean
    
    If (Trim(txt_wname.Text) = "") Then
    ok2 = False
     lbl_wnam.Visible = True
     Else
     ok2 = True
     End If
    If (Trim(txt_waddr.Text) = "") Then
    ok3 = False
     lbl_waddr.Visible = True
    Else
    ok3 = True
    End If
    If (Trim(txt_cont.Text = "")) Then
    ok4 = False
    lbl_cont.Visible = True
    Else
    ok4 = True
    End If
    If (Trim(txt_uname.Text) = "") Then
    ok5 = False
    lbl_unam.Visible = True
    Else
    ok5 = True
    End If
    If (Trim(txt_pword.Text) = "") Then
    ok6 = False
    lbl_pword.Visible = True
    Else
    ok6 = True
    End If
 If (ok2 And ok3 And ok4 And ok5 And ok6 = True) Then
 ok = True
 Else
 ok = False
 End If
 fnValidation = ok
    End Function
     
Public Sub cleardata()
'txt_uid.Text = " "
txt_wname.Text = ""
txt_waddr.Text = ""
txt_uname.Text = ""
txt_pword.Text = ""
txt_cont.Text = ""
End Sub

Private Sub subsetgrid()
gridadduser.Cols = 3
gridadduser.Rows = 2
gridadduser.FixedRows = 1
gridadduser.TextMatrix(0, 1) = "NAME"
gridadduser.TextMatrix(0, 2) = "TYPE"
gridadduser.ColWidth(0) = 0
gridadduser.ColWidth(1) = 1000
gridadduser.ColWidth(2) = 1730
End Sub



Private Sub cmd_update_Click()
If fnValidation And fnMobileValidation Then
If Option1.Value = True Then
ty = "Tester"
Else
ty = "Machanic"
End If

STRSQL = "update tbl_user set u_name='" & txt_wname.Text & "',u_add='" & txt_waddr.Text & "', " _
                     & "u_cont ='" & txt_cont.Text & "',usernam='" & txt_uname.Text & "',password='" & txt_pword & "',usertype='" & ty & "' where " _
                    & " u_id = '" & gridadduser.TextMatrix(gridadduser.RowSel, 0) & "'"
        adocn.Execute (STRSQL)
        MsgBox "User details updated . . ."
        subaddtogrid
        cmd_submit.Enabled = True
        cmd_update.Enabled = False
        cmd_delete.Enabled = False
        cmd_cancel.Enabled = True
        cleardata
        Else
        MsgBox "Field Required"
        End If
End Sub

Private Sub Form_Load()
Option1.Value = True
fillworker
subsetgrid
subaddtogrid
cleardata
clrlabel
End Sub


Private Sub txt_cont_Change()
clrlabel
End Sub

Private Sub txt_cont_KeyPress(KeyAscii As Integer)
If Len(txt_cont.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_cont.Text = ""
End If
If KeyAscii > 48 And KeyAscii > 58 And KeyAscii <> 8 And KeyAscii <> 13 Then
MsgBox "Only integer is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_cont_LostFocus()
fnMobileValidation
End Sub

Private Sub txt_pword_Change()
If (Trim(txt_pword.Text) = "") Then
lbl_pword.Visible = True
    Else
        lbl_pword.Visible = False
   End If
End Sub

Private Sub txt_pword_KeyPress(KeyAscii As Integer)
If Len(txt_pword.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_pword.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_uname_Change()
If (Trim(txt_uname.Text) = "") Then
lbl_unam.Visible = True
    Else
        lbl_unam.Visible = False
   End If
End Sub

Private Sub txt_uname_KeyPress(KeyAscii As Integer)
If Len(txt_uname.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_uname.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_waddr_Change()
If (Trim(txt_waddr.Text) = "") Then
lbl_waddr.Visible = True
    Else
        lbl_waddr.Visible = False
   End If
End Sub

Private Sub txt_waddr_KeyPress(KeyAscii As Integer)
If Len(txt_waddr.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_waddr.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_wname_Change()
If (Trim(txt_wname.Text) = "") Then
lbl_wnam.Visible = True
    Else
        lbl_wnam.Visible = False
   End If
End Sub
Public Sub subaddtogrid()
'gridadduser.Clear
subsetgrid
STRSQL = "select * from tbl_user"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridadduser.TextMatrix(i, 0) = RS!u_id
gridadduser.TextMatrix(i, 1) = RS!u_name
gridadduser.TextMatrix(i, 2) = RS!usertype
gridadduser.Rows = gridadduser.Rows + 1
SLNO = SLNO + 1
RS.MoveNext
i = i + 1
Wend
End If
gridadduser.Rows = gridadduser.Rows - 1
End Sub
Private Sub gridadduser_click()
'subclearlabel

If gridadduser.Rows > 1 Then
cmd_update.Enabled = True
cmd_delete.Enabled = True
cmd_submit.Enabled = False
STRSQL = "select * from tbl_user where  U_Id = '" & gridadduser.TextMatrix(gridadduser.RowSel, 0) & "'"
Set RS = adocn.Execute(STRSQL)
txt_uid.Text = RS!u_id
txt_wname.Text = RS!u_name
txt_waddr.Text = RS!u_add
txt_cont.Text = RS!u_cont
txt_uname.Text = RS!usernam
txt_pword.Text = RS!password
Test = "Tester"
 If (RS!usertype = Test) Then
            Option1.Value = True
        Else
            Option2.Value = True
        End If
        End If
End Sub
Public Function check()
Dim l As Boolean
STRSQL = "select usernam from tbl_user where usernam='" & txt_uname & "'"
 Set RS = adocn.Execute(STRSQL)
 If RS.RecordCount > 0 Then

 l = True
 Else
 l = False
 End If
check = l
End Function
Public Sub clrlabel()
lbl_wnam.Visible = False
lbl_waddr.Visible = False
lbl_cont.Visible = False
lbl_unam.Visible = False
lbl_pword.Visible = False
End Sub
Public Function fnMobileValidation()
        Dim Mobile As String
        Mobile = txt_cont.Text
        Dim ok As Boolean
        If (IsNumeric(Mobile) And Len(Mobile) = "10") Then
            ok = True
            lbl_cont.Visible = False
        Else
            ok = False
            lbl_cont.Caption = "* Invalid Mobile Number"
            lbl_cont.Visible = True
        End If
        fnMobileValidation = ok
End Function


Private Sub txt_wname_KeyPress(KeyAscii As Integer)
If Len(txt_wname.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_wname.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub
