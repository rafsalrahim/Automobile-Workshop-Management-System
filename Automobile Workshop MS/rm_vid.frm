VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_body 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Body Work Registration"
   ClientHeight    =   8070
   ClientLeft      =   3600
   ClientTop       =   1845
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   13320
   Begin VB.Frame fram_body 
      Height          =   7815
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   13095
      Begin VB.CheckBox Check4 
         Caption         =   "OTHER"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   26
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "PAINT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   25
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "GLASS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   24
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "BODY"
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
         Left            =   2640
         TabIndex        =   23
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   720
         TabIndex        =   17
         Top             =   6240
         Width           =   4575
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   2880
            TabIndex        =   21
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   110886913
            CurrentDate     =   42181
         End
         Begin VB.TextBox txt_fdate 
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
            Left            =   960
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "TO"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   20
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "FROM"
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
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "UPDATE"
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
         Left            =   3600
         TabIndex        =   15
         Top             =   7200
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid gridvhservice 
         Height          =   5175
         Left            =   5760
         TabIndex        =   14
         Top             =   1560
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   9128
         _Version        =   393216
      End
      Begin VB.TextBox txt_model 
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
         TabIndex        =   1
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txt_comp 
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
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
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
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   7200
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "CANCAL"
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
         TabIndex        =   10
         Top             =   7200
         Width           =   1215
      End
      Begin VB.TextBox txt_work 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   5160
         Width           =   2415
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
         Left            =   2640
         TabIndex        =   2
         Top             =   3480
         Width           =   2415
      End
      Begin VB.ComboBox comb_vid 
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
         Left            =   2640
         TabIndex        =   7
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lbl_bdw 
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
         Left            =   3480
         TabIndex        =   27
         Top             =   4920
         Width           =   1530
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Type of Work"
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
         Left            =   840
         TabIndex        =   22
         Top             =   4080
         Width           =   1305
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "VEHICLE ON SERVICE"
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
         Left            =   7320
         TabIndex        =   16
         Top             =   600
         Width           =   3150
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "MODEL"
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
         Left            =   840
         TabIndex        =   13
         Top             =   2760
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   " COMPANY"
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
         Left            =   840
         TabIndex        =   12
         Top             =   2160
         Width           =   1260
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Body Works"
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
         Left            =   840
         TabIndex        =   9
         Top             =   5280
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Registration NO."
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
         Left            =   840
         TabIndex        =   8
         Top             =   3480
         Width           =   1545
      End
      Begin VB.Label Label2 
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
         Left            =   960
         TabIndex        =   5
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "BODY WORKS"
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
         Left            =   1680
         TabIndex        =   4
         Top             =   600
         Width           =   2010
      End
   End
End
Attribute VB_Name = "frm_body"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Dim STRSQL1 As String
Dim RS1 As ADODB.Recordset
Dim tdate As Date
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Dim e As Date
Dim f As Date

Public Sub combofill()
tdate = DateValue(Now)
STRSQL = "select * from tbl_vehicleregistration where date = '" & tdate & "' and  status='REGISTERED' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb_vid.AddItem RS!vh_id
        RS.MoveNext
    Loop
End Sub

Private Sub comb_vid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_vid_Click()
subAutofill
End Sub

Private Sub Command1_Click()
subcleardata
MsgBox "Cancelled.. "
End Sub

Private Sub Command2_Click()
If fnValidation Then
STRSQL1 = "select * from tbl_bodywork where vh_id='" & comb_vid.List(comb_vid.ListIndex) & "'"
Set RS1 = adocn.Execute(STRSQL1)
If RS1.RecordCount < 1 Then
insert
MsgBox "inserted"
Unload Me
MDIForm1.Show
Else
MsgBox "Already Registered"
Unload Me
MDIForm1.Show
End If
Else
MsgBox "Not Registered"
End If
End Sub
Private Sub Command3_Click()
If fnValidation Then
If Check1.Value = 1 Then
a = "SELECTED"
Else
a = "NOT SELECTED"
End If
If Check2.Value = 1 Then
b = "SELECTED"
Else
b = "NOT SELECTED"
End If

If Check3.Value = 1 Then
c = "SELECTED"
Else
c = "NOT SELECTED"
End If
If Check4.Value = 1 Then
d = "SELECTED"
Else
d = "NOT SELECTED"
End If
e = DTPicker1.Value
STRSQL = "update tbl_bodywork set vh_compnam='" & txt_comp.Text & "',vh_model='" & txt_model.Text & "', " _
                     & "regi_no ='" & txt_rno.Text & "',description='" & txt_work.Text & "',status=' REGISTREAD ', " _
                     & " fdate='" & txt_fdate.Text & "',tdate='" & e & "gear='" & b & " engine='" & a & " oil='" & c & " others='" & d & "' where " _
                    & " UserId = '" & gridadduser.TextMatrix(gridadduser.RowSel, 0) & "'"
        adocn.Execute (STRSQL)
        MsgBox "User details updated . . ."
        subaddtogrid
        cmdSave.Enabled = True
        cmdUpdate.Enabled = False
        cmdDelete.Enabled = False
        subcleardata
        Else
        MsgBox "Field Needed"
        End If
End Sub

Private Sub Form_Load()
combofill
txt_fdate = DateValue(Now)
subaddtogrid
'subaddtogrid1
subcleardata
clearlabl
End Sub
Private Sub subAutofill()
    STRSQL = "select vh_compnam,vh_model,regi_no from tbl_vehicleregistration  where vh_id= " _
             & " '" & comb_vid.List(comb_vid.ListIndex) & "'"
   Set RS = adocn.Execute(STRSQL)
    txt_comp = RS!vh_compnam
    txt_model = RS!vh_model
    txt_rno = RS!regi_no
End Sub

Public Sub insert()
If Check1.Value = 1 Then
a = "SELECTED"
Else
a = "NOT SELECTED"
End If
If Check2.Value = 1 Then
b = "SELECTED"
Else
b = "NOT SELECTED"
End If

If Check3.Value = 1 Then
c = "SELECTED"
Else
c = "NOT SELECTED"
End If
If Check4.Value = 1 Then
d = "SELECTED"
Else
d = "NOT SELECTED"
End If
'e = DTPicker1.Value
f = txt_fdate
STRSQL = "insert into tbl_bodywork(bdywork,glass,paint,others,description,tdate,fdate,status,vh_compnam,vh_model,regi_no,vh_id)values " _
        & " ('" & a & "' , '" & b & "' , '" & c & "' , '" & d & "' , '" & txt_work.Text & "' , '" & DTPicker1.Value & "' , '" & f & "' , " _
        & " ' REGISTREAD ' ,'" & txt_comp & "','" & txt_model & "','" & txt_rno & "','" & comb_vid.List(comb_vid.ListIndex) & "')"
    Set RS = adocn.Execute(STRSQL)
    
End Sub
Public Sub subcleardata()
txt_comp = " "
txt_model = " "
txt_rno = " "
'Check1 = " "
'Check2 = ""
'Check3 = ""
'Check4 = ""
txt_work = ""
'combo_status = ""

End Sub
Public Sub billstart()
a = Check1.Value
b = Check2.Value
c = Check3.Value
d = Check4.Value
If a = True Then
STRSQL = " select * from tbl_rate where c_name = '" & a & "'"
 Set RS = adocn.Execute(STRSQL)
 s = s + RS!rate
 End If
 
End Sub
Private Sub subsetgrid()
gridvhservice.Cols = 3
gridvhservice.Rows = 2
gridvhservice.FixedRows = 1
gridvhservice.TextMatrix(0, 1) = "COMPANY"
gridvhservice.TextMatrix(0, 2) = "REGI. NO"
gridvhservice.ColWidth(0) = 0
gridvhservice.ColWidth(1) = 1000
gridvhservice.ColWidth(2) = 1730
End Sub
Public Sub subaddtogrid()
subsetgrid
STRSQL = "select * from tbl_bodywork"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridvhservice.TextMatrix(i, 0) = RS!cust_id2
gridvhservice.TextMatrix(i, 1) = RS!vh_compnam
gridvhservice.TextMatrix(i, 2) = RS!regi_no
gridvhservice.Rows = gridvhservice.Rows + 1
SLNO = SLNO + 1
RS.MoveNext
i = i + 1
Wend
End If
gridvhservice.Rows = gridvhservice.Rows - 1
End Sub

'Private Sub subsetgrid1()
'griddeleverd.Cols = 3
'griddeleverd.Rows = 2
'griddeleverd.FixedRows = 1
'griddeleverd.TextMatrix(0, 1) = "COMPANY"
'griddeleverd.TextMatrix(0, 2) = "REGI. NO"
'griddeleverd.ColWidth(0) = 0
'griddeleverd.ColWidth(1) = 1000
'griddeleverd.ColWidth(2) = 1730
'End Sub
'Public Sub subaddtogrid1()
'subsetgrid1
'STRSQL = "select * from tbl_bodywork where status= 'FINISHED' "
'Set RS = adocn.Execute(STRSQL)
'If RS.RecordCount > 0 Then
'i = 1
'SLNO = 1
'While Not RS.EOF
'griddeleverd.TextMatrix(i, 0) = RS!cust_id2
'griddeleverd.TextMatrix(i, 1) = RS!vh_compnam
'griddeleverd.TextMatrix(i, 2) = RS!regi_no
'griddeleverd.Rows = griddeleverd.Rows + 1
'SLNO = SLNO + 1
'RS.MoveNext
'i = i + 1
'Wend
'End If
'griddeleverd.Rows = griddeleverd.Rows - 1
'End Sub


'Private Sub griddeleverd_Click()
'If griddeleverd.Rows > 1 Then
''cmd_update.Enabled = True
''cmd_delete.Enabled = True
''cmd_submit.Enabled = False
'STRSQL = "select * from tbl_bodywork where  cust_id2 = '" & griddeleverd.TextMatrix(griddeleverd.RowSel, 0) & "'"
'Set RS = adocn.Execute(STRSQL)
'txt_cstatus.Text = RS!Status
'txt_comp.Text = RS!vh_compnam
'txt_model.Text = RS!vh_model
'txt_rno.Text = RS!regi_no
'txt_work.Text = RS!Description
'txt_fdate.Text = RS!fdate
'DTPicker1.Value = RS!tdate
'If RS!bdywork = "SELECTED" Then
'Check1.Value = 1
'Else
'Check1.Value = 0
'End If
'If RS!glass = "SELECTED" Then
'Check2.Value = 1
'Else
'Check2.Value = 0
'End If
'If RS!Paint = "SELECTED" Then
'Check3.Value = 1
'Else
'Check3.Value = 1
'End If
'If RS!others = "SELECTED" Then
'Check4.Value = 1
'Else
'Check4.Value = 1
'End If
'End If
'End Sub

Private Sub gridvhservice_Click()
If gridvhservice.Rows > 1 Then
'cmd_update.Enabled = True
'cmd_delete.Enabled = True
'cmd_submit.Enabled = False
STRSQL = "select * from tbl_bodywork where  cust_id2 = '" & gridvhservice.TextMatrix(gridvhservice.RowSel, 0) & "'"
Set RS = adocn.Execute(STRSQL)
'comb_vid.AddItem = RS!vh_id
'txt_cstatus.Text = RS!Status
txt_comp.Text = RS!vh_compnam
txt_model.Text = RS!vh_model
txt_rno.Text = RS!regi_no
txt_work.Text = RS!Description
txt_fdate.Text = RS!fdate
DTPicker1.Value = RS!tdate
If RS!bdywork = "SELECTED" Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If RS!glass = "SELECTED" Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If RS!Paint = "SELECTED" Then
Check3.Value = 1
Else
Check3.Value = 1
End If
If RS!others = "SELECTED" Then
Check4.Value = 1
Else
Check4.Value = 1
End If
End If
End Sub
Public Function fnValidation()
    Dim ok, ok1, ok2 As Boolean
    If Trim(txt_work.Text) = "" Then
        lbl_bdw.Visible = True
        ok1 = False
    Else
    ok1 = True
     End If
    If (Check1.Value = 1 Or Check2.Value = 1 Or Check3.Value = 1 Or Check4.Value = 1) Then
      ok2 = True
    Else
          ok2 = False
        MsgBox "Select Any Process"
    End If
    If (ok1 = True And ok2 = True) Then
    ok = True
    Else
    ok = False
    End If
    fnValidation = ok
End Function

Public Sub clearlabl()
lbl_bdw.Visible = False
End Sub

Private Sub txt_work_Change()
 If Trim(txt_work.Text) = "" Then
        lbl_bdw.Visible = True
    Else
        lbl_bdw.Visible = False
    End If
End Sub

Private Sub txt_work_KeyPress(KeyAscii As Integer)
If Len(txt_work.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_work.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub
