VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_mrepair 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mechanical Repair"
   ClientHeight    =   7770
   ClientLeft      =   2790
   ClientTop       =   1650
   ClientWidth     =   15345
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fram_mrepair 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   1200
         TabIndex        =   21
         Top             =   5760
         Width           =   5055
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   3000
            TabIndex        =   25
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
            Format          =   110559233
            CurrentDate     =   42182
         End
         Begin VB.TextBox txt_fdate 
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
            Left            =   840
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "To"
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
            Left            =   2400
            TabIndex        =   24
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "From"
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
            TabIndex        =   22
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         TabIndex        =   20
         Top             =   6720
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid gridvhservice 
         Height          =   5535
         Left            =   7560
         TabIndex        =   19
         Top             =   1080
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   9763
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
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Others"
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
         Left            =   4680
         TabIndex        =   18
         Top             =   4080
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Break "
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
         Left            =   3000
         TabIndex        =   17
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Oil change"
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
         Left            =   4680
         TabIndex        =   16
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Engine"
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
         Left            =   3000
         TabIndex        =   15
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txt_comp 
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
         Left            =   3000
         TabIndex        =   13
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txt_model 
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
         Left            =   3000
         TabIndex        =   11
         Top             =   2280
         Width           =   3255
      End
      Begin VB.CommandButton cmd_submit 
         Caption         =   "SUBMIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancel 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   8
         Top             =   6720
         Width           =   1335
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
         Height          =   735
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   4800
         Width           =   3255
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
         Left            =   3000
         TabIndex        =   5
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txt_rno 
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
         Left            =   3000
         TabIndex        =   4
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label lbl_mwrk 
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
         Left            =   4560
         TabIndex        =   27
         Top             =   4560
         Width           =   1530
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "VEHICLE REGESTERED"
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
         Left            =   8520
         TabIndex        =   26
         Top             =   360
         Width           =   3360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Type of  Work"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   3600
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle company"
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
         Left            =   1080
         TabIndex        =   12
         Top             =   1680
         Width           =   1560
      End
      Begin VB.Label lbl_vmodel 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle model"
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
         Left            =   1080
         TabIndex        =   10
         Top             =   2280
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Works"
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
         Left            =   1080
         TabIndex        =   6
         Top             =   4920
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Registration No."
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
         Left            =   1080
         TabIndex        =   3
         Top             =   2880
         Width           =   1485
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "MECHANICAL REPAIR WORKS"
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
         Left            =   1965
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frm_mrepair"
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
STRSQL = "insert into tbl_repare(vh_id,engine,brake,oil,others,discription,frm_date,to_date,status,vh_company,vh_model,regi_no)values " _
        & " ('" & comb_vid.List(comb_vid.ListIndex) & "','" & a & "' , '" & b & "' , '" & c & "' , '" & d & "' , '" & txt_work.Text & "' , '" & txt_fdate & "' , '" & e & "' , " _
        & " 'REGESTERED' ,'" & txt_comp & "','" & txt_model & "','" & txt_rno & "')"
    Set RS = adocn.Execute(STRSQL)
End Sub
Public Sub combofill()
frm_date = DateValue(Now)
STRSQL = "select * from tbl_vehicleregistration where date = '" & frm_date & "' and status='REGISTERED' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb_vid.AddItem RS!vh_id
        RS.MoveNext
    Loop
End Sub

Private Sub cmd_cancel_Click()
subcleardata
MsgBox "Cancelled.. "
End Sub

Private Sub cmd_submit_Click()
If fnValidation Then
STRSQL1 = "select * from tbl_repare where vh_id='" & comb_vid.List(comb_vid.ListIndex) & "'"
Set RS1 = adocn.Execute(STRSQL1)
If RS1.RecordCount < 1 Then
insert
subaddtogrid
MsgBox "Registered Succesfully"
Unload Me
MDIForm1.Show
Else
MsgBox "Vehicle Already Registered"
End If
Else
MsgBox "A problem occured while Registration process!!"
End If
End Sub

Private Sub comb_vid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_vid_Click()
subAutofill
End Sub

Private Sub Command1_Click()
If fnValidation Then
a = Check1.Value
b = Check2.Value
c = Check3.Value
d = Check4.Value
e = DTPicker1.Value
STRSQL = "update tbl_repare set vh_compnam='" & txt_comp.Text & "',vh_model='" & txt_model.Text & "', " _
                     & "regi_no ='" & txt_rno.Text & "',description='" & txt_work.Text & "',status='REGISTREAD', " _
                     & " fdate='" & txt_fdate.Text & "',tdate='" & e & "paint='" & b & " bdywork='" & a & " glass='" & c & " others='" & d & "' where " _
                    & " UserId = '" & gridadduser.TextMatrix(gridadduser.RowSel, 0) & "'"
        adocn.Execute (STRSQL)
        MsgBox "User details updated . . ."
        subaddtogrid
        'cmdSave.Enabled = True
        'cmdUpdate.Enabled = False
        'cmdDelete.Enabled = False
        subcleardata
        Else
        MsgBox "Field Required"
        End If
End Sub


Private Sub Form_Load()
combofill
txt_fdate.Text = DateValue(Now)
subaddtogrid
subcleardata
clear
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
'txt_fdate = ""
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
STRSQL = "select * from tbl_repare where status= 'REGESTERED'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridvhservice.TextMatrix(i, 0) = RS!rid
gridvhservice.TextMatrix(i, 1) = RS!vh_company
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
'STRSQL = "select * from tbl_repare where status= 'FINISHED' "
'Set RS = adocn.Execute(STRSQL)
'If RS.RecordCount > 0 Then
'i = 1
'SLNO = 1
'While Not RS.EOF
'griddeleverd.TextMatrix(i, 0) = RS!rid
'griddeleverd.TextMatrix(i, 1) = RS!vh_company
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
'STRSQL = "select * from tbl_repare where  rid = '" & griddeleverd.TextMatrix(griddeleverd.RowSel, 0) & "'"
'Set RS = adocn.Execute(STRSQL)
''comb_vid.AddItem = RS!vh_id
'txt_comp.Text = RS!vh_compnam
'txt_model.Text = RS!vh_model
'txt_rno.Text = RS!regi_no
'txt_work.Text = RS!Discription
'txt_fdate.Text = RS!frm_date
'DTPicker1.Value = RS!to_date
'If RS!engine = "SELECTED" Then
'Check1.Value = 1
'Else
'Check1.Value = 0
'End If
'If RS!oil = "SELECTED" Then
'Check2.Value = 1
'Else
'Check2.Value = 0
'End If
'If RS!brake = "SELECTED" Then
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
STRSQL = "select * from tbl_repare where  rid = '" & gridvhservice.TextMatrix(gridvhservice.RowSel, 0) & "'"
Set RS = adocn.Execute(STRSQL)
'comb_vid.AddItem = RS!vh_id
txt_comp.Text = RS!vh_company
txt_model.Text = RS!vh_model
txt_rno.Text = RS!regi_no
txt_work.Text = RS!Discription
txt_fdate.Text = RS!frm_date
DTPicker1.Value = RS!to_date
If RS!engine = "SELECTED" Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If RS!oil = "SELECTED" Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If RS!brake = "SELECTED" Then
Check3.Value = 1
Else
Check3.Value = 0
End If
If RS!others = "SELECTED" Then
Check4.Value = 1
Else
Check4.Value = 0
End If
End If
End Sub

Private Sub txt_work_Change()
If Trim(txt_work.Text) = "" Then
        lbl_mwrk.Visible = True
    Else
        lbl_mwrk.Visible = False
    End If
End Sub

Public Function fnValidation()
    Dim ok1, ok2, ok As Boolean
    If Trim(txt_work.Text) = "" Then
        lbl_mwrk.Visible = True
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
    'End If
    fnValidation = ok
End Function
Public Sub clear()
  lbl_mwrk.Visible = False
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

