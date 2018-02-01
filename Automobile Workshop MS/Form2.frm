VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_wservice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2Water Service"
   ClientHeight    =   7740
   ClientLeft      =   3525
   ClientTop       =   1845
   ClientWidth     =   12060
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   12060
   Begin VB.Frame Frame1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin MSFlexGridLib.MSFlexGrid gridvhservice 
         Height          =   4935
         Left            =   5640
         TabIndex        =   21
         Top             =   1680
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   8705
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2520
         TabIndex        =   20
         Top             =   6120
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   55246849
         CurrentDate     =   42200
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
         Left            =   2520
         TabIndex        =   18
         Top             =   2880
         Width           =   2415
      End
      Begin VB.CheckBox Check3 
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
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   4680
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Chaise"
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
         TabIndex        =   16
         Top             =   4080
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Full Body "
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
         Left            =   2520
         TabIndex        =   15
         Top             =   4080
         Width           =   1215
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
         Left            =   2520
         TabIndex        =   14
         Text            =   "---NOT SELECT---"
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CommandButton cmd_submit 
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
         Left            =   1560
         TabIndex        =   12
         Top             =   6960
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancel 
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
         Left            =   3480
         TabIndex        =   11
         Top             =   6960
         Width           =   1215
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
         Left            =   2520
         TabIndex        =   8
         Top             =   5520
         Width           =   2415
      End
      Begin VB.TextBox txt_regino 
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
         TabIndex        =   5
         Top             =   3480
         Width           =   2415
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
         Height          =   375
         Left            =   2520
         TabIndex        =   3
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label11 
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
         Left            =   6480
         TabIndex        =   22
         Top             =   720
         Width           =   3150
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   480
         TabIndex        =   19
         Top             =   2880
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Vahile ID"
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
         TabIndex        =   13
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label7 
         Caption         =   "To Date"
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
         TabIndex        =   10
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
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
         TabIndex        =   9
         Top             =   5520
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Type of Service"
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
         TabIndex        =   7
         Top             =   4080
         Width           =   1440
      End
      Begin VB.Label Label4 
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
         Left            =   480
         TabIndex        =   6
         Top             =   3480
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Width           =   60
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   2
         Top             =   2400
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "WATER SERVICE"
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
         Left            =   1560
         TabIndex        =   1
         Top             =   720
         Width           =   2430
      End
   End
End
Attribute VB_Name = "frm_wservice"
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
Dim d As Integer
Dim e As Date
Dim cmdat As Date

Private Sub subAutofill()
    STRSQL = "select vh_compnam,vh_model,regi_no from tbl_vehicleregistration  where vh_id= " _
             & " '" & comb_vid.List(comb_vid.ListIndex) & "'"
   Set RS = adocn.Execute(STRSQL)
    txt_comp = RS!vh_compnam
    txt_model = RS!vh_model
    txt_regino = RS!regi_no
End Sub
Public Sub combofill()
t_date = DateValue(Now)
STRSQL = "select * from tbl_vehicleregistration where date = '" & t_date & "' and status='REGISTERED' "
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
Public Sub subcleardata()
txt_comp.Text = " "
txt_model.Text = " "
txt_regino.Text = " "
'txt_fdate.Text = " "
'txt_tdate.Text = " "
comb_vid.Text = "---NOT SELECT---"
End Sub
Private Sub cmd_submit_Click()
If val Then
STRSQL1 = "select * from tbl_water where vh_id='" & comb_vid.List(comb_vid.ListIndex) & "'"
Set RS1 = adocn.Execute(STRSQL1)
If RS1.RecordCount < 1 Then
insert
subaddtogrid
MsgBox "Inserted Successfully"
subcleardata
Unload Me
MDIForm1.Show
Else
MsgBox "Vehicle Already Registered"
Unload Me
MDIForm1.Show
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
e = DTPicker1.Value
 cmdat = CDate(txt_fdate)
STRSQL = "insert into tbl_water(vh_id,fulbdy,chase,Others,t_date,f_date,status,vh_compnam,vh_model,regi_no)values " _
        & " ('" & comb_vid.List(comb_vid.ListIndex) & "','" & a & "' , '" & b & "' , '" & c & "' , '" & e & "' , '" & cmdat & "' , " _
        & " 'REGESTERED' ,'" & txt_comp & "','" & txt_model & "','" & txt_regino & "')"
    Set RS = adocn.Execute(STRSQL)
End Sub

Private Sub Form_Load()
combofill
txt_fdate = DateValue(Now)
subaddtogrid
subcleardata
'subaddtogrid1
End Sub
Public Sub sumrate()
a = Check1.Value
b = Check2.Value
c = Check3.Value

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
STRSQL = "select * from tbl_water"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
gridvhservice.TextMatrix(i, 0) = RS!w_id
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
'STRSQL = "select * from tbl_water where status= 'FINISHED' "
'Set RS = adocn.Execute(STRSQL)
'If RS.RecordCount > 0 Then
'i = 1
'SLNO = 1
'While Not RS.EOF
'griddeleverd.TextMatrix(i, 0) = RS!w_id
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
'STRSQL = "select * from tbl_water where  w_id = '" & griddeleverd.TextMatrix(griddeleverd.RowSel, 0) & "'"
'Set RS = adocn.Execute(STRSQL)
''comb_vid.AddItem = RS!vh_id
'txt_comp.Text = RS!vh_compnam
'txt_model.Text = RS!vh_model
'txt_rno.Text = RS!regi_no
'txt_work.Text = RS!Description
'txt_fdate.Text = RS!f_date
'DTPicker1.Value = RS!t_date
'If RS!fulbdy = "SELECTED" Then
'Check1.Value = 1
'Else
'Check1.Value = 0
'End If
'If RS!chase = "SELECTED" Then
'Check2.Value = 1
'Else
'Check2.Value = 0
'End If
'If RS!others = "SELECTED" Then
'Check3.Value = 1
'Else
'Check3.Value = 1
'End If
'End If
'End Sub

Private Sub gridvhservice_Click()
If gridvhservice.Rows > 1 Then
'cmd_update.Enabled = True
'cmd_delete.Enabled = True
'cmd_submit.Enabled = False
STRSQL = "select * from tbl_water where  w_id = '" & gridvhservice.TextMatrix(gridvhservice.RowSel, 0) & "'"
Set RS = adocn.Execute(STRSQL)
'comb_vid.AddItem = RS!vh_id
txt_comp.Text = RS!vh_compnam
txt_model.Text = RS!vh_model
txt_rno.Text = RS!regi_no
txt_work.Text = RS!Description
txt_fdate.Text = RS!f_date
DTPicker1.Value = RS!t_date
If RS!fulbdy = "SELECTED" Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If RS!chase = "SELECTED" Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If RS!others = "SELECTED" Then
Check3.Value = 1
Else
Check3.Value = 1
End If
End If
End Sub
Public Function val()
Dim ok As Boolean
 If (Check1.Value = 1 Or Check2.Value = 1 Or Check3.Value = 1) Then
        ok = True
        
    Else
    ok = False
    MsgBox "Select Any Process"
    End If
    val = ok
End Function
