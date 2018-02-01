VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm_wstatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Worker Status"
   ClientHeight    =   5535
   ClientLeft      =   3525
   ClientTop       =   2025
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9135
   Begin VB.Frame Frame1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   3240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
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
         CurrentDate     =   42206
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
         Left            =   3120
         TabIndex        =   11
         Top             =   3960
         Width           =   1095
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
         Left            =   1440
         TabIndex        =   10
         Top             =   3960
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid griddistriuser 
         Height          =   1335
         Left            =   4920
         TabIndex        =   9
         Top             =   1200
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2355
         _Version        =   393216
         Enabled         =   0   'False
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
         Left            =   2160
         TabIndex        =   7
         Text            =   "__Select One__"
         Top             =   2640
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
         Left            =   2160
         TabIndex        =   5
         Top             =   1920
         Width           =   2055
      End
      Begin VB.ComboBox comb_uid 
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
         Left            =   2160
         TabIndex        =   3
         Text            =   "__Select One__"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label txt_date 
         AutoSize        =   -1  'True
         Caption         =   "Date"
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
         Left            =   600
         TabIndex        =   8
         Top             =   3360
         Width           =   435
      End
      Begin VB.Label Label4 
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
         Left            =   600
         TabIndex        =   6
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label3 
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
         Left            =   600
         TabIndex        =   4
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Worker ID"
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
         Left            =   600
         TabIndex        =   2
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "WORK DISTRIBUTION"
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
         Left            =   2160
         TabIndex        =   1
         Top             =   240
         Width           =   3120
      End
   End
End
Attribute VB_Name = "frm_wstatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Dim STRSQL1 As String
Dim RS1 As ADODB.Recordset
Dim tdate As Date
Dim SLNO As Integer

Private Sub cmd_submit_Click()
If validation Then
insert
MsgBox "Work Alloted"
subaddtogrid
Else
MsgBox "Field Required"
End If
End Sub

Public Sub combvhid()
tdate = DateValue(Now)
STRSQL = "select * from tbl_vehicleregistration where date = '" & tdate & "' and  status='REGISTERED' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb_vid.AddItem RS!vh_id
        RS.MoveNext
    Loop
End Sub



Private Sub comb_uid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_uid_Click()
subAutofill1
End Sub
Private Sub subAutofill1()
    STRSQL = "select u_name from tbl_user  where u_id= " _
             & " '" & comb_uid.List(comb_uid.ListIndex) & "'"
   Set RS = adocn.Execute(STRSQL)
    txt_name.Text = RS!u_name
 
End Sub

Private Sub subAutofill()
    STRSQL = "select date from tbl_vehicleregistration  where vh_id= " _
             & " '" & comb_vid.List(comb_vid.ListIndex) & "'"
   Set RS = adocn.Execute(STRSQL)
    DTPicker1.Value = RS!Date
 
End Sub


Private Sub comb_vid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_vid_Click()
subAutofill
End Sub

Private Sub Form_Load()
combuid
combvhid
subaddtogrid
End Sub
Public Sub combuid()
 STRSQL = "select * from tbl_user where usertype = 'Machanic' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb_uid.AddItem RS!u_id
        RS.MoveNext
    Loop
End Sub

Public Sub insert()
'String str="select * from tbl_wstatus where "
STRSQL = "insert into tbl_wstatus(u_id,name,vh_id,f_date,status) values ('" & comb_uid.List(comb_uid.ListIndex) & "','" & txt_name & "','" & comb_vid.List(comb_vid.ListIndex) & "','" & DTPicker1.Value & "','ASSIGNED')"
       Set RS = adocn.Execute(STRSQL)
STRSQL1 = "update tbl_vehicleregistration set  status= 'ON PROCESS' where vh_id='" & comb_vid.List(comb_vid.ListIndex) & "' "
Set RS1 = adocn.Execute(STRSQL1)
End Sub
Private Sub subsetgrid()
griddistriuser.Cols = 3
griddistriuser.Rows = 2
griddistriuser.FixedRows = 1
griddistriuser.TextMatrix(0, 1) = "NAME"
griddistriuser.TextMatrix(0, 2) = "STATUS"
griddistriuser.ColWidth(0) = 0
griddistriuser.ColWidth(1) = 1000
griddistriuser.ColWidth(2) = 1730
End Sub
Public Sub subaddtogrid()
subsetgrid
STRSQL = "select * from tbl_wstatus where status='ASSIGNED'"
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
i = 1
SLNO = 1
While Not RS.EOF
griddistriuser.TextMatrix(i, 0) = RS!ws
griddistriuser.TextMatrix(i, 1) = RS!Name
griddistriuser.TextMatrix(i, 2) = RS!Status
griddistriuser.Rows = griddistriuser.Rows + 1
SLNO = SLNO + 1
RS.MoveNext
i = i + 1
Wend
End If
griddistriuser.Rows = griddistriuser.Rows - 1
End Sub

Public Function validation()
Dim ok As Boolean
If (comb_uid = "__Select One__") Then
ok = False
Else
If (comb_vid = "__Select One__") Then
ok = False
Else
ok = True
End If
End If
validation = ok

End Function
