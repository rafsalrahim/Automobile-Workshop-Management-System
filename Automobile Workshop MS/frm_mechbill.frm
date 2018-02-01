VERSION 5.00
Begin VB.Form frm_mechbill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mechanical Repair Bill"
   ClientHeight    =   10770
   ClientLeft      =   5205
   ClientTop       =   750
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   495
      Left            =   8040
      TabIndex        =   20
      Top             =   9600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   6240
      TabIndex        =   19
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "HOTE WHEELS"
         BeginProperty Font 
            Name            =   "FoughtKnight"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   555
         Left            =   2640
         TabIndex        =   22
         Top             =   360
         Width           =   3600
      End
      Begin VB.Label lbl_company 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <company>"
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
         Left            =   7680
         TabIndex        =   21
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Line Line4 
         X1              =   1680
         X2              =   1680
         Y1              =   2760
         Y2              =   8280
      End
      Begin VB.Line Line3 
         X1              =   7200
         X2              =   7200
         Y1              =   2760
         Y2              =   8280
      End
      Begin VB.Label lbl_tot 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <tot amt>"
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
         Left            =   8400
         TabIndex        =   18
         Top             =   8400
         Width           =   1335
      End
      Begin VB.Label lbl_total 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total"
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
         Left            =   240
         TabIndex        =   17
         Top             =   8520
         Width           =   1335
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   9960
         Y1              =   8280
         Y2              =   8280
      End
      Begin VB.Label lbl_sparamt 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <spr amt>"
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
         Left            =   8400
         TabIndex        =   16
         Top             =   6360
         Width           =   1335
      End
      Begin VB.Label lbl_others 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <other amts>"
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
         Left            =   8400
         TabIndex        =   15
         Top             =   5400
         Width           =   1335
      End
      Begin VB.Label lbl_brake 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <break amt>"
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
         Left            =   8400
         TabIndex        =   14
         Top             =   4800
         Width           =   1335
      End
      Begin VB.Label lbl_oil 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <oil amt>"
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
         Left            =   8400
         TabIndex        =   13
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label lbl_eng 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <eng amt>"
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
         Left            =   8400
         TabIndex        =   12
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lbl_spare 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <spareparts >"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   3120
         TabIndex        =   11
         Top             =   6360
         Width           =   3495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Spareparts"
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
         TabIndex        =   10
         Top             =   6240
         Width           =   990
      End
      Begin VB.Label lbl_5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   5640
         Width           =   615
      End
      Begin VB.Label lbl_4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Brake"
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
         TabIndex        =   8
         Top             =   4920
         Width           =   555
      End
      Begin VB.Label lbl_3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Oil Change"
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
         TabIndex        =   7
         Top             =   4200
         Width           =   1020
      End
      Begin VB.Label lbl_2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Engin"
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
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lbl_date 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <date>"
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
         Left            =   7680
         TabIndex        =   5
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9840
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label lbl_regno 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <reg no>"
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
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lbl_addr 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <address>"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lbl_name 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <name>"
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
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "BILL FOR REPAIRE WORK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2760
         TabIndex        =   1
         Top             =   1440
         Width           =   3735
      End
   End
End
Attribute VB_Name = "frm_mechbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Dim v2 As Integer

Private Sub Command1_Click()
frm_mechbill.PrintForm
End Sub

Private Sub Form_Load()
LBL_DATE.Caption = DateValue(Now)
STRSQL = "select a.engine,a.brake,a.oil,a.others,a.tot_amt,a.spare, a.bas_amt,a.spr_amt,b.vh_id,b.name,b.status,b.address,b.mob_no,b.vh_compnam,b.regi_no,b.vh_id,b.vh_model from tbl_repbill a inner join tbl_vehicleregistration b on a.vh_id=b.vh_id where b.vh_id='" & VECH_ID & "' and b.status='FINISHED'"
'Set RS = adocn.Execute(STRSQL)
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
lbl_name.Caption = RS!Name
lbl_addr.Caption = RS!address
lbl_regno.Caption = RS!regi_no
lbl_company.Caption = RS!vh_compnam
lbl_spare.Caption = RS!spare
lbl_sparamt.Caption = RS!spr_amt
'lbl_spareamt.Caption = RS!bas_amt
lbl_tot.Caption = RS!tot_amt
fill2
End If
End Sub

Public Sub fill2()
s = 0
STRSQL = "select * from tbl_repare where vh_id= '" & VECH_ID & "' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
If RS!engine = "SELECTED" Then
v2 = 1
lbl_eng.Caption = rate1(v2)
Else
lbl_eng.Caption = " "
End If
If RS!brake = "SELECTED" Then
v2 = 2
lbl_brake.Caption = rate1(v2)
Else
lbl_brake.Caption = " "
End If
If RS!oil = "SELECTED" Then
v2 = 3
lbl_oil.Caption = rate1(v2)
Else
lbl_oil.Caption = " "
End If
If RS!others = "SELECTED" Then
v2 = 4
lbl_others.Caption = rate1(v2)

Else
lbl_others.Caption = " "
End If
        RS.MoveNext
    Loop

End Sub
Public Function rate1(v2 As Integer)
Dim s1 As Integer
Dim str As String
Dim RS1 As ADODB.Recordset
str = "select rate from tbl_reprate where r_id = '" & v2 & "' "
Set RS1 = adocn.Execute(str)
s1 = RS1!rate
rate1 = s1
RS1.Close
End Function
