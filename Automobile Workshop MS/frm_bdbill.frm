VERSION 5.00
Begin VB.Form frm_bdbill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Body work Bill"
   ClientHeight    =   10620
   ClientLeft      =   4710
   ClientTop       =   1230
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10620
   ScaleWidth      =   14310
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   615
      Left            =   7800
      TabIndex        =   22
      Top             =   9960
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9735
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  HOTE WHEELS "
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
         Left            =   2160
         TabIndex        =   23
         Top             =   120
         Width           =   4095
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   9240
         Y1              =   7800
         Y2              =   7800
      End
      Begin VB.Label lbl_totamt 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <Tot_amt>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   21
         Top             =   8040
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   285
         Left            =   360
         TabIndex        =   20
         Top             =   7920
         Width           =   465
      End
      Begin VB.Label lbl_spareamt 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <spare amt>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   19
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Spare amount"
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
         Top             =   7200
         Width           =   1260
      End
      Begin VB.Label lbl_otheramt 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <other_amt>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   17
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label Label7 
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
         Left            =   360
         TabIndex        =   16
         Top             =   5040
         Width           =   615
      End
      Begin VB.Label lbl_paintamt 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <paint_amt>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   15
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Paint"
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
         TabIndex        =   14
         Top             =   4440
         Width           =   450
      End
      Begin VB.Label lbl_glassamt 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <glass_amt>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   13
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Glass Work"
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
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label lbl_bdamt 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <bd_amt>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   10
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label lbl_spare 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  <Spareparts>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3360
         TabIndex        =   9
         Top             =   5520
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "SPARE PARTS"
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
         TabIndex        =   8
         Top             =   5760
         Width           =   1500
      End
      Begin VB.Label lbl_bdwork 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Body work"
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
         Top             =   3000
         Width           =   1035
      End
      Begin VB.Label lbl_company 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <company>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         TabIndex        =   6
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lbl_regno 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <reg no>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9360
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label lbl_id 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  <address>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbl_name 
         BackColor       =   &H00FFFFFF&
         Caption         =   "    <name>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lbl_date 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <time>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7080
         TabIndex        =   2
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "BILL FOR BODY WORK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2760
         TabIndex        =   1
         Top             =   1920
         Width           =   3135
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
End
Attribute VB_Name = "frm_bdbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Dim v1 As Integer

Private Sub Command1_Click()
Command1.Visible = False
frm_bdbill.PrintForm
up
End Sub

Private Sub Form_Load()
LBL_DATE.Caption = DateValue(Now)
STRSQL = "select a.bdywork,a.paint,a.glass,a.others,a.tot_amt,a.spare, a.bas_amt,a.spr_amt,b.vh_id,b.name,b.status,b.address,b.mob_no,b.vh_compnam,b.regi_no,b.vh_id,b.vh_model from tbl_bdbill a inner join tbl_vehicleregistration b on a.vh_id=b.vh_id where b.vh_id='" & VECH_ID & "' and b.status='FINISHED'"
'Set RS = adocn.Execute(STRSQL)
Set RS = adocn.Execute(STRSQL)
If RS.RecordCount > 0 Then
lbl_name.Caption = RS!Name
lbl_id.Caption = RS!address
lbl_regno.Caption = RS!regi_no
lbl_company.Caption = RS!vh_compnam
lbl_spare.Caption = RS!spare
lbl_spareamt.Caption = RS!spr_amt
'lbl_spareamt.Caption = RS!bas_amt
lbl_totamt.Caption = RS!tot_amt
fill
End If
End Sub

Private Sub fill()
STRSQL = "select * from tbl_bodywork where vh_id= '" & VECH_ID & "' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        If RS!bdywork = "SELECTED" Then
v1 = 1
lbl_bdamt.Caption = rate(v1)
Else
lbl_bdamt.Caption = " "
End If
If RS!glass = "SELECTED" Then
v1 = 2
lbl_glassamt.Caption = rate(v1)
Else
lbl_glassamt.Caption = " "
End If
If RS!Paint = "SELECTED" Then
v1 = 3
lbl_paintamt.Caption = rate(v1)
Else
lbl_paintamt.Caption = " "
End If
If RS!others = "SELECTED" Then
v1 = 4
lbl_otheramt.Caption = rate(v1)
Else
lbl_otheramt.Caption = " "
End If
        RS.MoveNext
    Loop
    
'if Check1.Value=True
End Sub
Public Function rate(v1 As Integer)
Dim s1 As Integer
Dim str As String
Dim RS1 As ADODB.Recordset
str = "select rate from tbl_bdwrkrate where b_id = '" & v1 & "' "
Set RS1 = adocn.Execute(str)
s1 = RS1!rate
rate = s1
RS1.Close
End Function
Public Sub up()
Dim RS1 As ADODB.Recordset
Dim STRSQL1 As String
STRSQL1 = "update tbl_bodywork set status='BILL PRINTED'"
Set RS1 = adocn.Execute(STRSQL1)
End Sub

