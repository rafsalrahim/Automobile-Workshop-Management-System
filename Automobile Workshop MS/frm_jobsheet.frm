VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_jobsheet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Sheet"
   ClientHeight    =   9270
   ClientLeft      =   6915
   ClientTop       =   1230
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   8580
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
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
      Left            =   5760
      TabIndex        =   20
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   5760
         TabIndex        =   21
         Top             =   7560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Format          =   111214593
         CurrentDate     =   42330
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6120
         TabIndex        =   15
         Top             =   6240
         Width           =   615
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6120
         TabIndex        =   13
         Top             =   5520
         Width           =   615
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   6240
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   " "
         Height          =   495
         Left            =   2640
         TabIndex        =   9
         Top             =   5400
         Width           =   375
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delivery date"
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
         Left            =   360
         TabIndex        =   19
         Top             =   7680
         Width           =   1215
      End
      Begin VB.Label lbl_company 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   " <comp>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   780
      End
      Begin VB.Label lbl_regno 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   " <regno>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label vh_id1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "   <no>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7320
         TabIndex        =   16
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "w allignment"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   6240
         Width           =   1230
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Engine"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   12
         Top             =   5520
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "W service"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   6360
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Body"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   5520
         Width           =   510
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Complaint :"
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
         Left            =   360
         TabIndex        =   7
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label lbl_addr 
         BackColor       =   &H00FFFFFF&
         Caption         =   " <Adress>"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5280
         TabIndex        =   6
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Adress"
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
         TabIndex        =   5
         Top             =   3360
         Width           =   645
      End
      Begin VB.Label lbl_name 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  <name>"
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
         Left            =   5280
         TabIndex        =   4
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   " Name"
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
         TabIndex        =   3
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lbl_fdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  <f_date>"
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
         Left            =   5880
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "JOB SHEET"
         BeginProperty Font 
            Name            =   "Abaddon ll"
            Size            =   20.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2640
         TabIndex        =   1
         Top             =   840
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frm_jobsheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL1 As String
Dim RS1 As ADODB.Recordset

Private Sub Command1_Click()
Command1.Visible = False
frm_jobsheet.PrintForm
End Sub

Private Sub Form_Load()
lbl_fdate.Caption = DateValue(Now)
STRSQL1 = "select * from tbl_vehicleregistration where vh_id = '" & J_ID & "' "
'Set RS1 = adocn.Execute(STRSQL1)
Set RS1 = adocn.Execute(STRSQL1)
If RS1.RecordCount > 0 Then
lbl_name.Caption = RS1!Name
lbl_addr.Caption = RS1!address
lbl_regno.Caption = RS1!regi_no
lbl_company.Caption = RS1!vh_compnam
vh_id1.Caption = RS1!vh_id
End If
End Sub
