VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_report 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Report"
   ClientHeight    =   5100
   ClientLeft      =   6270
   ClientTop       =   2205
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton Command1 
         Caption         =   "VIEW"
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
         Left            =   2520
         TabIndex        =   6
         Top             =   3480
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   2280
         Width           =   2295
         _ExtentX        =   4048
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
         CurrentDate     =   42240
      End
      Begin VB.ComboBox comb_report 
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
         ItemData        =   "frm_report.frx":0000
         Left            =   3360
         List            =   "frm_report.frx":0010
         TabIndex        =   3
         Text            =   "----Select One----"
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label3 
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
         TabIndex        =   4
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Report Type"
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
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "REPORT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frm_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub comb_report_Change()
MsgBox "Should Select One"
End Sub

Private Sub Command1_Click()
If (comb_report.List(comb_report.ListIndex) = "VEHICLE REGISTERED") Then
DataEnvironment1.veh_regi DTPicker1.Value
DataReport1.Show
ElseIf (comb_report.List(comb_report.ListIndex) = "VEHICLE DELEVERED") Then
DataEnvironment1.veh_dele DTPicker1.Value
DataReport2.Show
ElseIf (comb_report.List(comb_report.ListIndex) = "VEHICLE ON PROCESS") Then
DataEnvironment1.veh_onpro DTPicker1.Value
DataReport3.Show
ElseIf (comb_report.List(comb_report.ListIndex) = "VEHICLE PENDING") Then
DataEnvironment1.veh_pending DTPicker1.Value
DataReport4.Show
End If
End Sub

Private Sub Form_Load()
DTPicker1 = DateValue(Now)
End Sub

