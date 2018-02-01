VERSION 5.00
Begin VB.Form frm_sparepartsreq 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   6990
   ClientTop       =   2205
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fram_workers 
      Height          =   7335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6015
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
         Height          =   495
         Left            =   3240
         TabIndex        =   15
         Top             =   6120
         Width           =   1215
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
         Height          =   495
         Left            =   1320
         TabIndex        =   14
         Top             =   6120
         Width           =   1335
      End
      Begin VB.TextBox txt_vmodel 
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox txt_vcomp 
         Height          =   375
         Left            =   2880
         TabIndex        =   11
         Top             =   3000
         Width           =   2415
      End
      Begin VB.TextBox txt_request 
         Height          =   1215
         Left            =   2880
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   4200
         Width           =   2415
      End
      Begin VB.TextBox txt_wname 
         Height          =   375
         Left            =   2880
         TabIndex        =   7
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txt_wid 
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txt_vid 
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label7 
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
         Left            =   720
         TabIndex        =   12
         Top             =   3600
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Vehicle Company"
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
         Left            =   720
         TabIndex        =   10
         Top             =   2880
         Width           =   1620
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Request"
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
         Left            =   720
         TabIndex        =   8
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label4 
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
         Left            =   720
         TabIndex        =   6
         Top             =   1920
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Worker name"
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
         Left            =   720
         TabIndex        =   4
         Top             =   2400
         Width           =   1260
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
         Left            =   720
         TabIndex        =   2
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "WORKERS REQUEST"
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
         Left            =   1440
         TabIndex        =   1
         Top             =   240
         Width           =   2955
      End
   End
End
Attribute VB_Name = "frm_sparepartsreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
