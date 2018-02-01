VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_billgen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Bill Generation"
   ClientHeight    =   8835
   ClientLeft      =   3975
   ClientTop       =   1965
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      Begin TabDlg.SSTab SSTab1 
         Height          =   7455
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   13150
         _Version        =   393216
         Tabs            =   4
         Tab             =   1
         TabsPerRow      =   4
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "BODY WORK"
         TabPicture(0)   =   "frm_billgen.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "MECHANICAL WORK"
         TabPicture(1)   =   "frm_billgen.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "WATER SERVICE"
         TabPicture(2)   =   "frm_billgen.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Image1"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame4"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "WHEEL ALLIGNMENT"
         TabPicture(3)   =   "frm_billgen.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame5"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Image4"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).ControlCount=   2
         Begin VB.Frame Frame5 
            Height          =   5295
            Left            =   -72480
            TabIndex        =   33
            Top             =   720
            Width           =   6015
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
               Height          =   375
               Left            =   2400
               TabIndex        =   60
               Top             =   1920
               Width           =   2055
            End
            Begin VB.CommandButton Command5 
               Caption         =   "Save"
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
               Left            =   2280
               TabIndex        =   40
               Top             =   4320
               Width           =   1455
            End
            Begin VB.TextBox txt_tot4 
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
               Height          =   495
               Left            =   2400
               TabIndex        =   39
               Top             =   3000
               Width           =   2055
            End
            Begin VB.ComboBox comb3_vid 
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
               Left            =   2400
               TabIndex        =   37
               Top             =   960
               Width           =   2175
            End
            Begin VB.Label Label10 
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
               TabIndex        =   59
               Top             =   1920
               Width           =   720
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Total amt."
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
               TabIndex        =   38
               Top             =   3120
               Width           =   915
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Vehcile ID"
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
               TabIndex        =   36
               Top             =   960
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            ForeColor       =   &H000000FF&
            Height          =   6855
            Left            =   -74760
            TabIndex        =   11
            Top             =   480
            Width           =   10335
            Begin VB.CommandButton Command7 
               Caption         =   "TOTAL"
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
               Left            =   7200
               TabIndex        =   54
               Top             =   6240
               Width           =   1215
            End
            Begin VB.TextBox txt_det1 
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
               Left            =   2400
               MultiLine       =   -1  'True
               TabIndex        =   53
               Top             =   3480
               Width           =   2775
            End
            Begin VB.TextBox txt_tot1 
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
               Left            =   8280
               TabIndex        =   51
               Top             =   4800
               Width           =   1215
            End
            Begin VB.TextBox txt_spare1 
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
               Left            =   5280
               TabIndex        =   49
               Top             =   4800
               Width           =   1095
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
               TabIndex        =   18
               Top             =   720
               Width           =   2775
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Body"
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
               Height          =   255
               Left            =   2520
               TabIndex        =   17
               Top             =   1920
               Width           =   975
            End
            Begin VB.CheckBox Check2 
               Caption         =   "Paint"
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
               Left            =   4200
               TabIndex        =   16
               Top             =   1920
               Width           =   1095
            End
            Begin VB.CheckBox Check3 
               Caption         =   "Glass"
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
               Height          =   255
               Left            =   2520
               TabIndex        =   15
               Top             =   2640
               Width           =   1215
            End
            Begin VB.CheckBox Check4 
               Caption         =   "Others"
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
               Height          =   255
               Left            =   4200
               TabIndex        =   14
               Top             =   2640
               Width           =   975
            End
            Begin VB.TextBox txt_basic1 
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
               Left            =   2400
               TabIndex        =   13
               Top             =   4800
               Width           =   1095
            End
            Begin VB.CommandButton Command1 
               Caption         =   "SAVE"
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
               Left            =   8640
               TabIndex        =   12
               Top             =   6240
               Width           =   1335
            End
            Begin VB.Image Image2 
               Height          =   3975
               Left            =   5520
               Picture         =   "frm_billgen.frx":0070
               Top             =   480
               Width           =   4650
            End
            Begin VB.Label lbl_spramt1 
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
               Left            =   5280
               TabIndex        =   56
               Top             =   4560
               Width           =   1530
            End
            Begin VB.Label lbl_spr1 
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
               Left            =   3600
               TabIndex        =   55
               Top             =   3240
               Width           =   1530
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Spareparts Used"
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
               TabIndex        =   52
               Top             =   3600
               Width           =   1530
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Total amt."
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
               Left            =   6840
               TabIndex        =   50
               Top             =   4800
               Width           =   915
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Spare amt."
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
               Left            =   3960
               TabIndex        =   48
               Top             =   4800
               Width           =   990
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "VEHICLE "
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
               TabIndex        =   21
               Top             =   720
               Width           =   1035
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "WORK"
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
               TabIndex        =   20
               Top             =   2280
               Width           =   735
            End
            Begin VB.Label Label4 
               Caption         =   "Basic amt"
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
               Left            =   600
               TabIndex        =   19
               Top             =   4920
               Width           =   1335
            End
         End
         Begin VB.Frame Frame3 
            Height          =   6735
            Left            =   240
            TabIndex        =   3
            Top             =   480
            Width           =   10095
            Begin VB.CommandButton Command6 
               Caption         =   "CALCULATE"
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
               Left            =   6360
               TabIndex        =   47
               Top             =   5760
               Width           =   1935
            End
            Begin VB.TextBox txt_det2 
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   2760
               MultiLine       =   -1  'True
               TabIndex        =   46
               Top             =   3360
               Width           =   3015
            End
            Begin VB.TextBox txt_tot2 
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
               Left            =   8160
               TabIndex        =   44
               Top             =   4800
               Width           =   1215
            End
            Begin VB.TextBox txt_spare2 
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
               Left            =   5400
               TabIndex        =   42
               Top             =   4800
               Width           =   1095
            End
            Begin VB.CommandButton Command2 
               Caption         =   "SAVE"
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
               Left            =   8520
               TabIndex        =   25
               Top             =   5760
               Width           =   1215
            End
            Begin VB.CheckBox Check8 
               Caption         =   "Others"
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
               Height          =   255
               Left            =   4680
               TabIndex        =   23
               Top             =   2640
               Width           =   975
            End
            Begin VB.TextBox txt_basic2 
               Enabled         =   0   'False
               Height          =   375
               Left            =   2640
               TabIndex        =   8
               Top             =   4800
               Width           =   975
            End
            Begin VB.CheckBox Check5 
               Caption         =   "Engine"
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
               Left            =   2880
               TabIndex        =   7
               Top             =   1800
               Width           =   1095
            End
            Begin VB.CheckBox Check6 
               Caption         =   "Break"
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
               Left            =   4680
               TabIndex        =   6
               Top             =   1800
               Width           =   975
            End
            Begin VB.CheckBox Check7 
               Caption         =   "Oil change"
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
               Left            =   2880
               TabIndex        =   5
               Top             =   2640
               Width           =   1575
            End
            Begin VB.ComboBox comb1_vid 
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
               Left            =   2760
               TabIndex        =   4
               Top             =   960
               Width           =   2895
            End
            Begin VB.Image Image3 
               Height          =   3975
               Left            =   6000
               Picture         =   "frm_billgen.frx":9F19
               Top             =   480
               Width           =   3780
            End
            Begin VB.Label lbl_spramt2 
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
               Left            =   5400
               TabIndex        =   58
               Top             =   4560
               Width           =   1530
            End
            Begin VB.Label lbl_spr2 
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
               Left            =   4200
               TabIndex        =   57
               Top             =   3120
               Width           =   1530
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "Spareparts used"
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
               TabIndex        =   45
               Top             =   3480
               Width           =   1470
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Total amt."
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
               Left            =   6840
               TabIndex        =   43
               Top             =   4800
               Width           =   915
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Spare amt"
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
               Left            =   4080
               TabIndex        =   41
               Top             =   4800
               Width           =   930
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Basic amt."
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
               Left            =   720
               TabIndex        =   24
               Top             =   4800
               Width           =   1170
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Vehicle id"
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
               Left            =   600
               TabIndex        =   10
               Top             =   960
               Width           =   1110
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Type"
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
               Left            =   720
               TabIndex        =   9
               Top             =   2280
               Width           =   585
            End
         End
         Begin VB.Frame Frame4 
            Height          =   5055
            Left            =   -72720
            TabIndex        =   2
            Top             =   720
            Width           =   6375
            Begin VB.TextBox txt_tot3 
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
               TabIndex        =   35
               Top             =   3480
               Width           =   1215
            End
            Begin VB.CommandButton Command3 
               Caption         =   "SAVE"
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
               Left            =   1920
               TabIndex        =   32
               Top             =   4320
               Width           =   1215
            End
            Begin VB.CheckBox Check11 
               Caption         =   "Others"
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
               Left            =   3360
               TabIndex        =   31
               Top             =   2760
               Width           =   975
            End
            Begin VB.CheckBox Check10 
               Caption         =   "Chase"
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
               Left            =   4320
               TabIndex        =   30
               Top             =   1920
               Width           =   1335
            End
            Begin VB.CheckBox Check9 
               Caption         =   "Full Body"
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
               Height          =   255
               Left            =   2760
               TabIndex        =   29
               Top             =   2040
               Width           =   1215
            End
            Begin VB.ComboBox comb2_vid 
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
               Left            =   2760
               TabIndex        =   27
               Top             =   1080
               Width           =   2655
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Total amt."
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
               TabIndex        =   34
               Top             =   3480
               Width           =   915
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Type"
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
               Left            =   600
               TabIndex        =   28
               Top             =   2280
               Width           =   585
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Vehicle id"
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
               TabIndex        =   26
               Top             =   1080
               Width           =   1110
            End
         End
         Begin VB.Image Image4 
            Height          =   16200
            Left            =   -75000
            Picture         =   "frm_billgen.frx":11DBA
            Top             =   360
            Width           =   28800
         End
         Begin VB.Image Image1 
            Height          =   21600
            Left            =   -75000
            Picture         =   "frm_billgen.frx":57725
            Top             =   360
            Width           =   38400
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "BILL GENERATION"
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
         Left            =   4200
         TabIndex        =   22
         Top             =   240
         Width           =   2475
      End
   End
End
Attribute VB_Name = "frm_billgen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STRSQL As String
Dim RS As ADODB.Recordset
Dim s As Integer
Dim v1 As Integer
Dim v2 As Integer
Dim v3 As Integer
Dim v4 As String
Dim a As String
Dim b As String
Dim c As String
Dim d As String
Public Sub combofill2()
STRSQL = "select * from tbl_bodywork where status='FINISHED' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb_vid.AddItem RS!vh_id
        RS.MoveNext
    Loop

End Sub


Private Sub fill()
s = 0
STRSQL = "select * from tbl_bodywork where vh_id= '" & comb_vid.List(comb_vid.ListIndex) & "' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        If RS!bdywork = "SELECTED" Then
Check1.Value = 1
v1 = 1
s = s + rate(v1)
Else
Check1.Value = 0
End If
If RS!glass = "SELECTED" Then
Check2.Value = 1
v1 = 2
s = s + rate(v1)
Else
Check2.Value = 0
End If
If RS!Paint = "SELECTED" Then
Check3.Value = 1
v1 = 3
s = s + rate(v1)
Else
Check3.Value = 0
End If
If RS!others = "SELECTED" Then
Check4.Value = 1
v1 = 4
s = s + rate(v1)
Else
Check4.Value = 0
End If
        RS.MoveNext
    Loop
txt_basic1.Text = s
    
'if Check1.Value=True
End Sub

Private Sub comb_vid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb_vid_Click()
fill
End Sub

Private Sub comb1_vid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb1_vid_Click()
fill2
End Sub

Private Sub comb2_vid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb2_vid_Click()
fill3
End Sub

Private Sub comb3_vid_Change()
MsgBox "Should Select One"
End Sub

Private Sub comb3_vid_Click()
fillwheel
End Sub

Private Sub Command1_Click()
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
STRSQL = "insert into tbl_bdbill(vh_id,spare,bas_amt,spr_amt,tot_amt, " _
         & " bdywork,paint,glass,others) values ('" & comb_vid.List(comb_vid.ListIndex) & "','" & txt_det1.Text & "','" & txt_basic1.Text & "' ," _
         & " '" & txt_spare1.Text & "','" & txt_tot1.Text & "','" & a & "','" & b & "'," _
         & " '" & c & "','" & d & "')"

Set RS = adocn.Execute(STRSQL)
MsgBox "Datas insearted"
Else
MsgBox "Field Required"
End If
End Sub
Private Sub Command2_Click()
If fnvalidation1 Then
If Check5.Value = 1 Then
a = "SELECTED"
Else
a = "NOT SELECTED"
End If
If Check6.Value = 1 Then
b = "SELECTED"
Else
b = "NOT SELECTED"
End If

If Check7.Value = 1 Then
c = "SELECTED"
Else
c = "NOT SELECTED"
End If
If Check8.Value = 1 Then
d = "SELECTED"
Else
d = "NOT SELECTED"
End If
STRSQL = "insert into tbl_repbill(vh_id,spare,bas_amt,spr_amt,tot_amt, " _
         & " engine,brake,oil,others) values ('" & comb1_vid.List(comb1_vid.ListIndex) & "','" & txt_det2.Text & "','" & txt_basic2.Text & "' ," _
         & " '" & txt_spare2.Text & "','" & txt_tot2.Text & "','" & a & "','" & b & "'," _
         & " '" & c & "','" & d & "')"

Set RS = adocn.Execute(STRSQL)
MsgBox "Datas insearted"
Else
MsgBox "Field Required"
End If
End Sub

Private Sub Command3_Click()
If fnvalidation3 Then
If Check9.Value = 1 Then
a = "SELECTED"
Else
a = "NOT SELECTED"
End If
If Check10.Value = 1 Then
b = "SELECTED"
Else
b = "NOT SELECTED"
End If

If Check11.Value = 1 Then
c = "SELECTED"
Else
c = "NOT SELECTED"
End If
STRSQL = "insert into tbl_watere(vh_id,tot_amt,fulbdy,chase,others) values ('" & comb2_vid.List(comb2_vid.ListIndex) & "','" & txt_tot3.Text & "','" & a & "','" & b & "'," _
         & " '" & c & "')"

Set RS = adocn.Execute(STRSQL)
MsgBox "Datas insearted"
Else
MsgBox "Field Required"
End If
End Sub

Private Sub Command4_Click()
txt_tot3.Text = s + val(txt_spare3.Text)
End Sub

Private Sub Command5_Click()
If fnvalidation4 Then
STRSQL = "insert into tbl_wheelbill(vh_id,amt) values ('" & comb3_vid.List(comb3_vid.ListIndex) & "','" & txt_tot4.Text & "')"
Set RS = adocn.Execute(STRSQL)
MsgBox "Datas insearted"
Else
MsgBox "Field Required"
End If
End Sub

Private Sub Command6_Click()
If (txt_basic2.Text = "" Or txt_spare2.Text = "") Then
MsgBox "Need Rate To Calculate"
Else
txt_tot2.Text = s + val(txt_spare2.Text)
End If
End Sub

Private Sub Command7_Click()
If (txt_basic1.Text = "" Or txt_spare1.Text = "") Then
MsgBox "Need Rate To Calculate"
Else
txt_tot1.Text = s + val(txt_spare1.Text)
End If
End Sub

Private Sub Form_Load()
combofill2
combofill1
combofill
combofill0
clear
End Sub
Public Sub bdy()
STRSQL = "select * from tbl_bdwrkrate where vh_id= '" & comb_vid.List(comb_vid.ListIndex) & "' "
Set RS = adocn.Execute(STRSQL)
End Sub
Public Sub combofill1()
STRSQL = "select * from tbl_repare where status='FINISHED' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb1_vid.AddItem RS!vh_id
        RS.MoveNext
    Loop
End Sub
Public Sub fill2()
s = 0
STRSQL = "select * from tbl_repare where vh_id= '" & comb1_vid.List(comb1_vid.ListIndex) & "' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
If RS!engine = "SELECTED" Then
Check5.Value = 1
v2 = 1
s = s + rate1(v2)
Else
Check5.Value = 0
End If
If RS!brake = "SELECTED" Then
Check6.Value = 1
v2 = 2
s = s + rate1(v2)
Else
Check6.Value = 0
End If
If RS!oil = "SELECTED" Then
Check7.Value = 1
v2 = 3
s = s + rate1(v2)
Else
Check7.Value = 0
End If
If RS!others = "SELECTED" Then
v2 = 4
s = s + rate1(v2)
Check8.Value = 1
Else
Check8.Value = 0
End If
        RS.MoveNext
    Loop
txt_basic2.Text = s
End Sub
Private Sub fill3()
s = 0
STRSQL = "select * from tbl_water where vh_id= '" & comb2_vid.List(comb2_vid.ListIndex) & "' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
If RS!fulbdy = "SELECTED" Then
Check9.Value = 1
v3 = 1
s = s + rate3(v3)
Else
Check9.Value = 0
End If
If RS!chase = "SELECTED" Then
Check10.Value = 1
v3 = 2
s = s + rate3(v3)
Else
Check10.Value = 0
End If
If RS!others = "SELECTED" Then
Check11.Value = 1
v3 = 3
s = s + rate3(v3)
Else
Check11.Value = 0
End If
        RS.MoveNext
    Loop
txt_tot3.Text = s
End Sub
Public Sub combofill()
STRSQL = "select * from tbl_water where status='FINISHED' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb2_vid.AddItem RS!vh_id
        RS.MoveNext
    Loop
End Sub
'Private Sub fill1()
'STRSQL = "select * from tbl_water where vh_id= '" & comb2_vid.List(comb2_vid.ListIndex) & "' "
'Set RS = adocn.Execute(STRSQL)
' Do While Not RS.EOF
'        If RS!fulbdy = "SELECTED" Then
'Check9.Value = 1
'Else
'Check9.Value = 0
'End If
'If RS!chase = "SELECTED" Then
'Check10.Value = 1
'Else
'Check10.Value = 0
'End If
'If RS!others = "SELECTED" Then
'Check11.Value = 1
'Else
'Check11.Value = 0
'End If
'        RS.MoveNext
'    Loop
'
'End Sub
Public Sub combofill0()
STRSQL = "select * from tbl_wheel where status='FINISHED' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
        comb3_vid.AddItem RS!vh_id
        RS.MoveNext
    Loop
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

Public Function rate3(v3 As Integer)
Dim s1 As Integer
Dim str As String
Dim RS1 As ADODB.Recordset
str = "select rate from tbl_waterrate where wr_id = '" & v3 & "' "
Set RS1 = adocn.Execute(str)
s1 = RS1!rate
rate3 = s1
RS1.Close
End Function
'Public Function rate4(v4 As String)
'Dim s1 As Integer
'Dim str As String
'Dim RS1 As ADODB.Recordset
'str = "select rate from tbl_wheelrate where catagory = '" & v4 & "' "
'Set RS1 = adocn.Execute(str)
's1 = RS1!rate
'rate = s1
'RS1.Close
'End Function

Public Sub fillwheel()
Dim str As String
Dim RS1 As ADODB.Recordset
STRSQL = "select * from tbl_wheel where vh_id= '" & comb3_vid.List(comb3_vid.ListIndex) & "' "
Set RS = adocn.Execute(STRSQL)
 Do While Not RS.EOF
txt_model.Text = RS!vh_model
v4 = RS!vh_model
str = "select rate from tbl_wheelrate where service = '" & v4 & "' "
Set RS1 = adocn.Execute(str)
Do While Not RS1.EOF
txt_tot4.Text = RS1!rate
   RS1.MoveNext
Loop
RS1.Close
'         txt_tot4.Text = rate(v4)
        RS.MoveNext
    Loop
 
End Sub
Public Sub clear()
lbl_spr1.Visible = False
lbl_spramt1.Visible = False
lbl_spr2.Visible = False
lbl_spramt2.Visible = False
End Sub
Public Function fnValidation()
Dim ok As Boolean
If (Trim(txt_det1.Text) = "") Then
   lbl_spr1.Visible = True
ok = False
Else
 If (Trim(txt_spare1.Text) = "") Then
  lbl_spramt1.Visible = True
  ok = False
  Else
  ok = True
  End If
  End If
  fnValidation = ok
End Function



Private Sub txt_det1_Change()
If Trim(txt_det1.Text) = "" Then
        lbl_spr1.Visible = True
    Else
        lbl_spr1.Visible = False
    End If
End Sub

Private Sub txt_det1_KeyPress(KeyAscii As Integer)
If Len(txt_det1.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_det1.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_det2_Change()
If Trim(txt_det2.Text) = "" Then
        lbl_spr2.Visible = True
    Else
        lbl_spr2.Visible = False
    End If
End Sub

Private Sub txt_det2_KeyPress(KeyAscii As Integer)
If Len(txt_det2.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_det2.Text = ""
End If
If KeyAscii > 46 And KeyAscii < 56 Then
MsgBox "Only character is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_spare1_Change()
If Trim(txt_spare1.Text) = "" Then
        lbl_spramt1.Visible = True
    Else
        lbl_spramt1.Visible = False
    End If
End Sub

Private Sub txt_spare1_KeyPress(KeyAscii As Integer)
If Len(txt_spare1.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_spare1.Text = ""
End If
If KeyAscii > 48 And KeyAscii > 58 And KeyAscii <> 8 And KeyAscii <> 13 Then
MsgBox "Only integer is allowed"
KeyAscii = 0
End If
End Sub

Private Sub txt_spare2_Change()
If Trim(txt_spare2.Text) = "" Then
        lbl_spramt2.Visible = True
    Else
        lbl_spramt2.Visible = False
    End If
End Sub
Public Function fnvalidation1()
Dim ok1 As Boolean
If (Trim(txt_det2.Text) = "") Then
   lbl_spr2.Visible = True
ok1 = False
Else
 If (Trim(txt_spare2.Text) = "") Then
  lbl_spramt2.Visible = True
  ok1 = False
  Else
  ok1 = True
  End If
  End If
  fnvalidation1 = ok1
End Function
Public Function fnvalidation3()
Dim ok1 As Boolean
If (Trim(txt_tot3.Text) = "") Then
ok1 = False
Else
 If (Check9.Value = 1 Or Check10.Value = 1 Or Check11.Value = 1) Then
        ok1 = True
        
    Else
    ok1 = False

  End If
  End If
  fnvalidation3 = ok1
End Function
Public Function fnvalidation4()
Dim ok1 As Boolean
If (Trim(txt_model.Text) = "") Then
ok1 = False
Else
 If (Trim(txt_tot4.Text) = "") Then
  ok1 = False
  Else
  ok1 = True
  End If
  End If
  fnvalidation4 = ok1
End Function

Private Sub txt_spare2_KeyPress(KeyAscii As Integer)
If Len(txt_spare2.Text) = 50 Then
MsgBox "Limit Crossed", vhinformation
txt_spare2.Text = ""
End If
If KeyAscii > 48 And KeyAscii > 58 And KeyAscii <> 8 And KeyAscii <> 13 Then
MsgBox "Only integer is allowed"
KeyAscii = 0
End If
End Sub
