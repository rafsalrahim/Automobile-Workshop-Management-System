VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10650
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   20250
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.Menu mnu_service 
      Caption         =   "SERVICE"
      Begin VB.Menu mnu_mach 
         Caption         =   "MACHANICAL SERVICE"
      End
      Begin VB.Menu mnu_body 
         Caption         =   "BODY SERVICE"
      End
      Begin VB.Menu mnu_water 
         Caption         =   "WATER SERVICE"
      End
      Begin VB.Menu mnu_wheel 
         Caption         =   "WHEEL ALLAINMENT"
      End
   End
   Begin VB.Menu mnu_tester 
      Caption         =   "TESTER"
      Begin VB.Menu mnu_regi 
         Caption         =   "REGISTRATION"
      End
      Begin VB.Menu mnu_statup 
         Caption         =   "UPDATE STATUS"
      End
      Begin VB.Menu mnu_chagep 
         Caption         =   "CHANGE PASSWORD"
      End
      Begin VB.Menu mnu_bill 
         Caption         =   "BILL PRINT"
      End
      Begin VB.Menu mnu_delevary 
         Caption         =   "DELEVARY STATUS"
      End
   End
   Begin VB.Menu mnu_admin 
      Caption         =   "ADMINISTRATIOR"
      Begin VB.Menu mnu_add 
         Caption         =   "ADD USER"
      End
      Begin VB.Menu mnu_distwrk 
         Caption         =   "DISTRIBUTE WORK"
      End
      Begin VB.Menu mnu_vehclass 
         Caption         =   "ADD VEHICLE CLASS"
      End
   End
   Begin VB.Menu mnu_machanic 
      Caption         =   "MACHANIC"
      Begin VB.Menu mnu_view 
         Caption         =   "VIEW INFORMATION"
      End
      Begin VB.Menu mnu_changemech 
         Caption         =   "CHANGE PASSWORD"
      End
      Begin VB.Menu mnuu_billp 
         Caption         =   "BILL PRIPARATION"
      End
   End
   Begin VB.Menu mnu_uprate 
      Caption         =   "UPDATE RATE"
      Begin VB.Menu mnu_bdywrkrat 
         Caption         =   "BODY WORK"
      End
      Begin VB.Menu mnu_mechrat 
         Caption         =   "MECHANICAL WORK"
      End
      Begin VB.Menu mnu_wheelrat 
         Caption         =   "WHEEL ALLIGENMENT"
      End
      Begin VB.Menu mnu_waterrat 
         Caption         =   "WATER SERVICE"
      End
   End
   Begin VB.Menu mnu_report 
      Caption         =   "REPORT"
   End
   Begin VB.Menu mnu_logout 
      Caption         =   "LOGOUT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnu_add_Click()
FRM_WORKER.Show
End Sub

Private Sub mnu_bdywrkrat_Click()
frm_bdyrate.Show
End Sub

Private Sub mnu_bill_Click()
frm_billselect.Show
End Sub







Private Sub mnu_body_Click()
frm_body.Show
End Sub

Private Sub mnu_chagep_Click()
frm_chpswrd.Show
End Sub

Private Sub mnu_changemech_Click()
frm_chpswrd.Show
End Sub

Private Sub mnu_delevary_Click()
frm_delevary.Show
End Sub

Private Sub mnu_distwrk_Click()
frm_wstatus.Show
End Sub

Private Sub mnu_logout_Click()
 frm_login.txt_uname.Text = " "
 frm_login.txt_pword.Text = " "
 Unload MDIForm1
 Load frm_login
 frm_login.Show
End Sub

Private Sub mnu_mach_Click()
frm_mrepair.Show
End Sub

Private Sub mnu_mechrat_Click()
frm_reparerate.Show
End Sub

Private Sub mnu_regi_Click()
frm_registration.Show
End Sub



Private Sub mnu_report_Click()
frm_report.Show
End Sub

Private Sub mnu_statup_Click()
frm_staupdate.Show
End Sub

Private Sub mnu_vehclass_Click()
frm_vehclass.Show
End Sub

Private Sub mnu_view_Click()
frm_viewinformation.Show
End Sub

Private Sub mnu_water_Click()
frm_wservice.Show
End Sub
Private Sub mnu_waterrat_Click()
frm_waterrate.Show
End Sub

Private Sub mnu_wheel_Click()
frm_allan.Show
End Sub

Private Sub mnu_wheelrat_Click()
FRM_WHEELRATE.Show
End Sub

Private Sub mnuu_billp_Click()
frm_billgen.Show
End Sub
