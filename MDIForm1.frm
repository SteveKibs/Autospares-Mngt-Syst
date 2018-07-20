VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9360
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   12495
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "Open Form"
         Begin VB.Menu mnul88 
            Caption         =   "-"
         End
         Begin VB.Menu mnunewr 
            Caption         =   "Student Registration Form"
         End
         Begin VB.Menu mnuln80 
            Caption         =   "-"
         End
         Begin VB.Menu mnuteacher 
            Caption         =   "&Teacher Form"
         End
         Begin VB.Menu mnuln76 
            Caption         =   "-"
         End
         Begin VB.Menu mnufinance 
            Caption         =   "&Finance Form"
         End
         Begin VB.Menu mnuln75 
            Caption         =   "-"
         End
         Begin VB.Menu mnupurchase 
            Caption         =   "&Purchase Form"
         End
         Begin VB.Menu mnuln56 
            Caption         =   "-"
         End
      End
   End
   Begin VB.Menu mnu0 
      Caption         =   ""
   End
   Begin VB.Menu mnuview 
      Caption         =   "& Reports"
      Begin VB.Menu mnuoppen 
         Caption         =   "Open"
         Begin VB.Menu mnunewrpt 
            Caption         =   "New Students Report"
         End
         Begin VB.Menu mnuln87 
            Caption         =   "-"
         End
         Begin VB.Menu mnuteach 
            Caption         =   "Teacher Report"
         End
         Begin VB.Menu mnuln22 
            Caption         =   "-"
         End
         Begin VB.Menu mnupurchaserpt 
            Caption         =   "Purchase Report"
         End
         Begin VB.Menu mnuln33 
            Caption         =   "-"
         End
         Begin VB.Menu mnufinancerpt 
            Caption         =   "Finance Report"
         End
         Begin VB.Menu mnuln21 
            Caption         =   "-"
         End
      End
   End
   Begin VB.Menu mnu01 
      Caption         =   ""
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnunew_Click()
Frmnewitem.Show
End Sub

Private Sub mnufinance_Click()
Frmfinance.Show
End Sub

Private Sub mnufinancerpt_Click()
rptfinance.Show
End Sub

Private Sub mnunewr_Click()
frmregistration.Show
End Sub

Private Sub mnunewrpt_Click()
rptregister.Show
End Sub

Private Sub mnuoderrpt_Click()
rptorder.Show
End Sub

Private Sub mnuorder_Click()
Frmorder.Show
End Sub

Private Sub mnupurchase_Click()
Frmpurchase.Show
End Sub

Private Sub mnupurchaserpt_Click()
rptpurchase.Show
End Sub

Private Sub mnusalerpt_Click()
rptsales.Show
End Sub

Private Sub mnusales_Click()
Frmsales.Show
End Sub

Private Sub mnustock_Click()

End Sub

Private Sub mnuteach_Click()
rptteacher.Show
End Sub

Private Sub mnuteacher_Click()
Frmteacher.Show
End Sub
