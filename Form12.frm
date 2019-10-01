VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00400000&
   Caption         =   "MAIN FORM"
   ClientHeight    =   6675
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11220
   FillColor       =   &H00FF0000&
   LinkTopic       =   "Form12"
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   6675
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu c 
      Caption         =   "COMPANY"
      WindowList      =   -1  'True
      Begin VB.Menu cd 
         Caption         =   "COMPANY DETAILS"
      End
      Begin VB.Menu nc 
         Caption         =   "NEW COMPANY"
      End
   End
   Begin VB.Menu o 
      Caption         =   "ORDER"
      Begin VB.Menu or 
         Caption         =   "ORDERS"
      End
      Begin VB.Menu nor 
         Caption         =   "NEW ORDER"
      End
   End
   Begin VB.Menu e 
      Caption         =   "EMPLOYEE"
      Begin VB.Menu ne 
         Caption         =   "NEW EMPLOYEE"
      End
      Begin VB.Menu ee 
         Caption         =   "EXISTING EMPLOYEES"
      End
   End
   Begin VB.Menu CU 
      Caption         =   "CUSTOMER"
      Begin VB.Menu ncu 
         Caption         =   "NEW CUSTOMER"
      End
      Begin VB.Menu ecu 
         Caption         =   "EXISTING CUSTOMER"
      End
   End
   Begin VB.Menu p 
      Caption         =   "PRODUCTS"
      Begin VB.Menu npr 
         Caption         =   "NEW PRODUCTS"
      End
      Begin VB.Menu pr 
         Caption         =   "PRODUCTS"
      End
      Begin VB.Menu st 
         Caption         =   "STOCK"
      End
   End
   Begin VB.Menu B 
      Caption         =   "BILLS"
      Begin VB.Menu nb 
         Caption         =   "NEW BILL"
      End
      Begin VB.Menu BI 
         Caption         =   "BILLS"
      End
   End
   Begin VB.Menu R 
      Caption         =   "REPORTS"
   End
   Begin VB.Menu s 
      Caption         =   "SETTINGS"
      Begin VB.Menu A 
         Caption         =   "ABOUT"
      End
      Begin VB.Menu H 
         Caption         =   "HELP"
      End
   End
   Begin VB.Menu E2 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cd_Click()
Form14.Show

End Sub

Private Sub E2_Click()
End
End Sub

Private Sub ecu_Click()
Form17.Show

End Sub

Private Sub ee_Click()
Form13.Show

End Sub

Private Sub nc_Click()
Form3.Show

End Sub

Private Sub ncu_Click()
Form7.Show

End Sub

Private Sub ne_Click()
Form6.Show

End Sub

Private Sub nor_Click()
Form5.Show

End Sub

Private Sub npr_Click()
Form18.Show

End Sub

Private Sub or_Click()
Form2.Show
End Sub

Private Sub pr_Click()
Form10.Show

End Sub

Private Sub st_Click()
Form8.Show

End Sub
