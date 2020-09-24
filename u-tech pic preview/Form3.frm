VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   525
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1920
   LinkTopic       =   "Form3"
   ScaleHeight     =   525
   ScaleWidth      =   1920
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnu_preview 
      Caption         =   "preview"
      Begin VB.Menu mnu_clear 
         Caption         =   "clear"
      End
      Begin VB.Menu mnu_add_dir 
         Caption         =   "add dir"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu_add_dir_Click()
Call filebox_2_list(Form1.List1, Form1.File1)

Call filebox_2_list2(Form1.List2, Form1.Dir1, Form1.File1)
End Sub

Private Sub mnu_clear_Click()
Form1.List1.Clear
Form1.List2.Clear
End Sub
