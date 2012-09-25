VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Tradator_choose_qty_price 
   Caption         =   "Choose"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   OleObjectBlob   =   "frm_Tradator_choose_qty_price.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Tradator_choose_qty_price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_insert_in_tradator_Click()

Call tradator_insert_qty_price_from_form

End Sub
