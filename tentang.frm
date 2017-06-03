VERSION 5.00
Begin VB.Form ftentang 
   Caption         =   "Tentang"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "ftentang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

