VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MoveSortForm 
   Caption         =   "Move Sort"
   ClientHeight    =   1965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5745
   OleObjectBlob   =   "MoveSortForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MoveSortForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Move & Sort in active document (user form)
' Author: Boob / SMFSW (port to module with user form)
' Author: SMFSW (port to module with user form)


' Move button
Private Sub btnMove_Click()
    Call DeplacementEntrees
End Sub

' Sort button
Private Sub btnSort_Click()
    Call TriFeuille
End Sub

Private Sub chkHeadline_Click()
    Call ListColumns
End Sub

' Sheets combo box (for Move)
Private Sub cboxSheets_Change()
    ' do nothing
End Sub

' Colums combo box (for sort)
Private Sub cboxCol_Change()
    ' do nothing
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub MoveSortForm_Initialize()
    ' Call this sub to invoke Move Sort form (with access to sheets)
    With Me ' en bas à droite
      .StartUpPosition = 3
      .Top = Application.Height - Me.Height - 45
      .Left = Application.Width - Me.Width - 25
      .Repaint
      .Show 0
    End With
End Sub

