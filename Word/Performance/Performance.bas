Attribute VB_Name = "Performance"
' Performance
' Modify/Restore context for performance when running script
' Author: SMFSW, 2016
' Copyright: MIT
'


Dim perfView As Variant
Dim perfPagin As Boolean
Dim perfUpdate As Boolean


'Public Sub savePerfActDoc()
'    savePerfContext ActiveDocument
'End Sub

'Public Sub restorePerfActDoc()
'    restorePerfContext ActiveDocument
'End Sub


Public Sub savePerfContext(ByRef doc As Word.Document)
    perfView = doc.Windows(1).View
    doc.Windows(1).View = wdNormalView
    With doc.Application
        perfPagin = .Options.Pagination
        perfUpdate = .ScreenUpdating
        .Options.Pagination = False
        .ScreenUpdating = False
    End With
End Sub


Public Sub restorePerfContext(ByRef doc As Word.Document)
    doc.Windows(1).View = perfView
    With doc.Application
        .Options.Pagination = perfPagin
        .ScreenUpdating = perfUpdate
    End With
End Sub

