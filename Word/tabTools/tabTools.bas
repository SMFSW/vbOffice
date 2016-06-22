Attribute VB_Name = "tabTools"
' tabTools
' Tools for tables generation
' Version: 0.1
' Author: SMFSW, 2016
' Copyright: MIT
'


Public Function tabBuild(nx As Integer, ny As Integer, ByRef dataParam, ByRef headerParam) As Boolean
'
' tabBuild
' Table building in a new document
'
    Dim docNew As Document
    Dim tableNew As Table
    Dim txt As String
    Dim lenTab As Integer: lenTab = 0
    
    For index = 1 To UBound(dataParam, 1)
        If dataParam(index - 1, 0) <> "" Then lenTab = lenTab + 1
    Next index
    
    Set docNew = Documents.Add
    Set tableNew = docNew.Tables.Add(Selection.Range, lenTab, nx)
    
    ' Populate Tab
    With tableNew
        For idy = 1 To lenTab
            For idx = 1 To nx
                txt = dataParam(idy - 1, idx - 1)
                .Cell(idy, idx).Range.InsertAfter txt
            Next idx
        Next idy
    End With
    
    ' Set array grid
    With Selection.Tables(1)
        If .Style <> "Grille du tableau" Then .Style = "Grille du tableau"
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = False
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = True
        .AutoFitBehavior (wdAutoFitContent)
    End With
    
    ' Insert and format header row
    With Selection
        .InsertRowsAbove 1
        Options.DefaultBorderColor = wdColorAutomatic
        With Options
            .DefaultBorderLineStyle = wdLineStyleDouble
            .DefaultBorderLineWidth = wdLineWidth050pt
            .DefaultBorderColor = wdColorAutomatic
        End With
        With Selection.Borders(wdBorderBottom)
            .LineStyle = Options.DefaultBorderLineStyle
            .LineWidth = Options.DefaultBorderLineWidth
            .Color = Options.DefaultBorderColor
        End With
        .Font.Bold = wdToggle
        .Tables(1).Select
        .Font.Size = 10
        .Font.Name = "Arial"
        .Rows.HeightRule = wdRowHeightExactly
        .Rows.Height = CentimetersToPoints(0.5)
    End With
    
    ' Populate table headers
    With tableNew
        For idx = 1 To nx
            txt = headerParam(idx - 1)
            .Cell(1, idx).Range.InsertAfter txt
        Next idx
    End With
End Function

