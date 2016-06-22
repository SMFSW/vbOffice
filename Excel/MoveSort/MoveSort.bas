Attribute VB_Name = "MoveSort"
' Move & Sort in active document
' Author: Boob / SMFSW (port to module with user form)

Sub MoveSortShow()
    Call ListSheets
    Call ListColumns
    
    ' Call this sub to invoke Move Sort form (with access to sheets)
    ' Call MoveSortForm_Initialize
    With MoveSortForm ' en bas à droite
      .StartUpPosition = 3
      .Top = Application.Height - MoveSortForm.Height - 45
      .Left = Application.Width - MoveSortForm.Width - 25
      .Show 0
    End With
End Sub


Private Function ListSheets()
    FD = ActiveSheet.Name
        
    For i = 1 To ActiveWorkbook.Sheets.Count    ' boucle sur les feuilles du fichier
        FeuilAdd = Sheets(i).Name
        If ActiveSheet.Name <> Sheets(i).Name Then
            MoveSortForm.cboxSheets.AddItem FeuilAdd
        End If
    Next i
    
    Sheets(FD).Select
End Function

Function ListColumns()
    Dim ColumnAdd As String
    
    For i = 1 To MoveSortForm.cboxCol.ListCount
        MoveSortForm.cboxCol.RemoveItem 0
    Next i
    
    For i = 1 To Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
        If MoveSortForm.chkHeadline.Value = True Then
            ColumnAdd = Cells(1, i).Value
        Else
            ColumnAdd = Chr(64 + i)     ' 64 is A in ASCII
        End If
    
        MoveSortForm.cboxCol.AddItem ColumnAdd
        
        If i = 1 Then
            MoveSortForm.cboxCol.Value = ColumnAdd
        End If
    Next i
End Function

Sub DeplacementEntrees()
    Select Case MoveSortForm.btnMove.Caption
        Case "Move start"
            ConstrChkBoxes
            MoveSortForm.btnMove.Caption = "Move selection"
        
        Case "Move selection"
            If TestFDValue = 0 Then Exit Sub
            MsgTrs = TransfertEntrees(MoveSortForm.cboxSheets.Value)
            MsgSupp = SupprLignes
            SupprChkBoxes
            
            MoveSortForm.btnMove.Caption = "Move start"
            
            Msg10 = MsgTrs & " lignes copiées sur " & MoveSortForm.cboxSheets.Value & vbCrLf
            Msg20 = MsgSupp & " lignes supprimées de " & ActiveSheet.Name + vbCrLf
            MsgFin = Msg10 & Msg20
            MsgBox MsgFin
            
            MoveSortForm.cboxSheets.Text = ""
            Range("A2").Select
        
        Case Else
            MoveSortForm.btnMove.Caption = "Move start"
            Exit Sub
    End Select
End Sub


Function ConstrChkBoxes()    'compte les cellules et pose des case à cocher cellule liée en col I
    AC = ActiveCell.Row
    
    Application.CutCopyMode = False
    Cells.Select
    Selection.RowHeight = 15    'ligne hauteur 15
    'Rows(1).Select
    'Selection.RowHeight = 15
    
    'détermination de la dernière ligne selon la présence de données dans la colonne A
    lastrow = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    Position = 0
    
    ActiveSheet.CheckBoxes.Delete
    
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.ColumnWidth = 3.71
    
    If MoveSortForm.chkHeadline.Value = True Then
        tmp = 1
    Else
        tmp = 0
    End If

    For i = tmp To lastrow - 1 Step 1
        Position = (15 * (i))
        ActiveSheet.CheckBoxes.Add(0, Position, 20, 15).Select
        With Selection
            Pos = "$A$" & (i + 1)
            .LinkedCell = Pos
            .Name = "CB" & i
            .Characters.Text = ""
            .ShapeRange.Fill.Visible = msoTrue
            .ShapeRange.Fill.Solid
            .ShapeRange.Fill.ForeColor.SchemeColor = 69
            .ShapeRange.ZOrder msoBringToFront
        End With
    Next i
    Cells(AC, 2).Select
End Function

Function TestFDValue()
    MSGP = " Choisir une feuille de destination !!!"
    MSGT = " Erreur : Feuille de destination non définie !!!"
    TestFDValue = 0
    
    If MoveSortForm.cboxSheets.Value = "" Then
        MsgBox MSGP, vbExclamation, MSGT
    Else
        TestFDValue = MoveSortForm.cboxSheets.Value
    End If
End Function


Function TransfertEntrees(FeuilDest)
    NbCopy = 0
    FD = ActiveSheet.Name
    
    If MoveSortForm.chkHeadline.Value = True Then
        tmp = 2
    Else
        tmp = 1
    End If
    
    For i = tmp To Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
        If Range("A" & i) = True Then
            NbCopy = NbCopy + 1
            Colfin = Cells(i, 256).End(xlToLeft).Column
            Range(Cells(i, 2), Cells(i, Colfin)).Select
            Range(Cells(i, 2), Cells(i, Colfin)).Copy
            DeplacementEntree FeuilDest
            Sheets(FD).Select
        End If
    Next i
    TransfertEntrees = NbCopy
End Function


Function DeplacementEntree(FeuilDest)
    A = FeuilDest
    Sheets(A).Select
    CF1 = "A" & (Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row + 1)
    Range(CF1).Select
    ActiveSheet.Paste
End Function


Function SupprLignes()
    NbSuppr = 0
    
    If MoveSortForm.chkHeadline.Value = True Then
        tmp = 2
    Else
        tmp = 1
    End If
    
    For i = tmp To Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
        If Range("A" & i) = True Then
            Rows(i).Select
            Rows(i).Delete
            NbSuppr = NbSuppr + 1
            i = i - 1
        End If
    Next i
    SupprLignes = NbSuppr
End Function


Function SupprChkBoxes()
    ActiveSheet.CheckBoxes.Delete
    Columns("A:A").Delete
End Function


Function TriFeuille()
    ColMax = 0
    
    'code de tri de la feuille en cours
    Rowfin = Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
    For f = 1 To Rowfin - 1
        Colfin = Cells(f, 256).End(xlToLeft).Column
        If Colfin > ColMax Then
            ColMax = Colfin
        End If
    Next f
    
    Range(Cells(1, 1), Cells(Rowfin, ColMax)).Select
    Selection.Sort Key1:=Range("A2"), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
    Range("A2").Select
End Function


