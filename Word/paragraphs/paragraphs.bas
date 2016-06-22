Attribute VB_Name = "paragraphs"
' paragraphs
' Get headings and build a table with text, line numbers, paragraph number & depth
' Version: 0.4
' Author: SMFSW, 2016
' Copyright: MIT
'

' TODO: Being able to handle modifications tracking (no count of deleted titles)
' TODO: Find the end of document another way (script may not work on some messy documents)
' TODO: Find a way to handle next heading find not going forward sometimes (which causes end of script)



' par (x, 0) : line in paragraph scale (full line to dot & carriage return / aka paragraph for Word)
' par (x, 1) : pargraph numbering output as string
' par (x, 2) : depth of heading
' par (x, 3) : text of heading
Private par(500, 3) As Variant
Private parCpt As Integer
Private parInit As Boolean


' erase content of global variables & array
Public Sub parErase()
    For i = 0 To 500
        For j = 0 To 3
            par(i, j) = ""
        Next j
    Next i
    parCpt = 0
    parInit = False
End Sub


' return paragraph number of range r
Public Function parGetNum(r As Range) As Double
    Dim rParagraphs As Range
    Dim CurPos As Double
    'If parInit = False Then Call parBuild  ' par tab not needed in parGetNum
    
    r.Select
    CurPos = ActiveDocument.Bookmarks("\startOfSel").Start
    Set rParagraphs = ActiveDocument.Range(Start:=0, End:=CurPos)
    parGetNum = rParagraphs.paragraphs.Count    ' USE NAME OF THE FUNCTION AS RETURN VALUE
End Function


' return paragraph number of range r as formated string
Public Function parGetStr(r As Range) As String
    Dim CurPos As Double
    Dim tmp As String: tmp = ""
    If parInit = False Then Call parBuild
    
    r.Select
    CurPos = parGetNum(r)
    For j = 0 To parCpt
        If par(j, 0) >= CurPos Then
            If j <> 0 Then
                tmp = par(j - 1, 1)
            Else
                tmp = 0 ' If before 1st Header, return 0
            End If
            
            Exit For    ' Exit when found
        End If
    Next j
    parGetStr = tmp
End Function


' add par table in a new document
Public Sub parDraw()
    Dim txtHeaders As Variant
    txtHeaders = Array("line", _
                       "chapter", _
                       "depth", _
                       "txt")
    
    savePerfContext ActiveDocument
    If parInit = False Then Call parBuild
    Call tabBuild(4, 0, par, txtHeaders)
    restorePerfContext ActiveDocument
End Sub


Private Sub parBuild()
    Dim maxDepth As Integer
    Dim cpt As Integer
    
    Dim flag As Boolean: flag = True
    Dim memStr As String
    
    ' parInit set to True before everything else so next calls to parGetXXX will not call parBuild again
    parInit = True
    
    ' move to the first heading (to determine text to strip for title depth)
    Selection.GoTo What:=wdGoToHeading, Which:=wdGoToFirst
    Selection.StartOf Unit:=wdParagraph
    Selection.MoveEnd Unit:=wdParagraph
    
    ' find how is called Title in your Word application
    Dim splt
    splt = split(Selection.Range.Style.NameLocal, " ")
    stripHeader = splt(0)
    Erase splt  ' erase temporary var splt
    
    ' move back to the start of the document
    Selection.HomeKey Unit:=wdStory
    
    'loop until the end of the document is reached
    parCpt = 0
    While flag = True
        Selection.GoTo What:=wdGoToHeading, Which:=wdGoToNext
        Selection.StartOf Unit:=wdParagraph
        Selection.MoveEnd Unit:=wdParagraph
        
        'get the line data
        strLine = Selection.Range.Text
        
        'check if the end of the document has been reached
        If memStr Like strLine Then flag = False
        memStr = strLine
        
        par(parCpt, 0) = parGetNum(Selection.Range)
        par(parCpt, 1) = "" ' init to empty str for later loop
        par(parCpt, 2) = Val(Replace(Selection.Range.Style.NameLocal, stripHeader, ""))
        par(parCpt, 3) = strLine    ' pour s'y retrouver (à virer par la suite)

        ' Determining max depth of titles for later
        If par(parCpt, 2) > maxDepth Then maxDepth = par(parCpt, 0)
                    
        ' Handling junk lines
        If parCpt <> 0 Then
            ' if depth par n-1 is equal to n's & line number from n-1 is right before n's
            If par(parCpt - 1, 2) = par(parCpt, 2) And par(parCpt - 1, 0) + 1 = par(parCpt, 0) Then
                ' copy to n-1 & don't incr parCpt
                par(parCpt - 1, 0) = par(parCpt, 0)
                par(parCpt - 1, 1) = par(parCpt, 1)
                par(parCpt - 1, 2) = par(parCpt, 2)
                par(parCpt - 1, 3) = par(parCpt, 3)
            Else: parCpt = parCpt + 1
            End If
        Else: parCpt = parCpt + 1
        End If
    Wend
    
    For i = 1 To maxDepth                   ' Applying paragraph numbers depth by depth
        cpt = 0                                 ' init at 0 so turns to 1 first time, which is what needed
        For j = 0 To parCpt
            If i > par(j, 2) Then cpt = 0           ' a sub paragraph end reached (resetting current and follow)
            If i = par(j, 2) Then cpt = cpt + 1     ' a new paragraph is reached (increment current and follow)
            If i <= par(j, 2) Then                  ' paragraph depth need to be added to sting
                If i <> 1 Then par(j, 1) = par(j, 1) & "."      ' dot added only if sub paragraph
                par(j, 1) = par(j, 1) & cpt                     ' append paragraph number to string in tab
            End If
        Next j
    Next i
End Sub



