Attribute VB_Name = "Module1"
Option Explicit

' Kør denne fra knap (kører kun beregning – ingen RefreshAll)
Public Sub UpdateCaseComplexes()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error GoTo CleanFail
    BuildCaseComplexes

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    MsgBox "Fejl: " & Err.Number & vbCrLf & Err.Description, vbExclamation
    Resume CleanExit
End Sub


' Bygger CaseComplexes-arket (CaseComplexID, CaseID) og skriver CaseComplexID tilbage i tblCases
Public Sub BuildCaseComplexes()

    Dim tblCases As ListObject, tblRel As ListObject
    Dim wsOut As Worksheet
    Dim graph As Object, visited As Object
    Dim r As ListRow
    Dim a As String, b As String
    Dim key As Variant
    Dim complexNo As Long, outRow As Long

    Set graph = CreateObject("Scripting.Dictionary")   ' CaseID -> neighbors(dict)
    Set visited = CreateObject("Scripting.Dictionary") ' CaseID -> True

    ' 1) Find tabeller
    Set tblCases = FindTableByName("tblCases")
    If tblCases Is Nothing Then Err.Raise vbObjectError + 100, , "Kan ikke finde tabellen 'tblCases'."

    Set tblRel = FindTableByName("tblRelations")
    If tblRel Is Nothing Then Err.Raise vbObjectError + 101, , "Kan ikke finde tabellen 'tblRelations'."

    ' 2) Indlæs alle CaseID fra tblCases (så sager uden relationer kommer med)
    Dim colCaseID As Long
    colCaseID = GetColumnIndex(tblCases, "Case ID")
    If colCaseID = 0 Then Err.Raise vbObjectError + 102, , "Kolonnen 'Case ID' findes ikke i tblCases."

    For Each r In tblCases.ListRows
        a = Trim$(CStr(r.Range(1, colCaseID).Value))
        If Len(a) > 0 Then EnsureNode graph, a
    Next r

    ' 3) Indlæs relationer fra tblRelations
    Dim colA As Long, colB As Long
    colA = GetColumnIndex(tblRel, "CaseID")
    colB = GetColumnIndex(tblRel, "RelatedCaseID")
    If colA = 0 Or colB = 0 Then
        Err.Raise vbObjectError + 103, , "tblRelations skal have kolonnerne 'CaseID' og 'RelatedCaseID'."
    End If

    For Each r In tblRel.ListRows
        a = Trim$(CStr(r.Range(1, colA).Value))
        b = Trim$(CStr(r.Range(1, colB).Value))

        If Len(a) > 0 And Len(b) > 0 And a <> b Then
            EnsureNode graph, a
            EnsureNode graph, b

            ' eksplicitte objects for stabilitet
            Dim na As Object, nb As Object
            Set na = graph.Item(a)
            Set nb = graph.Item(b)
            na.Item(b) = True
            nb.Item(a) = True
        End If
    Next r

    ' 4) Outputark
    Set wsOut = GetOrCreateSheet("CaseComplexes")
    wsOut.Cells.Clear
    wsOut.Range("A1:B1").Value = Array("CaseComplexID", "CaseID")

    ' Status (dato/tid)
    wsOut.Range("D1").Value = "Last updated"
    wsOut.Range("E1").Value = Now
    wsOut.Range("E1").NumberFormat = "yyyy-mm-dd hh:mm:ss"

    outRow = 2
    complexNo = 0

    ' 5) Find sammenhængende komponenter (iterativ DFS)
    For Each key In graph.Keys
        If Not visited.Exists(CStr(key)) Then
            complexNo = complexNo + 1
            DFS_Iterative CStr(key), graph, visited, wsOut, complexNo, outRow
        End If
    Next key

    wsOut.Columns.AutoFit

    ' 6) Skriv CaseComplexID tilbage i tblCases
    WriteComplexIdBackToTblCases tblCases, wsOut

    MsgBox "Sagskomplekser opdateret: " & complexNo, vbInformation
End Sub


' Iterativ DFS (robust mod cirkler)
Private Sub DFS_Iterative(ByVal startCase As String, ByVal graph As Object, ByVal visited As Object, _
                          ByVal ws As Worksheet, ByVal complexNo As Long, ByRef outRow As Long)

    Dim stack As Collection
    Set stack = New Collection
    stack.Add startCase

    Do While stack.Count > 0
        Dim current As String
        current = CStr(stack(1))
        stack.Remove 1

        If Not visited.Exists(current) Then
            visited.Item(current) = True

            ws.Cells(outRow, 1).Value = complexNo
            ws.Cells(outRow, 2).Value = current
            outRow = outRow + 1

            Dim neigh As Object
            Set neigh = graph.Item(current)

            Dim n As Variant
            For Each n In neigh.Keys
                If Not visited.Exists(CStr(n)) Then
                    stack.Add CStr(n)
                End If
            Next n
        End If
    Loop
End Sub


' Skriver CaseComplexID ind i tblCases (opretter kolonnen hvis den mangler)
Private Sub WriteComplexIdBackToTblCases(ByVal tblCases As ListObject, ByVal wsOut As Worksheet)

    Dim map As Object
    Set map = CreateObject("Scripting.Dictionary") ' CaseID -> CaseComplexID

    ' Byg opslag fra CaseComplexes-arket
    Dim lastRow As Long
    lastRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row

    Dim i As Long, caseId As String, complexId As Variant
    For i = 2 To lastRow
        complexId = wsOut.Cells(i, 1).Value
        caseId = Trim$(CStr(wsOut.Cells(i, 2).Value))
        If Len(caseId) > 0 Then map(caseId) = complexId
    Next i

    ' Sikr at kolonnen CaseComplexID findes i tblCases
    Dim colComplex As Long
    colComplex = GetColumnIndex(tblCases, "CaseComplexID")
    If colComplex = 0 Then
        tblCases.ListColumns.Add.Name = "CaseComplexID"
        colComplex = GetColumnIndex(tblCases, "CaseComplexID")
    End If

    Dim colCaseID As Long
    colCaseID = GetColumnIndex(tblCases, "Case ID")
    If colCaseID = 0 Then Err.Raise vbObjectError + 104, , "Kolonnen 'Case ID' findes ikke i tblCases."

    ' Skriv værdier tilbage
    Dim r As ListRow
    For Each r In tblCases.ListRows
        caseId = Trim$(CStr(r.Range(1, colCaseID).Value))
        If Len(caseId) > 0 And map.Exists(caseId) Then
            r.Range(1, colComplex).Value = map(caseId)
        Else
            r.Range(1, colComplex).ClearContents
        End If
    Next r
End Sub


' -------- Hjælpere

Private Sub EnsureNode(ByVal graph As Object, ByVal nodeId As String)
    If Not graph.Exists(nodeId) Then
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        graph.Add nodeId, d
    End If
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Private Function FindTableByName(ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
    Set FindTableByName = Nothing
End Function

Private Function GetColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0
End Function


