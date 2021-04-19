Attribute VB_Name = "export_VD"
Sub Загрузка_ВД()
    Set Data_out = Application.ActiveWorkbook.ActiveSheet
    n = SheetGetSize(Data_out)
    n_row = n(1)
    n_col = n(2)
    spec = Data_out.Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, n_col))
    CSVfilename$ = ThisWorkbook.Path & "\Загрузка_ВД.txt"
    If n_row > 0 Then r = ExportArray2CSV(spec, CSVfilename, vbTab, vbNewLine)
End Sub

Function ExportArray2CSV(ByVal arr As Variant, ByVal CSVfilename As String, Optional ByVal ColumnsSeparator$ = ";", Optional ByVal RowsSeparator$ = vbNewLine) As String
    buffer$ = vbNullString
    For i = 1 To UBound(arr, 1)
        txt = vbNullString
        For j = 1 To UBound(arr, 2)
            If arr(i, j) <> vbNullString Then txt = txt & ColumnsSeparator$ & arr(i, j)
        Next j
        If txt <> vbNullString Then
            Range2CSV = Range2CSV & Mid(txt, Len(ColumnsSeparator$) + 1) & RowsSeparator$
            If Len(Range2CSV) > 50000 Then buffer$ = buffer$ & Range2CSV: Range2CSV = vbNullString
        End If
    Next i
    CSVtext$ = buffer$ & Range2CSV
    ExportArray2CSV = ExportSaveTXTfile(CSVfilename$, CSVtext$)
End Function

Function ExportSaveTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.CreateTextFile(filename, True)
    ts.Write txt: ts.Close
    ExportSaveTXTfile = Err = 0
    Set ts = Nothing: Set fso = Nothing
End Function

Private Function SheetGetSize(ByVal objLst As Variant) As Variant
    Dim out(2)
    On Error Resume Next
    t = objLst.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If Not IsEmpty(t) Then
        out(2) = t
    Else
        out(2) = 1
    End If
    On Error Resume Next
    t = objLst.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If Not IsEmpty(t) Then
        out(1) = t
    Else
        out(1) = 1
    End If
    SheetGetSize = out
    Erase out
End Function
