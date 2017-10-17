Attribute VB_Name = "common"
Option Compare Text
Option Base 1

Public Const common_version As String = "2.1"

Public Function GetLeghtByID(id As String, table As Range, n_col_id As Integer, n_col_l As Integer) As Variant
    Sum_l = 0
    For i = 1 To table.Rows.Count
        If table(i, n_col_id) = id Then Sum_l = Sum_l + table(i, n_col_l)
    Next i
    GetLeghtByID = Sum_l
End Function

Function GetHeightSheet() As Double
    Set sh = Application.ThisWorkbook.ActiveSheet
    sh.ResetAllPageBreaks
    lsize = GetSizeSheet(sh)
    n_row = lsize(1)
    h_sheet = 0
    For i = 1 To n_row
        h_row_point = sh.Rows(i).RowHeight
        h_row_mm = h_row_point / 72 * 25.4
        h_sheet = h_sheet + h_row_mm
    Next i
    GetHeightSheet = h_sheet
End Function

Function SetPageBreaks(ByVal h_list As Double, Optional ByVal n_first As Integer) As Boolean
    If GetHeightSheet() > h_list Then
        Set sh = Application.ThisWorkbook.ActiveSheet
        sh.ResetAllPageBreaks
        lsize = GetSizeSheet(sh)
        n_row = lsize(1)
        n_col = lsize(2)
        sh.VPageBreaks.Add Before:=sh.Cells(1, n_col)
        h_dop = 0
        For i = 1 To n_first
            h_row_point = sh.Rows(i).RowHeight
            h_row_mm = h_row_point / 72 * 25.4
            h_dop = h_dop + h_row_mm
        Next i
        h_max = h_list + h_dop
        h_t = 0
        For i = 1 To n_row + 1
            h_row_point = sh.Rows(i).RowHeight
            h_row_mm = h_row_point / 72 * 25.4
            h_t = h_t + h_row_mm
            If h_t >= h_max Then
                sh.HPageBreaks.Add Before:=sh.Range(sh.Cells(i, 1).MergeArea(1).Address)
                h_t = 0
            End If
        Next i
        SetPageBreaks = True
    Else
        SetPageBreaks = False
    End If
End Function

Function ArrayCol(ByVal array_in As Variant, ByVal col As Integer) As _
    Variant
    If IsEmpty(array_in) Then ArrayCol = Empty: Exit Function
    If ArrayIsSecondDim(array_in) = False Then ArrayCol = array_in: Exit _
        Function
    If UBound(array_in, 2) < row Then ArrayCol = Empty: Exit Function
    n = UBound(array_in, 1)
    Dim out(): ReDim out(n)
    For i = 1 To n
        out(i) = array_in(i, col)
    Next i
    ArrayCol = out
    Erase out
End Function

Function ArrayGetRowIndex(ByVal array_in As Variant, ByVal param As _
    Variant, Optional ByVal n_col As Integer) As Integer
    Index = Empty
    If IsEmpty(array_in) Then
        ArrayGetRowIndex = Index
        Exit Function
    End If
    If ArrayIsSecondDim(array_in) Then
        For i = 1 To UBound(array_in, 1)
            If array_in(i, n_col) = param Then
                Index = i
                Exit For
            End If
        Next i
    Else
        For i = 1 To UBound(array_in)
            If array_in(i) = param Then
                Index = i
                Exit For
            End If
        Next i
    End If
    ArrayGetRowIndex = Index
End Function
    
Function ArrayCombine(ByVal Arr1 As Variant, ByVal Arr2 As Variant) As _
    Variant
    If (Not IsArray(Arr1)) And IsArray(Arr2) Then ArrayCombine = Arr2: Exit _
        Function
    If (Not IsArray(Arr2)) And IsArray(Arr1) Then ArrayCombine = Arr1: Exit _
        Function
    If (Not IsArray(Arr2)) And (Not IsArray(Arr1)) Then ArrayCombine = _
        Empty: Exit Function
    On Error Resume Next: Err.Clear
    If Err.Number = 9 Then ArrayCombine = Empty: Exit Function
    If ArrayIsSecondDim(Arr1) And ArrayIsSecondDim(Arr2) Then
        If (LBound(Arr1, 2) <> LBound(Arr2, 2)) Or (UBound(Arr1, 2) <> _
            UBound(Arr2, 2)) Then ArrayCombine = Empty: Exit Function
        ReDim arr(1 To UBound(Arr1, 1) + UBound(Arr2, 1), LBound(Arr1, 2) _
            To UBound(Arr1, 2))
        For i = 1 To UBound(Arr1, 1)
            For j = LBound(Arr1, 2) To UBound(Arr1, 2)
                arr(i, j) = Arr1(i, j)
            Next j
        Next i
        For i = 1 To UBound(Arr2, 1)
            For j = LBound(Arr2, 2) To UBound(Arr2, 2)
                arr(i + UBound(Arr1, 1), j) = Arr2(i, j)
            Next j
        Next i
        ArrayCombine = arr
    Else
        If ArrayIsSecondDim(Arr1) Then ArrayCombine = Empty: Exit Function
        If ArrayIsSecondDim(Arr2) Then ArrayCombine = Empty: Exit Function
        ReDim arr_(1 To UBound(Arr1) + UBound(Arr2))
        For i = 1 To UBound(Arr1)
            arr_(i) = Arr1(i)
        Next i
        For i = 1 To UBound(Arr2, 1)
            arr_(i + UBound(Arr1)) = Arr2(i)
        Next i
        ArrayCombine = arr_
    End If
End Function

Function ArrayEmp2Space(ByRef array_in As Variant) As Variant
    If Not (IsEmpty(array_in)) Then
        seconddim = ArrayIsSecondDim(array_in)
        If Not (seconddim) Then
            For i = 1 To UBound(array_in, 1)
                If array_in(i) = "" Then array_in(i) = " "
                If array_in(i) = 0 Then array_in(i) = " "
                If IsNumeric(array_in(i)) And type_okrugl > 2 Then _
                    array_in(i) = Round(array_in(i), 4)
            Next
        Else
            For i = 1 To UBound(array_in, 1)
                For j = 1 To UBound(array_in, 2)
                    If array_in(i, j) = "" Then array_in(i, j) = " "
                    If array_in(i, j) = 0 Then array_in(i, j) = " "
                    If IsNumeric(array_in(i, j)) And type_okrugl > 2 Then _
                        array_in(i, j) = Round(array_in(i, j), 4)
                Next
            Next
        End If
    End If
    ArrayEmp2Space = array_in
End Function

Function ArrayIsSecondDim(ByVal array_in As Variant) As Boolean
    If IsEmpty(array_in) Or Not IsArray(array_in) Then
        ArrayIsSecondDim = False
        Exit Function
    Else
        temp = 0
        On Error Resume Next
        n = UBound(array_in)
        For i = 1 To 60
            On Error Resume Next
            Tmp = Tmp + UBound(array_in, i)
        Next
        If Tmp > n Then
            ArrayIsSecondDim = True
        Else
            ArrayIsSecondDim = False
        End If
    End If
    Erase array_in
End Function

Function ArrayRedim(ByVal array_in As Variant, ByVal n_row As Integer) As _
    Variant
    If IsEmpty(array_in) Then ArrayRedim = Empty: Exit Function
    If ArrayIsSecondDim(array_in) = False Then
        ReDim Preserve array_in(n_row)
        ArrayRedim = array_in
        Exit Function
    End If
    n_col = UBound(array_in, 2)
    array_in = ArrayTranspose(array_in)
    ReDim Preserve array_in(n_col, n_row)
    array_in = ArrayTranspose(array_in)
    ArrayRedim = array_in
    Erase array_in
End Function

Function ArrayRow(ByVal array_in As Variant, ByVal row As Integer) As _
    Variant
    If IsEmpty(array_in) Then ArrayRow = Empty: Exit Function
    If ArrayIsSecondDim(array_in) = False Then ArrayRow = array_in: Exit _
        Function
    If UBound(array_in, 1) < row Then ArrayRow = Empty: Exit Function
    n = UBound(array_in, 2)
    Dim out(): ReDim out(n)
    For i = 1 To n
        out(i) = array_in(row, i)
    Next i
    ArrayRow = out
    Erase out, array_in
End Function

Function ArraySelectParam(ByVal array_in As Variant, ByVal param1 As _
    Variant, ByVal n_col1 As Variant, Optional ByVal param2 As Variant, _
    Optional ByVal n_col2 As Variant) As Variant
    Dim arrout
    If IsEmpty(array_in) Then
        ArraySelectParam = Empty
        Exit Function
    End If
    If ArrayIsSecondDim(array_in) Then
        n_row = UBound(array_in, 1)
        n_param = UBound(array_in, 2)
        n_row_s = 0
        ReDim arrout(n_row, n_param)
        For j = 1 To n_row
        
            If Not IsMissing(n_col2) And Not IsMissing(param2) Then
                If array_in(j, n_col2) = param2 Then
                    flag2 = 1 'Записывать
                Else
                    flag2 = 0 'Не записывать
                End If
            Else
                flag2 = 1 'Обязательно записывать
            End If
            
            If array_in(j, n_col1) = param1 Then
                flag1 = 1 'Конечно, записывать
            Else
                flag1 = 0 'Не записывать ни в коем случае
            End If
            
            If flag1 And flag2 Then 'Если все согласны
                    n_row_s = n_row_s + 1
                    For k = 1 To n_param
                        arrout(n_row_s, k) = array_in(j, k)
                    Next k
            End If
        Next j
        arrout = ArrayTranspose(arrout)
        If n_param > 0 And n_row_s > 0 Then
            ReDim Preserve arrout(n_param, n_row_s)
            ArraySelectParam = ArrayTranspose(arrout)
            Exit Function
        Else
            ArraySelectParam = Empty
            Exit Function
        End If
    Else
        n_row = UBound(array_in, 1)
        n_row_s = 0
        ReDim arrout(n_row)
        For j = 1 To n_row
            If array_in(j) = param1 Then
                n_row_s = n_row_s + 1
                arrout(n_row_s) = array_in(j)
            End If
        Next j
        If n_row_s > 0 Then
            ReDim Preserve arrout(n_row_s)
            ArraySelectParam = arrout
            Exit Function
        Else
            ArraySelectParam = Empty
            Exit Function
        End If
    End If
    Erase array_in
End Function

Function ArraySort(ByVal array_in As Variant, ByVal nCol As Integer) As Variant
    If IsEmpty(array_in) Then ArraySort = Empty: Exit Function
End Function

Function ArraySortABC(ByVal array_in As Variant, ByVal nCol As Integer) As _
    Variant
    If IsEmpty(array_in) Then ArraySortABC = Empty: Exit Function
    If ArrayIsSecondDim(array_in) Then
        Dim tempArray As Variant: ReDim tempArray(1, UBound(array_in, 2))
        k = UBound(array_in, 1)
        For j = 1 To UBound(array_in, 1) - 1
            For i = 2 To k
                If array_in(i - 1, nCol) <> Empty And array_in(i, nCol) <> _
                    Empty Then
                    If Asc(UCase(array_in(i - 1, nCol))) > _
                        Asc(UCase(array_in(i, nCol))) Then
                        For m = 1 To UBound(array_in, 2)
                            tempArray(1, m) = array_in(i - 1, m)
                            array_in(i - 1, m) = array_in(i, m)
                            array_in(i, m) = tempArray(1, m)
                        Next m
                    End If
                End If
            Next i
            k = k - 1
        Next j
    Else
        k = UBound(array_in)
        For j = 1 To UBound(array_in) - 1
            For i = 2 To k
                If Not IsEmpty(array_in(i - 1)) And Not _
                    IsEmpty(array_in(i)) Then
                    If Asc(UCase(array_in(i - 1))) > _
                        Asc(UCase(array_in(i))) Then
                        v = array_in(i - 1)
                        array_in(i - 1) = array_in(i)
                        array_in(i) = v
                    End If
                End If
            Next i
            k = k - 1
        Next j
    End If
    ArraySortABC = array_in
    Erase array_in
End Function

Function ArraySortNum(ByVal array_in As Variant, ByVal nCol As Integer) As _
    Variant
    If IsEmpty(array_in) Then ArraySortNum = Empty: Exit Function
    If ArrayIsSecondDim(array_in) Then
        If nCol > UBound(array_in, 2) Or nCol < LBound(array_in, 2) Then _
           MsgBox "Нет такого столбца в массиве!", vbCritical: Exit Function
        Dim Check As Boolean, iCount As Integer, jCount As Integer, nCount As Integer
        ReDim tmparr(UBound(array_in, 2)) As Variant
        Do Until Check
            Check = True
            For iCount = LBound(array_in, 1) To UBound(array_in, 1) - 1
                If val(array_in(iCount, nCol)) > val(array_in(iCount + 1, nCol)) Then
                    For jCount = LBound(array_in, 2) To UBound(array_in, 2)
                        tmparr(jCount) = array_in(iCount, jCount)
                        array_in(iCount, jCount) = array_in(iCount + 1, jCount)
                        array_in(iCount + 1, jCount) = tmparr(jCount)
                        Check = False
                    Next
                End If
            Next
        Loop
    Else
        n = UBound(array_in)
        For i = 1 To n
            For j = 1 To (n - i)
                If array_in(j) > array_in(j + 1) Then
                    Tmp = array_in(j)
                    array_in(j) = array_in(j + 1)
                    array_in(j + 1) = Tmp
                End If
            Next j
        Next i
    End If
    ArraySortNum = array_in
    Erase array_in
End Function
Function ArrayTranspose(ByVal array_in As Variant) As Variant
    Dim tempArray As Variant
    If ArrayIsSecondDim(array_in) Then
        ReDim tempArray(LBound(array_in, 2) To UBound(array_in, 2), _
            LBound(array_in, 1) To UBound(array_in, 1))
        For x = LBound(array_in, 2) To UBound(array_in, 2)
            For Y = LBound(array_in, 1) To UBound(array_in, 1)
                tempArray(x, Y) = array_in(Y, x)
            Next Y
        Next x
    Else:
        ReDim tempArray(LBound(array_in, 1) To UBound(array_in, 1), _
            LBound(array_in, 1) To UBound(array_in, 1))
        For x = LBound(array_in, 1) To UBound(array_in, 1)
            tempArray(x, 1) = array_in(x)
        Next x
    End If
    ArrayTranspose = tempArray
    Erase tempArray
End Function

Function ArrayUniqValColumn(ByVal array_in As Variant, ByVal cols As Long) _
    As Variant
    Dim array_out()
    Dim array_temp()
    If IsEmpty(array_in) Or Not IsArray(array_in) Then
        ArrayUniqValColumn = Empty
        Exit Function
    End If
    n_num = 0: n_str = 0
    If ArrayIsSecondDim(array_in) Then
        ReDim array_out(UBound(array_in, 1))
        n_un = 1
        If cols = 0 Then cols = 1
        array_out(1) = array_in(1, cols)
        For i = 1 To UBound(array_in, 1)
            flag = 1
            For j = 1 To n_un
                If array_out(j) = array_in(i, cols) Then flag = 0
            Next
            If array_in(i, cols) = "" Then flag = 0
            If array_in(i, cols) = " " Then flag = 0
            If ConvTxt2Num(array_in(i, cols)) = 0 Then flag = 0
            If flag = 1 Then
                n_un = n_un + 1
                array_out(n_un) = array_in(i, cols)
                If IsNumeric(array_out(n_un)) Then
                    n_num = n_num + 1
                Else
                    n_str = n_str + 1
                End If
            End If
        Next
    Else
        ReDim array_out(UBound(array_in))
        n_un = 1
        If cols = 0 Then cols = 1
        array_out(1) = array_in(1)
        For i = 1 To UBound(array_in)
            flag = 1
            For j = 1 To n_un
                If array_out(j) = array_in(i) Then flag = 0
            Next
            If array_in(i) = "" Then flag = 0
            If array_in(i) = " " Then flag = 0
            If ConvTxt2Num(array_in(i)) = 0 Then flag = 0
            If flag = 1 Then
                n_un = n_un + 1
                array_out(n_un) = array_in(i)
                If IsNumeric(array_out(n_un)) Then
                    n_num = n_num + 1
                Else
                    n_str = n_str + 1
                End If
            End If
        Next
    End If
    ReDim Preserve array_out(n_un)
    If (n_num > n_str) Then
        array_out = ArraySortNum(array_out, 1)
    Else
        array_out = ArraySortABC(array_out, 1)
    End If
    ArrayUniqValColumn = array_out
    Erase array_out
End Function

Function OutEnded() As Boolean
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    OutEnded = True
End Function

Function OutPrepare() As Boolean
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    OutPrepare = True
End Function

Function SpecGetType(ByVal nm As String) As String
    On Error Resume Next
    form = ActiveWorkbook.VBProject.VBComponents("UserForm2").Name
    If IsEmpty(form) Then
        SpecGetType = 7
        Exit Function
    End If
    If Left(nm, 1) <> "!" And Left(nm, 1) <> "|" Then
        If InStr(nm, "_") > 0 Then
            type_spec = Split(nm, "_")
            Select Case type_spec(1)
                Case "гр"
                    spec = 1
                Case "об"
                    spec = 2
                Case "км"
                    spec = 4
                Case "кж"
                    spec = 5
                Case "поз"
                    spec = 6
                Case "спец"
                    spec = 7
                Case "поз"
                    spec = 9
                Case "правила"
                    spec = 10
                Case "вед"
                    spec = 11
                Case "экспл"
                    spec = 12
                End Select
        Else
            spec = 3
            If InStr(nm, "Фас") > 0 Then spec = 8
        End If
    Else
        spec = 0
    End If
    SpecGetType = spec
End Function


Function FormatClear() As Boolean
    With Cells
        .FormatConditions.Delete
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
    FormatClear = True
End Function

Function FormatFont(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean

    arr_bold = Array("шт.)", ", на ")
    For Each txt In arr_bold
        Data_out.FormatConditions.Add Type:=xlTextString, String:=txt, TextOperator:=xlContains
        Data_out.FormatConditions(Data_out.FormatConditions.Count).SetFirstPriority
        Data_out.FormatConditions(1).Font.Bold = True
    Next
    
    arr_underline = Array(" Материалы ", " Сборочные единицы ", " Прокат ", " Изделия ")
    For Each txt In arr_underline
        Data_out.FormatConditions.Add Type:=xlTextString, String:=txt, TextOperator:=xlContains
        Data_out.FormatConditions(Data_out.FormatConditions.Count).SetFirstPriority
        Data_out.FormatConditions(1).Font.Underline = xlUnderlineStyleSingle
    Next
    
    arr_warning = Array("!!!!")
    For Each txt In arr_warning
        Data_out.FormatConditions.Add Type:=xlTextString, String:=txt, TextOperator:=xlContains
        Data_out.FormatConditions(Data_out.FormatConditions.Count).SetFirstPriority
        Data_out.FormatConditions(1).Font.Color = -16751204
    Next

    With Data_out.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    With Data_out.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Data_out.Borders(xlDiagonalDown).LineStyle = xlNone
    Data_out.Borders(xlDiagonalUp).LineStyle = xlNone
    With Data_out.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Data_out.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Data_out.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Data_out.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Data_out.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Data_out.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Data_out.Borders(xlDiagonalDown).LineStyle = xlNone
    Data_out.Borders(xlDiagonalUp).LineStyle = xlNone
    With Data_out.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Data_out.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Data_out.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Data_out.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Data_out.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Data_out.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Data_out.Font
        .Name = "ISOCPEUR"
        .FontStyle = "обычный"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    h = 1
    If h = 0 Then
        For i = 1 To n_row
            For j = 1 To n_col
                n = Data_out.Cells(i, j)
                On Error Resume Next
                If IsNumeric(Data_out.Cells(i, j)) And Data_out.Cells(i, j) <> 0 Then
                    With Data_out.Cells(i, j)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = False
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = True
                        .ReadingOrder = xlContext
                    End With
                Else
                    With Data_out.Cells(i, j)
                        .HorizontalAlignment = xlCenter
                        .VerticalAlignment = xlCenter
                        .WrapText = True
                        .Orientation = 0
                        .AddIndent = False
                        .IndentLevel = 0
                        .ShrinkToFit = False
                        .ReadingOrder = xlContext
                    End With
                End If
            Next j
        Next i
    Else
        With Data_out
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
        End With
    End If
    FormatFont = True
End Function


Function GetShortNameForGOST(ByVal gost As String) As String
    If IsEmpty(name_gost) Then r = ReadMetall()
    For i = 1 To UBound(name_gost, 1)
        If name_gost(i, 1) = gost Then
            GetShortNameForGOST = " " & name_gost(i, 3) & " "
            Exit Function
        End If
    Next
End Function

Function GetSizeSheet(ByVal objLst As Variant) As Variant
    Dim out(2)
    Dim rc As Long
    On Error Resume Next
    t = objLst.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If Not IsEmpty(t) Then
        out(2) = t
    Else
        out(2) = 1
    End If
    
    On Error Resume Next
    t = objLst.Cells.Find("*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    If Not IsEmpty(t) Then
        out(1) = t
    Else
        out(1) = 1
    End If
    GetSizeSheet = out
    Erase out
End Function

Function SheetExist(ByVal NameSheet As String) As Boolean
    On Error Resume Next
    Dim objWh As Excel.Worksheet
    Dim NameLst As String
    For Each objWh In ThisWorkbook.Worksheets
        NameLst = objWh.Name
        If NameLst = NameSheet Then
            SheetExist = True
            Exit Function
        End If
    Next objWh
    SheetExist = False
End Function

Function Var2Str(ByVal var As Variant) As String
    If IsNumeric(var) Then
        If var = 0 Then
            Var2Str = ""
        Else
            Var2Str = CStr(var)
        End If
    Else
        Var2Str = var
    End If
End Function

Function list2CSV(ByRef ra As Variant, ByVal CSVfilename As String, Optional ByVal ColumnsSeparator$ = ";", _
                   Optional ByVal RowsSeparator$ = vbNewLine) As String
    If ra.Cells.Count = 1 Then Range2CSV = ra.Value & RowsSeparator$: Exit Function
    If ra.Areas.Count > 1 Then
        Dim ar As Range
        For Each ar In ra.Areas
            Range2CSV = Range2CSV & Range2CSV(ar, ColumnsSeparator$, RowsSeparator$)
        Next ar
        Exit Function
    End If
    arr = ra.Value
    buffer$ = ""
    For i = 1 To UBound(arr, 1)
        txt = ""
        For j = 1 To UBound(arr, 2)
            If arr(i, j) <> "" Then txt = txt & ColumnsSeparator$ & arr(i, j)
        Next j
        If txt <> "" Then
            Range2CSV = Range2CSV & Mid(txt, Len(ColumnsSeparator$) + 1) & RowsSeparator$
            If Len(Range2CSV) > 50000 Then buffer$ = buffer$ & Range2CSV: Range2CSV = ""
        End If
    Next i
    CSVtext$ = buffer$ & Range2CSV
    list2CSV = SaveTXTfile(CSVfilename$, CSVtext$)
End Function

Function SaveTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.CreateTextFile(filename, True)
    ts.Write txt: ts.Close
    SaveTXTfile = Err = 0
    Set ts = Nothing: Set FSO = Nothing
End Function
