Attribute VB_Name = "surf"
Option Compare Text
Option Base 1

Public Const surf_version As String = "0.9"

Public Const col_s_numb_zone As Integer = 1
Public Const col_s_name_zone As Integer = 2
Public Const col_s_area_zone As Integer = 3
Public Const col_s_perim_zone As Integer = 4
Public Const col_s_perimhole_zone As Integer = 5
Public Const col_s_h_zone As Integer = 6
Public Const col_s_freelen_zone As Integer = 7
Public Const col_s_walllen_zone As Integer = 8
Public Const col_s_doorlen_zone As Integer = 9
Public Const col_s_hpan_zone As Integer = 10
Public Const col_s_mpot_zone As Integer = 11
Public Const col_s_mpan_zone As Integer = 12
Public Const col_s_mwall_zone As Integer = 13
Public Const col_s_mcolumn_zone As Integer = 14
Public Const col_s_area_wall As Integer = 15
Public Const col_s_type As Integer = 16
Public Const col_s_mat_wall As Integer = 17
Public Const col_s_layer As Integer = 18
Public Const col_s_type_el As Integer = 19
Public Const col_s_type_pol As Integer = 20
Public Const col_s_area_pol As Integer = 21
Public Const col_s_perim_pol As Integer = 22
Public Const col_s_n_mun_zone As Integer = 23
Public Const col_s_mun_zone As Integer = 24
Public Const col_s_tipverh_l As Integer = 25
Public Const col_s_tipl_l As Integer = 26
Public Const col_s_tipniz_l As Integer = 27
Public Const col_s_tippl_l As Integer = 28
Public Const col_s_areaverh_l As Integer = 29
Public Const col_s_areal_l As Integer = 30
Public Const col_s_areaniz_l As Integer = 31
Public Const col_s_areapl_l As Integer = 32
Public Const max_s_col As Integer = 32
Public Const n_round_area As Integer = 1

Function DataIsOtd(ByVal array_in As Variant) As Boolean
    n_col = UBound(array_in, 2)
    If array_in(1, col_s_type) = "ЗОНА" And (n_col = col_s_layer Or n_col = col_s_mun_zone Or n_col = col_s_areapl_l) Then DataIsOtd = True Else DataIsOtd = False
End Function

Function FormatSpec_Rule(ByVal Data_out As Range) As Boolean
    Data_out.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "ISOCPEUR"
        .FontStyle = "обычный"
        .Size = 11
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
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Rows.AutoFit
    Columns("C:C").ColumnWidth = 60
    Columns("B:B").ColumnWidth = 40
    Columns("A:A").ColumnWidth = 60
    Rows("1:1").RowHeight = 45
    Range("A1:C1").Select
    Selection.Font.Bold = True
 End Function
 
Function FormatSpec_Pol(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
    CSVfilename$ = ThisWorkbook.path & "\list\Спец_" & ActiveWorkbook.ActiveSheet.Name & ".txt"
    n = list2CSV(Data_out, CSVfilename$)
    FormatSpec_Pol = True
End Function

Function FormatSpec_Ved(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
    Cells.UnMerge
    Cells.NumberFormat = "@"
    Range(Data_out.Cells(1, 1), Data_out.Cells(2, 1)).Merge
    Range(Data_out.Cells(1, 2), Data_out.Cells(2, 2)).Merge
    Range(Data_out.Cells(1, 3), Data_out.Cells(1, n_col - 1)).Merge
    Range(Data_out.Cells(1, n_col), Data_out.Cells(2, n_col)).Merge

    For i = 1 To n_row
        If InStr(Data_out.Cells(i, 1), "Общяя площадь отделки, кв.м.") > 0 Then
            n_all = n_row
            n_row = i - 1
            n_start_all = i
        End If
    Next i

    n_start = 3
    n_end = 3
    For i = 3 To n_row
        temp = Trim(Data_out.Cells(i, 1))
        If temp = Empty Then n_end = i
        If temp <> Empty Then
            If n_end > n_start Then
                Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, 1)).Merge
                Range(Data_out.Cells(n_start, 2), Data_out.Cells(n_end, 2)).Merge
                Range(Data_out.Cells(n_start, n_col), Data_out.Cells(n_end, n_col)).Merge
            End If
            n_start = i
        End If
        If i = n_row And temp = Empty Then
            n_end = i
            Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, 1)).Merge
            Range(Data_out.Cells(n_start, 2), Data_out.Cells(n_end, 2)).Merge
            Range(Data_out.Cells(n_start, n_col), Data_out.Cells(n_end, n_col)).Merge
        End If
    Next i
    
    For n_c = 3 To n_col - 1
        n_start = 3
        n_end = 3
        For i = 3 To n_row
            temp = Trim(Data_out.Cells(i, n_c))
            If temp = Empty Then n_end = i
            If temp <> Empty Then
                If n_end > n_start Then Range(Data_out.Cells(n_start, n_c), Data_out.Cells(n_end, n_c)).Merge
                n_start = i
            End If
            If i = n_row And temp = Empty Then
                n_end = i
                Range(Data_out.Cells(n_start, n_c), Data_out.Cells(n_end, n_c)).Merge
            End If
        Next i
    Next n_c

    For n_c = 3 To n_col - 1
        If InStr(Data_out.Cells(2, n_c), "Площадь") = 0 And InStr(Data_out.Cells(2, n_c), "Высота") = 0 Then
            temp_1 = Data_out.Cells(n_start, n_c).Value
            n_start = 3
            n_end = 3
            For i = 3 To n_row
                temp_2 = Data_out.Cells(i, n_c).Value
                If temp_1 <> temp_2 And temp_2 <> Empty Then
                    temp_1 = temp_2
                    If n_end > n_start Then Range(Data_out.Cells(n_start, n_c), Data_out.Cells(n_end, n_c)).Merge
                    n_start = i
                Else
                    n_end = i
                End If
                If i = n_row And temp_1 = temp_2 And temp_2 <> Empty Then
                    n_end = i
                    Range(Data_out.Cells(n_start, n_c), Data_out.Cells(n_end, n_c)).Merge
                End If
            Next i
        End If
    Next n_c
    
    Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_start_all, 4)).Merge
    For i = n_start_all + 1 To n_all
        Range(Data_out.Cells(i, 1), Data_out.Cells(i, 3)).Merge
    Next i
    
    s_mat = 200
    s_ar = 30
    s1 = 30
    s2 = 120
    sp = 100
    sall = s1 + s2 + s_mat * 2 + s_ar * 2 + sp
    

    If Data_out.Cells(2, 7).Value = "Колонн" Then
        sall = sall + s_mat + s_ar
        If Data_out.Cells(2, 9).Value = "Низа стен/колонн" Then
            sall = sall + s_mat + s_ar * 2
        End If
    Else
        If Data_out.Cells(2, 7).Value = "Низа стен/колонн" Then
            sall = sall + s_mat + s_ar * 2
        End If
    End If
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Data_out.Borders(xlDiagonalDown).LineStyle = xlNone
    Data_out.Borders(xlDiagonalUp).LineStyle = xlNone
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlDiagonalDown).LineStyle = xlNone
    Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlDiagonalUp).LineStyle = xlNone
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlDiagonalDown).LineStyle = xlNone
    Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlDiagonalUp).LineStyle = xlNone
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlDiagonalDown).LineStyle = xlNone
    Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlDiagonalUp).LineStyle = xlNone
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_all, 4)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    
    
    With Data_out.Font
        .Name = "ISOCPEUR"
        .FontStyle = "обычный"
        .Size = 10
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
    
     With Range(Data_out.Cells(1, 1), Data_out.Cells(2, n_col)).Font
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
    
    Range(Data_out.Cells(1, 1), Data_out.Cells(1, 1)).ColumnWidth = s1
    koeff = s1 / Range(Data_out.Cells(1, 1), Data_out.Cells(1, 1)).Width
    
    dblPoints = Application.CentimetersToPoints(1)
    Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, n_col)).RowHeight = dblPoints * 0.9
    Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, n_col)).Rows.AutoFit
    Range(Data_out.Cells(1, 1), Data_out.Cells(1, n_col)).RowHeight = dblPoints * 0.5
    
    Range(Data_out.Cells(1, 1), Data_out.Cells(1, 1)).ColumnWidth = s1 * koeff
    Range(Data_out.Cells(1, 2), Data_out.Cells(1, 2)).ColumnWidth = s2 * koeff
    Range(Data_out.Cells(1, 3), Data_out.Cells(1, 3)).ColumnWidth = s_mat * koeff
    Range(Data_out.Cells(1, 4), Data_out.Cells(1, 4)).ColumnWidth = s_ar * koeff
    Range(Data_out.Cells(1, 5), Data_out.Cells(1, 5)).ColumnWidth = s_mat * koeff
    Range(Data_out.Cells(1, 6), Data_out.Cells(1, 6)).ColumnWidth = s_ar * koeff
    If Data_out.Cells(2, 7).Value = "Колонн" Then
        Range(Data_out.Cells(1, 7), Data_out.Cells(1, 7)).ColumnWidth = s_mat * koeff
        Range(Data_out.Cells(1, 8), Data_out.Cells(1, 8)).ColumnWidth = s_ar * koeff
        If Data_out.Cells(2, 9).Value = "Низа стен/колонн" Then
            Range(Data_out.Cells(1, 9), Data_out.Cells(1, 9)).ColumnWidth = s_mat * koeff
            Range(Data_out.Cells(1, 10), Data_out.Cells(1, 10)).ColumnWidth = s_ar * koeff
            Range(Data_out.Cells(1, 11), Data_out.Cells(1, 11)).ColumnWidth = s_ar * koeff
        End If
    Else
        If Data_out.Cells(2, 7).Value = "Низа стен/колонн" Then
            Range(Data_out.Cells(1, 7), Data_out.Cells(1, 7)).ColumnWidth = s_mat * koeff
            Range(Data_out.Cells(1, 8), Data_out.Cells(1, 8)).ColumnWidth = s_ar * koeff
            Range(Data_out.Cells(1, 9), Data_out.Cells(1, 9)).ColumnWidth = s_ar * koeff
        End If
    End If
    For n_c = 1 To n_col
        g = Data_out.Cells(2, n_c).Value
        If InStr(g, "Площадь") Or InStr(g, "Высота") Then
            Data_out.Cells(2, n_c).Orientation = 90
            Range(Data_out.Cells(3, n_c), Data_out.Cells(n_row, n_c)).Font.Size = 11
            Range(Data_out.Cells(3, n_c), Data_out.Cells(n_row, n_c)).ShrinkToFit = True
        End If
        g = Data_out.Cells(1, n_c).Value
        If InStr(g, "Номер") Then Data_out.Cells(1, n_c).Orientation = 90
    Next n_c
    Range(Data_out.Cells(1, n_col), Data_out.Cells(1, n_col)).ColumnWidth = sp * koeff
    FormatSpec_Ved = True
End Function

Function GetRules(ByVal nm As String) As Variant
    nm_rule = ""
    nm = Split(nm, "_")(0)
    listsheet = GetListOfSheet(ThisWorkbook)
    For Each nlist In listsheet
        spec_type = SpecGetType(nlist)
        name_list = Split(nlist, "_")
        If spec_type = 10 Then
            If name_list(0) = nm Then nm_rule = nlist
        End If
    Next
    If nm_rule <> "" Then
        Set rule_sheet = Application.ThisWorkbook.Sheets(nm_rule)
        lsize = GetSizeSheet(rule_sheet)
        n_row = lsize(1)
        n_col = lsize(2)
        If n_row = 1 Then n_row = 2
        Set Data_out = rule_sheet.Range(rule_sheet.Cells(1, 1), rule_sheet.Cells(n_row, n_col))
        Worksheets(nm_rule).Activate
        r = FormatClear()
        r = FormatSpec_Rule(Data_out)
        Dim rules: ReDim rules(n_row - 1, 3)
        For i = 2 To n_row
            rules(i - 1, 1) = Data_out(i, 1)
            rules(i - 1, 2) = Data_out(i, 2)
            rules(i - 1, 3) = Data_out(i, 3)
        Next i
        GetRules = rules
        Erase rules
    Else
        GetRules = Empty
        r = NewListRules(nm)
        MsgBox ("Не найден лист с правилами отделки (оканчивается на '_правила')")
    End If
End Function

Function NewListRules(ByVal nm As String) As Boolean
    ThisWorkbook.Worksheets.Add.Name = nm & "_правила"
    Worksheets(nm & "_правила").Activate
    Cells(1, 1).Value = "Имя многослойной конструкции (целиком или часть имени)"
    Cells(1, 2).Value = "Слой"
    Cells(1, 3).Value = "Черновая отделка"
    
    Cells(2, 1).Value = "ЖБ"
    Cells(2, 2).Value = "Колонны"
    Cells(2, 3).Value = "Затирка, шпатлёвка ж/б колонн"
    
    Cells(3, 1).Value = "П1"
    Cells(3, 2).Value = "Потолки"
    Cells(3, 3).Value = "Армстронг; Без отделки"
    
    Columns("A:A").ColumnWidth = 50
    Columns("B:B").ColumnWidth = 30
    Columns("C:C").ColumnWidth = 60
    Rows("1:1").EntireRow.AutoFit
End Function

Function AddRules(ByVal nm As String, ByVal add_rule As Variant) As Boolean
    nm_rule = ""
    nm = Split(nm, "_")(0)
    If UBound(add_rule, 1) < 1 Then Exit Function
    If UBound(Split(add_rule(0), ";"), 1) < 1 Then Exit Function
    listsheet = GetListOfSheet(ThisWorkbook)
    For Each nlist In listsheet
        spec_type = SpecGetType(nlist)
        name_list = Split(nlist, "_")
        If spec_type = 10 Then
            If name_list(0) = nm Then nm_rule = nlist
        End If
    Next
    If nm_rule <> "" Then
        Set rule_sheet = Application.ThisWorkbook.Sheets(nm_rule)
        lsize = GetSizeSheet(rule_sheet)
        n_row_sheet = lsize(1) + 1
        n_col = lsize(2)
        n_row = UBound(add_rule, 1)
        Worksheets(nm_rule).Activate
        For i = n_row_sheet To n_row_sheet + n_row
            t = add_rule(i - n_row_sheet)
            m = Split(t, ";")
            rule_sheet.Cells(i, 1) = m(0)
            rule_sheet.Cells(i, 2) = m(1)
            rule_sheet.Cells(i, 3) = m(2)
        Next i
        Set Data_out = rule_sheet.Range(rule_sheet.Cells(1, 1), rule_sheet.Cells(n_row_sheet + n_row, n_col))
        r = FormatClear()
        r = FormatSpec_Rule(Data_out)
        AddRules = True
    Else
        AddRules = False
        r = NewListRules(nm)
        MsgBox ("Не найден лист с правилами отделки (оканчивается на '_правила')")
    End If
End Function

Function ReadVed(ByVal lastfilespec As String) As Variant
    lastfilespec = Split(lastfilespec, "_")(0)
    out_data = ReadFile(lastfilespec & ".txt")
    If Not DataIsOtd(out_data) Then
        MsgBox ("Неверный формат файла")
        ReadVed = Empty
        Exit Function
    End If
    rules = GetRules(lastfilespec)
    Set add_rule = CreateObject("Scripting.Dictionary")
    add_rule.comparemode = 1
    If IsEmpty(rules) Or IsEmpty(out_data) Then
        ReadVed = Empty
        Exit Function
    End If
    n_row_a = UBound(out_data, 1) - 2
    n_col_a = UBound(out_data, 2)
    n_zone = 99
    For i = 1 To n_row_a
        If out_data(i, col_s_numb_zone) = 0 Then
            out_data(i, col_s_numb_zone) = n_zone
        Else
            n_zone = Var2Str(out_data(i, col_s_numb_zone))
            out_data(i, col_s_numb_zone) = n_zone
        End If
        If n_col_a > col_s_mun_zone Then ' Если есть лестницы
            out_data(i, col_s_tipverh_l) = Var2Str(out_data(i, col_s_tipverh_l))
            out_data(i, col_s_tipniz_l) = Var2Str(out_data(i, col_s_tipniz_l))
            out_data(i, col_s_tippl_l) = Var2Str(out_data(i, col_s_tippl_l))
            out_data(i, col_s_tipl_l) = Var2Str(out_data(i, col_s_tipl_l))
        End If
        If out_data(i, col_s_type) = "СТЕНА" Then
            layer = out_data(i, col_s_layer)
            material = out_data(i, col_s_mat_wall)
            name_mat = NameMat(layer, material, rules)
            If InStr(name_mat, "ОШИБКА") > 0 Then
                If Not add_rule.exists(name_mat) Then add_rule.Item(name_mat) = name_mat
                out_data(i, col_s_mat_wall) = "ОШИБКА"
            Else
                out_data(i, col_s_mat_wall) = name_mat
            End If
        End If
        If n_col_a > col_s_layer Then 'Если есть пол или потолок
            out_data(i, col_s_type_pol) = Var2Str(out_data(i, col_s_type_pol))
            If out_data(i, col_s_type) = "ОБЪЕКТ" And out_data(i, col_s_type_el) = "Потолок" Then
                layer = "Потолок"
                material = out_data(i, col_s_type_pol)
                out_data(i, col_s_type_pol) = NameMat(layer, material, rules)
            End If
            If out_data(i, col_s_mun_zone) = 1 Then
                n_mun = Var2Str(out_data(i, col_s_n_mun_zone))
                If n_mun <> "" Then out_data(i, col_s_numb_zone) = Var2Str(out_data(i, col_s_n_mun_zone))
            End If
        End If
        For j = 1 To n_col_a
            If IsNumeric(out_data(i, j)) = 0 Then
                If out_data(i, j) = "" Then out_data(i, j) = 0
            End If
        Next j
    Next i
    Dim pos_out: ReDim pos_out(2)
    If add_rule.Count = 0 Then
        pos_out(1) = out_data
        pos_out(2) = rules
    Else
        r = AddRules(lastfilespec, add_rule.Keys)
        pos_out(1) = Empty
        pos_out(2) = Empty
    End If
    ReadVed = pos_out
End Function

Function NameMat(ByVal layer As String, ByVal material As String, ByRef rules As Variant) As String
    name_m = ""
    flag = 0
    For i = 1 To UBound(rules, 1)
        m = rules(i, 1)
        l = rules(i, 2)
        If layer = l Or layer = "" Then
            If m = material Then
                name_m = rules(i, 3)
                flag = flag + 1
            Else
                If InStr(material, m) > 0 Then
                    name_m = rules(i, 3)
                    flag = flag + 1
                End If
            End If
        End If
    Next i
    If flag = 1 Then
        If InStr(name_m, "ез отделк") > 0 Then
            If InStr(name_m, ";") > 0 Then name_m = Split(name_m, ";")(0)
            name_m = name_m + "="
        End If
        NameMat = name_m
    Else
        NameMat = material + ";" + layer + ";ОШИБКА"
        If flag > 1 Then
            MsgBox ("Несколько правил для одного материала - " + material + " слой" + layer)
        End If
    End If
End Function

Function ReadPol(ByVal lastfilespec As String) As Variant
    lastfilespec = Split(lastfilespec, "_")(0)
    out_data = ReadFile(lastfilespec & ".txt")
    If IsEmpty(out_data) Then
        ReadPol = Empty
        Exit Function
    End If
    If Not DataIsOtd(out_data) Then
        MsgBox ("Неверный формат файла")
        ReadPol = Empty
        Exit Function
    End If
    n_row_a = UBound(out_data, 1) - 1
    n_col_a = UBound(out_data, 2)
    If n_col_a <= col_s_layer Then
        ReadPol = Empty
        Exit Function
    End If
    Dim add_pol: ReDim add_pol(1, 1)
    If UBound(out_data, 2) >= col_s_tipverh_l Then
        ReDim add_pol(col_s_areapl_l, n_row_a)
        n_add = 0
    End If
    n_zone = 9999
    For i = 1 To n_row_a
        If out_data(i, col_s_numb_zone) = 0 Then
            out_data(i, col_s_numb_zone) = n_zone
        Else
            n_zone = Var2Str(out_data(i, col_s_numb_zone))
            out_data(i, col_s_numb_zone) = n_zone
        End If
        If out_data(i, col_s_numb_zone) = 0 Then
            out_data(i, col_s_numb_zone) = n_zone
        Else
            n_zone = Var2Str(out_data(i, col_s_numb_zone))
            out_data(i, col_s_numb_zone) = n_zone
        End If
        If n_col_a >= col_s_type_el Then out_data(i, col_s_type_pol) = Var2Str(out_data(i, col_s_type_pol))
        If n_col_a >= col_s_tipverh_l Then
            out_data(i, col_s_tipverh_l) = Var2Str(out_data(i, col_s_tipverh_l))
            out_data(i, col_s_tipniz_l) = Var2Str(out_data(i, col_s_tipniz_l))
            out_data(i, col_s_tippl_l) = Var2Str(out_data(i, col_s_tippl_l))
            out_data(i, col_s_tipl_l) = Var2Str(out_data(i, col_s_tipl_l))
            If out_data(i, col_s_tipverh_l) <> "" Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_tipniz_l) <> "" Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_tippl_l) <> "" Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_tipl_l) <> "" Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_type_el) = "Лестница" Then
                For k = col_s_tipverh_l To col_s_tippl_l
                    n_add = n_add + 1
                    add_pol(col_s_numb_zone, n_add) = out_data(i, col_s_numb_zone)
                    add_pol(col_s_type, n_add) = "ОБЪЕКТ"
                    add_pol(col_s_type_el, n_add) = "Пол"
                    add_pol(col_s_type_pol, n_add) = out_data(i, k)
                    add_pol(col_s_area_pol, n_add) = out_data(i, k + 4)
                    add_pol(col_s_perim_pol, n_add) = (Sqr(out_data(i, k + 4)) * 2 + 0.5) * 1000
                Next k
            End If
        End If
        For j = 1 To n_col_a
            If Not IsNumeric(out_data(i, j)) Then
                If out_data(i, j) = "" Then out_data(i, j) = 0
            End If
        Next j
    Next i
    zone = ArraySelectParam(out_data, "ЗОНА", col_s_type)
    out_data = ArraySelectParam(out_data, "ОБЪЕКТ", col_s_type, "Пол", col_s_type_el)
    If n_add > 0 Then
        ReDim Preserve add_pol(col_s_areapl_l, n_add)
        add_pol = ArrayTranspose(add_pol)
        out_data = ArrayCombine(out_data, add_pol)
    End If
    out_data = ArrayCombine(out_data, zone)
    ReadPol = out_data
End Function

Function Spec_POL(ByRef out_data As Variant) As Variant
    pol = ArraySelectParam(out_data, "Пол", col_s_type_el)
    un_pol = ArrayUniqValColumn(pol, col_s_type_pol)
    n_type_pol = UBound(un_pol, 1)
    Dim pos_out: ReDim pos_out(n_type_pol, 4)
    n_row_tot = 0
    For i = 1 To n_type_pol
        un_pol(i) = ConvTxt2Num(un_pol(i))
    Next i
    un_pol = ArraySort(un_pol, 1)
    For i = 1 To n_type_pol
        un_pol(i) = Var2Str(un_pol(i))
    Next i
    For j = 1 To n_type_pol
        type_pol = un_pol(j)
        t_pol = ArraySelectParam(pol, type_pol, col_s_type_pol)
        t_un_zone = ArrayUniqValColumn(t_pol, col_s_numb_zone)
        pol_area = 0
        pol_perim = 0
        For i = 1 To UBound(t_pol, 1)
            pol_area = pol_area + t_pol(i, col_s_area_pol)
            pol_perim = pol_perim + t_pol(i, col_s_perim_pol)
        Next i
        t_zone = ""
        For i = 1 To UBound(t_un_zone, 1) - 1
            t_zone = t_zone + t_un_zone(i) + ", "
        Next i
        t_zone = t_zone + t_un_zone(i)
        pos_out(j, 1) = type_pol
        pos_out(j, 2) = t_zone
        pos_out(j, 3) = Round_w(pol_area * k_zap_total, n_round_area)
        pos_out(j, 4) = Round_w(pol_perim / 1000, n_round_area)
    Next j
    Spec_POL = pos_out
End Function

Function Spec_VED(ByRef all_data As Variant) As Variant
    out_data = all_data(1)
    rules = all_data(2)
    Erase all_data
    If IsEmpty(out_data) Or IsEmpty(rules) Then
        Spec_VED = Empty
        Exit Function
    End If
    Set zone = CreateObject("Scripting.Dictionary")
    Set materials = CreateObject("Scripting.Dictionary")
    zone.comparemode = 1
    materials.comparemode = 1
    n_un_mat = 0
    un_n_zone = ArrayUniqValColumn(out_data, col_s_numb_zone)
    zone.Item("list") = un_n_zone
    is_pan = False
    is_column = False
    n_row_tot = 0
    For i = 1 To UBound(un_n_zone, 1)
        un_n_zone(i) = ConvTxt2Num(un_n_zone(i))
    Next i
    un_n_zone = ArraySort(un_n_zone, 1)
    For i = 1 To UBound(un_n_zone, 1)
        un_n_zone(i) = Var2Str(un_n_zone(i))
    Next i
    spec_type = 1 'И лестницы, и полы
    If UBound(out_data, 2) < col_s_tipverh_l Then spec_type = 2 'Без лестниц
    If UBound(out_data, 2) < col_s_type_el Then spec_type = 3 'Только зоны
    For Each num In un_n_zone
        n_row_p = 0
        n_row_w = 0
        If IsNumeric(num) Then num = CStr(num)
        zone_el = ArraySelectParam(out_data, num, col_s_numb_zone, "ЗОНА", col_s_type)
        If Not IsEmpty(zone_el) Then
            If UBound(zone_el, 1) > 1 Then MsgBox ("Зоны с одинаковыми именами считаются не правильно - " + num)
            zone.Item(num + ";name") = zone_el(1, col_s_name_zone)
            area_total = zone_el(1, col_s_area_zone)
            perim_total = zone_el(1, col_s_perim_zone) / 1000
            perim_hole = zone_el(1, col_s_perimhole_zone) / 1000
            h_zone = zone_el(1, col_s_h_zone) / 1000
            free_len = zone_el(1, col_s_freelen_zone) / 1000
            h_pan = zone_el(1, col_s_hpan_zone) / 1000
            fin_pot = zone_el(1, col_s_mpot_zone)
            fin_pan = zone_el(1, col_s_mpan_zone)
            fin_wall = zone_el(1, col_s_mwall_zone)
            fin_column = zone_el(1, col_s_mcolumn_zone)
            wall = ArraySelectParam(out_data, num, col_s_numb_zone, "СТЕНА", col_s_type)
            n_wall = UBound(wall, 1)
            wall_len = 0
            door_len = 0
            For i = 1 To n_wall
                door_len = door_len + wall(i, col_s_doorlen_zone) / 1000
                wall_len = wall_len + wall(i, col_s_walllen_zone) / 1000
            Next i
            '----------------------
            '        КОЛОННЫ
            '----------------------
            'Добавить возможность выбора наличия отверстий
            perim_column_in_wall = perim_total - wall_len - free_len - perim_hole
            perim_column = perim_hole
            perim_column_total = perim_column + perim_column_in_wall
            area_pan_column = perim_column_total * h_pan
            area_pan_column = Round(area_pan_column * 10, 0) / 10
            area_column = perim_column_total * (h_zone - h_pan)
            area_column = Round(area_column * 10, 0) / 10
            name_mat_column = "ЖБ"
            c = NameMat("Колонны", name_mat_column, rules)
            cn = ""
            pn = ""
            If InStr(c, "=") > 0 Then
                cn = Left(c, Len(c) - 1)
                pn = Left(c, Len(c) - 1)
            Else
                cn = c + "; " + fin_column
                pn = c + "; " + fin_pan
            End If
            zone.Item(num + ";c") = name_mat_column
            zone.Item(num + ";cn") = cn
            zone.Item(num + ";ca;") = area_column
            
            n_mat = cn
            area_mat = area_column
            If area_mat > 0 Then
                If InStr(n_mat, ";") > 0 Then
                    mat_1 = Trim(Split(n_mat, ";")(0))
                    mat_2 = Trim(Split(n_mat, ";")(1))
                Else
                    mat_1 = n_mat
                    mat_2 = ""
                End If
                If materials.exists(mat_1) Then
                    materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                Else
                    n_un_mat = n_un_mat + 1
                    materials.Item(mat_1) = mat_1
                    materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                End If
                If mat_2 <> "" Then
                    If materials.exists(mat_2) Then
                        materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                    Else
                        n_un_mat = n_un_mat + 1
                        materials.Item(mat_2) = mat_2
                        materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                    End If
                End If
            End If
                
            If area_pan_column > 0 Then
                If zone.exists(num + ";pn;" + c) Then
                    zone.Item(num + ";pa;" + c) = zone.Item(num + ";pa;" + c) + area_pan_column
                Else
                    zone.Item(num + ";pn;" + c) = pn
                    zone.Item(num + ";pa;" + c) = area_pan_column
                End If
                zone.Item(num + ";ph;" + c) = h_pan
                
                n_mat = pn ' + " на высоту h=" + Str(h_pan) + "м."
                area_mat = area_pan_column
                If area_mat > 0 Then
                    If InStr(n_mat, ";") > 0 Then
                        mat_1 = Trim(Split(n_mat, ";")(0))
                        mat_2 = Trim(Split(n_mat, ";")(1))
                    Else
                        mat_1 = n_mat
                        mat_2 = ""
                    End If
                    If materials.exists(mat_1) Then
                        materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                    Else
                        n_un_mat = n_un_mat + 1
                        materials.Item(mat_1) = mat_1
                        materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                    End If
                    If mat_2 <> "" Then
                        If materials.exists(mat_2) Then
                            materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                        Else
                            n_un_mat = n_un_mat + 1
                            materials.Item(mat_2) = mat_2
                            materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                        End If
                    End If
                End If
            End If
            If area_column > 0 Then is_column = True
            
            '----------------------
            '        СТЕНЫ
            '----------------------
            un_wall = ArrayUniqValColumn(wall, col_s_mat_wall)
            zone.Item(num + ";w") = un_wall
            For Each w In un_wall
                wall_by_key = ArraySelectParam(wall, w, col_s_mat_wall)
                n_wall = UBound(wall_by_key, 1)
                wall_len = 0
                door_len = 0
                wall_area = 0
                For i = 1 To n_wall
                    door_len = door_len + wall_by_key(i, col_s_doorlen_zone) / 1000
                    wall_len = wall_len + wall_by_key(i, col_s_walllen_zone) / 1000
                    wall_area = wall_area + wall_by_key(i, col_s_area_wall)
                Next i
                pan_area = (wall_len - door_len) * h_pan
                wall_area = wall_area - pan_area
                wn = ""
                pn = ""
                If InStr(w, "=") > 0 Then
                    wn = Left(w, Len(w) - 1)
                    pn = Left(w, Len(w) - 1)
                Else
                    wn = w + "; " + fin_wall
                    pn = w + "; " + fin_pan
                End If
                zone.Item(num + ";wn;" + w) = wn
                zone.Item(num + ";wa;" + w) = wall_area
                
                n_mat = wn
                area_mat = wall_area
                If area_mat > 0 Then
                    If InStr(n_mat, ";") > 0 Then
                        mat_1 = Trim(Split(n_mat, ";")(0))
                        mat_2 = Trim(Split(n_mat, ";")(1))
                    Else
                        mat_1 = n_mat
                        mat_2 = ""
                    End If
                    If materials.exists(mat_1) Then
                        materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                    Else
                        n_un_mat = n_un_mat + 1
                        materials.Item(mat_1) = mat_1
                        materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                    End If
                    If mat_2 <> "" Then
                        If materials.exists(mat_2) Then
                            materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                        Else
                            n_un_mat = n_un_mat + 1
                            materials.Item(mat_2) = mat_2
                            materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                        End If
                    End If
                End If

                If pan_area > 0 Then
                    If zone.exists(num + ";pn;" + w) Then
                        zone.Item(num + ";pa;" + w) = zone.Item(num + ";pa;" + w) + pan_area
                    Else
                        zone.Item(num + ";pa;" + w) = pan_area
                        zone.Item(num + ";pn;" + w) = pn
                    End If
                    zone.Item(num + ";ph;" + w) = h_pan
                    n_mat = pn ' + " на высоту h=" + Str(h_pan) + "м."
                    area_mat = pan_area
                    If area_mat > 0 Then
                        If InStr(n_mat, ";") > 0 Then
                            mat_1 = Trim(Split(n_mat, ";")(0))
                            mat_2 = Trim(Split(n_mat, ";")(1))
                        Else
                            mat_1 = n_mat
                            mat_2 = ""
                        End If
                        If materials.exists(mat_1) Then
                            materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                        Else
                            n_un_mat = n_un_mat + 1
                            materials.Item(mat_1) = mat_1
                            materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                        End If
                        If mat_2 <> "" Then
                            If materials.exists(mat_2) Then
                                materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                            Else
                                n_un_mat = n_un_mat + 1
                                materials.Item(mat_2) = mat_2
                                materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                            End If
                        End If
                    End If
                End If
            Next
            n_row_w = n_row_w + UBound(un_wall, 1)
            If h_pan > 0 Then
                is_pan = True
                If area_pan_column > 0 Then
                    un_panel = ArrayUniqValColumn(ArrayCombine(Array(c), un_wall), 1)
                    zone.Item(num + ";p") = un_panel
                    n_row_w = n_row_w + 1
                Else
                    zone.Item(num + ";p") = un_wall
                End If
            End If
            '----------------------
            '        ПОТОЛКИ
            '----------------------
            n_row_p = 0
            If spec_type < 3 Then
                area_total_pot = 0
                pot = ArraySelectParam(out_data, num, col_s_numb_zone, "Потолок", col_s_type_el)
                un_pot = ArrayUniqValColumn(pot, col_s_type_pol)
                zone.Item(num + ";pot") = un_pot
                If Not IsEmpty(un_pot) Then
                    For Each p In un_pot
                        pot_by_type = ArraySelectParam(pot, p, col_s_type_pol)
                        n_pot = UBound(pot_by_type, 1)
                        pot_area = 0
                        pot_perim = 0
                        For i = 1 To n_pot
                            pot_area = pot_area + pot_by_type(i, col_s_area_pol)
                            pot_perim = pot_perim + pot_by_type(i, col_s_perim_pol) / 1000
                            area_total_pot = area_total_pot + pot_area
                        Next
                        potn = ""
                        If InStr(p, "=") > 0 Then
                            potn = Left(p, Len(p) - 1)
                        Else
                            potn = p + "; " + fin_pot
                        End If
                        zone.Item(num + ";potn;" + p) = potn
                        zone.Item(num + ";pota;" + p) = pot_area
                        zone.Item(num + ";potp;" + p) = pot_perim
                        n_mat = potn
                        area_mat = pot_area
                        If area_mat > 0 Then
                            If InStr(n_mat, ";") > 0 Then
                                mat_1 = Trim(Split(n_mat, ";")(0))
                                mat_2 = Trim(Split(n_mat, ";")(1))
                            Else
                                mat_1 = n_mat
                                mat_2 = ""
                            End If
                            If materials.exists(mat_1) Then
                                materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                            Else
                                n_un_mat = n_un_mat + 1
                                materials.Item(mat_1) = mat_1
                                materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                            End If
                            If mat_2 <> "" Then
                                If materials.exists(mat_2) Then
                                    materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                                Else
                                    n_un_mat = n_un_mat + 1
                                    materials.Item(mat_2) = mat_2
                                    materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                                End If
                            End If
                        End If
                    Next
                    n_row_p = UBound(un_pot, 1)
                    diff_area_pot = area_total - area_total_pot
                Else
                    zone.Item(num + ";pot") = Array(fin_pot)
                    zone.Item(num + ";potn;" + fin_pot) = fin_pot
                    zone.Item(num + ";pota;" + fin_pot) = area_total
                    zone.Item(num + ";potp;" + fin_pot) = perim_total
                    n_mat = fin_pot
                    area_mat = area_total
                    If area_mat > 0 Then
                        If InStr(n_mat, ";") > 0 Then
                            mat_1 = Trim(Split(n_mat, ";")(0))
                            mat_2 = Trim(Split(n_mat, ";")(1))
                        Else
                            mat_1 = n_mat
                            mat_2 = ""
                        End If
                        If materials.exists(mat_1) Then
                            materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                        Else
                            n_un_mat = n_un_mat + 1
                            materials.Item(mat_1) = mat_1
                            materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                        End If
                        If mat_2 <> "" Then
                            If materials.exists(mat_2) Then
                                materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                            Else
                                n_un_mat = n_un_mat + 1
                                materials.Item(mat_2) = mat_2
                                materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                            End If
                        End If
                    End If
                End If

                If Abs(diff_area_pot) > 0.1 Then
                    MsgBox ("Разница площади помещения и потолка в помещении " & num & " = " & Str(diff_area_pot))
                End If
            Else
                zone.Item(num + ";pot") = Array(fin_pot)
                zone.Item(num + ";potn;" + fin_pot) = fin_pot
                zone.Item(num + ";pota;" + fin_pot) = area_total
                zone.Item(num + ";potp;" + fin_pot) = perim_total
                n_mat = fin_pot
                area_mat = area_total
                If area_mat > 0 Then
                    If InStr(n_mat, ";") > 0 Then
                        mat_1 = Trim(Split(n_mat, ";")(0))
                        mat_2 = Trim(Split(n_mat, ";")(1))
                    Else
                        mat_1 = n_mat
                        mat_2 = ""
                    End If
                    If materials.exists(mat_1) Then
                        materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                    Else
                        n_un_mat = n_un_mat + 1
                        materials.Item(mat_1) = mat_1
                        materials.Item(mat_1 + ";a") = materials.Item(mat_1 + ";a") + area_mat
                    End If
                    If mat_2 <> "" Then
                        If materials.exists(mat_2) Then
                            materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                        Else
                            n_un_mat = n_un_mat + 1
                            materials.Item(mat_2) = mat_2
                            materials.Item(mat_2 + ";a") = materials.Item(mat_2 + ";a") + area_mat
                        End If
                    End If
                End If
            End If
            '----------------------
            '        ПОЛЫ
            '----------------------
            If spec_type < 3 Then
                area_total_pol = 0
                pol = ArraySelectParam(out_data, num, col_s_numb_zone, "Пол", col_s_type_el)
                If Not IsEmpty(pol) Then
                    n_pol = UBound(pol, 1)
                    For i = 1 To n_pol
                        area_total_pol = area_total_pol + pol(i, col_s_area_pol)
                    Next
                    diff_area_pol = area_total - area_total_pol
                End If
            End If
            n_row_tot = n_row_tot + Application.WorksheetFunction.Max(1, n_row_p, n_row_w)
        Else
            MsgBox ("Номер зоны в элементе записан не правильно - " + num)
            Spec_VED = Empty
            Exit Function
        End If
    Next
    Erase out_data
    n_col_out = 7
    If is_pan Then n_col_out = n_col_out + 3
    If is_column Then n_col_out = n_col_out + 2
    Dim pos_out: ReDim pos_out(3 + n_row_tot + n_un_mat, n_col_out)
    pos_out(1, 1) = "Номер помещения"
    pos_out(1, 2) = "Наименование помещения"
    For i = 3 To n_col_out - 1
        pos_out(1, i) = "Ведомость отделки элементов интерьера"
    Next i
    pos_out(2, 3) = "Потолка"
    pos_out(2, 4) = "Площадь, кв.м."
    pos_out(2, 5) = "Стены и перегородки"
    pos_out(2, 6) = pos_out(2, 4)
    col_end = 6
    If is_column Then
        pos_out(2, 7) = "Колонн"
        pos_out(2, 8) = pos_out(2, 4)
        col_end = 8
    End If
    If is_pan Then
        pos_out(2, col_end + 1) = "Низа стен/колонн"
        pos_out(2, col_end + 2) = pos_out(2, 4)
        pos_out(2, col_end + 3) = "Высота, м."
    End If
    pos_out(1, n_col_out) = "Примечание"
    summ_area_pot = 0
    n_row = 3
    For Each num In un_n_zone
        n_row_p = n_row
        n_row_w = n_row
        n_row_c = n_row
        n_row_pan = n_row
        pos_out(n_row, 1) = num
        pos_out(n_row, 2) = zone.Item(num + ";name")
        '-- ПОТОЛКИ ---
        pot = zone.Item(num + ";pot")
        If Not IsEmpty(pot) Then
            For Each p In pot
                pos_out(n_row_p, 3) = zone.Item(num + ";potn;" + p)
                pos_out(n_row_p, 4) = Round_w(zone.Item(num + ";pota;" + p) * k_zap_total, n_round_area)
                summ_area_pot = summ_area_pot + pos_out(n_row_p, 4)
                n_row_p = n_row_p + 1
            Next p
        Else
            pos_out(n_row_p, 3) = "-"
            pos_out(n_row_p, 4) = "-"
            n_row_p = n_row_p + 1
        End If
        '-- СТЕНЫ ---
        wall = zone.Item(num + ";w")
        If Not IsEmpty(wall) Then
            For Each w In wall
                pos_out(n_row_w, 5) = zone.Item(num + ";wn;" + w)
                pos_out(n_row_w, 6) = Round_w(zone.Item(num + ";wa;" + w) * k_zap_total, n_round_area)
                n_row_w = n_row_w + 1
            Next w
        Else
            pos_out(n_row_w, 3) = "-"
            pos_out(n_row_w, 4) = "-"
            n_row_w = n_row_w + 1
        End If
         '-- КОЛОННЫ ---
        If is_column Then
            If zone.Item(num + ";ca;") > 0 Then
                pos_out(n_row_c, 7) = zone.Item(num + ";cn")
                pos_out(n_row_c, 8) = Round_w(zone.Item(num + ";ca;") * k_zap_total, n_round_area)
            Else
                pos_out(n_row_c, 7) = "-"
                pos_out(n_row_c, 8) = "-"
            End If
            n_row_c = n_row_c + 1
        End If
         '-- ПАНЕЛИ ---
        If is_pan Then
            If zone.exists(num + ";p") Then
                For Each p In zone.Item(num + ";p")
                    pos_out(n_row_pan, col_end + 1) = zone.Item(num + ";pn;" + p)
                    pos_out(n_row_pan, col_end + 2) = Round_w(zone.Item(num + ";pa;" + p) * k_zap_total, n_round_area)
                    pos_out(n_row_pan, col_end + 3) = zone.Item(num + ";ph;" + p)
                    n_row_pan = n_row_pan + 1
                Next p
            Else
                pos_out(n_row_pan, col_end + 1) = "-"
                pos_out(n_row_pan, col_end + 2) = "-"
                pos_out(n_row_pan, col_end + 3) = "-"
                n_row_pan = n_row_pan + 1
            End If
        End If
        n_row = Application.WorksheetFunction.Max(n_row_p - 1, n_row_w - 1, n_row_c - 1, n_row_pan - 1) + 1
    Next
    pos_out(n_row, 1) = "Общяя площадь отделки, кв.м."
    n_mat = 0
    For Each mat In ArraySort(materials.Keys())
        If InStr(mat, ";a") = 0 Then
            n_mat = n_mat + 1
            pos_out(n_row + n_mat, 1) = materials.Item(mat)
            pos_out(n_row + n_mat, 4) = Round_w(materials.Item(mat + ";a") * k_zap_total, n_round_area)
        End If
    Next
  
    MsgBox ("Сумма площади потолка " & CStr(summ_area_pot))
    Spec_VED = pos_out
End Function
