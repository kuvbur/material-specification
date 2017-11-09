Attribute VB_Name = "manual"
Option Compare Text
Option Base 1
Public Const manual_version As String = "2.2"
'-------------------------------------------------------
'Описание файла сортамента
Public Const col_gost_spec As Integer = 1
Public Const col_klass_spec As Integer = 2
Public Const col_diametr_spec As Integer = 3
Public Const col_area_spec As Integer = 4
Public Const col_weight_spec As Integer = 5
'-------------------------------------------------------
'Столбцы ручной спецификации (суффикс "_спец")
'Общие
Public Const col_man_subpos As Integer = 1
Public Const col_man_pos As Integer = 2
Public Const col_man_obozn As Integer = 3
Public Const col_man_naen As Integer = 4
Public Const col_man_qty As Integer = 5
Public Const col_man_weight As Integer = 6
Public Const col_man_prim As Integer = 7
Public Const col_man_komment As Integer = 18
'Арматура
Public Const col_man_length As Integer = 8
Public Const col_man_diametr As Integer = 9
Public Const col_man_klass As Integer = 10
'Прокат
Public Const col_man_pr_length As Integer = 11
Public Const col_man_pr_gost_pr As Integer = 12
Public Const col_man_pr_prof As Integer = 13
Public Const col_man_pr_type As Integer = 14
Public Const col_man_pr_st As Integer = 15
Public Const col_man_pr_okr As Integer = 16
Public Const col_man_pr_ogn As Integer = 17
Public Const max_col_man As Integer = col_man_komment
'-------------------------------------------------------
Public Const t_arm As Integer = 10
Public Const t_prokat As Integer = 20
Public Const t_mat As Integer = 30
Public Const t_mat_spc As Integer = 35
Public Const t_izd As Integer = 40
Public Const t_subpos As Integer = 45
Public Const t_else As Integer = 50
Public Const t_error As Integer = -1
Public Const t_sys As Integer = -10

Public gost2fklass, name_gost, reinforcement_specifications As Variant 'Разные массивы
Public pr_adress As Variant

Function CeilAlert(ByVal ceil As Variant, ByVal txt As String)
    ceil.AddComment (txt)
    ceil.Comment.Shape.TextFrame.AutoSize = True
    ceil.Comment.Visible = False
    With ceil.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Function

Function CeilSetValue(ByRef ceil As Variant, ByVal val As Variant, ByVal mode As String)
    If ceil.Value <> val Then
        If mode = "add" Then
            nColor = 49407
        Else
            nColor = 65535
        End If
        ceil.Value = val
        With ceil.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = nColor
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
End Function

Function ManualCheck() As Boolean
    'Проверка корректности заполнения ручной спецификации
    If IsEmpty(pr_adress) Then r = ReadPrSortament()
    r = OutPrepare()
    nm = Application.ThisWorkbook.ActiveSheet.Name
    If SpecGetType(nm) <> 7 Then
        MsgBox ("Перейдите на лист с ручной спецификацией" & vbLf & "(заканчивается на _спец) и повторите")
        r = OutEnded()
        Exit Function
    End If
    Set Data_out = Application.ThisWorkbook.Sheets(nm)
    r = FormatClear()
    Data_out.Cells.ClearFormats
    Data_out.Cells.ClearComments
    n_row = GetSizeSheet(Data_out)(1)
    col = max_col_man
    spec = Data_out.Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, max_col_man))
    r = FormatFont(Data_out.Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, max_col_man)), n_row, max_col_man)
    n_err = 0
    For i = 3 To n_row
        row = ArrayRow(spec, i)
        type_el = ManualType(row)
        subpos = row(col_man_subpos)  ' Марка элемента
        pos = row(col_man_pos)  ' Поз.
        obozn = row(col_man_obozn) ' Обозначение
        naen = row(col_man_naen) ' Наименование
        qty = row(col_man_qty) ' Кол-во на один элемент
        Weight = row(col_man_weight) ' Масса, кг
        prim = row(col_man_prim) ' Примечание (на лист)
        
        If subpos = Empty Then
            'subpos = "-"
            'Data_out.Cells(i, col_man_subpos).value = subpos
        End If
        If pos = Empty Then
            'pos = " "
            'Data_out.Cells(i, col_man_pos).value = pos
        End If
        
        Select Case type_el
            Case t_sys 'Отмечаем вспомогательные строки
                With Data_out.Range(Data_out.Cells(i, 1), Data_out.Cells(i, max_col_man)).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = XlRgbColor.rgbLightGrey
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            Case t_arm 'Правила для арамтуры
                Length = row(col_man_length) ' Арматура
                diametr = row(col_man_diametr) ' Диаметр
                klass = row(col_man_klass) ' Класс
                'Массу п.м. посчитаем автоматом
                Data_out.Cells(i, col_man_weight).Value = GetWeightForDiametr(diametr, klass)
                Data_out.Cells(i, col_man_weight).Interior.Color = XlRgbColor.rgbLightGrey
                If qty = Empty And prim <> "п.м." Then
                    r = CeilAlert(Data_out.Cells(i, col_man_qty), "Необходимо указать количество" & vbLf & "или добавить примечание п.м.")
                    n_err = n_err + 1
                End If
                If Length > 11800 And prim <> "п.м." Then
                    r = CeilAlert(Data_out.Cells(i, col_man_length), "Стержни длиной выше 11,8 должны идти в п.м.")
                    n_err = n_err + 1
                End If
                If Length < 100 Then
                    r = CeilAlert(Data_out.Cells(i, col_man_length), "Подозрительно малая длина.")
                    n_err = n_err + 1
                End If
            Case t_mat
                If (prim <> "кв.м." And prim <> "куб.м.") Then
                    r = CeilAlert(Data_out.Cells(i, col_man_prim), "Проверьте единицы измерения.")
                    n_err = n_err + 1
                End If
            Case t_prokat
                pr_length = row(col_man_pr_length) ' Прокат
                pr_gost_pr = row(col_man_pr_gost_pr) ' ГОСТ профиля
                pr_prof = row(col_man_pr_prof) ' Профиль
                pr_type = row(col_man_pr_type) ' Тип конструкции
                pr_st = row(col_man_pr_st) ' Сталь
                Data_out.Cells(i, col_man_weight).Interior.Color = XlRgbColor.rgbLightGrey
                If Length > 11800 And prim <> "п.м." Then
                    r = CeilAlert(Data_out.Cells(i, col_man_length), "Стержни длиной выше 11,8 должны идти в п.м.")
                    n_err = n_err + 1
                End If
                
                If prim = "п.м." Then
                    If Not IsEmpty(qty) Then
                        r = CeilAlert(Data_out.Cells(i, col_man_qty), "Количество для элементов в п.м. не указывается")
                        n_err = n_err + 1
                    Else
                        Data_out.Cells(i, col_man_qty).Interior.Color = XlRgbColor.rgbLightGrey
                    End If
                Else
                    If IsEmpty(qty) Then
                        r = CeilAlert(Data_out.Cells(i, col_man_qty), "Необходимо указать количество" & vbLf & "или добавить примечание п.м.")
                        n_err = n_err + 1
                    End If
                End If

            Case t_subpos 'Правила для маркировки сборок
                With Data_out.Range(Data_out.Cells(i, 1), Data_out.Cells(i, max_col_man)).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = XlRgbColor.rgbLightCoral
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            Case t_error
                r = CeilAlert(Data_out.Cells(i, col_man_length), "Проверьте правильность заполнения.")
                r = CeilAlert(Data_out.Cells(i, col_man_pr_length), "Проверьте правильность заполнения.")
                n_err = n_err + 1
        End Select
    Next
    Range("A1").Select
    r = FormatManual(nm)
    r = OutEnded()
    If (n_err) Then
        MsgBox ("Обнаружено " & Str(n_err) & " ошибок, см. примечания к ячейкам")
        ManualCheck = False
    Else
        ManualCheck = True
    End If
    r = Write2log(nm, "check", Str(n_err))
End Function


Function ManualType(ByVal row As Variant) As Integer

    For i = 1 To max_col_man
        If IsError(row(i)) Then
            ManualType = t_sys
            Exit Function
        End If
    Next i

    subpos = row(col_man_subpos)  ' Марка элемента
    pos = row(col_man_pos)  ' Поз.
    obozn = row(col_man_obozn) ' Обозначение
    naen = row(col_man_naen) ' Наименование
    qty = row(col_man_qty) ' Кол-во на один элемент
    Weight = row(col_man_weight) ' Масса, кг
    prim = row(col_man_prim) ' Примечание (на лист)
        
    Length = row(col_man_length) ' Арматура
    diametr = row(col_man_diametr) ' Диаметр
    klass = row(col_man_klass) ' Класс
        
    pr_length = row(col_man_pr_length) ' Прокат
    pr_gost_pr = row(col_man_pr_gost_pr) ' ГОСТ профиля
    pr_prof = row(col_man_pr_prof) ' Профиль
    pr_type = row(col_man_pr_type) ' Тип конструкции
    pr_st = row(col_man_pr_st) ' Сталь
    
    type_el = t_izd ' По умолчанию - изделие
    If IsEmpty(naen) Then type_el = 0 'Пустая строка
 
    isSys = (InStr(subpos, "!") > 0 Or InStr(pos, "!") > 0)
    isSPos = ((subpos = pos) And Not IsEmpty(subpos) And (InStr(subpos, "!") = 0) And (InStr(pos, "!") = 0))
    isArm = (Not IsEmpty(Length) Or Not IsEmpty(diametr) Or Not IsEmpty(klass))
    isProkat = (Not IsEmpty(pr_length) Or Not IsEmpty(pr_gost_pr) Or Not IsEmpty(pr_prof) Or Not IsEmpty(pr_prof))
    isMat = (InStr(prim, "кв.м.") > 0 Or InStr(prim, "куб.м.") > 0 Or InStr(naen, "Бетон") > 0)
    isEr = ((isSPos And isArm) Or (isSPos And isProkat) Or (isSPos And isMat) Or (isArm And isProkat) Or (isArm And isMat) Or (isProkat And isMat)) 'Проверим - не подходит ли элемент к нескольким типам
    
    If isSys Then type_el = t_sys
    If isSPos Then type_el = t_subpos
    If isArm Then type_el = t_arm
    If isProkat Then type_el = t_prokat
    If isMat Then type_el = t_mat
    If isEr Then type_el = t_error
    
    ManualType = type_el
    Erase row
End Function


Function FormatManual(ByVal nm As String) As Boolean
    'Наведение красоты на листе с ручной спецификацией
    r = OutPrepare()
    Set Data_out = Application.ThisWorkbook.Sheets(nm)
    Data_out.Cells.UnMerge
    'Имена диапазонов для каждого столбца
    r_all = FormatManuallitera(1) & ":" & FormatManuallitera(max_col_man)
    r_subpos = FormatManualrange(col_man_subpos)
    r_pos = FormatManualrange(col_man_pos)
    r_obozn = FormatManualrange(col_man_obozn)
    r_naen = FormatManualrange(col_man_naen)
    r_qty = FormatManualrange(col_man_qty)
    r_weight = FormatManualrange(col_man_weight)
    r_prim = FormatManualrange(col_man_prim)
    r_komment = FormatManualrange(col_man_komment)
        
    r_length = FormatManualrange(col_man_length)
    r_length = FormatManualrange(col_man_length)
    r_diametr = FormatManualrange(col_man_diametr)
    r_klass = FormatManualrange(col_man_klass)
    
    r_pr_length = FormatManualrange(col_man_pr_length)
    r_pr_gost_pr = FormatManualrange(col_man_pr_gost_pr)
    r_pr_prof = FormatManualrange(col_man_pr_prof)
    r_pr_type = FormatManualrange(col_man_pr_type)
    r_pr_st = FormatManualrange(col_man_pr_st)
    r_pr_okr = FormatManualrange(col_man_pr_okr)
    r_pr_ogn = FormatManualrange(col_man_pr_ogn)

    Data_out.Cells(1, col_man_subpos) = "Марка" & vbLf & "элемента"
    Data_out.Cells(1, col_man_pos) = "Поз."
    Data_out.Cells(1, col_man_obozn) = "Обозначение"
    Data_out.Cells(1, col_man_naen) = "Наименование"
    Data_out.Cells(1, col_man_qty) = "Кол-во" & vbLf & "на один элемент"
    Data_out.Cells(1, col_man_weight) = "Масса, кг"
    Data_out.Cells(1, col_man_prim) = "Примечание" & vbLf & "(на лист)"
    Data_out.Cells(1, col_man_komment) = "Комментарий"
    
    Data_out.Cells(1, col_man_length) = "Арматура"
    Data_out.Cells(2, col_man_length) = "Длина, мм"
    Data_out.Cells(2, col_man_diametr) = "Диаметр"
    Data_out.Cells(2, col_man_klass) = "Класс"
    
    Data_out.Cells(1, col_man_pr_length) = "Прокат"
    Data_out.Cells(2, col_man_pr_length) = "Длина" & vbLf & "(площадь кв.мм для пластин), мм"
    Data_out.Cells(2, col_man_pr_gost_pr) = "ГОСТ профиля"
    Data_out.Cells(2, col_man_pr_prof) = "Профиль"
    Data_out.Cells(2, col_man_pr_type) = "Тип конструкции"
    Data_out.Cells(2, col_man_pr_st) = "Сталь"
    Data_out.Cells(2, col_man_pr_okr) = "Окраска"
    Data_out.Cells(2, col_man_pr_ogn) = "Огнезащита"
    
    Range(r_all).ClearOutline
    Range(FormatManuallitera(col_man_obozn) & ":" & FormatManuallitera(col_man_prim)).Columns.Group
    Range(FormatManuallitera(col_man_diametr) & ":" & FormatManuallitera(col_man_klass)).Columns.Group
    Range(FormatManuallitera(col_man_pr_gost_pr) & ":" & FormatManuallitera(col_man_pr_ogn)).Columns.Group
    
    Columns(r_all).Validation.Delete

    Range("A1:A2").Merge
    Range("B1:B2").Merge
    Range("C1:C2").Merge
    Range("D1:D2").Merge
    Range("E1:E2").Merge
    Range("F1:F2").Merge
    Range("G1:G2").Merge
    
    Range("H1:J1").Merge
    Range("K1:Q1").Merge
    Range("R1:R2").Merge
    
    Data_out.Cells(1, col_man_subpos).ColumnWidth = 8
    Data_out.Cells(1, col_man_pos).ColumnWidth = 8
    Data_out.Cells(1, col_man_obozn).ColumnWidth = 25
    Data_out.Cells(1, col_man_naen).ColumnWidth = 25
    Data_out.Cells(1, col_man_qty).ColumnWidth = 8
    Data_out.Cells(1, col_man_weight).ColumnWidth = 8
    Data_out.Cells(1, col_man_prim).ColumnWidth = 15
    Data_out.Cells(2, col_man_length).ColumnWidth = 10
    Data_out.Cells(2, col_man_diametr).ColumnWidth = 10
    Data_out.Cells(2, col_man_klass).ColumnWidth = 10
    Data_out.Cells(1, col_man_komment).ColumnWidth = 15
    
    Data_out.Cells(1, col_man_pr_length).ColumnWidth = 15
    Data_out.Cells(2, col_man_pr_length).ColumnWidth = 15
    Data_out.Cells(2, col_man_pr_gost_pr).ColumnWidth = 34
    Data_out.Cells(2, col_man_pr_prof).ColumnWidth = 11
    Data_out.Cells(2, col_man_pr_type).ColumnWidth = 15
    Data_out.Cells(2, col_man_pr_st).ColumnWidth = 8
    Data_out.Cells(2, col_man_pr_okr).ColumnWidth = 8
    Data_out.Cells(2, col_man_pr_ogn).ColumnWidth = 8

    
    'Создаём столбец с марками элементов и добавим раскрывающийся список
    sheet_subpos_name = Left(nm, Len(nm) - 5) & "_поз"
    If SheetExist(sheet_subpos_name) Then
        Set subpos_sheet = Application.ThisWorkbook.Sheets(sheet_subpos_name)
        istart = istart + 1
        row = GetSizeSheet(subpos_sheet)(1)
        pos = subpos_sheet.Range(subpos_sheet.Cells(3, 1), subpos_sheet.Cells(row, 1))
        If IsArray(pos) Then
            un_pos = ArraySort(ArrayUniqValColumn(pos, 1), 1)
        Else
            un_pos = Array(pos)
        End If
        If Not IsEmpty(un_pos) Then
            iend = UBound(un_pos, 1)
            'Data_out.range(Data_out.Cells(1, istart), Data_out.Cells((iEnd + 3) * 500, istart)).ClearContents
            For i = 1 To iend
                Data_out.Range(Data_out.Cells(i, istart), Data_out.Cells(i, istart)) = un_pos(i)
            Next
            Range(r_subpos).Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = False
            End With
            Data_out.Range(Data_out.Cells(1, istart), Data_out.Cells(iend, istart)).Select
            With Selection.Font
                .Name = "Calibri"
                .Size = 8
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            With Range(r_subpos).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & Selection.Address
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = ""
                .ShowInput = True
                .ShowError = False
            End With
        End If
    End If

    With Range(r_prim).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("Примечания")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    With Range(r_klass).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & pr_adress.Item("Классы")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    With Range(r_pr_gost_pr).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & pr_adress.Item("ГОСТпрокат")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    With Range(r_pr_st).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" & pr_adress.Item("Марки стали")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    For i = 1 To 500
        gost = Cells(i, col_man_pr_gost_pr).Value
        addr = pr_adress.Item(gost)
        If Not IsEmpty(addr) Then
            With Cells(i, col_man_pr_prof).Validation
                            .Delete
                            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                            xlBetween, Formula1:="=" & addr(1)
                            .IgnoreBlank = True
                            .InCellDropdown = True
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .ShowInput = True
                            .ShowError = True
            End With
        End If
        
        klass = Cells(i, col_man_klass).Value
        addr = pr_adress.Item(klass)
        If Not IsEmpty(addr) Then
            With Cells(i, col_man_diametr).Validation
                            .Delete
                            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                            xlBetween, Formula1:="=" & addr
                            .IgnoreBlank = True
                            .InCellDropdown = True
                            .InputTitle = ""
                            .ErrorTitle = ""
                            .InputMessage = ""
                            .ErrorMessage = ""
                            .ShowInput = True
                            .ShowError = True
            End With
        End If
    Next i
    
    Range(r_all).Rows.AutoFit
    FormatManual = True
End Function

Function FormatManualrange(ByVal col As Integer) As String
    litera = FormatManuallitera(col)
    out = litera & "3:" & litera & "500"
    FormatManualrange = out
End Function

Function FormatManuallitera(ByVal col As Integer) As String
    If col > 0 Then
        litera = Split(Cells(1, col).Address, "$")(1)
    Else
        litera = "A"
    End If
    FormatManuallitera = litera
End Function

Function SetPath()
    On Error Resume Next
    form = ActiveWorkbook.VBProject.VBComponents("UserForm2").Name
    isFormExistis = 0
    If Not IsEmpty(form) Then isFormExistis = CBool(Len(form))
    If isFormExistis Then
        SortamentPath = UserForm2.SortamentPath
    Else
        SortamentPath = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) & "sort\"
    End If
    SetPath = SortamentPath
End Function

Function ReadPrSortament()
    r = OutPrepare()
    Set Sh = Application.ThisWorkbook.Sheets("!System") 'На этом скрытом листе будем хранить данные для списков
    Sh.Cells.Clear
    Set tpr_adress = CreateObject("Scripting.Dictionary") 'В этом словаре будем хранить адреса
    'Сначала - металл
    SortamentPath = SetPath()
    file = SortamentPath & "Сортаменты.txt"
    f_list_sort = ReadTxt(file, 1, vbTab, vbNewLine)
    f_list_file = ArrayCol(f_list_sort, 3)
    f_list_gost = ArrayCol(f_list_sort, 2)
    n_sort = UBound(f_list_file)
    tpr_adress.Item("ГОСТпрокат") = "'!System'!" & Sh.Range(Sh.Cells(1, 1), Sh.Cells(1, n_sort)).Address
    Dim tmp_arr(3)
    For n_col = 2 To n_sort
        file = f_list_file(n_col)
        Sh.Cells(1, n_col - 1) = file
        f_prof = ReadTxt(SortamentPath & file & ".txt", 1, vbTab, vbNewLine)
        f_list_prof = ArrayCol(f_prof, 2)
        f_list_weight = ArrayCol(f_prof, 3)
        n_prof = UBound(f_list_prof) + 1
        Sh.Range(Sh.Cells(2, n_col - 1), Sh.Cells(n_prof, n_col - 1)) = ArrayTranspose(f_list_prof)
        tmp_arr(1) = "'!System'!" & Sh.Range(Sh.Cells(3, n_col - 1), Sh.Cells(n_prof, n_col - 1)).Address
        tmp_arr(2) = f_list_gost(n_col)
        tpr_adress.Item(file) = tmp_arr
        For j = 2 To n_prof - 1
            prof = f_list_prof(j)
            tmp_arr(1) = f_list_weight(j) 'Вес
            tmp_arr(2) = j 'Периметр
            tmp_arr(3) = j 'Площадь сечения
            tpr_adress.Item(file & prof) = tmp_arr
        Next j
    Next
    n_start = n_sort + 1
    
    'Теперь арматура
    file = SortamentPath & "Сортамент_арматуры.txt"
    f_list_sort = ReadTxt(file)
    f_klass = ArrayUniqValColumn(f_list_sort, 2)
    n_klass = UBound(f_klass)
    n_end = n_start + n_klass
    tpr_adress.Item("Классы") = "'!System'!" & Sh.Range(Sh.Cells(1, n_start + 1), Sh.Cells(1, n_end)).Address
    For i = 1 To n_klass
        klass = f_klass(i)
        If klass <> "Класс" Then
            Sh.Cells(1, n_start + i) = klass
            row = ArrayGetRowIndex(f_list_sort, klass, 2)
            diam = ArrayTranspose(ArrayUniqValColumn(ArraySelectParam(f_list_sort, klass, 2), 3))
            n_diam = UBound(diam)
            Sh.Range(Sh.Cells(2, n_start + i), Sh.Cells(n_diam + 1, n_start + i)) = diam
            tpr_adress.Item(klass) = "'!System'!" & Sh.Range(Sh.Cells(2, n_start + i), Sh.Cells(n_diam + 1, n_start + i)).Address
        End If
    Next
    
    'Теперь марки стали
    n_start = n_end
    file = SortamentPath & "Сталь.txt"
    f_list_stal = ReadTxt(file, 1, vbTab, vbNewLine)
    n_stal = UBound(f_list_stal, 1)
    n_end = n_start + n_stal + 1
    tpr_adress.Item("Марки стали") = "'!System'!" & Sh.Range(Sh.Cells(1, n_start + 1), Sh.Cells(1, n_end)).Address
    For i = 1 To n_stal
        stal = f_list_stal(i, 1)
        Sh.Cells(1, n_start + i) = stal
        Sh.Cells(2, n_start + i) = f_list_stal(i, 2)
        tpr_adress.Item(stal) = f_list_stal(i, 2)
    Next
    
    n_start = n_end + 1
    'Теперь всякие вспомогательные элементы
    Sh.Cells(1, n_start) = "*"
    Sh.Cells(2, n_start) = "п.м."
    Sh.Cells(3, n_start) = "кв.м."
    Sh.Cells(4, n_start) = "куб.м."
    tpr_adress.Item("Примечания") = "'!System'!" & Sh.Range(Sh.Cells(1, n_start), Sh.Cells(4, n_start)).Address
    
    r = OutEnded()
    Set pr_adress = tpr_adress
    ReadPrSortament = True
End Function

Function CatchChange(ByVal Target As Range)
    If IsEmpty(Target) Then Exit Function
    If IsEmpty(pr_adress) Then r = ReadPrSortament()
    nr = 0
    For Each ceil In Target.Cells
        type_izm = Empty
        nr = nr + 1
        If nr > 50 Then Exit Function
        n_colum = ceil.Column
        n_row = ceil.row
        name_colum = Cells(2, n_colum).Value
        If name_colum = "ГОСТ профиля" Then
            'Cells(n_row, col_man_pr_prof).ClearContents
            gost = ceil.Value
            addr = pr_adress.Item(gost)
            If Not IsEmpty(addr) Then
                With Cells(n_row, col_man_pr_prof).Validation
                                .Delete
                                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                xlBetween, Formula1:="=" & addr(1)
                                .IgnoreBlank = True
                                .InCellDropdown = True
                                .InputTitle = ""
                                .ErrorTitle = ""
                                .InputMessage = ""
                                .ErrorMessage = ""
                                .ShowInput = True
                                .ShowError = True
                End With
                type_izm = "Прокат"
            End If
        End If
        If name_colum = "Профиль" Then type_izm = "Прокат"
        
        If name_colum = "Класс" Then
            'Cells(n_row, col_man_diametr).ClearContents
            klass = ceil.Value
            addr = pr_adress.Item(klass)
            If Not IsEmpty(addr) Then
                With Cells(n_row, col_man_diametr).Validation
                                .Delete
                                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                                xlBetween, Formula1:="=" & addr
                                .IgnoreBlank = True
                                .InCellDropdown = True
                                .InputTitle = ""
                                .ErrorTitle = ""
                                .InputMessage = ""
                                .ErrorMessage = ""
                                .ShowInput = True
                                .ShowError = True
                End With
            End If
            type_izm = "Арматура"
        End If
        If name_colum = "Диаметр" Then type_izm = "Арматура"
        
        If type_izm = "Арматура" Then
            diametr = Cells(n_row, col_man_diametr)
            klass = Cells(n_row, col_man_klass)
            If Not IsEmpty(klass) Then
                Cells(n_row, col_man_obozn) = GetGOSTForKlass(klass)
                If Not IsEmpty(diametr) And IsNumeric(diametr) Then Cells(n_row, col_man_weight) = GetWeightForDiametr(diametr, klass)
            End If
        End If
        
        If type_izm = "Прокат" Then
            gost = Cells(n_row, col_man_pr_gost_pr)
            prof = Cells(n_row, col_man_pr_prof)
            If Not IsEmpty(pr_adress.Item(gost)) Then
                Cells(n_row, col_man_obozn) = pr_adress.Item(gost)(2)
                If Not IsEmpty(prof) Then
                    If Not IsEmpty(pr_adress.Item(gost & prof)) Then
                        Cells(n_row, col_man_weight) = pr_adress.Item(gost & prof)(1)
                        Cells(n_row, col_man_naen) = GetNameForGOST(pr_adress.Item(gost)(2)) & " " & prof
                    Else
                        Cells(n_row, col_man_pr_prof).ClearContents
                        Cells(n_row, col_man_weight).ClearContents
                    End If
                End If
            End If
        End If
    Next
End Function

Function ReadMetall() As Boolean
    SortamentPath = SetPath()
    nf_prof = SortamentPath & "Имена профилей.csv"
    If Len(Dir$(nf_prof)) > 0 Then
        name_gost = ReadTxt(nf_prof)
    Else
        MsgBox ("Нет файла с именами профилей")
    End If
End Function

Function GetNameForGOST(ByVal gost As String) As String
    If IsEmpty(name_gost) Then r = ReadMetall()
    For i = 1 To UBound(name_gost, 1)
        If name_gost(i, 1) = gost Then
            GetNameForGOST = name_gost(i, 2) & vbLf & gost
            Exit Function
        End If
    Next
    GetNameForGOST = gost
End Function

Function GetGOSTForKlass(ByVal klass As String) As String
    If IsEmpty(gost2fklass) Then r = ReadReinforce()
    GetGOSTForKlass = gost2fklass.Item(klass)
End Function

Function GetWeightForDiametr(ByVal diametr As Integer, ByVal klass As String) As Double
    If IsEmpty(reinforcement_specifications) Then r = ReadReinforce()
    For i = 1 To UBound(reinforcement_specifications, 1)
        diametr_r = reinforcement_specifications(i, col_diametr_spec)
        klass_r = reinforcement_specifications(i, col_klass_spec)
        If klass_r = klass And diametr_r = diametr Then
            GetWeightForDiametr = CDbl(reinforcement_specifications(i, col_weight_spec))
            Exit Function
        End If
    Next
    MsgBox ("Отсутвует вес для " & diametr & " " & klass)
    GetWeightForDiametr = 1
End Function

Function ReadReinforce() As Boolean
    'Чтение сортамента
    SortamentPath = SetPath()
    nf_sort = SortamentPath & "Сортамент_арматуры.txt"
    If Len(Dir$(nf_sort)) > 0 Then
        reinforcement_specifications = ReadTxt(nf_sort)
    Else
        MsgBox ("Нет файла сортамента арматуры")
        Exit Function
    End If
    Set gost2fklass = CreateObject("Scripting.Dictionary")
    'Массив соответсвия классов и гостов
    For i = 1 To UBound(reinforcement_specifications, 1)
        klass = reinforcement_specifications(i, col_klass_spec)
        gost = reinforcement_specifications(i, col_gost_spec)
        If InStr(klass, "Класс") = 0 Then gost2fklass.Item(klass) = gost
    Next i
End Function

Function ReadTxt(ByVal filename$, Optional ByVal FirstRow& = 1, Optional ByVal ColumnsSeparator$ = ";", Optional ByVal RowsSeparator$ = vbNewLine) As Variant
    On Error Resume Next
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.OpenTextFile(filename$, 1, True): txt$ = ts.ReadAll: ts.Close
    Set ts = Nothing: Set FSO = Nothing
    txt = Trim(txt): Err.Clear
    If txt Like "*" & RowsSeparator$ Then txt = Left(txt, Len(txt) - Len(RowsSeparator$))
    If FirstRow& > 1 Then
       txt = Split(txt, RowsSeparator$, FirstRow&)(FirstRow& - 1)
    End If
    Err.Clear: tmpArr1 = Split(txt, RowsSeparator$): RowsCount = UBound(tmpArr1) + 1
    ColumnsCount = 0
    For i = i To RowsCount - 1
        ColumnsCount = Application.WorksheetFunction.Max(ColumnsCount, UBound(Split(tmpArr1(i), ColumnsSeparator$)) + 1, max_col)
    Next i
    ReDim arr(1 To RowsCount, 1 To ColumnsCount)
    For i = LBound(tmpArr1) To UBound(tmpArr1)
        tmpArr2 = Split(Trim(tmpArr1(i)), ColumnsSeparator$)
        For j = 1 To UBound(tmpArr2) + 1
            arr(i + 1, j) = ConvTxt2Num(Trim(tmpArr2(j - 1)))
        Next j
    Next i
    ReadTxt = arr
    Erase arr
End Function


Function ConvNum2Txt(x As Variant) As Variant
    ConvNum2Txt = Replace(CStr(x), ",", ".")
End Function

Function ConvTxt2Num(ByVal x As Variant) As Variant
    If IsNumeric(x) Then
        out = CDbl(x)
    Else
        x = Trim(x)
        x = Replace(x, ".", ",")
        If IsNumeric(x) Then
            out = CDbl(x)
        Else
            x = Replace(x, ",", ".")
            If IsNumeric(x) Then
                out = CDbl(x)
            Else
                If x = "0" Then
                    out = 0
                Else
                    out = x
                End If
            End If
        End If
    End If
    ConvTxt2Num = out
End Function

