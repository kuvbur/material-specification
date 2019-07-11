Attribute VB_Name = "calc"
Option Compare Text
Option Base 1
Public Const macro_version As String = "3.5"
'-------------------------------------------------------
'Типы элементов (столбец col_type_el)
Public Const t_arm As Integer = 10
Public Const t_prokat As Integer = 20
Public Const t_mat As Integer = 30
Public Const t_mat_spc As Integer = 35
Public Const t_izd As Integer = 40
Public Const t_subpos As Integer = 45
Public Const t_else As Integer = 50
Public Const t_wind As Integer = 60
Public Const t_perem_m As Integer = 70
Public Const t_perem As Integer = 71
Public Const t_error As Integer = -1 'Ошибка распознавания типов
Public Const t_sys As Integer = -10 'Вспомогательный тип
Public Const t_syserror As Integer = -20 'Строки с ошибками
'Столбцы общие
Public Const col_marka As Integer = 1
Public Const col_sub_pos As Integer = 2
Public Const col_type_el As Integer = 3
Public Const col_pos As Integer = 4
Public Const col_qty As Integer = 8
Public Const col_chksum As Integer = 12
Public Const col_parent As Integer = 15
'Столбцы арматуры (t_arm)
Public Const col_klass As Integer = 5
Public Const col_diametr As Integer = 6
Public Const col_length As Integer = 7
Public Const col_fon As Integer = 9
Public Const col_mp As Integer = 10
Public Const col_gnut As Integer = 11
'Столбцы проката (t_prokat)
Public Const col_pr_type_konstr As Integer = 5
Public Const col_pr_gost_st As Integer = 6
Public Const col_pr_st As Integer = 7
Public Const col_pr_gost_prof As Integer = 9
Public Const col_pr_prof As Integer = 10
Public Const col_pr_length As Integer = 11
Public Const col_pr_weight As Integer = 13
Public Const col_pr_naen As Integer = 14
'Столбцы материалов и изделий (t_izd, t_mat, t_subpos)
Public Const col_m_obozn As Integer = 5
Public Const col_m_naen As Integer = 6
Public Const col_m_weight As Integer = 7
Public Const col_m_edizm As Integer = 9
'Столбцы окон, дверей
Public Const col_w_obozn As Integer = 5
Public Const col_w_naen As Integer = 6
Public Const col_w_weight As Integer = 7
Public Const col_w_prim As Integer = 9
Public Const col_w_nfloor As Integer = 10
Public Const col_w_floor As Integer = 11
'Общее количество столбцов во входном массиве
Public Const max_col As Integer = 15
'-------------------------------------------------------
'Описание таблицы с именами сборок (суффикс "_поз")
Public Const col_add_pos As Integer = 1
Public Const col_add_obozn As Integer = 2
Public Const col_add_naen As Integer = 3
Public Const col_add_qty As Integer = 4
Public Const col_add_prim As Integer = 5
'-------------------------------------------------------
'Описание файла с отделкой
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
Public Const col_s_type_otd As Integer = 18
Public Const col_s_layer As Integer = 19
Public Const col_s_type_el As Integer = 20
Public Const col_s_type_pol As Integer = 21
Public Const col_s_area_pol As Integer = 22
Public Const col_s_perim_pol As Integer = 23
Public Const col_s_n_mun_zone As Integer = 24
Public Const col_s_mun_zone As Integer = 25
Public Const col_s_tipverh_l As Integer = 26
Public Const col_s_tipl_l As Integer = 27
Public Const col_s_tipniz_l As Integer = 28
Public Const col_s_tippl_l As Integer = 29
Public Const col_s_areaverh_l As Integer = 30
Public Const col_s_areal_l As Integer = 31
Public Const col_s_areaniz_l As Integer = 32
Public Const col_s_areapl_l As Integer = 33
Public Const max_s_col As Integer = 33
Public Const fin_str As String = "!!AA_"
Public zone_error As Variant
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
Public Const col_man_ank As Integer = 19
Public Const col_man_nahl As Integer = 20
Public Const col_man_dgib As Integer = 21
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
Public Const max_col_man As Integer = col_man_dgib
'-------------------------------------------------------
Public symb_diam As String 'Символ диаметра в спецификацию
Public gost2fklass, name_gost, reinforcement_specifications As Variant 'Разные массивы
Public pr_adress As Variant
Public k_zap_total As Double
Public w_format As String 'Формат вывода в техничку
Public pos_data As Variant
Public sheet_option As Variant
Public concrsubpos As Variant
'-------------------------------------------------------
'-----Переменные, читаемые из INI-----------------------
'-------------------------------------------------------
    'Тип округления
    ' 1 - округление в большую сторону
    ' 2 - округление стандартным round
    ' 3 - округление отключено
Public isINIset As Boolean
Public type_okrugl As Integer
Public n_round_l As Integer 'Длина
Public n_round_w As Integer 'Вес
Public n_round_wkzh As Integer 'Вес в ведомости расхода стали
Public sum_row_wkzh As Boolean
Public show_bet_wkzh As Boolean
Public show_sum_prim As Boolean
Public del_dor_perim As Boolean
Public type_perim As Integer
Public del_freelen_perim As Boolean
Public add_holes_perim As Boolean
Public show_mat_area As Boolean
Public show_surf_area As Boolean
Public show_perim As Boolean
Public zonenum_pot As Boolean
Public delim_by_sheet As Boolean
Public delim_group_ved As Boolean
Public n_round_area As Integer 'Площадь в ведомость отделки
Public ignore_pos As String 'Игнорировать элементы, содержащих ЭТО в позиции или марке
Public subpos_delim As String 'Разделитель основной и вложенной сборки
Public izd_sheet_name As String
Public inx_name As String
Public isErrorNoFin As Boolean 'Выводить ошибку, если в зоне не задана отделка
Public hole_in_zone As Boolean 'Считать отверстия в зонах колоннами
Public mem_option As Boolean 'Запоминать и восстанавливать нстройки листов
Public check_on_active As Boolean 'Проверять листы с ручной спецификацией при переходе на них
Public inx_on_new As Boolean 'Обновлять содежрание после создания нового листа
Public def_decode As Boolean 'Декодировать независимо от настроек
Public Debug_mode As Boolean 'Режим отладки
Public check_version As Boolean 'Проверять версию при загрузке

Function INISet()
    sIniFile = UserForm2.CodePath & "setting.ini"
    If CBool(Len(Dir$(sIniFile))) Then
        type_okrugl = INIReadKeyVal("РАСЧЁТЫ", "type_okrugl")
        n_round_l = INIReadKeyVal("РАСЧЁТЫ", "n_round_l")
        n_round_w = INIReadKeyVal("РАСЧЁТЫ", "n_round_w")
        n_round_wkzh = INIReadKeyVal("РАСЧЁТЫ", "n_round_wkzh")
        ignore_pos = INIReadKeyVal("РАСЧЁТЫ", "ignore_pos")
        subpos_delim = INIReadKeyVal("РАСЧЁТЫ", "subpos_delim")
        n_round_area = INIReadKeyVal("ОТДЕЛКА", "n_round_area")
        hole_in_zone = INIReadKeyVal("ОТДЕЛКА", "hole_in_zone")
        isErrorNoFin = INIReadKeyVal("ОТДЕЛКА", "isErrorNoFin")
        izd_sheet_name = INIReadKeyVal("ЛИСТЫ", "izd_sheet_name")
        inx_name = INIReadKeyVal("ЛИСТЫ", "inx_name")
        mem_option = INIReadKeyVal("ЛИСТЫ", "mem_option")
        inx_on_new = INIReadKeyVal("ЛИСТЫ", "inx_on_new")
        check_on_active = INIReadKeyVal("ЛИСТЫ", "check_on_active")
        def_decode = INIReadKeyVal("ЛИСТЫ", "def_decode")
        Debug_mode = INIReadKeyVal("DEBUG", "Debug_mode")
        check_version = INIReadKeyVal("DEBUG", "check_version")
        del_dor_perim = INIReadKeyVal("ОТДЕЛКА", "del_dor_perim")
        type_perim = INIReadKeyVal("ОТДЕЛКА", "type_perim")
        del_freelen_perim = INIReadKeyVal("ОТДЕЛКА", "del_freelen_perim")
        add_holes_perim = INIReadKeyVal("ОТДЕЛКА", "add_holes_perim")
        show_mat_area = INIReadKeyVal("ОТДЕЛКА", "show_mat_area")
        show_surf_area = INIReadKeyVal("ОТДЕЛКА", "show_surf_area")
        show_perim = INIReadKeyVal("ОТДЕЛКА", "show_perim")
        zonenum_pot = INIReadKeyVal("ОТДЕЛКА", "zonenum_pot")
        delim_by_sheet = INIReadKeyVal("ОТДЕЛКА", "delim_by_sheet")
        sum_row_wkzh = INIReadKeyVal("ЛИСТЫ", "sum_row_wkzh")
        show_bet_wkzh = INIReadKeyVal("ЛИСТЫ", "show_bet_wkzh")
        delim_group_ved = INIReadKeyVal("ЛИСТЫ", "delim_group_ved")
        show_sum_prim = INIReadKeyVal("ЛИСТЫ", "show_sum_prim")
        flag = False
    Else
        flag = True
    End If
    '-----Значения по умолчанию-----------------------------
    If IsEmpty(type_okrugl) Or flag Then type_okrugl = 1
    If IsEmpty(n_round_l) Or flag Then n_round_l = 2
    If IsEmpty(n_round_w) Or flag Then n_round_w = 2
    If IsEmpty(n_round_wkzh) Or flag Then n_round_wkzh = 1
    If IsEmpty(n_round_area) Or flag Then n_round_area = 1
    If IsEmpty(ignore_pos) Or flag Then ignore_pos = "!!"
    If IsEmpty(subpos_delim) Or flag Then subpos_delim = "'"
    If IsEmpty(izd_sheet_name) Or flag Then izd_sheet_name = "Изделия"
    If IsEmpty(inx_name) Or flag Then inx_name = "|Содержание|"
    If IsEmpty(isErrorNoFin Or flag) Then isErrorNoFin = True
    If IsEmpty(hole_in_zone) Or flag Then hole_in_zone = False
    If IsEmpty(mem_option) Or flag Then mem_option = True
    If IsEmpty(inx_on_new) Or flag Then inx_on_new = True
    If IsEmpty(check_on_active) Or flag Then check_on_active = True
    If IsEmpty(def_decode) Or flag Then def_decode = False
    If IsEmpty(check_version) Or flag Then check_version = True
    If IsEmpty(del_dor_perim) Or flag Then del_dor_perim = False
    If IsEmpty(type_perim) Or flag Then type_perim = 1
    If IsEmpty(del_freelen_perim) Or flag Then del_freelen_perim = False
    If IsEmpty(add_holes_perim) Or flag Then add_holes_perim = False
    If IsEmpty(show_mat_area) Or flag Then show_mat_area = True
    If IsEmpty(show_surf_area) Or flag Then show_surf_area = True
    If IsEmpty(show_perim) Or flag Then show_perim = True
    If IsEmpty(zonenum_pot) Or flag Then zonenum_pot = False
    If IsEmpty(delim_by_sheet) Or flag Then delim_by_sheet = False
    If IsEmpty(sum_row_wkzh) Or flag Then sum_row_wkzh = True
    If IsEmpty(show_bet_wkzh) Or flag Then show_bet_wkzh = False
    If IsEmpty(delim_group_ved) Or flag Then delim_group_ved = False
    If IsEmpty(show_sum_prim) Or flag Then show_sum_prim = True
    '----Запись умолчаний, если файл не найден
    If flag Then
        t = INIWriteKeyVal("РАСЧЁТЫ", "type_okrugl", type_okrugl)
        t = INIWriteKeyVal("РАСЧЁТЫ", "n_round_l", n_round_l)
        t = INIWriteKeyVal("РАСЧЁТЫ", "n_round_w", n_round_w)
        t = INIWriteKeyVal("РАСЧЁТЫ", "n_round_wkzh", n_round_wkzh)
        t = INIWriteKeyVal("РАСЧЁТЫ", "ignore_pos", ignore_pos)
        t = INIWriteKeyVal("РАСЧЁТЫ", "subpos_delim", subpos_delim)
        t = INIWriteKeyVal("ОТДЕЛКА", "n_round_area", n_round_area)
        t = INIWriteKeyVal("ОТДЕЛКА", "hole_in_zone", hole_in_zone)
        t = INIWriteKeyVal("ОТДЕЛКА", "isErrorNoFin", isErrorNoFin)
        t = INIWriteKeyVal("ЛИСТЫ", "izd_sheet_name", izd_sheet_name)
        t = INIWriteKeyVal("ЛИСТЫ", "inx_name", inx_name)
        t = INIWriteKeyVal("ЛИСТЫ", "mem_option", mem_option)
        t = INIWriteKeyVal("ЛИСТЫ", "inx_on_new", inx_on_new)
        t = INIWriteKeyVal("ЛИСТЫ", "check_on_active", check_on_active)
        t = INIWriteKeyVal("ЛИСТЫ", "def_decode", def_decode)
        t = INIWriteKeyVal("DEBUG", "Debug_mode", False)
        t = INIWriteKeyVal("DEBUG", "check_version", True)
        t = INIWriteKeyVal("ОТДЕЛКА", "del_dor_perim", False)
        t = INIWriteKeyVal("ОТДЕЛКА", "type_perim", 1)
        t = INIWriteKeyVal("ОТДЕЛКА", "del_freelen_perim", False)
        t = INIWriteKeyVal("ОТДЕЛКА", "add_holes_perim", False)
        t = INIWriteKeyVal("ОТДЕЛКА", "show_mat_area", True)
        t = INIWriteKeyVal("ОТДЕЛКА", "show_surf_area", True)
        t = INIWriteKeyVal("ОТДЕЛКА", "show_perim", True)
        t = INIWriteKeyVal("ОТДЕЛКА", "zonenum_pot", False)
        t = INIWriteKeyVal("ОТДЕЛКА", "delim_by_sheet", False)
        t = INIWriteKeyVal("ЛИСТЫ", "sum_row_wkzh", True)
        t = INIWriteKeyVal("ЛИСТЫ", "show_bet_wkzh", True)
        t = INIWriteKeyVal("ЛИСТЫ", "show_sum_prim", True)
        t = INIWriteKeyVal("ЛИСТЫ", "delim_group_ved", False)
    End If
    '----Принудительное включение
    delim_by_sheet = True
    isINIset = True
End Function

Function INIReadKeyVal(ByVal sSection As String, ByVal sKey As String) As Variant
    sIniFile = UserForm2.CodePath & "setting.ini"
    sIniFileContent = ""
    bSectionExists = False
    bKeyExists = False
    sIniFileContent = INIReadFile(sIniFile)    'Read the file into memory
    aIniLines = Split(sIniFileContent, vbCrLf)
    For i = 0 To UBound(aIniLines)
        sLine = Trim(aIniLines(i))
        If bSectionExists = True And Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
            Exit For    'Start of a new section
        End If
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
        End If
        If bSectionExists = True Then
            If Len(sLine) > Len(sKey) Then
                If Left(sLine, Len(sKey) + 1) = sKey & "=" Then
                    bKeyExists = True
                    INIReadKeyVal = Mid(sLine, InStr(sLine, "=") + 1)
                End If
            End If
        End If
    Next i
    If InStr(INIReadKeyVal, "#") > 0 Then INIReadKeyVal = Trim(Split(INIReadKeyVal, "#")(0))
End Function

Function INIWriteKeyVal(ByVal sSection As String, ByVal sKey As String, ByVal sValue As String) As Boolean
    sIniFile = UserForm2.CodePath & "setting.ini"
    sIniFileContent = ""
    bSectionExists = False
    bKeyExists = False
    sIniFileContent = INIReadFile(sIniFile)    'Read the file into memory
    aIniLines = Split(sIniFileContent, vbCrLf)    'Break the content into individual lines
    sIniFileContent = ""    'Reset it
    For i = 0 To UBound(aIniLines)    'Loop through each line
        sNewLine = ""
        sLine = Trim(aIniLines(i))
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
            bInSection = True
        End If
        If bInSection = True Then
            If sLine <> "[" & sSection & "]" _
               And Left(sLine, 1) = "[" And Right(sLine, 1) = "]" Then
                'Our section exists, but the key wasn't found, so append it
                bInSection = False    ' we're switching section
            End If
            If Len(sLine) > Len(sKey) Then
                If Left(sLine, Len(sKey) + 1) = sKey & "=" Then
                    sNewLine = sKey & "=" & sValue
                    bKeyExists = True
                    bKeyAdded = True
                End If
            End If
        End If
        If Len(sIniFileContent) > 0 Then sIniFileContent = sIniFileContent & vbCrLf
        If sNewLine = "" Then
            sIniFileContent = sIniFileContent & sLine
        Else
            sIniFileContent = sIniFileContent & sNewLine
        End If
    Next i
    'if not found, add it to the end
    If bSectionExists = False Then
        If Len(sIniFileContent) > 0 Then sIniFileContent = sIniFileContent & vbCrLf
        sIniFileContent = sIniFileContent & "[" & sSection & "]"
    End If
    If bKeyAdded = False Then
        sIniFileContent = sIniFileContent & vbCrLf & sKey & "=" & sValue
    End If
    'Write to the ini file the new content
    r = ExportSaveTXTfile(sIniFile, sIniFileContent)
    Ini_WriteKeyVal = True
End Function

Function INIReadFile(ByVal strFile As String) As String
    Dim FileNumber  As Integer
    Dim sFile       As String 'Variable contain file content
    FileNumber = FreeFile
    Open strFile For Binary Access Read As FileNumber
    sFile = Space(LOF(FileNumber))
    Get #FileNumber, , sFile
    Close FileNumber
    INIReadFile = sFile
End Function

Function ArrayCol(ByVal array_in As Variant, ByVal col As Integer) As Variant
    If IsEmpty(array_in) Then ArrayCol = Empty: Exit Function
    If ArrayIsSecondDim(array_in) = False Then ArrayCol = array_in: Exit Function
    If UBound(array_in, 2) < row Then ArrayCol = Empty: Exit Function
    n = UBound(array_in, 1)
    Dim out(): ReDim out(n)
    For i = 1 To n
        out(i) = array_in(i, col)
    Next i
    ArrayCol = out
    Erase out
End Function
    
Function ArrayCombine(ByVal Arr1 As Variant, ByVal Arr2 As Variant) As Variant
    If (Not IsArray(Arr1)) And IsArray(Arr2) Then ArrayCombine = Arr2: Exit Function
    If (Not IsArray(Arr2)) And IsArray(Arr1) Then ArrayCombine = Arr1: Exit Function
    If (Not IsArray(Arr2)) And (Not IsArray(Arr1)) Then ArrayCombine = Empty: Exit Function
    On Error Resume Next: Err.Clear
    If Err.Number = 9 Then ArrayCombine = Empty: Exit Function
    n_rec_row = 1: n_rec_col = 1
    If ArrayIsSecondDim(Arr1) And ArrayIsSecondDim(Arr2) Then
        If (LBound(Arr1, 2) <> LBound(Arr2, 2)) Or (UBound(Arr1, 2) <> UBound(Arr2, 2)) Then ArrayCombine = Empty: Exit Function
        n_row = (UBound(Arr1, 1) - LBound(Arr1, 1) + 1) + (UBound(Arr2, 1) - LBound(Arr2, 1) + 1)
        n_col = (UBound(Arr1, 2) - LBound(Arr1, 2) + 1)
        ReDim arr(n_row, n_col)
        For i = LBound(Arr1, 1) To UBound(Arr1, 1)
            n_rec_col = 1
            For j = LBound(Arr1, 2) To UBound(Arr1, 2)
                arr(n_rec_row, n_rec_col) = Arr1(i, j)
                n_rec_col = n_rec_col + 1
            Next j
            n_rec_row = n_rec_row + 1
        Next i
        For i = LBound(Arr2, 1) To UBound(Arr2, 1)
            n_rec_col = 1
            For j = LBound(Arr2, 2) To UBound(Arr2, 2)
                arr(n_rec_row, n_rec_col) = Arr2(i, j)
                n_rec_col = n_rec_col + 1
            Next j
            n_rec_row = n_rec_row + 1
        Next i
        ArrayCombine = arr
    Else
        If ArrayIsSecondDim(Arr1) Then ArrayCombine = Empty: Exit Function
        If ArrayIsSecondDim(Arr2) Then ArrayCombine = Empty: Exit Function
        n_row = (UBound(Arr1) - LBound(Arr1) + 1) + (UBound(Arr2) - LBound(Arr2) + 1)
        ReDim arr_(n_row)
        For i = LBound(Arr1) To UBound(Arr1)
            arr_(n_rec_row) = Arr1(i)
            n_rec_row = n_rec_row + 1
        Next i
        For i = LBound(Arr2) To UBound(Arr2)
            arr_(n_rec_row) = Arr2(i)
            n_rec_row = n_rec_row + 1
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
                If IsNumeric(array_in(i)) And type_okrugl > 2 Then array_in(i) = Round(array_in(i), 4)
            Next
        Else
            For i = 1 To UBound(array_in, 1)
                For j = 1 To UBound(array_in, 2)
                    If array_in(i, j) = "" Then array_in(i, j) = " "
                    If array_in(i, j) = 0 Then array_in(i, j) = " "
                    If IsNumeric(array_in(i, j)) And type_okrugl > 2 Then array_in(i, j) = Round(array_in(i, j), 4)
                Next
            Next
        End If
    End If
    ArrayEmp2Space = array_in
End Function

Function ArrayGetRowIndex(ByVal array_in As Variant, ByVal param As Variant, Optional ByVal n_col As Integer) As Integer
    index = Empty
    If IsEmpty(array_in) Then
        ArrayGetRowIndex = index
        Exit Function
    End If
    If ArrayIsSecondDim(array_in) Then
        For i = 1 To UBound(array_in, 1)
            If array_in(i, n_col) = param Then
                index = i
                Exit For
            End If
        Next i
    Else
        For i = 1 To UBound(array_in)
            If array_in(i) = param Then
                index = i
                Exit For
            End If
        Next i
    End If
    ArrayGetRowIndex = index
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

Function ArrayRedim(ByVal array_in As Variant, ByVal n_row As Integer) As Variant
    If IsEmpty(array_in) Then ArrayRedim = Empty: Exit Function
    If n_row < 1 Then ArrayRedim = Empty: Exit Function
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

Function ArrayRow(ByVal array_in As Variant, ByVal row As Integer) As Variant
    If IsEmpty(array_in) Then ArrayRow = Empty: Exit Function
    If ArrayIsSecondDim(array_in) = False Then ArrayRow = array_in: Exit Function
    If UBound(array_in, 1) < row Then ArrayRow = Empty: Exit Function
    If row < LBound(array_in, 1) Then ArrayRow = Empty: Exit Function
    n = UBound(array_in, 2)
    Dim out(): ReDim out(n)
    For i = 1 To n
        out(i) = array_in(row, i)
    Next i
    ArrayRow = out
    Erase out, array_in
End Function

Function ArraySelectParam(ByVal array_in As Variant, ByVal param1 As Variant, ByVal n_col1 As Variant, Optional ByVal param2 As Variant, Optional ByVal n_col2 As Variant) As Variant
    Dim arrout
    If IsEmpty(array_in) Then
        ArraySelectParam = Empty
        Exit Function
    End If
    If IsArray(param1) Or IsArray(param2) Then
        ArraySelectParam = ArraySelectParam_2(array_in, param1, n_col1, param2, n_col2)
        Exit Function
    End If
    If ArrayIsSecondDim(array_in) Then
        n_row = UBound(array_in, 1)
        n_param = UBound(array_in, 2)
        n_row_s = 0
        If n_col1 > n_param Then
            ArraySelectParam = Empty
            Exit Function
        End If
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
        If n_param > 0 And n_row_s > 0 Then
            arrout = ArrayTranspose(arrout)
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
Function ArraySelectParam_2(ByVal array_in As Variant, ByVal param1 As Variant, ByVal n_col1 As Variant, Optional ByVal param2 As Variant, Optional ByVal n_col2 As Variant) As Variant
    Dim arrout
    If IsEmpty(array_in) Then
        ArraySelectParam_2 = Empty
        Exit Function
    End If
    If Not IsArray(param1) Then
        param1 = Array(param1)
    End If
    If Not IsMissing(param2) Then
        If Not IsArray(param2) Then param2 = Array(param2)
    End If
    If ArrayIsSecondDim(array_in) Then
        n_row = UBound(array_in, 1)
        n_param = UBound(array_in, 2)
        n_row_s = 0
        If n_col1 > n_param Then
            ArraySelectParam_2 = Empty
            Exit Function
        End If
        ReDim arrout(n_row, n_param)
        For j = 1 To n_row
            flag1 = 0 'Не записывать ни в коем случае
            For Each tparam1 In param1
                If array_in(j, n_col1) = tparam1 Then
                    flag1 = 1 'Конечно, записывать
                Else
                    If InStr(tparam1, "?") > 0 Then
                        tpar = array_in(j, n_col1)
                        If IsNumeric(tpar) Then tparam1 = ConvNum2Txt(tpar)
                        If Right(tparam1, 1) = "?" And Left(tparam1, 1) = "?" Then
                            tparam1 = Trim(Replace(tparam1, "?", ""))
                            If InStr(tpar, tparam1) > 0 Then flag1 = 1
                        End If
                        If Left(tparam1, 1) = "?" Then
                            tparam1 = Trim(Replace(tparam1, "?", ""))
                            If Right(tpar, Len(tparam1)) = tparam1 Then flag1 = 1
                        End If
                        If Right(tparam1, 1) = "?" Then
                            tparam1 = Trim(Replace(tparam1, "?", ""))
                            If Left(tpar, Len(tparam1)) = tparam1 Then flag1 = 1
                        End If
                    End If
                End If
                If flag1 = 1 Then Exit For
            Next
            If flag1 = 1 Then
                If Not IsMissing(n_col2) And Not IsMissing(param2) Then
                    flag2 = 0
                    For Each tparam2 In param2
                        If array_in(j, n_col2) = tparam2 Then
                            flag2 = 1 'Записывать
                        Else
                            If InStr(tparam2, "?") > 0 Then
                                tpar = array_in(j, n_col2)
                                If Right(tparam2, 1) = "?" And Left(tparam2, 1) = "?" Then
                                    tparam2 = Trim(Replace(tparam2, "?", ""))
                                    If InStr(tpar, tparam2) > 0 Then flag2 = 1
                                End If
                                If Left(tparam2, 1) = "?" Then
                                    tparam2 = Trim(Replace(tparam2, "?", ""))
                                    If Right(tpar, Len(tparam2)) = tparam2 Then flag2 = 1
                                End If
                                If Right(tparam2, 1) = "?" Then
                                    tparam2 = Trim(Replace(tparam2, "?", ""))
                                    If Left(tpar, Len(tparam2)) = tparam2 Then flag2 = 1
                                End If
                            End If
                        End If
                        If flag2 = 1 Then Exit For
                    Next
                Else
                    flag2 = 1 'Обязательно записывать
                End If
            End If
            If flag1 And flag2 Then 'Если все согласны
                n_row_s = n_row_s + 1
                For k = 1 To n_param
                    arrout(n_row_s, k) = array_in(j, k)
                Next k
            End If
        Next j
        If n_param > 0 And n_row_s > 0 Then
            arrout = ArrayTranspose(arrout)
            ReDim Preserve arrout(n_param, n_row_s)
            ArraySelectParam_2 = ArrayTranspose(arrout)
            Exit Function
        Else
            ArraySelectParam_2 = Empty
            Exit Function
        End If
    Else
        n_row = UBound(array_in, 1)
        n_row_s = 0
        ReDim arrout(n_row)
        For j = 1 To n_row
            For Each tparam1 In param1
                If array_in(j) = tparam1 Then
                    n_row_s = n_row_s + 1
                    arrout(n_row_s) = array_in(j)
                    Exit For
                End If
            Next
        Next j
        If n_row_s > 0 Then
            ReDim Preserve arrout(n_row_s)
            ArraySelectParam_2 = arrout
            Exit Function
        Else
            ArraySelectParam_2 = Empty
            Exit Function
        End If
    End If
    Erase array_in
End Function
Function ArraySort_2(ByVal array_in As Variant, ByVal nCol1 As Integer, ByVal nCol2 As Integer) As Variant
    If IsEmpty(array_in) Then
        ArraySort_2 = Empty
        Exit Function
    End If
    If Not ArrayIsSecondDim(array_in) Then
        ArraySort_2 = Empty
        Exit Function
    End If
    n_row = UBound(array_in, 1)
    n_col = UBound(array_in, 2)
    If n_col1 > n_col Or n_col2 > n_col Then
        ArraySort_2 = Empty
        Exit Function
    End If
    Dim array_out As Variant
    sort_key = ArrayUniqValColumn(array_in, nCol1)
    For Each stkey In sort_key
        array_by_key = ArraySelectParam(array_in, stkey, nCol1)
        array_by_key = ArraySort(array_by_key, nCol2)
        array_out = ArrayCombine(array_out, array_by_key)
    Next
    ArraySort_2 = array_out
End Function
Function ArraySort(ByVal array_in As Variant, Optional ByVal nCol As Integer = 1) As Variant
    If IsEmpty(array_in) Then
        ArraySort = Empty
        Exit Function
    End If
    Dim array_in_str As Variant
    Dim array_in_num As Variant
    If ArrayIsSecondDim(array_in) Then
        n_row = UBound(array_in, 1)
        n_col = UBound(array_in, 2)
        If LBound(array_in, 1) = 0 Then n_row = n_row + 1
        If LBound(array_in, 2) = 0 Then n_col = n_col + 1
        ReDim array_in_str(n_col, n_row)
        ReDim array_in_num(n_col, n_row)
        n_str = 0
        n_num = 0
        For i = LBound(array_in, 1) To UBound(array_in, 1)
            If IsNumeric(ConvTxt2Num(array_in(i, nCol))) Then
                n_num = n_num + 1
                For j = LBound(array_in, 2) To UBound(array_in, 2)
                    array_in_num(j, n_num) = array_in(i, j)
                Next j
            Else
                n_str = n_str + 1
                For j = LBound(array_in, 2) To UBound(array_in, 2)
                    array_in_str(j, n_str) = array_in(i, j)
                Next j
            End If
        Next i
        If n_num > 0 Then
            ReDim Preserve array_in_num(n_col, n_num)
            array_in_num = ArraySortNum(ArrayTranspose(array_in_num), nCol)
        End If
        If n_str > 0 Then
            ReDim Preserve array_in_str(n_col, n_str)
            array_in_str = ArraySortABC(ArrayTranspose(array_in_str), nCol)
            If n_num > 0 Then
                array_in = ArrayCombine(array_in_num, array_in_str)
            Else
                array_in = array_in_str
            End If
        Else
            array_in = array_in_num
        End If
    Else
        n_row = UBound(array_in)
        If LBound(array_in) = 0 Then n_row = n_row + 1
        If n_row <= 0 Then
            ArraySort = Empty
            Exit Function
        End If
        ReDim array_in_str(n_row)
        ReDim array_in_num(n_row)
        n_str = 0
        n_num = 0
        For i = LBound(array_in) To UBound(array_in)
            If IsNumeric(ConvTxt2Num(array_in(i))) Then
                n_num = n_num + 1
                array_in_num(n_num) = array_in(i)
            Else
                n_str = n_str + 1
                array_in_str(n_str) = array_in(i)
            End If
        Next i
        If n_num > 0 Then
            ReDim Preserve array_in_num(n_num)
            array_in_num = ArraySortNum(array_in_num, nCol)
        End If
        If n_str > 0 Then
            ReDim Preserve array_in_str(n_str)
            array_in_str = ArraySortABC(array_in_str, nCol)
            If n_num > 0 Then
                array_in = ArrayCombine(array_in_num, array_in_str)
            Else
                array_in = array_in_str
            End If
        Else
            array_in = array_in_num
        End If
    End If
    ArraySort = array_in
End Function

Function ArraySortABC(ByVal array_in As Variant, ByVal nCol As Integer) As Variant
    If IsEmpty(array_in) Then ArraySortABC = Empty: Exit Function
    If ArrayIsSecondDim(array_in) Then
        Dim tempArray As Variant: ReDim tempArray(1, UBound(array_in, 2))
        k = UBound(array_in, 1)
        For j = LBound(array_in, 1) To UBound(array_in, 1) - 1
            For i = 2 To k
                If array_in(i - 1, nCol) <> Empty And array_in(i, nCol) <> Empty Then
                    If StrComp(array_in(i - 1, nCol), array_in(i, nCol), vbTextCompare) = 1 Then
                    'If Asc(UCase(array_in(i - 1, nCol))) > Asc(UCase(array_in(i, nCol))) Then
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
        For j = LBound(array_in) To UBound(array_in) - 1
            For i = 2 To k
                If Not IsEmpty(array_in(i - 1)) And Not IsEmpty(array_in(i)) And Not Len(array_in(i)) = 0 And Not Len(array_in(i - 1)) = 0 Then
                    If StrComp(array_in(i - 1), array_in(i), vbTextCompare) = 1 Then
                    'If Asc(UCase(array_in(i - 1))) > Asc(UCase(array_in(i))) Then
                        V = array_in(i - 1)
                        array_in(i - 1) = array_in(i)
                        array_in(i) = V
                    End If
                End If
            Next i
            k = k - 1
        Next j
    End If
    ArraySortABC = array_in
    Erase array_in
End Function

Function ArraySortNum(ByVal array_in As Variant, ByVal nCol As Integer) As Variant
    If IsEmpty(array_in) Then ArraySortNum = Empty: Exit Function
    If ArrayIsSecondDim(array_in) Then
        If nCol > UBound(array_in, 2) Or nCol < LBound(array_in, 2) Then MsgBox "Нет такого столбца в массиве!", vbCritical: Exit Function
        Dim Check As Boolean, iCount As Integer, jCount As Integer
        ReDim tmpArr(UBound(array_in, 2)) As Variant
        Do Until Check
            Check = True
            For iCount = LBound(array_in, 1) To UBound(array_in, 1) - 1
                If val(array_in(iCount, nCol)) > val(array_in(iCount + 1, nCol)) Then
                    For jCount = LBound(array_in, 2) To UBound(array_in, 2)
                        tmpArr(jCount) = array_in(iCount, jCount)
                        array_in(iCount, jCount) = array_in(iCount + 1, jCount)
                        array_in(iCount + 1, jCount) = tmpArr(jCount)
                        Check = False
                    Next
                End If
            Next
        Loop
    Else
        n = UBound(array_in)
        For i = 1 To n
            For j = 1 To (n - i)
                If val(array_in(j)) > val(array_in(j + 1)) Then
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
    If IsEmpty(array_in) Then
        ArrayTranspose = Empty
        Exit Function
    End If
    Dim tempArray As Variant
    If ArrayIsSecondDim(array_in) Then
        ReDim tempArray(LBound(array_in, 2) To UBound(array_in, 2), LBound(array_in, 1) To UBound(array_in, 1))
        For x = LBound(array_in, 2) To UBound(array_in, 2)
            For Y = LBound(array_in, 1) To UBound(array_in, 1)
                tempArray(x, Y) = array_in(Y, x)
            Next Y
        Next x
    Else:
        ReDim tempArray(LBound(array_in, 1) To UBound(array_in, 1), LBound(array_in, 1) To UBound(array_in, 1))
        For x = LBound(array_in, 1) To UBound(array_in, 1)
            tempArray(x, 1) = array_in(x)
        Next x
    End If
    ArrayTranspose = tempArray
    Erase tempArray
End Function

Function ArrayUniqValColumn(ByVal array_in As Variant, ByVal cols As Long) As Variant
    Dim array_out()
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
            If IsError(array_in(i, cols)) Then Exit For
            For j = 1 To n_un
                If array_out(j) = array_in(i, cols) Then
                    flag = 0
                    Exit For
                End If
            Next
            If IsEmpty(array_in(i, cols)) Then flag = 0
            If Len(array_in(i, cols)) = 0 Then flag = 0
            If array_in(i, cols) = " " Then flag = 0
            If ConvTxt2Num(array_in(i, cols)) = 0 Then flag = 0
            If flag = 1 Then
                n_un = n_un + 1
                array_out(n_un) = array_in(i, cols)
            End If
        Next
        ReDim Preserve array_out(n_un)
    Else
        n_un = 1
        ReDim array_out(n_un)
        If cols = 0 Then cols = 1
        array_out(1) = array_in(LBound(array_in))
        For i = LBound(array_in) To UBound(array_in)
            flag = 1
            If IsError(array_in(i)) Then Exit For
            For j = 1 To n_un
                If array_out(j) = array_in(i) Then
                    flag = 0
                    Exit For
                End If
            Next
            If IsEmpty(array_in(i)) Then flag = 0
            If Len(array_in(i)) = 0 Then flag = 0
            If array_in(i) = " " Then flag = 0
            If ConvTxt2Num(array_in(i)) = 0 Then flag = 0
            If flag = 1 Then
                n_un = n_un + 1
                ReDim Preserve array_out(n_un)
                array_out(n_un) = array_in(i)
            End If
        Next
    End If
    array_out = ArraySort(array_out, 1)
    ArrayUniqValColumn = array_out
    Erase array_out
End Function

Function ControlSumAddVar(ByVal var As Variant) As String
    If IsNumeric(var) Then var = Trim(Str(var))
    If var = "_" Then
        ControlSumAddVar = "_"
    Else
        For Each deltxt In Array(" ", "--", "x", "х", "-")
            var = Trim(Replace(var, deltxt, ""))
        Next
        ControlSumAddVar = var
    End If
End Function

Function ControlSumEl(ByVal array_in As Variant) As String
    Dim param
    isel = 0
    If ArrayIsSecondDim(array_in) Then
        Dim t: ReDim t(UBound(array_in, 2))
        For i = 1 To UBound(array_in, 2)
            t(i) = array_in(1, i)
        Next i
        array_in = t
        Erase t
    End If
    'marka = array_in(col_marka)
    subpos = array_in(col_sub_pos)
    type_el = array_in(col_type_el)
    pos = array_in(col_pos)
    qty = array_in(col_qty)
    chksum = array_in(col_chksum)
    sparent = array_in(col_parent)
    If sparent = 0 Then sparent = "-"
    If subpos = 0 Then subpos = "-"
    If pos = 0 Then pos = "-"
    Select Case type_el
        Case t_arm
            isel = 1
            klass = array_in(col_klass)
            diametr = array_in(col_diametr)
            Length = array_in(col_length)
            fon = array_in(col_fon)
            mp = array_in(col_mp)
            gnut = array_in(col_gnut)

            ReDim param(12)
            param(1) = diametr
            param(2) = klass
            param(3) = "_"
            param(4) = subpos
            param(5) = sparent
            param(6) = "_"
            param(7) = pos
            param(8) = "_"
            If fon Then
                param(9) = "l"
                param(10) = "f1"
                param(11) = 0
                param(12) = "g" + ConvNum2Txt((gnut = 1) * 3)
            Else
                param(9) = "l" + ConvNum2Txt(Int(Length))
                param(10) = "f0"
                param(11) = 0
                param(12) = "g" + ConvNum2Txt((gnut = 1) * 3)
            End If
        Case t_prokat
            isel = 1
            type_konstr = array_in(col_pr_type_konstr)
            gost_st = array_in(col_pr_gost_st)
            st = array_in(col_pr_st)
            gost_prof = array_in(col_pr_gost_prof)
            prof = array_in(col_pr_prof)
            naen = array_in(col_pr_naen)
            Length = array_in(col_pr_length)
            'Weight = array_in(col_pr_weight)
            
            ReDim param(11)
            param(1) = prof
            param(2) = gost_prof
            param(3) = st
            param(4) = "_"
            param(5) = subpos
            param(6) = sparent
            param(7) = "_"
            param(8) = pos
            param(9) = "_"
            param(10) = type_konstr
            param(11) = "l" + ConvNum2Txt(Int(Length)) + naen
            
        Case t_mat
            isel = 1
            obozn = array_in(col_m_obozn)
            naen = array_in(col_m_naen)
            'Weight = array_in(col_m_weight)
            qty = array_in(col_qty)
            edizm = array_in(col_m_edizm)
            chksum = array_in(col_chksum)
            
            ReDim param(9)
            param(1) = obozn
            param(2) = naen
            param(3) = edizm
            param(4) = "_"
            param(5) = subpos
            param(6) = sparent
            param(7) = "_"
            param(8) = pos
        Case t_izd
            isel = 1
            obozn = array_in(col_m_obozn)
            naen = array_in(col_m_naen)
            edizm = array_in(col_m_edizm)
            Weight = array_in(col_m_weight)
            
            ReDim param(10)
            param(1) = obozn
            param(2) = naen
            param(3) = Weight
            param(4) = "_"
            param(5) = subpos
            param(6) = sparent
            param(7) = "_"
            param(8) = pos
            param(9) = "_"
            param(10) = " "
        Case t_subpos
            isel = 1
            obozn = array_in(col_m_obozn)
            naen = array_in(col_m_naen)
            Weight = array_in(col_m_weight)
            edizm = array_in(col_m_edizm)
            
            ReDim param(6)
            param(1) = subpos
            param(2) = "_"
            param(3) = subpos
            param(4) = sparent
            param(5) = "_"
            param(6) = subpos
        Case t_wind
            isel = 1
            obozn = array_in(col_w_obozn)
            naen = array_in(col_w_naen)
            Floor = array_in(col_w_floor)
            
            ReDim param(7)
            param(1) = pos
            param(2) = subpos
            param(3) = "_"
            param(4) = Floor
            param(5) = "_"
            param(6) = obozn
            param(7) = naen
        Case t_perem
            isel = 1
            obozn = array_in(col_m_obozn)
            naen = array_in(col_m_naen)
            Weight = array_in(col_m_weight)
            edizm = array_in(col_m_edizm)
            
            ReDim param(6)
            param(1) = subpos
            param(2) = "_"
            param(3) = subpos
            param(4) = sparent
            param(5) = "_"
            param(6) = subpos
    End Select
    control_sum = ""
    If isel Then
        For i = 1 To UBound(param, 1)
            var = param(i)
            cs = ControlSumAddVar(var)
            control_sum = control_sum & ControlSumAddVar(var)
        Next i
        'If chksum <> Empty And control_sum <> chksum Then
            'Debug.Print (subpos & ", " & ", " & pos & "-> chksum")
        'End If
    End If
    ControlSumEl = control_sum
End Function

Function ConvNum2Txt(ByVal var As Variant) As String
    txt = ""
    If IsNumeric(var) Then
        If var = 0 Then
            txt = ""
        Else
            txt = Trim(CStr(var))
            If Left(txt, 1) = "." Or Left(txt, 1) = "," Then txt = "0" + txt
        End If
        txt = Replace(txt, ".", ",")
    Else
        txt = var
    End If
    ConvNum2Txt = txt
End Function

Function ConvTxt2Num(ByVal x As Variant) As Variant
    If IsNumeric(x) Then
        out = CDbl(x)
    Else
        x_tmp = x
        x = Replace(x, " ", "")
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
                    out = x_tmp
                End If
            End If
        End If
    End If
    ConvTxt2Num = out
End Function

Function DataAddNullSubpos(ByVal array_in As Variant) As Variant
    'TODO переделать под новую систему
    'Если в массиве есть элементы, состоящие в сборках, но маркировки сборок (t_subpos) нет - добавляет строки маркировок сборок
    If IsEmpty(array_in) Then
        DataAddNullSubpos = Empty
        Exit Function
    End If
    Dim add_subpos
    Dim out_subpos
    Set name_subpos = DataNameSubpos(exist_subpos) 'Получим для них имена
    arr_subpos = ArrayUniqValColumn(array_in, col_sub_pos)
    add_txt = Empty
    For Each current_subpos In arr_subpos
        If current_subpos <> "-" And InStr(current_subpos, "zap") = 0 Then
            'Проеряем - есть ли маркировка для главных сборок
            seach_subpos = ArraySelectParam(array_in, current_subpos, col_sub_pos, t_subpos, col_type_el)
            If IsEmpty(seach_subpos) Then
                If IsEmpty(add_txt) Then
                    add_txt = current_subpos
                Else
                    add_txt = current_subpos & ", " & add_txt
                End If
                If name_subpos.Exists(current_subpos) Then
                    naen = name_subpos(current_subpos)(1)
                    obozn = name_subpos(current_subpos)(2)
                Else
                    naen = current_subpos
                    obozn = "!!!"
                End If
                ReDim add_subpos(1, max_col)
                add_subpos(1, col_sub_pos) = current_subpos
                add_subpos(1, col_type_el) = t_subpos
                add_subpos(1, col_pos) = current_subpos
                add_subpos(1, col_m_naen) = Replace(naen, "@", subpos_delim)
                add_subpos(1, col_m_obozn) = obozn
                add_subpos(1, col_qty) = 1
                add_subpos(1, col_chksum) = ControlSumEl(add_subpos)
                out_subpos = ArrayCombine(out_subpos, add_subpos)
            End If
        End If
    Next
    If Not IsEmpty(add_txt) Then MsgBox ("Добавлена маркировка " & add_txt)
    DataAddNullSubpos = DataCheck(out_subpos)
    Erase array_in
End Function

Function DataCheck(ByVal array_in As Variant) As Variant
    If IsEmpty(array_in) Then DataCheck = Empty: Exit Function
    n_col = UBound(array_in, 2)
    n_ingore = 0
    Dim out_data: ReDim out_data(UBound(array_in, 1), n_col): n_row = 0
    For i = 1 To UBound(array_in, 1)
        type_el = array_in(i, col_type_el)
        'Вложенные сборки определяем по разделителю subpos_delim в первом столбце
        'Если это строка относится к вложенной сборке, то формат будет ИмяГлавнойСборки subpos_delim ПозЭлемента
        array_in(i, col_parent) = Empty
        array_in(i, col_marka) = Replace(array_in(i, col_marka), ",", ".")
        If InStr(array_in(i, col_marka), subpos_delim) Then
            parent_subpos = Split(array_in(i, col_marka), subpos_delim)(0)
            pos = Split(array_in(i, col_marka), subpos_delim)(1)
            array_in(i, col_parent) = parent_subpos
            array_in(i, col_marka) = parent_subpos
        End If
'        Также проверяем поле сборки, сравнивая его с маркой
'        Если это вложенная сброрка, то в графе сборки будет МАРКА subpos_delim СБОРКА
        If InStr(array_in(i, col_sub_pos), subpos_delim) Then
            parent_subpos = Split(array_in(i, col_sub_pos), subpos_delim)(0)
            subpos = Split(array_in(i, col_sub_pos), subpos_delim)(1)
            array_in(i, col_sub_pos) = subpos
            array_in(i, col_parent) = parent_subpos
            array_in(i, col_marka) = parent_subpos
            If type_el = t_subpos Then array_in(i, col_pos) = subpos
        End If
        
        ignore_flag = 0
        If InStr(array_in(i, col_sub_pos), ignore_pos) Then ignore_flag = 1
        If InStr(array_in(i, col_parent), ignore_pos) Then ignore_flag = 1
        If InStr(array_in(i, col_marka), ignore_pos) Then ignore_flag = 1
        If ignore_flag Then
            type_el = t_error
            array_in(i, col_type_el) = t_error
            n_ingore = n_ingore + 1
        End If
    
        If Len(type_el) > 0 Then
            Select Case type_el
                Case t_arm
                    'Для арматуры в п.м. количество должно быть 1.
                    If array_in(i, col_fon) = 1 And array_in(i, col_qty) > 1 Then
                        l_spec = array_in(i, col_length) * array_in(i, col_qty)
                        array_in(i, col_length) = l_spec
                        array_in(i, col_qty) = 1
                    End If
                    klass = array_in(i, col_klass)
                    diametr = array_in(i, col_diametr)
                    weight_pm = GetWeightForDiametr(diametr, klass)
                    length_pos = Round_w(array_in(i, col_length) / 1000, n_round_l)

                Case t_prokat
                    If Not IsNumeric(array_in(i, col_pr_weight)) Then
                        array_in(i, col_pr_weight) = 0.01
                    End If
                    length_pos = Round_w(array_in(i, col_pr_length) / 1000, n_round_l)
                Case t_mat
                
                Case t_izd
                
                Case t_subpos
                Case t_mat_spc
                    If InStr(array_in(i, col_marka), subpos_delim) Then
                        array_in(i, col_sub_pos) = Split(array_in(i, col_marka), subpos_delim)(1)
                    Else
                        array_in(i, col_sub_pos) = array_in(i, col_marka)
                    End If
                    array_in(i, col_pos) = Empty
                    array_in(i, col_type_el) = t_mat
                    array_in(i, col_m_weight) = "-"
            End Select
            'Пометим заполнение окон, дверей, а также перемычки - для них нет нужны выделять элементы сборок
            If (type_el = t_perem) Or (type_el = t_perem_m) Or (type_el = t_wind) Then
                 array_in(i, col_sub_pos) = "zap" + array_in(i, col_sub_pos)
            End If
            
            If array_in(i, col_sub_pos) = "" Then array_in(i, col_sub_pos) = "-"
            If array_in(i, col_sub_pos) = " " Then array_in(i, col_sub_pos) = "-"
            If array_in(i, col_sub_pos) = 0 Then array_in(i, col_sub_pos) = "-"
            If array_in(i, col_sub_pos) = "-" Then array_in(i, col_parent) = "-"
            If IsEmpty(array_in(i, col_parent)) Then array_in(i, col_parent) = "-"
            array_in(i, col_sub_pos) = Replace(array_in(i, col_sub_pos), "@", subpos_delim)
            array_in(i, col_parent) = Replace(array_in(i, col_parent), "@", subpos_delim)
            array_in(i, col_marka) = Replace(array_in(i, col_marka), "@", subpos_delim)
            array_in(i, col_pos) = Replace(array_in(i, col_pos), "@", subpos_delim)
            array_in(i, col_pos) = Replace(array_in(i, col_pos), ",", ".")
            'Вычисление и проверка контрольных сумм
            array_in(i, col_chksum) = ControlSumEl(ArrayRow(array_in, i))
            n_row = n_row + 1
            For j = 1 To n_col
                out_data(n_row, j) = array_in(i, j)
            Next j
        End If
    Next i
    If n_ingore > 0 Then r = LogWrite("Строк, содержащих " & ignore_pos & " пропущено", "", n_ingore)
    If n_row Then
        out_data = ArrayTranspose(out_data)
        ReDim Preserve out_data(n_col, n_row)
        out_data = ArrayTranspose(out_data)
        DataCheck = out_data
    Else
        DataCheck = Empty
    End If
    Erase array_in, out_data
End Function

Function DataIsOtd(ByVal array_in As Variant) As Boolean
    If IsEmpty(array_in) Then
        DataIsOtd = False
        Exit Function
    End If
    n_col = UBound(array_in, 2)
    If array_in(1, col_s_type) = "ЗОНА" And (n_col = col_s_layer Or n_col = col_s_mun_zone Or n_col = col_s_areapl_l) Then
        DataIsOtd = True
    Else
        DataIsOtd = False
    End If
End Function

Function DataIsShort(ByVal array_in As Variant) As Boolean
'Если номер столбца с типами элементов отличается от col_type_el - то первый столбец, скорее всего - количество элементов
    colum = 0
    n_row = Int(UBound(array_in, 1) / 2) + 1
    For j = col_type_el To col_type_el + 1
        n = 0
        For i = 1 To n_row
            If t_arm = array_in(i, j) Then n = n + 1
            If t_izd = array_in(i, j) Then n = n + 1
            If t_subpos = array_in(i, j) Then n = n + 1
            If t_mat = array_in(i, j) Then n = n + 1
            If t_prokat = array_in(i, j) Then n = n + 1
            If t_else = array_in(i, j) Then n = n + 1
            If t_mat_spc = array_in(i, j) Then n = n + 1
        Next i
        If n > 0 And colum = 0 Then colum = j
    Next j
    res = False
    If colum <> col_type_el Then res = True
    DataIsShort = res
End Function

Function DataIsSpec(ByVal array_in As Variant) As Boolean
    n_row = Int(UBound(array_in, 1) / 2) + 1
    n = 0
    For i = 1 To n_row
        If t_arm = array_in(i, col_type_el) Then n = n + 1
        If t_izd = array_in(i, col_type_el) Then n = n + 1
        If t_subpos = array_in(i, col_type_el) Then n = n + 1
        If t_mat = array_in(i, col_type_el) Then n = n + 1
        If t_prokat = array_in(i, col_type_el) Then n = n + 1
        If t_else = array_in(i, col_type_el) Then n = n + 1
        If t_mat_spc = array_in(i, col_type_el) Then n = n + 1
        If t_perem = array_in(i, col_type_el) Then n = n + 1
        If t_perem_m = array_in(i, col_type_el) Then n = n + 1
        If t_wind = array_in(i, col_type_el) Then n = n + 1
    Next i
    If n > 0 Then DataIsSpec = True Else DataIsSpec = False
End Function

Function DataIsWall(ByVal nm As String) As Variant
    array_in = ReadTxt(ThisWorkbook.path & "\import\" & nm, 1, vbTab, vbNewLine)
    n_row = UBound(array_in, 1)
    Dim pos_out: ReDim pos_out(n_row - 1, max_col)
    For i = 2 To n_row
        subpos = Replace(array_in(i, 1), subpos_delim, "@")
        naen = array_in(i, 2)
        obozn = "-"
        p_start = 0
        If InStr(naen, "(ТУ") Then p_start = InStr(naen, "(ТУ") - 1
        If InStr(naen, "(ГОСТ") Then p_start = InStr(naen, "(ГОСТ") - 1
        If InStr(naen, "(СТО") Then p_start = InStr(naen, "(СТО") - 1
        If p_start > 0 Then
            p_end = InStr(naen, ")") + 1
            t_start = Trim(Mid(naen, 1, p_start))
            t_end = Trim(Mid(naen, p_end, Len(naen)))
            obozn = Trim(Mid(naen, p_start + 2, p_end - p_start - 3))
            naen = t_start & " " & t_end
        End If
        t_sl = array_in(i, 3)
        If t_sl > 0.1 Then naen = naen & " t=" & ConvNum2Txt(t_sl) & "мм."
        qty = array_in(i, 4)
        prim = "кв.м."
        
        n_row_out = i - 1
        pos_out(n_row_out, col_sub_pos) = subpos
        pos_out(n_row_out, col_type_el) = t_mat
        pos_out(n_row_out, col_qty) = qty
        pos_out(n_row_out, col_m_obozn) = obozn
        pos_out(n_row_out, col_m_naen) = naen
        pos_out(n_row_out, col_m_weight) = "-"
        pos_out(n_row_out, col_m_edizm) = prim
    Next i
    DataIsWall = pos_out
End Function

Function DataNameSubpos(ByVal sub_pos_arr As Variant) As Object
    Set name_subpos = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(sub_pos_arr) Then
        For i = 1 To UBound(sub_pos_arr, 1)
            subpos = sub_pos_arr(i, col_sub_pos)
            naen = sub_pos_arr(i, col_m_naen)
            obozn = sub_pos_arr(i, col_m_obozn)
            name_subpos.Item(subpos) = Array(naen, obozn)
        Next i
    End If
    nm = ThisWorkbook.ActiveSheet.Name
    type_sheet = SpecGetType(nm)
    If Not IsEmpty(type_sheet) And type_sheet <> 3 Then
        sheet = Split(nm, "_")(0) & "_поз"
    Else
        sheet = nm & "_поз"
    End If
    If SheetExist(sheet) Then
        array_in = ReadPos(sheet)
        all_subpos_in_sheet = ArraySelectParam(array_in, t_subpos, col_type_el)
        For i = 1 To UBound(all_subpos_in_sheet, 1)
            subpos = all_subpos_in_sheet(i, col_sub_pos)
            naen = all_subpos_in_sheet(i, col_m_naen)
            obozn = all_subpos_in_sheet(i, col_m_obozn)
            If InStr(naen, "!!!") = 0 And InStr(obozn, "!!!") = 0 Then
                If name_subpos.Exists(subpos) Then name_subpos.Remove subpos
                name_subpos.Item(subpos) = Array(naen, obozn)
            End If
        Next i
    End If
    Set DataNameSubpos = name_subpos
End Function

Function DataRead(ByVal nm As String) As Variant
    errread = 0
    Select Case SpecGetType(nm)
        Case 7
            'Читаем с листа
            out_data = ManualSpec(nm)
        Case Else
            'Проверим - есть ли такой файл
            listFile = GetListFile("*.txt")
            If InStr(nm, "_") > 0 Then
                type_spec = Split(nm, "_")
                nsfile = type_spec(0)
            Else
                nsfile = nm
            End If
            file = ArraySelectParam(listFile, nsfile, 1)
            If IsEmpty(file) Then
                'Если файла нет - поищем листы с суффиксом "_спец"
                listsheet = GetListOfSheet(ThisWorkbook)
                If IsEmpty(type_spec) Then
                    nsht = nm & "_спец"
                Else
                    nsht = type_spec(0) & "_спец"
                End If
                sheet = ArraySelectParam(listsheet, nsht, 1)
                If IsEmpty(sheet) Then
                    'Нет ни файла, ни листа.
                    errread = 1
                Else
                    'Читаем с листа
                    r = ManualCheck(nsht)
                    out_data = ManualSpec(nsht)
                End If
            Else
                'Читаем из файла
                out_data = ReadFile(file(1, 1) & ".txt")
                If InStr(out_data(1, 1), "Площадь Слоя/Компонента") Then
                    out_data = DataIsWall(file(1, 1) & ".txt")
                End If
            End If
    End Select
    If IsEmpty(out_data) Then DataRead = Empty: Exit Function
    If errread Then
        MsgBox ("Лист или файл отсутствуют")
    Else
        If DataIsShort(out_data) Then out_data = DataShort(out_data)
        Dim out: ReDim out(UBound(out_data, 1), max_col)
        For i = 1 To UBound(out_data, 1)
            For j = 1 To max_col
                If j <= UBound(out_data, 2) Then
                    out(i, j) = out_data(i, j)
                End If
            Next j
        Next i
        out_data = out
        Erase out
        If Not DataIsSpec(out_data) And SpecGetType(nm) <> 7 Then
            MsgBox ("Неверный формат файла")
            r = LogWrite(nm, "", "Неверный формат файла")
            DataRead = Empty
            Exit Function
        End If
        out_data = DataCheck(out_data) 'Проверяем и корректируем
        add_subpos = DataAddNullSubpos(out_data)
        If Not IsEmpty(add_subpos) Then out_data = ArrayCombine(add_subpos, out_data)
        out_data = DataSumByControlSum(out_data) 'Объединяем все позиции с одинаковой контрольной суммой
        Set pos_data = DataUniqParent(ArraySelectParam(out_data, t_subpos, col_type_el))
        Set pos_data.Item("weight") = DataWeightSubpos(out_data)
        If Not IsEmpty(ArraySelectParam(out_data, "-", col_sub_pos)) Then
            If pos_data.Exists("-") Then
                pos_data.Item("-").Item("-") = 1
            Else
                Set dfirst = CreateObject("Scripting.Dictionary")
                dfirst.Item("-") = 1
                Set pos_data.Item("-") = dfirst
            End If
        End If
        DataRead = out_data
        If Not IsEmpty(out_data) Then Erase out_data
    End If
End Function

Function DataShort(ByRef array_in As Variant) As Variant
    'Домножаем количество элементов на число в первом столбце
    rows_array_in = UBound(array_in, 1)
    cols_array_in = UBound(array_in, 2)
    cols_out = cols_array_in - 1
    ReDim out(1 To rows_array_in, 1 To cols_out)
    n_row = 0
    For i = 1 To rows_array_in
        If IsNumeric(array_in(i, 1)) Then
            n_row = n_row + 1
            
            For j = 2 To cols_array_in
                out(n_row, j - 1) = array_in(i, j)
            Next j
            qty = array_in(i, 1)
            out(n_row, col_qty) = out(n_row, col_qty) * array_in(i, 1)
            'Select Case out(n_row, col_type_el)
                'Case t_arm
                    'out(n_row, col_qty) = out(n_row, col_qty) * qty
                'Case t_prokat
                    'out(n_row, col_qty) = out(n_row, col_qty) * qty
                'Case t_else
                    'out(n_row, col_qty) = out(n_row, col_qty) * qty
                'Case t_izd
                    'out(n_row, col_qty) = out(n_row, col_qty) * qty
                'Case t_mat
                    'out(n_row, col_qty) = out(n_row, col_qty) * qty
            'End Select
        End If
    Next i
    DataShort = out
End Function

Function DataSumByControlSum(ByVal array_in As Variant)
    'Суммирует количество элементов с одинаковой контрольной суммой
    If IsEmpty(array_in) Then
        DataSumByControlSum = Empty
        Exit Function
    End If
    n_row = UBound(array_in, 1)
    n_col = UBound(array_in, 2)
    Dim out_data
    Dim sum_by_type
    For Each t_el In Array(t_arm, t_prokat, t_mat, t_izd, t_subpos, t_wind, t_perem)
        arr_by_type = ArraySelectParam(array_in, t_el, col_type_el)
        If Not IsEmpty(arr_by_type) Then
            un_controlsum_type = ArrayUniqValColumn(arr_by_type, col_chksum)
            ReDim sum_by_type(UBound(un_controlsum_type), n_col)
            For i = 1 To UBound(un_controlsum_type)
                For j = 1 To UBound(arr_by_type, 1)
                    If arr_by_type(j, col_chksum) = un_controlsum_type(i) Then
                        'Полностью переписываем все столбцы, если контрольная сумма пустая
                        If sum_by_type(i, col_chksum) <> un_controlsum_type(i) Then
                            For k = 1 To n_col
                                sum_by_type(i, k) = arr_by_type(j, k)
                            Next
                        Else
                            Select Case t_el
                                Case t_arm
                                    If arr_by_type(j, col_fon) Then
                                        If arr_by_type(j, col_qty) <> 1 Then MsgBox ("Ошибка суммирования арматуры в п.м.")
                                        sum_by_type(i, col_qty) = 1
                                        l_pos = arr_by_type(j, col_qty) * arr_by_type(j, col_length)
                                        sum_by_type(i, col_length) = sum_by_type(i, col_length) + l_pos
                                    Else
                                        sum_by_type(i, col_qty) = sum_by_type(i, col_qty) + arr_by_type(j, col_qty)
                                    End If
                                Case Else
                                    'Суммируем
                                    sum_by_type(i, col_qty) = sum_by_type(i, col_qty) + arr_by_type(j, col_qty)
                            End Select
                        End If
                    End If
                Next j
            Next i
            out_data = ArrayCombine(sum_by_type, out_data)
        End If
    Next
    DataSumByControlSum = out_data
    Erase array_in, out_data, sum_by_type
End Function

Function DataUniqParent(ByVal sub_pos_arr As Variant) As Variant
    'Возвращает словарь с главными сборками и входящими в них вложенными
    Set dparent = CreateObject("Scripting.Dictionary")
    Set dchild = CreateObject("Scripting.Dictionary")
    Set dqty = CreateObject("Scripting.Dictionary")
    Set dfirst = CreateObject("Scripting.Dictionary")
    Set out = CreateObject("Scripting.Dictionary")
    dparent.comparemode = 1
    dchild.comparemode = 1
    dqty.comparemode = 1
    dfirst.comparemode = 1
    out.comparemode = 1
    If Not IsEmpty(sub_pos_arr) Then
        un_subpos = ArrayUniqValColumn(sub_pos_arr, col_sub_pos) 'Уникальные сборки
        For Each subpos In un_subpos
            If subpos <> "-" Then 'Элементы в сборках
                flag = 1
                For i = 1 To UBound(sub_pos_arr, 1)
                    spos = sub_pos_arr(i, col_sub_pos)
                    tparent = sub_pos_arr(i, col_parent)
                    qty = sub_pos_arr(i, col_qty)
                    If ((spos = subpos) And (tparent <> "-")) Then 'Найден элемент второго уровня
                        flag = 0
                        If Not dchild.Exists(subpos) Then dchild.Item(subpos) = 1
                    End If
                Next i
                If flag And (Not dparent.Exists(subpos)) Then dparent.Item(subpos) = 1
            End If
        Next
        For i = 1 To UBound(sub_pos_arr, 1)
            spos = sub_pos_arr(i, col_sub_pos)
            tparent = sub_pos_arr(i, col_parent)
            qty = sub_pos_arr(i, col_qty)
            If dqty.Exists(tparent & "_" & spos) Then
                dqty.Item(tparent & "_" & spos) = dqty.Item(tparent & "_" & spos) + qty
            Else
                dqty.Item(tparent & "_" & spos) = qty
            End If
            If Not (tparent = "-" And dparent.Exists(spos)) Then
                If dqty.Exists("all" & spos) Then
                    dqty.Item("all" & spos) = dqty.Item("all" & spos) + qty
                Else
                    dqty.Item("all" & spos) = qty
                End If
            End If
            If tparent = "-" And dchild.Exists(spos) Then
                If Not dfirst.Exists(spos) Then dfirst.Item(spos) = 1
            End If
        Next i
    End If
    Set out.Item("parent") = dparent
    Set out.Item("child") = dchild
    Set out.Item("qty") = dqty
    If dfirst.Count Then Set out.Item("-") = dfirst
    Set out.Item("name") = DataNameSubpos(sub_pos_arr)
    Set DataUniqParent = out
End Function

Function DataWeightSubpos(ByVal array_in As Variant) As Variant
    Set dweight = CreateObject("Scripting.Dictionary")
    dweight.comparemode = 1
    Dim tweight As Double
    If (UBound(pos_data.Item("parent").keys()) < 0) Then
        Set DataWeightSubpos = dweight
        Exit Function
    End If
    'Общий вес всех элементов сборки
    For i = 1 To UBound(array_in, 1)
        subpos = array_in(i, col_sub_pos)
        type_el = array_in(i, col_type_el)
        If (subpos <> "-") Then
            tweight = 0
            Select Case type_el
                Case t_arm
                    qty = array_in(i, col_qty)
                    If (qty = 0) Or IsEmpty(qty) Then qty = 1
                    klass = array_in(i, col_klass)
                    diametr = array_in(i, col_diametr)
                    weight_pm = GetWeightForDiametr(diametr, klass)
                    length_pos = Round_w(array_in(i, col_length) / 1000, n_round_l)
                    tweight = Round_w(weight_pm * length_pos * k_zap_total, n_round_w) * qty
                Case t_prokat
                    qty = array_in(i, col_qty)
                    If (qty = 0) Or IsEmpty(qty) Then qty = 1
                    weight_pm = array_in(i, col_pr_weight)
                    length_pos = Round_w(array_in(i, col_pr_length) / 1000, n_round_l)
                    name_pr = GetShortNameForGOST(array_in(i, col_pr_gost_prof))
                    If InStr(1, name_pr, "Лист") Then
                        naen_plate = SpecMetallPlate(array_in(i, col_pr_prof), array_in(i, col_pr_naen), length_pos, weight_pm)
                        tweight = Round_w(naen_plate(4) * k_zap_total, n_round_w) * qty
                        ll = 1
                    Else
                        tweight = Round_w(weight_pm * length_pos * k_zap_total, n_round_w) * qty
                    End If
                    ll = 1
                Case t_izd
                    qty = array_in(i, col_qty)
                    If (qty = 0) Or IsEmpty(qty) Then qty = 1
                    If Not IsNumeric(array_in(i, col_m_weight)) Then
                        tweight = 0
                    Else
                        tweight = Round_w(array_in(i, col_m_weight) * k_zap_total, n_round_w) * qty
                    End If
            End Select
            If tweight Then dweight.Item(subpos) = dweight.Item(subpos) + tweight
        End If
    Next
    'Делим на количество вхождений, чтоб получить массу одной шт.
    For Each subpos In dweight.keys()
        If pos_data.Item("child").Exists(subpos) Then
            nSubPos = pos_data.Item("qty").Item("all" & subpos)
        Else
            nSubPos = pos_data.Item("qty").Item("-_" & subpos)
        End If
        If nSubPos < 1 Then
            MsgBox ("Не определено кол-во сборок " & subpos & ", принято 1 шт.")
            r = LogWrite(subpos, "", "Не определено кол-во сборок")
            nSubPos = 1
        End If
        w = (dweight.Item(subpos) / nSubPos)
        dweight.Item(subpos) = w
    Next
    'Для сборок первого уровня учтём вхождения сборок второго уровня
    For Each subpos In pos_data.Item("parent").keys()
        For Each tchild In pos_data.Item("child").keys()
            If pos_data.Item("qty").Exists(subpos & "_" & tchild) Then
                qty = pos_data.Item("qty").Item(subpos & "_" & tchild) / pos_data.Item("qty").Item("-_" & subpos)
                tweight = dweight.Item(tchild)
                dweight.Item(subpos) = dweight.Item(subpos) + qty * tweight
            End If
        Next
    Next
    Set DataWeightSubpos = dweight
End Function

Function DebugOut(ByVal pos_out As Variant, Optional ByVal module_name As String) As Boolean
    sh_name = "DEBUG"
    If SheetExist(sh_name) Then
        Set Sh = Application.ThisWorkbook.Sheets("DEBUG")
        If module_name = "clear" Then
            Sh.Cells.Clear
            Sh.Cells.ClearFormats
            Sh.Cells.ClearContents
            Sh.Cells.NumberFormat = "@"
        Else
            lsize = SheetGetSize(Sh)
            n_row_s = lsize(1)
            n_col_s = lsize(2)
            Sh.Cells(n_row_s + 2, 1) = module_name
            If Not IsEmpty(pos_out) Then
                If ArrayIsSecondDim(pos_out) Then
                    n_row_a = UBound(pos_out, 1)
                    n_col_a = UBound(pos_out, 2)
                Else
                    n_row_a = 1
                    n_col_a = UBound(pos_out)
                End If
                Sh.Range(Sh.Cells(n_row_s + 3, 1), Sh.Cells(n_row_s + n_row_a + 2, n_col_a)) = pos_out
            Else
                Sh.Cells(n_row_s + 2, 3) = "EMPTY"
            End If
        End If
    End If
End Function

Function ExportList2CSV(ByRef ra As Variant, ByVal CSVfilename As String, Optional ByVal ColumnsSeparator$ = ";", Optional ByVal RowsSeparator$ = vbNewLine) As String
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
    ExportList2CSV = ExportSaveTXTfile(CSVfilename$, CSVtext$)
End Function

Function ExportSaveTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.CreateTextFile(filename, True)
    ts.Write txt: ts.Close
    ExportSaveTXTfile = Err = 0
    Set ts = Nothing: Set fso = Nothing
End Function

Function ExportSetPageBreaks(ByRef Sh As Variant, ByVal h_list As Double, Optional ByVal n_first As Integer, Optional ByVal page_delim As String) As Boolean
    h_sheet = GetHeightSheet(Sh)
    If IsMissing(page_delim) Or Len(page_delim) < 2 Then
        If h_sheet > h_list Then
            lsize = SheetGetSize(Sh)
            n_row = lsize(1)
            n_col = lsize(2)
            Sh.ResetAllPageBreaks
            Sh.VPageBreaks.Add Before:=Sh.Cells(1, n_col)
            h_dop = 0
            For i = 1 To n_first
                h_row_point = Sh.Rows(i).RowHeight
                h_row_mm = h_row_point / 72 * 25.4
                h_dop = h_dop + h_row_mm
            Next i
            h_max = h_list + h_dop
            h_t = 0
            For i = 1 To n_row + 1
                h_row_point = Sh.Rows(i).RowHeight
                h_row_mm = h_row_point / 72 * 25.4
                h_t = h_t + h_row_mm
                If h_t >= h_max Then
                    Sh.HPageBreaks.Add Before:=Sh.Range(Sh.Cells(i, 1).MergeArea(1).Address)
                    h_t = 0
                End If
            Next i
            ExportSetPageBreaks = True
        Else
            ExportSetPageBreaks = False
        End If
    Else
        lsize = SheetGetSize(Sh)
        n_row = lsize(1)
        n_col = lsize(2)
        Sh.ResetAllPageBreaks
        Sh.VPageBreaks.Add Before:=Sh.Cells(1, n_col)
        For i = 2 To n_row + 1
            If InStr(Sh.Cells(i, 1).Text, page_delim) > 0 Then
                Sh.HPageBreaks.Add Before:=Sh.Range(Sh.Cells(i, 1).Address)
            End If
        Next i
        Sh.HPageBreaks.Add Before:=Sh.Range(Sh.Cells(n_row + 1, 1).MergeArea(1).Address)
    End If
End Function

Function ExportSheet(nm)
    type_spec = SpecGetType(nm)
    If type_spec <> 7 And type_spec > 0 And Len(nm) > 1 Then
        Set Sh = Application.ThisWorkbook.Sheets(nm)
        lsize = SheetGetSize(Sh)
        n_row = lsize(1)
        n_col = lsize(2)
        Set Data_out = Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col))
        r = SheetSetOption(nm)
        r = SetKzap()
        filename$ = ThisWorkbook.path & "\list\Спец_" & nm & "_" & ConvNum2Txt(k_zap_total * 10) & ".pdf"
        If Dir(filename) <> "" Then
            If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\list\old\") Then
                MkDir (ThisWorkbook.path & "\list\old\")
            End If
            tdate = Right(Str(DatePart("yyyy", Now)), 2) & Str(DatePart("m", Now)) & Str(DatePart("d", Now))
            stamp = "=" + tdate + "=" + Str(DatePart("h", Now)) + Str(DatePart("n", Now)) + Str(DatePart("s", Now))
            stamp = Replace(stamp, " ", "")
            Set fs = CreateObject("Scripting.FileSystemObject")
            fs.CopyFile filename, ThisWorkbook.path & "\list\old\Спец_" & nm & ConvNum2Txt(k_zap_total * 10) & stamp & ".pdf"
            Set fs = Nothing
        End If
        type_print = 0
        hh = GetHeightSheet(Sh)
        bb = GetWidthSheet(Sh)
        If SpecGetType(nm) = 11 Then type_print = 1
        If delim_group_ved And type_spec = 1 Then
            r = ExportSetPageBreaks(Sh, 420, 2, "Марка" & vbLf & "изделия.")
        Else
            If GetHeightSheet(Sh) > 420 Then r = ExportSetPageBreaks(Sh, 420, 2)
        End If
        r = ExportSheet2Pdf(Data_out, filename, type_print)
        r = LogWrite(filename, "PDF", "ОК")
    End If
End Function

Function ExportSheet2Pdf(ByVal Data_out As Range, ByVal filename As String, Optional ByVal type_print As Integer = 0) As Boolean
    Data_out.Select
    On Error Resume Next
    'Application.PrintCommunication = False
    ActiveSheet.PageSetup.PrintArea = Data_out.Address
    Select Case type_print
        Case 0
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0)
                .RightMargin = Application.InchesToPoints(0)
                .TopMargin = Application.InchesToPoints(0)
                .BottomMargin = Application.InchesToPoints(0)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 72
                .CenterHorizontally = False
                .CenterVertically = False
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA3
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                If delim_group_ved Then
                    .FitToPagesTall = False
                Else
                    .FitToPagesTall = 1
                End If
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = True
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
        Case 1
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0)
                .RightMargin = Application.InchesToPoints(0)
                .TopMargin = Application.InchesToPoints(0)
                .BottomMargin = Application.InchesToPoints(0)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 72
                .CenterHorizontally = False
                .CenterVertically = False
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA3
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = True
                .Zoom = Auto
                .FitToPagesWide = 1
                .PrintErrors = xlPrintErrorsBlank
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = True
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
    End Select
    On Error Resume Next
    'Application.PrintCommunication = True
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=filename$, Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ExportSheet2Pdf = True
End Function

Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal mask As String = "", Optional ByVal SearchDeep As Long = 2) As Collection
    Set FilenamesCollection = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetAllFileNamesUsingFSO FolderPath, mask, fso, FilenamesCollection, SearchDeep
    Set fso = Nothing: Application.StatusBar = False
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
    
    arr_warning = Array("!!!", "ИЗ ФАЙЛА", "С ЛИСТА")
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
    H = 1
    If H = 0 Then
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

Function FormatManual(ByVal nm As String) As Boolean
    'Наведение красоты на листе с ручной спецификацией
    If IsEmpty(pr_adress) Then r = ReadPrSortament()
    Set Data_out = Application.ThisWorkbook.Sheets(nm)
    size_sh = SheetGetSize(Data_out)
    nrow = size_sh(1)
    nsols = size_sh(2)
    'Имена диапазонов для каждого столбца
    r_all = FormatManuallitera(1) & ":" & FormatManuallitera(max_col_man)
    r_subpos = FormatManualrange(col_man_subpos, nrow)
    r_pos = FormatManualrange(col_man_pos, nrow)
    r_obozn = FormatManualrange(col_man_obozn, nrow)
    r_naen = FormatManualrange(col_man_naen, nrow)
    r_qty = FormatManualrange(col_man_qty, nrow)
    r_weight = FormatManualrange(col_man_weight, nrow)
    r_prim = FormatManualrange(col_man_prim, nrow)
    r_komment = FormatManualrange(col_man_komment, nrow)
        
    r_ank = FormatManualrange(col_man_ank, nrow)
    r_nahl = FormatManualrange(col_man_nahl, nrow)
    r_dgib = FormatManualrange(col_man_dgib, nrow)
            
    r_length = FormatManualrange(col_man_length, nrow)
    r_diametr = FormatManualrange(col_man_diametr, nrow)
    r_klass = FormatManualrange(col_man_klass, nrow)
    
    r_pr_length = FormatManualrange(col_man_pr_length, nrow)
    r_pr_gost_pr = FormatManualrange(col_man_pr_gost_pr, nrow)
    r_pr_prof = FormatManualrange(col_man_pr_prof, nrow)
    r_pr_type = FormatManualrange(col_man_pr_type, nrow)
    r_pr_st = FormatManualrange(col_man_pr_st, nrow)
    r_pr_okr = FormatManualrange(col_man_pr_okr, nrow)
    r_pr_ogn = FormatManualrange(col_man_pr_ogn, nrow)

    Columns(r_all).Validation.Delete
    Range(r_all).ClearOutline
    Data_out.Cells.UnMerge
    
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
    
    Data_out.Cells(1, col_man_ank) = "Всё в мм"
    Data_out.Cells(2, col_man_ank) = "Анкеровка"
    Data_out.Cells(2, col_man_nahl) = "Нахлёст"
    Data_out.Cells(2, col_man_dgib) = "Радиус оправки"
    
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
    Range("S1:U1").Merge
    
    
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
    
    Data_out.Cells(2, col_man_ank).ColumnWidth = 8
    Data_out.Cells(2, col_man_nahl).ColumnWidth = 8
    Data_out.Cells(2, col_man_dgib).ColumnWidth = 8
    
    
    Range(r_all).FormatConditions.Add Type:=xlExpression, Formula1:="=ЕОШИБКА(A1)"
    Range(r_all).FormatConditions(Range(r_all).FormatConditions.Count).SetFirstPriority
    With Range(r_all).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10040319
        .TintAndShade = 0
    End With
    Range(r_all).FormatConditions(1).StopIfTrue = False

    
    'Создаём столбец с марками элементов и добавим раскрывающийся список
    sheet_subpos_name = Left(nm, Len(nm) - 5) & "_поз"
    If SheetExist(sheet_subpos_name) Then
        Set subpos_sheet = Application.ThisWorkbook.Sheets(sheet_subpos_name)
        row = SheetGetSize(subpos_sheet)(1)
        pos = subpos_sheet.Range(subpos_sheet.Cells(3, 1), subpos_sheet.Cells(row, 1))
        If IsArray(pos) Then
            un_pos = ArrayUniqValColumn(pos, 1)
        Else
            un_pos = Array(pos)
        End If
        If Not IsEmpty(un_pos) Then
            istart = max_col_man + 1
            iend = UBound(un_pos, 1)
            'Data_out.range(Data_out.Cells(1, istart), Data_out.Cells((iEnd + 3) * 500, istart)).ClearContents
            For i = 1 To iend
                Data_out.Range(Data_out.Cells(i, istart), Data_out.Cells(i, istart)) = un_pos(i)
            Next
            With Range(r_subpos).Validation
                .Delete
                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                .IgnoreBlank = True
                .InCellDropdown = True
                .ShowInput = True
                .ShowError = False
            End With
            With Data_out.Range(Data_out.Cells(1, istart), Data_out.Cells(iend, istart)).Font
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
            With Data_out.Range(Data_out.Cells(1, istart), Data_out.Cells(iend, istart)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With Range(r_subpos).Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & Data_out.Range(Data_out.Cells(1, istart), Data_out.Cells(iend, istart)).Address
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
        On Error Resume Next
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("Примечания")
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
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("Классы")
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
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("ГОСТпрокат")
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
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("Марки стали")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    
    With Range(r_pr_okr).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("Окраска")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    For i = 1 To nrow + 100
        gost = Cells(i, col_man_pr_gost_pr).Value
        addr = pr_adress.Item(gost)
        If Not IsEmpty(addr) And Not IsEmpty(gost) Then
            With Cells(i, col_man_pr_prof).Validation
                            .Delete
                            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & addr(1)
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
        If Not IsEmpty(addr) And Not IsEmpty(klass) Then
            With Cells(i, col_man_diametr).Validation
                            .Delete
                            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & addr
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
    Columns("K:Q").Group
    Columns("K:Q").EntireColumn.Hidden = True
    FormatManual = True
End Function

Function FormatManuallitera(ByVal col As Integer) As String
    If col > 0 Then
        litera = Split(Cells(1, col).Address, "$")(1)
    Else
        litera = "A"
    End If
    FormatManuallitera = litera
End Function

Function FormatManualrange(ByVal col As Integer, ByVal nrow As Integer) As String
    litera = FormatManuallitera(col)
    out = litera & "3:" & litera & Trim(Str(nrow))
    FormatManualrange = out
End Function

Function FormatColWidth(ByVal dblWidthCm As Double, ByRef rngTarget As Range)
    dblWidthPoint = Application.CentimetersToPoints(dblWidthCm)
    If dblWidthPoint >= 255 Then dblWidthPoint = 254
    For Each col In rngTarget.Columns
        With col
            
            While .Width > dblWidthPoint
                .ColumnWidth = .ColumnWidth - 0.1
            Wend
            While .Width < dblWidthPoint
                .ColumnWidth = .ColumnWidth + 0.1
            Wend
        End With
    Next col
End Function

Function FormatRowHigh(ByVal dblHightCm As Double, ByRef rngTarget As Range)
    dblHightPoint = Application.CentimetersToPoints(dblHightCm)
    For Each row In rngTarget.Rows
        row.AutoFit
        If row.RowHeight < dblHightPoint Then
            row.RowHeight = dblHightPoint
        End If
    Next row
End Function

Function FormatRowPrint(ByRef Data_out As Range, ByVal n_row As Integer)
    Application.PrintCommunication = False
    With Application.ThisWorkbook.Sheets(Data_out.Parent.Name).PageSetup
        .PrintTitleRows = "$1:$" + CStr(n_row)
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
End Function

Function FormatSpec_AS(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
        For i = 2 To n_row
            If InStr(Data_out(i, 1), ", на ") > 0 Or InStr(Data_out(i, 1), ",**") > 0 Then Range(Cells(i, 1), Cells(i, 6)).Merge
            If InStr(Data_out(i, 1), ",**") > 0 Then Data_out(i, 1) = Left(Data_out(i, 1), Len(Data_out(i, 1)) - Len(",**"))
            If InStr(Data_out(i, 1), " Прочие") > 0 Then Range(Cells(i, 1), Cells(i, 6)).Merge
'            If IsNumeric(ConvTxt2Num(Data_out(i, 1).Value)) Then
'                Data_out.Cells(i, 1) = "'" + Replace(CStr(Data_out(i, 1)), ",", ".")
'            End If
        Next i
        s1 = 15
        s2 = 50
        s3 = 60
        s4 = 15
        s5 = 20
        s6 = 25
        sall = s1 + s2 + s3 + s4 + s5
        koeff = (sall / 209) * 100
        dblPoints = Application.CentimetersToPoints(1)
        r = FormatFont(Data_out, n_row, n_col)
        
        Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, n_col)).Rows.AutoFit
        If Range(Data_out.Cells(1, 1), Data_out.Cells(1, n_col)).RowHeight < dblPoints * 0.8 Then
            Range(Data_out.Cells(1, 1), Data_out.Cells(1, n_col)).RowHeight = dblPoints * 0.8
        End If
        Range(Data_out.Cells(1, 1), Data_out.Cells(1, 1)).ColumnWidth = (s1 / sall) * koeff
        Range(Data_out.Cells(1, 2), Data_out.Cells(1, 2)).ColumnWidth = (s2 / sall) * koeff
        Range(Data_out.Cells(1, 3), Data_out.Cells(1, 3)).ColumnWidth = (s3 / sall) * koeff
        Range(Data_out.Cells(1, 4), Data_out.Cells(1, 4)).ColumnWidth = (s4 / sall) * koeff
        Range(Data_out.Cells(1, 5), Data_out.Cells(1, 5)).ColumnWidth = (s5 / sall) * koeff
        Range(Data_out.Cells(1, 6), Data_out.Cells(1, 6)).ColumnWidth = (s6 / sall) * koeff
End Function

Function FormatSpec_ASGR(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
        n_sb = n_col - 6
        s1 = 15
        s2 = 50
        s3 = 60
        ssb = 15
        s5 = 20
        s6 = 25
        sall = s1 + s2 + s3 + ssb * n_sb + s5
        koeff = (sall / 209) * 100
        dblPoints = Application.CentimetersToPoints(1)
        r = FormatFont(Data_out, n_row, n_col)
        For i = 1 To 3
            Range(Data_out.Cells(1, i), Cells(2, i)).Merge
        Next i
        Range(Data_out.Cells(1, 4), Cells(1, n_col - 3)).Merge
        For i = n_col - 2 To n_col
            Range(Data_out.Cells(1, i), Cells(2, i)).Merge
        Next i
        Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, n_col)).Rows.AutoFit
        Range(Data_out.Cells(1, 1), Data_out.Cells(1, 1)).ColumnWidth = (s1 / sall) * koeff
        Range(Data_out.Cells(1, 2), Data_out.Cells(1, 2)).ColumnWidth = (s2 / sall) * koeff
        Range(Data_out.Cells(1, 3), Data_out.Cells(1, 3)).ColumnWidth = (s3 / sall) * koeff
        For i = 4 To n_col - 2
            Range(Data_out.Cells(1, i), Data_out.Cells(1, i)).ColumnWidth = (ssb / sall) * koeff
        Next i
        Range(Data_out.Cells(1, n_col - 1), Data_out.Cells(1, n_col - 1)).ColumnWidth = (s5 / sall) * koeff
        Range(Data_out.Cells(1, n_col), Data_out.Cells(1, n_col)).ColumnWidth = (s6 / sall) * koeff
End Function

Function FormatSpec_Fas(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
    If n_col < 5 Or n_row < 2 Then
        If n_col < 5 Then n_col = 5
        If n_row < 2 Then n_row = 2
        Set Data_out = Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, n_col))
    End If
    Data_out.Cells(1, 1) = "Поз." & vbLf & "отделки"
    Data_out.Cells(1, 2) = "Наименование" & vbLf & "элементов фасада"
    Data_out.Cells(1, 3) = "Наименование материала отделки"
    Data_out.Cells(1, 4) = "Наименование и номер эталона цвета или образец колера"
    Data_out.Cells(1, 5) = "Примечание"
    
    s1 = 20
    s2 = 45
    s3 = 65
    s4 = 30
    s5 = 25
    sall = s1 + s2 + s3 + s4 + s5
    koeff = (sall / 207.5) * 100
    dblPoints = Application.CentimetersToPoints(1)
    
    Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, n_col)).Rows.AutoFit
    If Range(Data_out.Cells(1, 1), Data_out.Cells(1, n_col)).RowHeight < dblPoints * 1.5 Then
        Range(Data_out.Cells(1, 1), Data_out.Cells(1, n_col)).RowHeight = dblPoints * 1.5
    End If
    Range(Data_out.Cells(1, 1), Data_out.Cells(1, 1)).ColumnWidth = (s1 / sall) * koeff
    Range(Data_out.Cells(1, 2), Data_out.Cells(1, 2)).ColumnWidth = (s2 / sall) * koeff
    Range(Data_out.Cells(1, 3), Data_out.Cells(1, 3)).ColumnWidth = (s3 / sall) * koeff
    Range(Data_out.Cells(1, 4), Data_out.Cells(1, 4)).ColumnWidth = (s4 / sall) * koeff
    Range(Data_out.Cells(1, 5), Data_out.Cells(1, 5)).ColumnWidth = (s5 / sall) * koeff
    r = FormatFont(Data_out, n_row, n_col)
End Function

Function FormatSpec_GR(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
    start_cell = 1
        For j = 2 To n_row - 1
            If (Data_out(j - 1, 1) <> Data_out(j, 1)) Then
                EndCell = j - 1
                Range(Cells(start_cell, 1), Cells(EndCell, 1)).Merge
                Range(Cells(start_cell, 6), Cells(EndCell, 6)).Merge
                start_cell = j
            End If
            If j = n_row - 1 Then
                EndCell = j
                Range(Cells(start_cell, 1), Cells(EndCell, 1)).Merge
                Range(Cells(start_cell, 6), Cells(EndCell, 6)).Merge
            End If
            If InStr(Data_out(j, 1), "* расход на ") > 0 Then Range(Cells(j, 1), Cells(j, 6)).Merge
        Next j
    Range(Cells(n_row, 1), Cells(n_row, 6)).Merge
    koeff = (185 / 208) * 100
    r = FormatFont(Data_out, n_row, n_col)
    dblPoints = Application.CentimetersToPoints(1)
    Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, n_col)).Rows.AutoFit
    
    If Range(Data_out.Cells(1, 1), Data_out.Cells(1, n_col)).RowHeight < dblPoints * 1.5 Then
        Range(Data_out.Cells(1, 1), Data_out.Cells(1, n_col)).RowHeight = dblPoints * 1.5
    End If
    Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, 1)).Columns.AutoFit
    Range(Data_out.Cells(1, 2), Data_out.Cells(1, 2)).ColumnWidth = 0.07 * koeff
    Range(Data_out.Cells(1, 3), Data_out.Cells(1, 3)).ColumnWidth = 0.5 * koeff
    Range(Data_out.Cells(1, 4), Data_out.Cells(1, 4)).ColumnWidth = 0.07 * koeff
    Range(Data_out.Cells(1, 5), Data_out.Cells(1, 5)).ColumnWidth = 0.1 * koeff
    Range(Data_out.Cells(1, 6), Data_out.Cells(1, 6)).ColumnWidth = 0.1 * koeff
End Function

Function FormatSpec_KM(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
    start_cell = 0
    For i = 1 To 2
        If start_cell = 0 Then start_cell = 1
            For j = 2 To n_row
                If (Data_out.Cells(j - 1, i) <> Data_out.Cells(j, i)) Then
                    EndCell = j - 1
                    Range(Data_out.Cells(start_cell, i), Data_out.Cells(EndCell, i)).Merge
                    start_cell = j
                End If
            Next j
        start_cell = 0
    Next i
    For i = 1 To n_row
        k = 0
        For j = 1 To 3
            If Data_out.Cells(i, j) = " " Then k = k + 1
        Next j
        If k = 2 Then Range(Data_out.Cells(i, 1), Data_out.Cells(i, 3)).Merge
        If Cells(i, 2) = "Итого" Then Range(Data_out.Cells(i, 2), Data_out.Cells(i, 3)).Merge
        If Cells(i, 1) = "Всего масса металла:" Then r_obsh = i
        If Cells(i, 1) = "Антикоррозийная окраска" Then
            Range(Data_out.Cells(i, 1), Data_out.Cells(i, n_col)).Merge
            r_okr = i
        End If
        If i > r_okr And r_okr <> 0 Then Range(Cells(i, 4), Cells(i, n_col - 1)).Merge
    Next i
    Range(Cells(1, 3), Cells(2, 3)).Merge
    Range(Cells(1, 4), Cells(2, 4)).Merge
    Range(Cells(1, 5), Cells(1, n_col - 1)).Merge
    Range(Cells(1, n_col), Cells(2, n_col)).Merge
    
    r = FormatFont(Data_out, n_row, n_col)
    If Not IsEmpty(r_okr) Then
        Range(Cells(4, 5), Cells(r_okr, n_col)).NumberFormat = w_format
        Range(Cells(r_okr, 5), Cells(n_row, n_col)).NumberFormat = "0.00"
    End If
    
    dblPoints = Application.CentimetersToPoints(1)
    Range(Data_out.Cells(1, 1), Data_out.Cells(2, n_col)).RowHeight = dblPoints * 1.5
    Range(Data_out.Cells(3, 1), Data_out.Cells(3, n_col)).RowHeight = dblPoints * 0.4
    Range(Data_out.Cells(4, 1), Data_out.Cells(n_row, n_col)).Rows.AutoFit
    koeff = 5
    Range(Data_out.Cells(1, 1), Data_out.Cells(2, 3)).ColumnWidth = 3 * koeff
    Range(Data_out.Cells(1, 4), Data_out.Cells(2, 4)).ColumnWidth = 1 * koeff
    Range(Data_out.Cells(1, 5), Data_out.Cells(2, n_col - 1)).ColumnWidth = 1.5 * koeff
    Range(Data_out.Cells(1, n_col), Data_out.Cells(2, n_col)).ColumnWidth = 2.5 * koeff

    Set MyRange = Range(Cells(r_obsh, n_col), Cells(r_obsh, n_col))
    MyRange.Font.Bold = True
    With MyRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
End Function

Function FormatSpec_KZH(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
    n_col_bet = 0
    For i = 1 To n_col
        If InStr(Data_out(2, i), "етон") > 0 Then
            n_col = n_col - 1
            n_col_bet = n_col_bet + 1
        End If
    Next i
    If n_col_bet > 0 Then
        If n_col_bet > 1 Then
            n_col_bet = n_col_bet + 1
            n_col = n_col - 1
            Range(Cells(1, n_col + 1), Cells(1, n_col + n_col_bet - 1)).Merge
            Range(Cells(1, n_col + n_col_bet), Cells(5, n_col + n_col_bet)).Merge
        Else
            If InStr(Data_out(2, n_col + n_col_bet), "етон") > 0 Then Range(Cells(2, n_col + n_col_bet), Cells(5, n_col + n_col_bet)).Merge
        End If
        For i = n_col + 1 To n_col + n_col_bet - 1
            If InStr(Data_out(2, i), "етон") > 0 Then Range(Cells(2, i), Cells(5, i)).Merge
        Next i
    End If
    If Not IsEmpty(Cells(2, 1).Value) Then
        Range(Cells(1, 1), Cells(5, 1)).Merge 'Объединение марки
        Range(Cells(1, n_col), Cells(5, n_col)).Merge 'Объединение изделий
        start_cell = 2
        For i = 1 To 4
            start_cell = 2
            For j = 2 To n_col
                If (Cells(i, j).Value <> Cells(i, start_cell).Value) Then
                    end_cell = j - 1
                        If end_cell <> start_cell Then
                            Range(Cells(i, start_cell), Cells(i, end_cell)).Merge
                        End If
                    start_cell = j
                End If
            Next j
        Next i
        
        For i = 2 To n_col
            If Data_out(2, i) = "Всего" Then Range(Cells(2, i), Cells(5, i)).Merge
        Next i
    End If
    r = FormatRowHigh(0.8, Data_out)
    r = FormatColWidth(1.5, Data_out)
    r = FormatColWidth(3, Data_out.Columns(1))
    r = FormatFont(Data_out.Range(Data_out.Cells(1, 1), Data_out.Cells(5, n_col)), 5, n_col)
    r = FormatFont(Data_out.Range(Data_out.Cells(6, 1), Data_out.Cells(n_row, n_col)), n_row - 6, n_col)
    With Data_out.Range(Data_out.Cells(6, 2), Data_out.Cells(n_row, n_col))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    n_del = 0
    For i = 6 To n_row
        flag = 1
        For j = 2 To n_col
            If Data_out(i, j) <> "-" Then flag = 0
        Next j
        If flag Then
            Data_out.Rows(i).Delete Shift:=xlUp
            n_del = n_del + 1
        End If
    Next i
    n_row = n_row - n_del
    If n_row = 7 Then
        Data_out.Rows(7).Delete Shift:=xlUp
        n_row = 6
    End If
    Range(Data_out.Cells(n_row, 1), Data_out.Cells(n_row, n_col)).Font.Bold = True
    With Data_out.Cells(n_row, n_col).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    If n_col_bet > 0 Then
        r = FormatFont(Data_out.Range(Data_out.Cells(1, n_col + 1), Data_out.Cells(n_row, n_col + n_col_bet)), n_row, n_col + n_col_bet)
        With Data_out.Cells(n_row, n_col + n_col_bet).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent2
            .TintAndShade = 0.599993896298105
            .PatternTintAndShade = 0
        End With
        For i = n_col + 1 To n_col + n_col_bet - 1
            If InStr(Data_out(2, i), "етон") > 0 Then
                With Range(Cells(2, i), Cells(5, i))
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .WrapText = True
                    .Orientation = 90
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
            End If
        Next i
    End If
End Function

Function FormatSpec_NRM(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean

    r = FormatFont(Data_out, n_row, n_col)
    Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, n_col)).Rows.AutoFit
    
    Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, 1)).ColumnWidth = 15
    Range(Data_out.Cells(1, 2), Data_out.Cells(n_row, 2)).ColumnWidth = 25
    Range(Data_out.Cells(1, 3), Data_out.Cells(n_row, 5)).ColumnWidth = 15

    Range(Data_out.Cells(n_row, 1), Data_out.Cells(n_row, n_col)).Font.Bold = True
    Range(Data_out.Cells(1, 1), Data_out.Cells(1, n_col)).Font.Bold = True
    
    Data_out.Range(Data_out.Cells(2, 1), Data_out.Cells(n_row, n_col)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = True
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Data_out.Cells(n_row, n_col).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
   
    Data_out.Range(Data_out.Cells(2, 5), Data_out.Cells(n_row - 1, 5)).Select
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With

End Function
 
Function FormatSpec_Pol(ByVal Data_out As Range) As Boolean
    CSVfilename$ = ThisWorkbook.path & "\list\Спец_" & ThisWorkbook.ActiveSheet.Name & ".txt"
    n = ExportList2CSV(Data_out, CSVfilename$)
    FormatSpec_Pol = True
End Function

Function FormatSpec_Split(ByVal Data_out As Range) As Boolean
    Data_out.Range("A1").FormulaR1C1 = "Имя листа"
    Data_out.Range("B1").FormulaR1C1 = "Список значений параметров зоны"
    Data_out.Range("C1").FormulaR1C1 = "Номер столбца параметров"
    With Data_out
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
    With Data_out.Font
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
    Data_out.Borders(xlDiagonalDown).LineStyle = xlNone
    Data_out.Borders(xlDiagonalUp).LineStyle = xlNone
    Data_out.Borders(xlEdgeLeft).LineStyle = xlNone
    Data_out.Borders(xlEdgeTop).LineStyle = xlNone
    Data_out.Borders(xlEdgeBottom).LineStyle = xlNone
    Data_out.Borders(xlEdgeRight).LineStyle = xlNone
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
    Data_out.Columns("B:B").ColumnWidth = 35
    Data_out.Columns("C:C").ColumnWidth = 11.57
    Data_out.Cells.Rows.AutoFit
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
    Selection.FormatConditions.Add Type:=xlTextString, String:="Исключить", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="Добавить", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Columns("C:C").ColumnWidth = 60
    Columns("B:B").ColumnWidth = 40
    Columns("A:A").ColumnWidth = 60
    Rows("1:1").RowHeight = 45
    Range("A1:C1").Select
    Selection.Font.Bold = True
 End Function

Function FormatSpec_Ved(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
    s_mat = 4
    s_ar = 1.5
    s1 = 1
    s2 = 5
    sp = 3
    Cells.UnMerge
    Cells.NumberFormat = "@"
    Range(Data_out.Cells(1, 1), Data_out.Cells(2, 1)).Merge
    Range(Data_out.Cells(1, 2), Data_out.Cells(2, 2)).Merge
    Range(Data_out.Cells(1, 3), Data_out.Cells(1, n_col - 1)).Merge
    Range(Data_out.Cells(1, n_col), Data_out.Cells(2, n_col)).Merge
    For i = 1 To n_row
        If InStr(Data_out.Cells(i, 1), "Общяя площадь") > 0 Or InStr(Data_out.Cells(i, 5), "Общяя площадь") > 0 Then
            n_all = n_row
            n_row = i - 1
            n_start_all = i
        End If
    Next i
    If n_all = Empty Then
        n_all = n_row
        n_start_all = n_all
    End If
    n_start = 3
    n_end = 3
    For i = 3 To n_row
        temp = Trim(Data_out.Cells(i, 1).MergeArea.Cells(1, 1).Value)
        If temp = Empty Or temp = "-" Then n_end = i
        If temp <> Empty And temp <> "-" Then
            If n_end > n_start Then
                Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, 1)).Merge
                If zonenum_pot = False Then Range(Data_out.Cells(n_start, 2), Data_out.Cells(n_end, 2)).Merge
                Range(Data_out.Cells(n_start, n_col), Data_out.Cells(n_end, n_col)).Merge
                
                With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlMedium
                End With
                With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlInsideVertical)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlInsideHorizontal)
                    .LineStyle = xlContinuous
                    .ColorIndex = 0
                    .TintAndShade = 0
                    .Weight = xlThin
                End With
                
            End If
            n_start = i
        End If
        If i = n_row And temp = Empty Or temp = "-" Then
            n_end = i
            Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, 1)).Merge
            If zonenum_pot = False Then Range(Data_out.Cells(n_start, 2), Data_out.Cells(n_end, 2)).Merge
            Range(Data_out.Cells(n_start, n_col), Data_out.Cells(n_end, n_col)).Merge
            Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlDiagonalDown).LineStyle = xlNone
            Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlDiagonalUp).LineStyle = xlNone
            
            With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Range(Data_out.Cells(n_start, 1), Data_out.Cells(n_end, n_col)).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
        End If
    Next i
    If zonenum_pot = False Then
        n_cst = 3
    Else
        n_cst = 2
    End If
    For n_c = n_cst To n_col - 1
        n_start = 3
        n_end = 3
        For i = 3 To n_row
            temp = Trim(Data_out.Cells(i, n_c).MergeArea.Cells(1, 1).Value)
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

    If UserForm2.merge_material_CB.Value And Not UserForm2.otd_by_type_CB.Value Then
        For n_c = 3 To n_col - 1
            If InStr(Data_out.Cells(2, n_c), "Площадь") = 0 Then
                temp_1 = Data_out.Cells(n_start, n_c).MergeArea.Cells(1, 1).Value
                n_start = 3
                n_end = 3
                For i = 3 To n_row
                    temp_2 = Data_out.Cells(i, n_c).MergeArea.Cells(1, 1).Value
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
    End If
    
    If UserForm2.otd_by_type_CB.Value Then
        For n_c = 3 To n_col - 1
            If InStr(Data_out.Cells(2, n_c), "Высота") > 0 Then
                temp_1 = Data_out.Cells(n_start, n_c).MergeArea.Cells(1, 1).Value
                n_start = 3
                n_end = 3
                For i = 3 To n_row
                    temp_2 = Data_out.Cells(i, n_c).MergeArea.Cells(1, 1).Value
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
    End If
    
    If show_mat_area Then
        Range(Data_out.Cells(n_start_all, 1), Data_out.Cells(n_start_all, 4)).Merge
        Range(Data_out.Cells(n_start_all, 5), Data_out.Cells(n_start_all, 8)).Merge
        For i = n_start_all + 1 To n_all
            Range(Data_out.Cells(i, 1), Data_out.Cells(i, 3)).Merge
            Range(Data_out.Cells(i, 5), Data_out.Cells(i, 7)).Merge
        Next i
    End If
    With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        With Range(Data_out.Cells(1, 1), Data_out.Cells(n_start_all - 1, n_col)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(Data_out.Cells(1, 1), Data_out.Cells(2, n_col)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(Data_out.Cells(1, 1), Data_out.Cells(2, n_col)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(Data_out.Cells(1, 1), Data_out.Cells(2, n_col)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(Data_out.Cells(1, 1), Data_out.Cells(2, n_col)).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(Data_out.Cells(1, 1), Data_out.Cells(2, n_col)).Borders(xlInsideHorizontal)
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
    
    r = FormatRowHigh(0.5, Data_out.Rows(1))
    r = FormatRowHigh(0.8, Range(Data_out.Cells(2, 1), Data_out.Cells(n_row, n_col)))
    
    r = FormatColWidth(s1, Data_out.Columns(1))
    r = FormatColWidth(s2, Data_out.Columns(2))
    r = FormatColWidth(s_mat, Data_out.Columns(3))
    r = FormatColWidth(s_ar, Data_out.Columns(4))
    r = FormatColWidth(s_mat, Data_out.Columns(5))
    r = FormatColWidth(s_ar, Data_out.Columns(6))

    If Data_out.Cells(2, 7).Value = "Колонн" Then
        r = FormatColWidth(s_mat, Data_out.Columns(7))
        r = FormatColWidth(s_ar, Data_out.Columns(8))
        If Data_out.Cells(2, 9).Value = "Низа стен/колонн" Then
            r = FormatColWidth(s_mat, Data_out.Columns(9))
            r = FormatColWidth(s_ar, Data_out.Columns(10))
            r = FormatColWidth(s_ar, Data_out.Columns(11))
        End If
    Else
        If Data_out.Cells(2, 7).Value = "Низа стен/колонн" Then
            r = FormatColWidth(s_mat, Data_out.Columns(7))
            r = FormatColWidth(s_ar, Data_out.Columns(8))
            r = FormatColWidth(s_ar, Data_out.Columns(9))
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
        If InStr(g, "Тип") And UserForm2.otd_by_type_CB.Value Then Data_out.Cells(1, n_c).Orientation = 90
        If InStr(g, "Номер") And Not UserForm2.otd_by_type_CB.Value Then Data_out.Cells(1, n_c).Orientation = 90
    Next n_c
    r = FormatColWidth(sp, Data_out.Columns(n_col))
    Data_out.FormatConditions.Add Type:=xlTextString, String:="НЕТ ОТДЕЛКИ", TextOperator:=xlContains
    With Data_out.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Data_out.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    FormatSpec_Ved = True
End Function

Function FormatSpec_WIN(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
    s1 = 1.5
    s2 = 5.5
    s3 = 6.5
    sqty = 2
    sprim = 2.5
    
    If by_floor Then
        start_row = 3
    Else
        start_row = 2
    End If
    
    For n_c = 1 To 2
        n_start = start_row
        n_end = start_row
        temp_1 = Data_out.Cells(n_start, n_c).MergeArea.Cells(1, 1).Value
        For i = start_row To n_row
            temp_2 = Data_out.Cells(i, n_c).MergeArea.Cells(1, 1).Value
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
    Next n_c
    
    r = FormatColWidth(s1, Data_out.Columns(1))
    r = FormatColWidth(s2, Data_out.Columns(2))
    r = FormatColWidth(s3, Data_out.Columns(3))
    r = FormatColWidth(sprim, Range(Data_out.Cells(1, n_col - 1), Data_out.Cells(n_row, n_col)))
    If UserForm2.qtyOneSubpos_CB.Value Then
        For i = 1 To 3
            Range(Data_out.Cells(1, i), Data_out.Cells(2, i)).Merge
        Next i
        Range(Data_out.Cells(1, 4), Data_out.Cells(1, n_col - 3)).Merge
        For i = n_col - 2 To n_col
            Range(Data_out.Cells(1, i), Data_out.Cells(2, i)).Merge
        Next i
        r = FormatRowHigh(0.5, Data_out.Rows(1))
        r = FormatRowHigh(0.8, Data_out.Rows(2))
        r = FormatRowHigh(0.8, Range(Data_out.Cells(3, n_col), Data_out.Cells(n_row, n_col)))
        r = FormatColWidth(sqty, Range(Data_out.Cells(1, 4), Data_out.Cells(n_row, n_col - 2)))
        r = FormatRowPrint(Data_out, 2)
    Else
        r = FormatRowHigh(1.5, Data_out.Rows(1))
        r = FormatRowHigh(0.8, Range(Data_out.Cells(2, n_col), Data_out.Cells(n_row, n_col)))
        r = FormatColWidth(sqty, Range(Data_out.Cells(1, 4), Data_out.Cells(n_row, n_col - 2)))
        r = FormatRowPrint(Data_out, 1)
    End If
    r = FormatFont(Data_out, n_row, n_col)
End Function

Function FormatTable(ByVal nm As String, Optional ByVal pos_out As Variant) As Boolean
    Set Sh = Application.ThisWorkbook.Sheets(nm)
    If IsError(pos_out) Or IsEmpty(pos_out) Then
        lsize = SheetGetSize(Sh)
        n_row = lsize(1)
        n_col = lsize(2)
        Set Data_out = Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col))
    Else
        n_row = UBound(pos_out, 1)
        n_col = UBound(pos_out, 2)
        pos_out = ArrayEmp2Space(pos_out)
        Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col)) = pos_out
        Set Data_out = Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col))
    End If
    
    type_spec = SpecGetType(Sh.Name)
    If type_spec <> 7 Then r = FormatClear()
    Select Case type_spec
        Case 1
            r = FormatSpec_GR(Data_out, n_row, n_col)
        Case 2, 3
            r = FormatSpec_AS(Data_out, n_row, n_col)
        Case 4
            r = FormatSpec_KM(Data_out, n_row, n_col)
        Case 5
            r = FormatSpec_KZH(Data_out, n_row, n_col)
        Case 6
            r = FormatSpec_AS(Data_out, n_row, n_col)
        Case 8
            r = FormatSpec_Fas(Data_out, n_row, n_col)
        Case 11
            r = FormatSpec_Ved(Data_out, n_row, n_col)
        Case 12
            r = FormatSpec_Pol(Data_out)
        Case 13
            r = FormatSpec_ASGR(Data_out, n_row, n_col)
        Case 14
            r = FormatSpec_NRM(Data_out, n_row, n_col)
        Case 20
            r = FormatSpec_WIN(Data_out, n_row, n_col)
        Case 21
            r = FormatSpec_Split(Data_out)
    End Select
    FormatTable = True
End Function

Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal mask As String, ByRef fso, ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
    On Error Resume Next: Set curfold = fso.GetFolder(FolderPath)
    If Not curfold Is Nothing Then
        For Each fil In curfold.Files
            If fil.Name Like "*" & mask Then FileNamesColl.Add fil.path
        Next
        SearchDeep = SearchDeep - 1
        If SearchDeep Then
           For Each sfol In curfold.SubFolders
               GetAllFileNamesUsingFSO sfol.path, mask, fso, FileNamesColl, SearchDeep
            Next
        End If
        Set fil = Nothing: Set curfold = Nothing
    End If
End Function

Function GetAreaList(razm As String) As Double
    ab = Split(razm, "x")
    If UBound(ab) < 1 Then ab = Split(razm, "х")
    If UBound(ab) < 1 Then ab = Split(razm, "*")
    If UBound(ab) < 1 Then
        GetAreaList = 0
        Exit Function
    End If
    aa = ConvTxt2Num(ab(0))
    bb = ConvTxt2Num(ab(1))
    GetAreaList = aa * bb
End Function

Function GetGOSTForKlass(ByVal klass As String) As String
    If IsEmpty(gost2fklass) Then r = ReadReinforce()
    GetGOSTForKlass = gost2fklass.Item(klass)
End Function

Function GetHeightSheet(ByRef Sh As Variant) As Double
    n_row = SheetGetSize(Sh)(1)
    h_sheet = 0
    For i = 1 To n_row
        h_row_point = Sh.Rows(i).RowHeight
        h_row_mm = h_row_point / 72 * 25.4
        h_sheet = h_sheet + h_row_mm
    Next i
    GetHeightSheet = h_sheet
End Function

Function GetWidthSheet(ByRef Sh As Variant) As Double
    n_col = SheetGetSize(Sh)(2)
    w_sheet = 0
    For i = 1 To n_col
        w_col_point = Sh.Columns(i).Width
        w_col_mm = w_col_point / 72 * 25.4
        w_sheet = w_sheet + w_col_mm
    Next i
    GetWidthSheet = w_sheet
End Function

Function GetListFile(ByRef mask As String) As Variant
    path = ThisWorkbook.path & "\import"
    Set coll = FilenamesCollection(path, mask)
    Dim out(): ReDim out(coll.Count, 2)
    i = 0
    For Each fl In coll
        i = i + 1
        fname = RelFName(fl)
        out(i, 1) = fname
        out(i, 2) = fl
    Next
    out = ArraySort(out, 1)
    GetListFile = out
    Erase out
End Function

Function GetListOfSheet(ByRef objCBook As Variant) As Variant
    n = objCBook.Worksheets.Count
    Dim out(): ReDim out(1)
    For Each objWh In objCBook.Worksheets
        sNameLst = objWh.Name
        If InStr(sNameLst, "!") = 0 Then
            c_size = UBound(out)
            out(c_size) = sNameLst
            ReDim Preserve out(c_size + 1)
        End If
    Next
    ReDim Preserve out(c_size)
    out = ArraySort(out, 1)
    GetListOfSheet = out
    Erase out
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

Function GetNSubpos(ByVal subpos As String, ByVal type_spec As Integer) As Integer
    'Получаем количество сборок с именем = subpos
    Dim nSubPos As Integer
    If subpos <> "-" Then
        If type_spec = 1 Then
            nSubPos = pos_data.Item("qty").Item("all" & subpos)
            If nSubPos = 0 Then nSubPos = pos_data.Item("qty").Item("-_" & subpos)
        Else
            nSubPos = pos_data.Item("qty").Item("-_" & subpos)
        End If
        If nSubPos < 1 Then
            MsgBox ("Не определено кол-во сборок " & subpos & ", принято 1 шт.")
            r = LogWrite(subpos, "", "Не определено кол-во сборок")
            nSubPos = 1
        End If
    Else
        nSubPos = 1
    End If
    GetNSubpos = nSubPos
End Function

Function GetNumberConstr(ByVal unique_type_konstr As Variant, ByVal konstr As String) As Integer
    For i = 1 To UBound(unique_type_konstr)
        If unique_type_konstr(i) = konstr Then
            GetNumberConstr = i
        End If
    Next i
End Function

Function GetNumberStal(ByVal unique_stal As Variant, ByVal stal As String) As Integer
    For i = 1 To UBound(unique_stal)
        If unique_stal(i) = stal Then
            GetNumberStal = i
        End If
    Next i
End Function

Function GetSheetOfBook(ByRef objCloseBook As Variant, ByVal sName As String) As Worksheet
    Set GetSheetOfBook = objCloseBook.Sheets(sName)
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
    r = LogWrite("Ошибка арматуры", "", "Отсутвует вес для " & diametr & " " & klass)
    GetWeightForDiametr = 1
End Function

Private Function ins_row(ByRef arr_out As Variant, ByRef arr_tmp As Variant, ByVal i As Integer, ByVal n_col_sb As Integer, ByRef n_row_ex As Integer, ByVal nSubPos As Integer) As Boolean
    end_col = UBound(arr_out, 2)
    n_row_ins = 0
    If n_row_ex > 0 Then
        For j = 1 To n_row_ex
            flag = 0
            For k = 1 To 3
                If arr_out(j, k) = arr_tmp(i, k) Then flag = flag + 1
            Next k
            If flag = 3 Then
                n_row_ins = j
                Exit For
            End If
        Next j
    End If
    If n_row_ins = 0 Then
        n_row_ex = n_row_ex + 1
        n_row_ins = n_row_ex
        arr_out(n_row_ins, 1) = arr_tmp(i, 1)
        arr_out(n_row_ins, 2) = arr_tmp(i, 2)
        arr_out(n_row_ins, 3) = arr_tmp(i, 3)
        arr_out(n_row_ins, end_col - 1) = arr_tmp(i, 5)
        If IsNumeric(arr_tmp(i, 6)) Then
            arr_out(n_row_ins, end_col) = arr_tmp(i, 6) * nSubPos
        Else
            arr_out(n_row_ins, end_col) = arr_tmp(i, 6)
        End If
    Else
        If IsNumeric(arr_tmp(i, 6)) And IsNumeric(arr_out(n_row_ins, end_col)) Then
            arr_out(n_row_ins, end_col) = arr_out(n_row_ins, end_col) + arr_tmp(i, 6) * nSubPos
        End If
    End If
    arr_out(n_row_ins, n_col_sb) = arr_tmp(i, 4)
    arr_out(n_row_ins, end_col - 2) = arr_out(n_row_ins, end_col - 2) + arr_tmp(i, 4) * nSubPos
End Function

Function LogNewSheet(ByVal log_sheet_name As String)
    ThisWorkbook.Worksheets.Add.Name = log_sheet_name
    Set log_sheet = Application.ThisWorkbook.Sheets(log_sheet_name)
    ThisWorkbook.Worksheets(log_sheet_name).Move After:=ThisWorkbook.Sheets(1)
    Sheets(log_sheet_name).Visible = False
    i = 0
    i = i + 1: log_sheet.Cells(1, i) = "Время"
    i = i + 1: log_sheet.Cells(1, i) = "Логин"
    i = i + 1: log_sheet.Cells(1, i) = "Лист"
    i = i + 1: log_sheet.Cells(1, i) = "Тип"
    i = i + 1: log_sheet.Cells(1, i) = "Результат"
    i = i + 1: log_sheet.Cells(1, i) = "calc"
    i = i + 1: log_sheet.Cells(1, i) = "common"
    i = i + 1: log_sheet.Cells(1, i) = "form"
    i = i + 1: log_sheet.Cells(1, i) = "Коэфф. запаса"
    i = i + 1: log_sheet.Cells(1, i) = "Арм п.м."
    i = i + 1: log_sheet.Cells(1, i) = "Про п.м."
    i = i + 1: log_sheet.Cells(1, i) = "Поз"
    i = i + 1: log_sheet.Cells(1, i) = "На одну"
    i = i + 1: log_sheet.Cells(1, i) = "Сборки"
    i = i + 1: log_sheet.Cells(1, i) = "Игнор"
    n_col = i
    Range(log_sheet.Cells(1, 1), log_sheet.Cells(1, n_col)).RowHeight = 95
    Range(log_sheet.Cells(2, 1), log_sheet.Cells(1000, n_col)).RowHeight = 16
    i = 0
    i = i + 1: log_sheet.Cells(1, i).ColumnWidth = 16
    i = i + 1: log_sheet.Cells(1, i).ColumnWidth = 20
    i = i + 1: log_sheet.Cells(1, i).ColumnWidth = 40
    i = i + 1: log_sheet.Cells(1, i).ColumnWidth = 9
    i = i + 1: log_sheet.Cells(1, i).ColumnWidth = 20

    Range(log_sheet.Cells(1, i), log_sheet.Cells(1, n_col)).ColumnWidth = 8
    
    With Range(log_sheet.Cells(1, 1), log_sheet.Cells(1000, n_col))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range(log_sheet.Cells(1, i), log_sheet.Cells(1, n_col)).Orientation = 90
    Range(log_sheet.Cells(1, 1), log_sheet.Cells(1, n_col)).Font.Bold = True
End Function

Function LogWrite(ByVal sheet_name As String, ByVal suffix As String, ByVal rezult As String)
    log_sheet_name = "|Лог|"
    If Not SheetExist(log_sheet_name) Then r = LogNewSheet(log_sheet_name)
    Set log_sheet = Application.ThisWorkbook.Sheets(log_sheet_name)
    log_sheet.Visible = False
    j = SheetGetSize(log_sheet)(1) + 1
    i = 0
    i = i + 1: log_sheet.Cells(j, i) = Now
    i = i + 1: log_sheet.Cells(j, i) = Environ$("computername") & "-" & Environ$("username")
    i = i + 1: log_sheet.Cells(j, i) = sheet_name
    i = i + 1: log_sheet.Cells(j, i) = suffix
    i = i + 1: log_sheet.Cells(j, i) = rezult
    i = i + 1: log_sheet.Cells(j, i) = macro_version
    i = i + 1: log_sheet.Cells(j, i) = common_version
    i = i + 1: log_sheet.Cells(j, i) = UserForm2.form_ver.Caption
    i = i + 1: log_sheet.Cells(j, i) = k_zap_total
    i = i + 1: log_sheet.Cells(j, i) = UserForm2.arm_pm_CB.Value
    i = i + 1: log_sheet.Cells(j, i) = UserForm2.pr_pm_CB.Value
    i = i + 1: log_sheet.Cells(j, i) = UserForm2.keep_pos_CB
    i = i + 1: log_sheet.Cells(j, i) = UserForm2.qtyOneSubpos_CB
    i = i + 1: log_sheet.Cells(j, i) = UserForm2.show_subpos_CB
    i = i + 1: log_sheet.Cells(j, i) = UserForm2.ignore_subpos_CB
End Function

Function ManualAdd(ByVal lastfileadd As String) As Boolean
    nm = ActiveSheet.Name
    If SpecGetType(nm) <> 7 Then
        MsgBox ("Перейдите на лист с ручной спецификацией (заканчивается на _спец) и повторите")
        ManualAdd = False
        Exit Function
    End If
    If Right(lastfileadd, 4) = "_поз" Then
        add_array = ReadPos(lastfileadd)
    Else
        add_array = DataRead(lastfileadd)
    End If
    
    add_array = DataSumByControlSum(add_array)
    man_arr = DataRead(nm)

    For Each t_el In Array(t_arm, t_prokat, t_mat, t_izd, t_subpos)
        t = ManualDiff(add_array, man_arr, t_el)
        If IsArray(t) Then diff_array = ArrayCombine(diff_array, t)
    Next
    
    If Not IsEmpty(diff_array) Then
        For i = 1 To UBound(add_array, 1)
            For j = 1 To UBound(diff_array)
                If diff_array(j) = add_array(i, col_chksum) Then
                    add_array(i, col_marka) = "mod"
                End If
            Next j
        Next i
    End If
    sub_pos_arr = ArraySelectParam(add_array, t_subpos, col_type_el)
    Dim array_out(): ReDim array_out(UBound(add_array, 1), UBound(add_array, 2))
    n_row = 0
    If IsEmpty(sub_pos_arr) Then
        For j = 1 To UBound(add_array, 1)
            n_row = n_row + 1
            For k = 1 To UBound(add_array, 2)
                array_out(n_row, k) = add_array(j, k)
            Next k
        Next j
    Else
        For i = 1 To UBound(sub_pos_arr, 1)
            n_row = n_row + 1
            For k = 1 To UBound(add_array, 2)
                array_out(n_row, k) = sub_pos_arr(i, k)
            Next k
            
            For j = 1 To UBound(add_array, 1)
                If add_array(j, col_sub_pos) = sub_pos_arr(i, col_sub_pos) And add_array(j, col_type_el) <> t_subpos Then
                    n_row = n_row + 1
                    For k = 1 To UBound(add_array, 2)
                        array_out(n_row, k) = add_array(j, k)
                    Next k
                End If
            Next j
        Next i
    End If
    r = ManualSpec(nm, array_out)
    r = LogWrite(nm, "add", Str(UBound(add_array, 1)))
    ManualAdd = True
End Function

Function ManualCatchChange(ByVal Target As Range)
    If IsEmpty(Target) Then Exit Function
    If IsEmpty(pr_adress) Then r = ReadPrSortament()
    nr = 0
    For Each ceil In Target.Cells
        type_izm = Empty
        nr = nr + 1
        If nr > 200 Then Exit Function
        n_colum = ceil.Column
        n_row = ceil.row
        name_colum = Cells(2, n_colum).Value
        If name_colum = "ГОСТ профиля" Then
            gost = ceil.Value
            addr = pr_adress.Item(gost)
            If Not IsEmpty(addr) And Not IsEmpty(gost) Then
                With Cells(n_row, col_man_pr_prof).Validation
                                .Delete
                                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:="=" & addr(1)
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
            If IsArray(addr) Then addr = addr(1)
            If Not IsEmpty(addr) And Not IsEmpty(klass) Then
                With Cells(n_row, col_man_diametr).Validation
                                .Delete
                                On Error Resume Next
                                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:="=" & addr
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
            If Not IsEmpty(pr_adress.Item(gost)) And Not IsEmpty(gost) And Not IsEmpty(prof) Then
                Cells(n_row, col_man_obozn) = pr_adress.Item(gost)(2)
                If Not IsEmpty(prof) Then
                    If Not IsEmpty(pr_adress.Item(gost & prof)) Then
                        Cells(n_row, col_man_weight) = pr_adress.Item(gost & prof)(1)
                        If InStr(Cells(n_row, col_man_pr_gost_pr).Value, "Лист") Then
                            If GetAreaList(Cells(n_row, col_man_naen).Value) <> Cells(n_row, col_man_pr_length).Value Then
                                Cells(n_row, col_man_pr_length).Value = GetAreaList(Cells(n_row, col_man_naen).Value)
                                Cells(n_row, col_man_pr_length).Interior.Color = XlRgbColor.rgbLightGrey
                            End If
                        Else
                            Cells(n_row, col_man_naen) = GetNameForGOST(pr_adress.Item(gost)(2)) & " " & prof
                        End If
                    Else
                        Cells(n_row, col_man_pr_prof).ClearContents
                        Cells(n_row, col_man_weight).ClearContents
                    End If
                End If
            End If
        End If
    Next
End Function

Function ManualCeilAlert(ByVal ceil As Variant, ByVal txt As String)
    On Error Resume Next
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

Function ManualCeilSetValue(ByRef ceil As Variant, ByVal val As Variant, ByVal mode As String)
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

Function ManualCheck(nm) As Boolean
    'Проверка корректности заполнения ручной спецификации
    If IsEmpty(pr_adress) Then r = ReadPrSortament()
    If SpecGetType(nm) <> 7 And SpecGetType(nm) <> 15 Then
        MsgBox ("Перейдите на лист с ручной спецификацией" & vbLf & "(заканчивается на _спец) и повторите")
        Exit Function
    End If
    If Not SheetCheckName(nm) Then Exit Function
    Set Data_out = Application.ThisWorkbook.Sheets(nm)
    r = FormatClear()
    Data_out.Cells.ClearFormats
    Data_out.Cells.ClearComments
    n_row = SheetGetSize(Data_out)(1)
    col = max_col_man
    spec = Data_out.Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, max_col_man))
    r = FormatFont(Data_out.Range(Data_out.Cells(1, 1), Data_out.Cells(n_row, max_col_man)), n_row, max_col_man)
    n_err = 0
    Set concrsubpos = CreateObject("Scripting.Dictionary")
    Set dsubpos = CreateObject("Scripting.Dictionary")
    Set ank_subpos = CreateObject("Scripting.Dictionary")
    For i = 3 To n_row
        row = ArrayRow(spec, i)
        type_el = ManualType(row)
        If type_el <> t_syserror Then
            subpos = row(col_man_subpos)  ' Марка элемента
            pos = row(col_man_pos)  ' Поз.
            obozn = row(col_man_obozn) ' Обозначение
            naen = row(col_man_naen) ' Наименование
            qty = row(col_man_qty) ' Кол-во на один элемент
            Weight = row(col_man_weight) ' Масса, кг
            prim = row(col_man_prim) ' Примечание (на лист)
            
            If Not IsNumeric(qty) And Not IsEmpty(qty) Then
                qty = ConvTxt2Num(qty)
                If Not IsNumeric(qty) Then
                    r = ManualCeilAlert(Data_out.Cells(i, col_man_qty), "Проверьте разделитель")
                    n_err = n_err + 1
                Else
                    r = ManualCeilSetValue(Data_out.Cells(i, col_man_qty), qty, "check")
                End If
            End If
    
            Select Case type_el
                Case t_sys 'Отмечаем вспомогательные строки
                    If InStr(obozn, "сновной") > 0 And InStr(naen, "етон") > 0 And InStr(subpos, "!!") > 0 Then
                        r = ManualCeilAlert(Data_out.Cells(i, col_man_subpos), "Впишите марку элемента")
                        n_err = n_err + 1
                    End If
                    If (InStr(obozn, "ейсмика") > 0 Or InStr(naen, "ейсмика") > 0) And InStr(subpos, "!!") > 0 Then
                        r = ManualCeilAlert(Data_out.Cells(i, col_man_subpos), "Впишите марку элемента")
                        n_err = n_err + 1
                    End If
                    If InStr(obozn, "сновной") > 0 And InStr(naen, "етон") > 0 And InStr(subpos, "!!") = 0 Then ank_subpos.Item(subpos & "_бет") = Trim(Replace(Replace(Replace(naen, "Бетон", ""), "бетон", ""), "B", "В"))
                    If InStr(obozn, "ейсмика") > 0 Or InStr(naen, "ейсмика") > 0 And InStr(subpos, "!!") = 0 Then ank_subpos.Item(subpos & "_kseism") = 1.3
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
                        r = ManualCeilAlert(Data_out.Cells(i, col_man_qty), "Необходимо указать количество")
                        n_err = n_err + 1
                    End If
                    If Length > 11700 And prim <> "п.м." Then
                        r = ManualCeilAlert(Data_out.Cells(i, col_man_length), "Стержни длиной выше 11,7 должны идти в п.м.")
                        n_err = n_err + 1
                    End If
                    If Length < 100 Then
                        r = ManualCeilAlert(Data_out.Cells(i, col_man_length), "Подозрительно малая длина.")
                        n_err = n_err + 1
                    End If
                    If InStr(naen, "жатая") > 0 Then ank_subpos.Item(subpos & pos & "тип") = "сжатая"
                    If InStr(naen, "астянутая") > 0 Then ank_subpos.Item(subpos & pos & "тип") = "растянутая"
                    If InStr(naen, "войная") > 0 Then ank_subpos.Item(subpos & pos & "тип") = "двойная"
                Case t_mat
                    If (prim <> "кв.м." And prim <> "куб.м.") Then
                        r = ManualCeilAlert(Data_out.Cells(i, col_man_prim), "Проверьте единицы измерения.")
                        n_err = n_err + 1
                    End If
                    If InStr(naen, "Бетон") <> 0 Then
                        concrsubpos.Item(subpos) = True
                        concrsubpos.Item(subpos & "_" & naen) = i
                        With Data_out.Range(Data_out.Cells(i, 3), Data_out.Cells(i, col_man_qty)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = XlRgbColor.rgbLightBlue
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        Data_out.Cells(i, col_man_weight).Interior.Color = XlRgbColor.rgbLightGrey
                        If IsEmpty(obozn) Then
                            r = ManualCeilAlert(Data_out.Cells(i, col_man_obozn), "Отсутствует ГОСТ на бетон")
                            n_err = n_err + 1
                        End If
                        If prim <> "куб.м." Then
                            r = ManualCeilAlert(Data_out.Cells(i, col_man_obozn), "Бетон должен быть в куб.м.")
                            n_err = n_err + 1
                        End If
                        Data_out.Cells(i, col_man_weight).Value = "-"
                    End If
                Case t_prokat
                    pr_length = row(col_man_pr_length) ' Прокат
                    pr_gost_pr = row(col_man_pr_gost_pr) ' ГОСТ профиля
                    pr_prof = row(col_man_pr_prof) ' Профиль
                    pr_type = row(col_man_pr_type) ' Тип конструкции
                    pr_st = row(col_man_pr_st) ' Сталь
                    If IsEmpty(pr_st) Then
                        r = ManualCeilAlert(Data_out.Cells(i, col_man_pr_st), "Не указана марка стали.")
                        n_err = n_err + 1
                    End If
                    Data_out.Cells(i, col_man_weight).Interior.Color = XlRgbColor.rgbLightGrey
                    If IsEmpty(qty) Then
                        r = ManualCeilAlert(Data_out.Cells(i, col_man_qty), "Необходимо указать количество")
                        n_err = n_err + 1
                    End If
                    If InStr(Data_out.Cells(i, col_man_pr_gost_pr).Value, "Лист_") Then
                        If InStr(pr_prof, "--") = 0 Then
                            r = ManualCeilAlert(Data_out.Cells(i, col_man_pr_prof), "Проверьте толщину, должно начинаться с --")
                            n_err = n_err + 1
                        Else
                            If GetAreaList(Cells(i, col_man_naen).Value) <> Cells(i, col_man_pr_length).Value Then
                                Data_out.Cells(i, col_man_pr_length).Value = GetAreaList(Data_out.Cells(i, col_man_naen).Value)
                            End If
                            Data_out.Cells(i, col_man_pr_length).Interior.Color = XlRgbColor.rgbLightGrey
                        End If
                    End If
                    If Not IsEmpty(pr_adress.Item(pr_gost_pr)) Then Data_out.Cells(i, col_man_obozn) = pr_adress.Item(pr_gost_pr)(2)
                    If Not IsEmpty(pr_adress.Item(pr_gost_pr & pr_prof)) Then Data_out.Cells(i, col_man_weight) = pr_adress.Item(pr_gost_pr & pr_prof)(1)
                    If Not IsEmpty(pr_length) And Not IsEmpty(pr_gost_pr) And Not IsEmpty(pr_prof) And Not IsEmpty(qty) And Not IsEmpty(pr_st) Then
                        With Data_out.Range(Data_out.Cells(i, col_man_pos), Data_out.Cells(i, col_man_qty)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = XlRgbColor.rgbLightGoldenrodYellow
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End If
                Case t_subpos 'Правила для маркировки сборок
                    If qty = Empty Then
                        With Data_out.Range(Data_out.Cells(i, 1), Data_out.Cells(i, max_col_man)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = XlRgbColor.rgbLightGreen
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        suff = ""
                        If IsEmpty(obozn) Then
                            r = ManualCeilAlert(Data_out.Cells(i, col_man_obozn), "Нужна ссылка на лист")
                            n_err = n_err + 1
                        End If
                    Else
                        With Data_out.Range(Data_out.Cells(i, 1), Data_out.Cells(i, max_col_man)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = XlRgbColor.rgbLightCoral
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        suff = "_par"
                    End If
                    ky = pos & " " & obozn & " " & naen & suff
                    If dsubpos.Exists(ky) Then
                        dsubpos.Item(ky) = dsubpos.Item(ky) + 1
                        dsubpos.Item(ky + "_adr") = dsubpos.Item(ky + "_adr") + "+" + Data_out.Cells(i, 1).Address
                    Else
                        dsubpos.Item(ky) = 1
                        dsubpos.Item(ky + "_adr") = Data_out.Cells(i, 1).Address
                    End If
                Case t_error
                    r = ManualCeilAlert(Data_out.Cells(i, col_man_length), "Проверьте правильность заполнения.")
                    r = ManualCeilAlert(Data_out.Cells(i, col_man_pr_length), "Проверьте правильность заполнения.")
                    n_err = n_err + 1
                Case -2
                    r = ManualCeilAlert(Data_out.Cells(i, col_man_subpos), "Пустая строка")
                    n_err = n_err + 1
                Case 0
                    With Data_out.Range(Data_out.Cells(i, 1), Data_out.Cells(i, max_col_man)).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.08
                        .PatternTintAndShade = 0
                    End With
                    With Data_out.Range(Data_out.Cells(i, 1), Data_out.Cells(i, max_col_man))
                        .Borders(xlDiagonalDown).LineStyle = xlNone
                        .Borders(xlDiagonalUp).LineStyle = xlNone
                        .Borders(xlEdgeLeft).LineStyle = xlNone
                        .Borders(xlEdgeRight).LineStyle = xlNone
                        .Borders(xlInsideVertical).LineStyle = xlNone
                        .Borders(xlInsideHorizontal).LineStyle = xlNone
                    End With
            End Select
            If type_el <> t_arm Then
                Data_out.Cells(i, col_man_ank).ClearContents
                Data_out.Cells(i, col_man_nahl).ClearContents
                Data_out.Cells(i, col_man_dgib).ClearContents
            End If
            If type_el <> t_prokat Then
                Data_out.Range(Data_out.Cells(i, col_man_pr_length), Data_out.Cells(i, col_man_pr_ogn)).ClearContents
            End If
        Else
            With Data_out.Range(Data_out.Cells(i, 1), Data_out.Cells(i, max_col_man)).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = XlRgbColor.rgbLightGrey
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            n_err = n_err + 1
        End If
    Next
    
    For i = 3 To n_row
        row = ArrayRow(spec, i)
        type_el = ManualType(row)
        If type_el = t_arm Then
            subpos = row(col_man_subpos)  ' Марка элемента
            pos = row(col_man_pos)  ' Поз.
            diametr = row(col_man_diametr) ' Диаметр
            klass = row(col_man_klass) ' Класс
            r_opr = Арм_МинРадиус(diametr, klass) - 0.5 * diametr
            Data_out.Cells(i, col_man_dgib) = r_opr
            If Not ank_subpos.Exists(subpos & "_бет") Then
                Data_out.Cells(i, col_man_ank) = "НЕТ БЕТОНА"
                Data_out.Cells(i, col_man_nahl) = "НЕТ БЕТОНА"
            Else
                beton = ank_subpos.Item(subpos & "_бет")
                kseism = 1
                If ank_subpos.Exists(subpos & "_kseism") Then kseism = 1.3
                type_arm = "растянутая"
                If ank_subpos.Exists(subpos & pos & "тип") Then type_arm = ank_subpos.Item(subpos & pos & "тип")
                type_out = "L"
                l_ank = Арм_Анкеровка(diametr, klass, beton, kseism, type_arm, type_out)
                l_nahl = Арм_Нахлёст(diametr, klass, beton, kseism, type_arm, type_out)
                Data_out.Cells(i, col_man_ank) = l_ank
                Data_out.Cells(i, col_man_nahl) = l_nahl
            End If
        End If
    Next
    For Each subpos In ArrayUniqValColumn(spec, col_man_subpos)
        If ank_subpos.Exists(subpos & "_бет") And concrsubpos.Exists(subpos) Then
            flag_eq = 0
            i = 0
            bet_ank = ank_subpos.Item(subpos & "_бет")
            bet_ank = Replace(bet_ank, "B", "В")
            For Each bet In concrsubpos.keys()
                If InStr(bet, "_") > 0 And InStr(bet, subpos) > 0 Then
                    i = concrsubpos.Item(bet)
                    bet = Replace(bet, "B", "В")
                    If InStr(bet, bet_ank) > 0 Then flag_eq = 1
                End If
            Next
            If flag_eq = 0 And i > 0 Then
                r = ManualCeilAlert(Data_out.Cells(i, col_man_naen), "Марка отличается от марки для расчёта анкеровки (" + ank_subpos.Item(subpos & "_бет") + ")")
                n_err = n_err + 1
            Else
                concrsubpos.Item(subpos & "@бет") = bet_ank
            End If
        End If
    Next
    
    If SheetExist(izd_sheet_name + "_спец.И") And nm <> izd_sheet_name Then
        Set spec_izd_sheet = Application.ThisWorkbook.Sheets(izd_sheet_name + "_спец.И")
        spec_izd_size = SheetGetSize(spec_izd_sheet)
        n_izd_row = spec_izd_size(1)
        spec_izd = spec_izd_sheet.Range(spec_izd_sheet.Cells(3, 1), spec_izd_sheet.Cells(n_izd_row, max_col_man))
        For i = 1 To UBound(spec_izd, 1)
        row = ArrayRow(spec_izd, i)
            If ManualType(row) = t_subpos Then
                pos = row(col_man_pos)  ' Поз.
                obozn = row(col_man_obozn) ' Обозначение
                naen = row(col_man_naen) ' Наименование
                ky = pos & " " & obozn & " " & naen & ""
                If dsubpos.Exists(ky) Then
                    dsubpos.Item(ky) = dsubpos.Item(ky) + 1
                    dsubpos.Item(ky + "_adr") = ""
                Else
                    dsubpos.Item(ky) = 1
                    dsubpos.Item(ky + "_adr") = ""
                End If
            End If
        Next
    End If

    For Each ky In dsubpos.keys()
        If InStr(ky, "_adr") = 0 Then
            If dsubpos.Item(ky) > 1 Then
                For Each adr In Split(dsubpos.Item(ky + "_adr"), "+")
                    adr = Replace(adr, "$", "")
                    If InStr(ky, "_par") = 0 Then
                        r = ManualCeilAlert(Data_out.Range(adr), "Повторное определение вложенной сборки (" & dsubpos.Item(ky) & " раза)")
                        n_err = n_err + 1
                    Else
                        r = ManualCeilAlert(Data_out.Range(adr), "Эта сборка повторяется " & dsubpos.Item(ky) & " раза. Не ошибка, но подозрительно.")
                    End If
                Next
            End If
            For i = 3 To n_row
                row = ArrayRow(spec, i)
                type_el = ManualType(row)
                If type_el <> t_syserror Then
                    subpos = spec(i, col_man_subpos) ' Марка элемента
                    pos = spec(i, col_man_pos) ' Поз.
                    obozn = spec(i, col_man_obozn) ' Обозначение
                    naen = spec(i, col_man_naen) ' Наименование
                    qty = spec(i, col_man_qty) ' Кол-во на один элемент
                    prim = spec(i, col_man_prim) ' Примечание (на лист)
                    Weight = spec(i, col_man_weight) ' Масса, кг
                    kyt = pos & " " & obozn & " " & naen
                    If subpos <> pos And kyt = ky Then
                        'Всякие правила для вхождений сборок
                        With Data_out.Range(Data_out.Cells(i, col_man_pos), Data_out.Cells(i, col_man_qty)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = XlRgbColor.rgbBurlyWood
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        Data_out.Cells(i, col_man_weight).Interior.Color = XlRgbColor.rgbLightGrey
                        If Not IsEmpty(prim) Then
                            r = ManualCeilAlert(Data_out.Cells(i, col_man_prim), "Вхождения сборок - только в штуках. Удали " + prim)
                            n_err = n_err + 1
                        End If
                        If Int(qty) - qty <> 0 Then
                            r = ManualCeilAlert(Data_out.Cells(i, col_man_qty), "Дробное количество сборок")
                            n_err = n_err + 1
                        End If
                        If Not IsEmpty(Weight) Then
                            r = ManualCeilAlert(Data_out.Cells(i, col_man_weight), "Масса для сборки считается автоматически. Удали " + Str(Weight))
                            n_err = n_err + 1
                        End If

                    End If
                End If
            Next i
        End If
    Next
    Range("A1").Select
    r = FormatManual(nm)
    If (n_err) Then
        MsgBox ("Обнаружено " & Str(n_err) & " ошибок, см. примечания к ячейкам")
        ManualCheck = False
    Else
        ManualCheck = True
    End If
    r = LogWrite(nm, "check", Str(n_err))
End Function

Function ManualDiff(ByVal add_array As Variant, ByVal man_arr As Variant, ByVal type_el As Integer) As Variant
    arr_a = ArrayUniqValColumn(ArraySelectParam(add_array, type_el, col_type_el), col_chksum)
    If IsEmpty(arr_a) Then ManualDiff = Empty: Exit Function
    
    arr_m = ArrayUniqValColumn(ArraySelectParam(man_arr, type_el, col_type_el), col_chksum)
    If IsEmpty(arr_m) Then ManualDiff = Empty: Exit Function
    
    Dim change_man(): n_change = 0
    
    For i = 1 To UBound(arr_a)
        chck_a = arr_a(i)
        For j = 1 To UBound(arr_m)
            chck_m = arr_m(j)
            H = InStr(chck_m, chck_a)
            If H Then
                n_change = n_change + 1
                ReDim Preserve change_man(n_change)
                change_man(n_change) = chck_a
            End If
        Next j
    Next i
    If n_change > 0 Then
        change_man = ArrayUniqValColumn(change_man, 1)
        ManualDiff = change_man
    Else
        ManualDiff = Empty
    End If
    Erase add_array, man_arr, change_man
End Function

Function ManualPaste2Sheet(ByRef array_in As Variant) As Boolean
    Set Sh = Application.ThisWorkbook.ActiveSheet
    If SpecGetType(Sh.Name) <> 7 Then
        MsgBox ("Перейдите на лист с ручной спецификацией (заканчивается на _спец) и повторите")
        ManualPaste2Sheet = False
        Exit Function
    End If
    startpos = SheetGetSize(Sh)(1) + 2
    endpos = startpos + UBound(array_in, 1) - 1
    Sh.Range(Sh.Cells(startpos, 1), Sh.Cells(endpos, max_col_man)) = array_in
    r = ManualCheck(Sh.Name)
End Function

Function ManualUndoPos(ByVal nm As String) As Boolean
    istart = 2
    Set spec_sheet = Application.ThisWorkbook.Sheets(nm)
    sheet_size = SheetGetSize(spec_sheet)
    n_row = sheet_size(1)
    If n_row = istart Then n_row = n_row + 1
    Dim pos_out: ReDim pos_out(max_col)
    spec = spec_sheet.Range(spec_sheet.Cells(1, 1), spec_sheet.Cells(n_row, max_col_man + 1))
    For i = istart To n_row
        If spec(i, max_col_man + 1) <> Empty Then
            spec_sheet.Cells(i, col_man_pos) = spec(i, max_col_man + 1)
            spec_sheet.Cells(i, max_col_man + 1) = Empty
        End If
    Next
    ManualUndoPos = True
End Function

Function posarmsort(ByRef chksum_pos As Variant, ByVal arm As Variant, ByVal cur_pos As Integer, ByVal type_pos As Integer) As Integer
    For i = 1 To UBound(arm, 1)
        chksum_part = Split(arm(i, col_chksum), "_")
        chksum = Empty
        If type_pos = 1 Then chksum = chksum_part(0) + "_" + chksum_part(3) 'Убираем позицию и сборку из контрольной суммы
        If type_pos = 2 Then chksum = chksum_part(0) + "_" + chksum_part(1) + "_" + chksum_part(3) 'Убираем позицию из контрольной суммы
        If chksum = Empty Then chksum = arm(i, col_chksum)
        arm(i, col_chksum) = chksum
    Next i
    arm = DataSumByControlSum(arm)
    'Сначала берём весь фон
    arm_temp = ArraySelectParam_2(arm, 1, col_fon, 0, col_gnut)
    If Not IsEmpty(arm_temp) Then
        'Сортируем по диаметру и длине
        arm_temp = ArraySort_2(arm_temp, col_diametr, col_length)
        For i = UBound(arm_temp, 1) To LBound(arm_temp, 1) Step -1
            cur_pos = cur_pos + 1
            chksum_pos.Item(arm_temp(i, col_chksum)) = cur_pos
        Next i
    End If
    'Остальное сортируем по длине
    'Берём прямые стержни
    arm_temp = ArraySelectParam_2(arm, 0, col_fon, 0, col_gnut)
    If Not IsEmpty(arm_temp) Then
        arm_temp = ArraySort_2(arm_temp, col_diametr, col_length)
        For i = UBound(arm_temp, 1) To LBound(arm_temp, 1) Step -1
            cur_pos = cur_pos + 1
            chksum_pos.Item(arm_temp(i, col_chksum)) = cur_pos
        Next i
    End If
    'Теперь - гнутые
    arm_temp = ArraySelectParam_2(arm, 1, col_gnut)
    If Not IsEmpty(arm_temp) Then
        arm_temp = ArraySort_2(arm_temp, col_diametr, col_length)
        For i = UBound(arm_temp, 1) To LBound(arm_temp, 1) Step -1
            cur_pos = cur_pos + 1
            chksum_pos.Item(arm_temp(i, col_chksum)) = cur_pos
        Next i
    End If
    posarmsort = cur_pos
End Function

Function ManualPos(ByVal nm As String, ByVal type_pos As Integer) As Boolean
    istart = 2
    Set spec_sheet = Application.ThisWorkbook.Sheets(nm)
    sheet_size = SheetGetSize(spec_sheet)
    n_row = sheet_size(1)
    If n_row = istart Then n_row = n_row + 1
    Dim pos_out: ReDim pos_out(max_col)
    spec = spec_sheet.Range(spec_sheet.Cells(1, 1), spec_sheet.Cells(n_row, max_col_man + 1))
    If IsEmpty(pr_adress) Then r = ReadPrSortament()
    out_data = DataRead(nm)
    'Словарь, где будем хранить контрольную сумму и позицию
    Set chksum_pos = CreateObject("Scripting.Dictionary")
    'Лишние элементы убираем
    un_parent = ArraySort(ArrayCombine(pos_data.Item("parent").keys(), Array("-"))) 'для всех сборок и элементов вне сборок
    arm = ArraySelectParam_2(out_data, t_arm, col_type_el, un_parent, col_sub_pos)
    Select Case type_pos
    Case 1
        If Not IsEmpty(arm) Then cur_pos = posarmsort(chksum_pos, arm, 0, type_pos)
    Case 2
        'Индивидуальная разбивка
        cur_pos = 0
        For i = 1 To UBound(un_parent)
            subpos = un_parent(i)
            arm_temp = ArraySelectParam(arm, subpos, col_sub_pos)
            cur_pos = 0
            If Not IsEmpty(arm_temp) Then cur_pos = posarmsort(chksum_pos, arm_temp, cur_pos, 2)
        Next i
    End Select
    If cur_pos < 1 Then Exit Function
    For i = istart To n_row
        row = ArrayRow(spec, i)
        type_el = ManualType(row)
        If type_el = t_arm Or type_el = t_prokat Then
            subpos = Trim(Replace(row(col_man_subpos), subpos_delim, "@"))  ' Марка элемента
            If spec(i, max_col_man + 1) = Empty Then spec_sheet.Cells(i, max_col_man + 1) = spec_sheet.Cells(i, col_man_pos)
            pos = Trim(Replace(row(col_man_pos), subpos_delim, "@"))  ' Поз.
            obozn = Trim(row(col_man_obozn)) ' Обозначение
            naen = Trim(row(col_man_naen)) ' Наименование
            prim = Trim(row(col_man_prim)) ' Примечание (на лист)
            If qty = Empty Or qty <= 0 Then qty = 1
            pos_out(col_marka) = pos
            pos_out(col_sub_pos) = subpos
            pos_out(col_type_el) = type_el
            pos_out(col_pos) = pos
            Select Case type_el
            Case t_arm
                Length = row(col_man_length) ' Арматура
                diametr = row(col_man_diametr) ' Диаметр
                klass = row(col_man_klass) ' Класс
                r_arm = diametr / 2000
                gnut = 0: If InStr(prim, "*") > 0 Then gnut = 1  'Ага, гнутик
                fon = 0: If InStr(prim, "п.м.") > 0 Then fon = 1  'Ага, погонаж. А fon потому, что сложилось так.
                'Можно формировать строку для спецификации
                pos_out(col_klass) = klass
                pos_out(col_diametr) = diametr
                pos_out(col_length) = Length
                pos_out(col_fon) = fon
                pos_out(col_mp) = 0
                pos_out(col_gnut) = gnut
            Case t_prokat
                pr_length = row(col_man_pr_length) ' Прокат
                pr_gost_pr = row(col_man_pr_gost_pr) ' ГОСТ профиля
                pr_prof = row(col_man_pr_prof) ' Профиль
                pr_type = row(col_man_pr_type) ' Тип конструкции
                pr_st = row(col_man_pr_st) ' Сталь
                pos_out(col_pr_type_konstr) = pr_type
                pos_out(col_pr_gost_st) = pr_adress.Item(pr_st)
                pos_out(col_pr_st) = pr_st
                pos_out(col_pr_gost_prof) = pr_adress.Item(pr_gost_pr)(2)
                pos_out(col_pr_prof) = pr_prof
                koef_l = 1
                area_okr = -1
                If InStr(pr_gost_pr, "Лист_") Then
                    koef_l = 1000
                    ab = Split(naen, "x")
                    If UBound(ab) = 0 Then ab = Split(naen, "х")
                    If UBound(ab) = 0 Then ab = Split(naen, "*")
                    If UBound(ab) > 0 Then
                        aa = Application.WorksheetFunction.Max(ab(0), ab(1))
                        bb = Application.WorksheetFunction.Min(ab(0), ab(1))
                        pos_out(col_pr_naen) = pr_prof & "x" & aa & "x" & bb
                    End If
                Else
                    pos_out(col_pr_naen) = pr_prof & " L=" & pr_length * 1000 & "мм."
                End If
                pos_out(col_pr_length) = pr_length / koef_l
                pos_out(col_pr_weight) = Weight
            End Select
            current_sum = ControlSumEl(pos_out)
            chksum_part = Split(current_sum, "_")
            chksum = Empty
            If type_pos = 1 Then chksum = chksum_part(0) + "_" + chksum_part(3)  'Убираем позицию из контрольной суммы
            If type_pos = 2 Then chksum = chksum_part(0) + "_" + chksum_part(1) + "_" + chksum_part(3)
            pos = chksum_pos.Item(chksum)
            If Not IsEmpty(pos) Then spec_sheet.Cells(i, 2) = pos
        End If
    Next i
    ManualPos = True
End Function

Function ManualSpec(ByVal nm As String, Optional ByVal add_array As Variant) As Variant
    istart = 2 'Пропускаем шапку
    If IsArray(add_array) Then
        flag_add = 1
        mod_array = ArraySelectParam(add_array, "mod", col_marka)
    Else
        flag_add = 0
        mod_array = Empty
    End If
    Set spec_sheet = Application.ThisWorkbook.Sheets(nm)
    sheet_size = SheetGetSize(spec_sheet)
    n_row = sheet_size(1)
    If n_row = istart Then n_row = n_row + 1
    spec = spec_sheet.Range(spec_sheet.Cells(1, 1), spec_sheet.Cells(n_row, max_col_man))
    If SheetExist(izd_sheet_name + "_спец.И") And nm <> izd_sheet_name Then
        Set spec_izd_sheet = Application.ThisWorkbook.Sheets(izd_sheet_name + "_спец.И")
        spec_izd_size = SheetGetSize(spec_izd_sheet)
        n_izd_row = spec_izd_size(1)
        spec_izd = spec_izd_sheet.Range(spec_izd_sheet.Cells(3, 1), spec_izd_sheet.Cells(n_izd_row, max_col_man))
        unic_pos_mun = ArrayUniqValColumn(spec, col_man_pos)
        unic_subpos_izd = ArrayUniqValColumn(spec_izd, col_man_subpos)
        For i = 1 To UBound(unic_subpos_izd)
            flag_use = False
            For j = 1 To UBound(unic_pos_mun)
                If unic_subpos_izd(i) = unic_pos_mun(j) Then
                    flag_use = True
                    Exit For
                End If
            Next j
            If flag_use = False Then
                unic_subpos_izd(i) = Empty
                t = 1
            End If
        Next i
        For Each subpos_izd In unic_subpos_izd
            If Not IsEmpty(subpos_izd) Then
                subpos_spec_izd = ArraySelectParam(spec_izd, subpos_izd, col_man_subpos)
                spec = ArrayCombine(spec, subpos_spec_izd)
                n_row = n_row + UBound(subpos_spec_izd, 1)
            End If
        Next
    End If
    Dim pos_out: ReDim pos_out(n_row - istart, max_col): n_row_out = 0
    Dim param
    Dim add_okr_array
    n_add_okr = 0
    For i = istart To n_row
        If Not IsEmpty(spec(i, col_man_pr_okr)) And spec(i, col_man_pr_okr) <> "-" Then n_add_okr = n_add_okr + 1
    Next i
    ReDim add_okr_array(n_add_okr, max_col)
    n_add_okr = 0
    For i = istart To n_row
        row = ArrayRow(spec, i)
        type_el = ManualType(row)
        If type_el > 0 And type_el <> t_sys Then
            subpos = Trim(Replace(row(col_man_subpos), subpos_delim, "@"))  ' Марка элемента
            pos = Trim(Replace(row(col_man_pos), subpos_delim, "@"))  ' Поз.
            obozn = Trim(row(col_man_obozn)) ' Обозначение
            naen = Trim(row(col_man_naen)) ' Наименование
            qty = row(col_man_qty) ' Кол-во на один элемент
            Weight = row(col_man_weight) ' Масса, кг
            prim = Trim(row(col_man_prim)) ' Примечание (на лист)
            If qty = Empty Or qty <= 0 Then qty = 1
            If type_el = t_subpos Then nSubPos = qty
            If nSubPos = Empty Or nSubPos <= 0 Then nSubPos = 1
            n_row_out = n_row_out + 1
            pos_out(n_row_out, col_marka) = pos
            pos_out(n_row_out, col_sub_pos) = subpos
            pos_out(n_row_out, col_type_el) = type_el
            pos_out(n_row_out, col_pos) = pos
            pos_out(n_row_out, col_qty) = qty * nSubPos
            Select Case type_el
            Case t_arm
                Length = row(col_man_length) ' Арматура
                diametr = row(col_man_diametr) ' Диаметр
                klass = row(col_man_klass) ' Класс
                r_arm = diametr / 2000
                gnut = 0: If InStr(prim, "*") > 0 Then gnut = 1  'Ага, гнутик
                fon = 0: If InStr(prim, "п.м.") > 0 Then fon = 1  'Ага, погонаж. А fon потому, что сложилось так.
                'Можно формировать строку для спецификации
                pos_out(n_row_out, col_klass) = klass
                pos_out(n_row_out, col_diametr) = diametr
                pos_out(n_row_out, col_length) = Length
                pos_out(n_row_out, col_fon) = fon
                pos_out(n_row_out, col_mp) = 0
                pos_out(n_row_out, col_gnut) = gnut
                pos_out(n_row_out, col_chksum + 1) = Round_w(r_arm * r_arm * Length * 3.14, 3)
            Case t_prokat
                pr_length = row(col_man_pr_length) ' Прокат
                pr_gost_pr = row(col_man_pr_gost_pr) ' ГОСТ профиля
                pr_prof = row(col_man_pr_prof) ' Профиль
                pr_type = row(col_man_pr_type) ' Тип конструкции
                pr_st = row(col_man_pr_st) ' Сталь
                pr_okr = row(col_man_pr_okr) ' Окраска
                pos_out(n_row_out, col_pr_type_konstr) = pr_type
                pos_out(n_row_out, col_pr_gost_st) = pr_adress.Item(pr_st)
                pos_out(n_row_out, col_pr_st) = pr_st
                pos_out(n_row_out, col_pr_gost_prof) = pr_adress.Item(pr_gost_pr)(2)
                pos_out(n_row_out, col_pr_prof) = pr_prof
                koef_l = 1
                area_okr = -1
                If InStr(pr_gost_pr, "Лист_") Then
                    koef_l = 1000
                    ab = Split(naen, "x")
                    If UBound(ab) = 0 Then ab = Split(naen, "х")
                    If UBound(ab) = 0 Then ab = Split(naen, "*")
                    If UBound(ab) > 0 Then
                        aa = Application.WorksheetFunction.Max(ab(0), ab(1))
                        bb = Application.WorksheetFunction.Min(ab(0), ab(1))
                        pos_out(n_row_out, col_pr_naen) = pr_prof & "x" & aa & "x" & bb
                        perim_okr = 2 / 1000
                    End If
                Else
                    perim_okr = pr_adress.Item(pr_gost_pr & pr_prof)(2)
                    pos_out(n_row_out, col_pr_naen) = pr_prof & " L=" & pr_length * 1000 & "мм."
                End If
                pos_out(n_row_out, col_pr_length) = pr_length / koef_l
                pos_out(n_row_out, col_pr_weight) = Weight
                
                If Not IsEmpty(pr_okr) And pr_okr <> "-" Then
                    n_add_okr = n_add_okr + 1
                    area_okr = perim_okr * pr_length / 1000
                    add_okr_array(n_add_okr, col_marka) = pos
                    add_okr_array(n_add_okr, col_sub_pos) = subpos
                    add_okr_array(n_add_okr, col_type_el) = t_mat
                    add_okr_array(n_add_okr, col_pos) = "-"
                    add_okr_array(n_add_okr, col_qty) = qty * area_okr
                    add_okr_array(n_add_okr, col_m_obozn) = "см. примечания"
                    add_okr_array(n_add_okr, col_m_naen) = "Покрытие " & StrConv(pr_okr, vbLowerCase)
                    add_okr_array(n_add_okr, col_m_weight) = "-"
                    add_okr_array(n_add_okr, col_m_edizm) = "кв.м."
                End If
            Case t_mat
                pos_out(n_row_out, col_m_obozn) = obozn
                pos_out(n_row_out, col_m_naen) = naen
                pos_out(n_row_out, col_m_weight) = "-"
                pos_out(n_row_out, col_m_edizm) = prim
            Case t_izd
                pos_out(n_row_out, col_m_obozn) = obozn
                pos_out(n_row_out, col_m_naen) = naen
                pos_out(n_row_out, col_m_weight) = Weight
                pos_out(n_row_out, col_m_edizm) = prim
            Case t_subpos
                pos_out(n_row_out, col_m_obozn) = obozn
                pos_out(n_row_out, col_m_naen) = naen
                pos_out(n_row_out, col_m_weight) = Weight
                pos_out(n_row_out, col_m_edizm) = prim
                pos_out(n_row_out, col_qty) = qty
            End Select
            If flag_add And Not IsEmpty(mod_array) Then
                'Изменяем строки с одинаковыми контрольными суммами
                ReDim param(max_col)
                param = ArrayRow(pos_out, n_row_out)
                current_sum = ControlSumEl(param)
                For kk = 1 To UBound(mod_array, 1)
                    mod_sum = mod_array(kk, col_chksum)
                    If mod_sum = current_sum Then
                        r = ManualCeilSetValue(spec_sheet.Cells(i, col_man_qty), mod_array(kk, col_qty), "mod")
                    End If
                Next kk
            End If
        End If
    Next i
    
    n_row_out = n_row_out + n_add_okr
    pos_out = ArrayCombine(pos_out, add_okr_array)
    If flag_add Then
        'Добавим из add_array все новые элементы (в первом столбце нет значения "mod")
        end_row = n_row_out + istart + 1
        For i = 1 To UBound(add_array, 1)
            type_el = add_array(i, col_type_el)
            If add_array(i, col_marka) <> "mod" And type_el <> t_prokat Then
                end_row = end_row + 1
                r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_subpos), add_array(i, col_sub_pos), "add")
                r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_pos), add_array(i, col_pos), "add")
                r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_qty), add_array(i, col_qty), "add")
                Select Case type_el
                Case t_arm
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_naen), "Арматура", "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_obozn), "", "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_weight), "", "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_length), add_array(i, col_length), "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_diametr), add_array(i, col_diametr), "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_klass), add_array(i, col_klass), "add")
                    If add_array(i, col_fon) Then r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_prim), "*", "add")
                    If add_array(i, col_gnut) Then r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_prim), "п.м.", "add")
                Case t_prokat
                    'r = ManualCeilSetValue(spec_sheet.Cells( end_row, col_man_naen, "Прокат", "add")
                    'r = ManualCeilSetValue(spec_sheet.Cells( end_row, col_man_obozn), add_array(i, col_pr_gost_prof), "add")
                    'r = ManualCeilSetValue(spec_sheet.Cells( end_row, col_man_weight), add_array(i, col_pr_weight), "add")
                    'r = ManualCeilSetValue(spec_sheet.Cells( end_row, col_man_length), add_array(i, col_pr_length), "add")
                    'r = ManualCeilSetValue(spec_sheet.Cells( end_row, col_man_diametr), add_array(i, col_pr_prof), "add")
                    'r = ManualCeilSetValue(spec_sheet.Cells( end_row, col_man_klass), add_array(i, col_pr_st), "add")
                    'r = ManualCeilSetValue(spec_sheet.Cells( end_row, col_man_komment, GetShortNameForGOST(add_array(i, col_pr_gost_prof)), "add")
                Case t_izd
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_obozn), add_array(i, col_m_obozn), "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_naen), add_array(i, col_m_naen), "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_weight), add_array(i, col_m_weight), "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_prim), "", "add")
                Case Else
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_obozn), add_array(i, col_m_obozn), "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_naen), add_array(i, col_m_naen), "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_weight), "", "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_prim), add_array(i, col_m_edizm), "add")
                End Select
            End If
        Next i
    Else
        sub_pos_arr = ArraySelectParam(pos_out, t_subpos, col_type_el)
        If Not IsEmpty(sub_pos_arr) Then
            'Из за того, что в дальнейшем количество элементов в сборке делится на количество сборок - нужно домножить количества
            'Для этого сначала получим количество сборок
            'Далее будем вытаскивать элементы для каждой сборки и домножать
    'For j = 1 To UBound(sub_pos_arr, 1)
        'subpos = sub_pos_arr(j, col_sub_pos)
        'qty = sub_pos_arr(j, col_qty)
        'For i = 1 To UBound(pos_out, 1)
            'If pos_out(i, col_type_el) <> t_subpos And pos_out(i, col_sub_pos) = subpos Then
                'pos_out(i, col_qty) = pos_out(i, col_qty) * qty
            'End If
        'Next i
    'Next j
            'Проверим массив на наличие вложенных сборок
            'Признак вложенной сборки - её позиция и наименование встречаются в других сборках
            'На данный момент все вложенные сборки являются изделием (t_izd)
            izd = ArraySelectParam(pos_out, t_izd, col_type_el)
            Set subpos_el = CreateObject("Scripting.Dictionary")
            For i = 1 To UBound(sub_pos_arr, 1)
                pos = sub_pos_arr(i, col_pos)
                naen = sub_pos_arr(i, col_m_naen)
                'Есть ли изделия с таким же наименованием и позицией?
                'Если есть - это сборка второго уровня.
                tmp_subpos = ArraySelectParam(izd, pos, col_pos, naen, col_m_naen)
                If Not IsEmpty(tmp_subpos) Then
                    subpos_el.Item(pos & naen) = ArraySelectParam(pos_out, pos, col_sub_pos)
                    For j = 1 To UBound(pos_out, 1)
                        If pos_out(j, col_sub_pos) = pos Then pos_out(j, col_type_el) = ""
                    Next j
                End If
            Next i
            'Осталось ещё раз пройти по элементам и добавить элементы из сборок второго уровня,
            'Поменяв при этом обозначение вхождения сборки с t_izd на t_subpos
            Dim subarray: ReDim subarray(max_col, 1)
            For j = 1 To UBound(pos_out, 1)
                If pos_out(j, col_type_el) = t_izd Then
                    pos = pos_out(j, col_pos)
                    naen = pos_out(j, col_m_naen)
                    If subpos_el.Exists(pos & naen) Then
                        subpos = pos_out(j, col_sub_pos)
                        pos_out(j, col_marka) = subpos & subpos_delim & pos_out(j, col_pos)
                        pos_out(j, col_sub_pos) = pos_out(j, col_pos)
                        pos_out(j, col_type_el) = t_subpos
                        qty = pos_out(j, col_qty)
                        el = subpos_el.Item(pos & naen)
                        arr_sub = ArraySelectParam(el, t_subpos, col_type_el)
                        If IsEmpty(arr_sub) Then
                            qty_from_list = 1
                            r = LogWrite("Повторное определение сборки " & pos, "ERROR", 0)
                        Else
                            qty_from_list = arr_sub(1, col_qty)
                            If qty_from_list <> 1 Then r = LogWrite("Подозрительное количество сборок " & pos, "ERROR", qty_from_list)
                        End If
                        For k = 1 To UBound(el, 1)
                            If el(k, col_type_el) <> t_subpos Then
                                c_size = UBound(subarray, 2)
                                For i = 1 To max_col
                                    subarray(i, c_size) = el(k, i)
                                Next i
                                subarray(col_marka, c_size) = subpos & subpos_delim & el(k, col_pos)
                                subarray(col_qty, c_size) = el(k, col_qty) * qty / qty_from_list
                                ReDim Preserve subarray(max_col, c_size + 1)
                            End If
                        Next k
                    End If
                End If
            Next j
            pos_out = ArrayCombine(pos_out, ArrayTranspose(subarray))
        End If
    End If
    ManualSpec = pos_out
    'Erase pos_out
End Function

Function ManualSpec_batch(type_out)
    r = LogWrite("Автовывод", "Начало", "-")
    If mem_option Then r = LogWrite("Автовывод", "Включена автонастройка листов", "-")
    n_out = 0
    r = OutPrepare()
    For Each objWh In ThisWorkbook.Worksheets
        nm = objWh.Name
        type_spec = SpecGetType(nm)
        If type_spec = 7 Then
            For Each tspec In type_out
                If Not IsEmpty(tspec) Then
                    If mem_option Then r = SheetSetOption(nm)
                    sheet_out = Spec_Select(nm, tspec, True)
                    r = ExportSheet(sheet_out)
                    n_out = n_out + 1
                End If
            Next
        End If
    Next objWh
    r = SheetIndex()
    r = LogWrite("Автовывод", "Конец", Str(n_out))
    r = OutEnded()
End Function

Function ManualSpec_NewSubpos()
    r = OutPrepare()
    nm_out = izd_sheet_name + "_спец.И"
    If SheetExist(nm_out) Then
        Worksheets(nm_out).Activate
    Else
        ThisWorkbook.Worksheets.Add.Name = nm_out
    End If
    Worksheets(nm_out).Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set out_sheet = ThisWorkbook.Sheets(nm_out)
    r = FormatClear()
    r = FormatManual(nm_out)
    r = FormatManual(nm_out)
    n_last = SheetGetSize(out_sheet)(1) + 2
    flag = Empty
    If UserForm2.fromthiswbCB.Value Then
        For Each objWh In ThisWorkbook.Worksheets
            nm = objWh.Name
            type_spec = SpecGetType(nm)
            If type_spec = 7 Then
                Set spec_sheet = ThisWorkbook.Sheets(nm)
                n_row = SheetGetSize(spec_sheet)(1)
                spec = spec_sheet.Range(spec_sheet.Cells(1, 1), spec_sheet.Cells(n_row, max_col_man))
                For i = 1 To n_row
                    subpos = spec(i, col_man_subpos) ' Марка элемента
                    pos = spec(i, col_man_pos) ' Поз.
                    qty = spec(i, col_man_qty) ' Кол-во на один элемент
                    Weight = spec(i, col_man_weight) ' Масса, кг
                    If (subpos = pos) And IsEmpty(qty) And IsEmpty(Weight) And Not IsEmpty(pos) And InStr(pos, "!") = 0 Then
                        flag = pos
                        n_last = n_last + 1
                    End If
                    If subpos = flag And Not IsEmpty(flag) Then
                        For j = 1 To max_col_man
                            out_sheet.Cells(n_last, j).Value = spec(i, j)
                        Next j
                        spec_sheet.Cells(i, 1) = "!" & spec_sheet.Cells(i, 1)
                        spec_sheet.Cells(i, 2) = "!" & spec_sheet.Cells(i, 2)
                        n_last = n_last + 1
                    End If
                Next i
                If Not IsEmpty(flag) Then
                    n_last = n_last + 1
                    flag = Empty
                End If
            End If
        Next
    End If
    r = ManualCheck(nm_out)
    r = OutEnded()
End Function


Function ManualSpec_Merge()
    r = OutPrepare()
    nm_out = "Сводная_спец"
    If SheetExist(nm_out) Then Worksheets(nm_out).Delete
    ThisWorkbook.Worksheets.Add.Name = nm_out
    Worksheets(nm_out).Move After:=ThisWorkbook.Sheets(3)
    Set out_sheet = ThisWorkbook.Sheets(nm_out)
    Worksheets(nm_out).Cells.Clear
    r = FormatClear()
    r = FormatManual(nm_out)
    r = FormatManual(nm_out)
    n_row_out = 4
    If UserForm2.fromthiswbCB.Value Then
        For Each objWh In ThisWorkbook.Worksheets
            nm = objWh.Name
            type_spec = SpecGetType(nm)
            If type_spec = 7 And nm <> nm_out Then
                Set spec_sheet = ThisWorkbook.Sheets(nm)
                n_row = SheetGetSize(spec_sheet)(1) + 4
                spec = spec_sheet.Range(spec_sheet.Cells(2, 1), spec_sheet.Cells(n_row, max_col_man))
                n_row = n_row - 3
                n_row_out_start = n_row_out
                n_row_out_end = n_row_out + n_row
                out_sheet.Cells(n_row_out_start - 1, 1) = "!!!"
                out_sheet.Cells(n_row_out_start - 1, 2) = "!!!"
                out_sheet.Cells(n_row_out_start - 1, 3) = "C ЛИСТА"
                out_sheet.Cells(n_row_out_start - 1, 4) = nm
                out_sheet.Hyperlinks.Add anchor:=out_sheet.Cells(n_row_out_start - 1, 4), Address:="", SubAddress:="'" & nm & "'" & "!D3"
                
                out_sheet.Range(out_sheet.Cells(n_row_out_start, 1), out_sheet.Cells(n_row_out_end, max_col_man)) = spec
                n_row_out = n_row_out_end + 3
            End If
        Next
    End If
    If UserForm2.fromfileCB.Value Then
        Set coll = FilenamesCollection(ThisWorkbook.path, ".xlsm")
        For Each snm In coll
            If InStr(snm, "Спец") And InStr(snm, "~$") = 0 And InStr(snm, ThisWorkbook.Name) = 0 Then
                Set spec_book = GetObject(snm)
                snm_short = Split(snm, "\")(UBound(Split(snm, "\")))
                listsheet = GetListOfSheet(spec_book)
                For Each nm In listsheet
                    type_spec = SpecGetType(nm)
                    If type_spec = 7 Then
                        Set spec_sheet = spec_book.Sheets(nm)
                        n_row = SheetGetSize(spec_sheet)(1) + 4
                        spec = spec_sheet.Range(spec_sheet.Cells(2, 1), spec_sheet.Cells(n_row, max_col_man))
                        n_row = n_row - 3
                        n_row_out_start = n_row_out
                        n_row_out_end = n_row_out + n_row
                        out_sheet.Cells(n_row_out_start - 1, 1) = "!!!"
                        out_sheet.Cells(n_row_out_start - 1, 2) = "!!!"
                        out_sheet.Cells(n_row_out_start - 1, 3) = "ИЗ ФАЙЛА"
                        out_sheet.Cells(n_row_out_start - 1, 4) = snm_short & " - " & nm
                        out_sheet.Range(out_sheet.Cells(n_row_out_start, 1), out_sheet.Cells(n_row_out_end, max_col_man)) = spec
                        n_row_out = n_row_out_end + 3
                    End If
                Next
                spec_book.Close True
            End If
        Next
    End If
    n_row = SheetGetSize(out_sheet)(1)
    For i = n_row To 3 Step -1
        If Trim$(out_sheet.Cells(i, 1)) = Empty And Trim$(out_sheet.Cells(i, 4)) = Empty Then out_sheet.Rows(i).Delete Shift:=xlUp
        If Trim$(out_sheet.Cells(i, 4)) = Empty And out_sheet.Cells(i, 5) = 0 Then out_sheet.Rows(i).Delete Shift:=xlUp
        If InStr(out_sheet.Cells(i, 1), "!!") <> 0 And InStr(out_sheet.Cells(i, 3), "ИЗ ФАЙЛА") = 0 And InStr(out_sheet.Cells(i, 3), "C ЛИСТА") = 0 Then out_sheet.Rows(i).Delete Shift:=xlUp
    Next
    r = ManualCheck(nm_out)
    r = OutEnded()
End Function


Function ManualType(ByVal row As Variant) As Integer
    If IsEmpty(row) Then
        ManualType = t_syserror
        Exit Function
    End If
    tempt = 0
    For i = 1 To col_man_komment - 1
        If IsError(row(i)) Then
            ManualType = t_syserror
            Exit Function
        End If
        If IsEmpty(row(i)) Then tempt = tempt + 1
    Next i
    
    If tempt = col_man_komment - 1 Then
        type_el = 0
        ManualType = type_el
        Erase row
        Exit Function
    End If

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
    isSys = (InStr(subpos, "!") > 0 Or InStr(pos, "!") > 0)
    isSPos = ((subpos = pos) And Not IsEmpty(subpos) And Not isSys)
    isArm = ((Not IsEmpty(Length) Or Not IsEmpty(diametr) Or Not IsEmpty(klass)) And Not isSys)
    isProkat = ((Not IsEmpty(pr_length) Or Not IsEmpty(pr_gost_pr) Or Not IsEmpty(pr_prof) Or Not IsEmpty(pr_prof)) And Not isSys)
    isMat = ((InStr(prim, "кв.м.") > 0 Or InStr(prim, "куб.м.") > 0 Or InStr(naen, "Бетон") > 0) And Not isSys)
    isEr = ((isSPos And isArm) Or (isSPos And isProkat) Or (isSPos And isMat) Or (isArm And isProkat) Or (isArm And isMat) Or (isProkat And isMat)) 'Проверим - не подходит ли элемент к нескольким типам
    
    If Not isSys And tempt < 3 Then type_el = -2
    If isSys Then type_el = t_sys
    If isSPos Then type_el = t_subpos
    If isArm Then type_el = t_arm
    If isProkat Then type_el = t_prokat
    If isMat Then type_el = t_mat
    If isEr Then type_el = t_error
    
    ManualType = type_el
    Erase row
End Function

Function NRowOut(ByRef arr As Variant) As Variant
    n = 0
    If Not (ArrayIsSecondDim(arr)) Then
        n = 1
    Else
        n_row = UBound(arr, 1)
        n_col = UBound(arr, 2)
        For i = 1 To n_row
            fl = 0
            If i = 6 Then
            H = 1
            End If
            For j = 1 To n_col
                el = Trim(arr(i, j))
                If el = "" Or el = " " Or el = 0 Or IsEmpty(el) Then fl = fl + 1
                
                If i < n_row Then
                    next_el = Trim(arr(i + 1, j))
                    If el <> "" And el <> " " And el <> 0 And Not IsEmpty(el) Then fl = fl - 1
                End If

            Next j
            If fl < n_col Then n = n + 1
        Next i
    End If
    NRowOut = n
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
    On Error Resume Next
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    OutPrepare = True
End Function

Function ReadFile(ByRef mask As String) As Variant
    On Error Resume Next
    Set coll = FilenamesCollection(ThisWorkbook.path & "\import\", mask)
    For Each file In coll
        arr = ArrayCombine(arr, ReadTxt(file))
    Next
    ReadFile = arr
    Erase arr
End Function

Function ReadMetall() As Boolean
    SortamentPath = SetPath()
    nf_prof = SortamentPath & "Имена профилей.csv"
    If Len(Dir$(nf_prof)) > 0 Then
        name_gost = ReadTxt(nf_prof)
    Else
        MsgBox ("Нет файла с именами профилей")
        r = LogWrite("Ошибка профилей", "", "Нет файла с именами профилей")
    End If
End Function

Function ReadPos(ByVal lastfileadd As String) As Variant
    Set add_sheet = Application.ThisWorkbook.Sheets(lastfileadd)
    sheet_size = SheetGetSize(add_sheet)
    istart = 2
    n_row = sheet_size(1)
    n_col = 6
    spec = add_sheet.Range(add_sheet.Cells(1, 1), add_sheet.Cells(n_row, n_col))
    Dim add_array: ReDim add_array(n_row - istart + 1, max_col): n_row_out = 0
    For i = istart To n_row
        pos = spec(i, col_add_pos): subpos = pos
        obozn = spec(i, col_add_obozn)
        naen = spec(i, col_add_naen)
        qty = spec(i, col_add_qty): If qty = Empty Or qty <= 0 Then qty = 1
        prim = spec(i, col_add_prim)
        If pos <> Empty And naen <> Empty Then
            n_row_out = n_row_out + 1
            add_array(n_row_out, col_marka) = "add"
            add_array(n_row_out, col_sub_pos) = subpos
            add_array(n_row_out, col_type_el) = t_subpos
            add_array(n_row_out, col_pos) = subpos
            add_array(n_row_out, col_m_obozn) = obozn
            add_array(n_row_out, col_m_naen) = naen
            add_array(n_row_out, col_m_weight) = "-"
            add_array(n_row_out, col_qty) = qty
            add_array(n_row_out, col_m_edizm) = prim
        End If
    Next i
    add_array = DataCheck(add_array)
    ReadPos = add_array
    Erase add_array
End Function

Function ReadPrSortament()
    If Not SheetExist("!System") Then ThisWorkbook.Worksheets.Add.Name = "!System"
    Set Sh = Application.ThisWorkbook.Sheets("!System") 'На этом скрытом листе будем хранить данные для списков
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
        If IsEmpty(f_list_prof) Then
            MsgBox ("Ошибка чтения файла " + file)
            Exit Function
        End If
        n_prof = UBound(f_list_prof) + 1
        Sh.Range(Sh.Cells(2, n_col - 1), Sh.Cells(n_prof, n_col - 1)) = ArrayTranspose(f_list_prof)
        tmp_arr(1) = "'!System'!" & Sh.Range(Sh.Cells(3, n_col - 1), Sh.Cells(n_prof, n_col - 1)).Address
        tmp_arr(2) = f_list_gost(n_col)
        tpr_adress.Item(file) = tmp_arr
        type_prof = f_list_sort(n_col, 5)
        For j = 2 To n_prof - 1
            If Not IsEmpty(f_list_weight(j)) Then
                'Определяем периметр
                perim = 0
                Select Case type_prof
                    Case "Круглая труба"
                        Dd = f_prof(j, 5) / 1000
                        perim = 3.141592 * Dd
                    Case "Квадратная труба"
                        hh = f_prof(j, 5) / 1000
                        bb = f_prof(j, 6) / 1000
                        perim = 2 * hh + 2 * bb
                    Case "Швеллер", "Швеллер гнутый"
                        bb = f_prof(j, 6) / 1000
                        hh = f_prof(j, 5) / 1000
                        perim = 2 * hh + 4 * bb
                    Case "Двутавр"
                        hh = f_prof(j, 5) / 1000
                        bb = f_prof(j, 6) / 1000
                        tt = f_prof(j, 8) / 1000
                        perim = 2 * hh + 4 * bb + 4 * tt
                    Case "Уголок", "Уголок гнутый"
                        hh = f_prof(j, 5) / 1000
                        bb = f_prof(j, 6) / 1000
                        tt = f_prof(j, 7) / 1000
                        perim = 2 * hh + 2 * bb + 2 * tt
                    Case "Лист"
                        perim = -1
                End Select
                prof = f_list_prof(j)
                tmp_arr(1) = f_list_weight(j) 'Вес
                tmp_arr(2) = perim 'Периметр
                tmp_arr(3) = f_prof(j, 4) 'Площадь сечения
                tpr_adress.Item(file & prof) = tmp_arr
            End If
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
    Sh.Cells(5, n_start) = "кг."
    Sh.Cells(6, n_start) = "л."
    tpr_adress.Item("Примечания") = "'!System'!" & Sh.Range(Sh.Cells(1, n_start), Sh.Cells(6, n_start)).Address
    
    n_start = n_end + 2
    'Типы окраски
    Sh.Cells(1, n_start) = "-"
    Sh.Cells(2, n_start) = "Тип 1"
    Sh.Cells(3, n_start) = "Тип 2"
    Sh.Cells(4, n_start) = "Тип 3"
    Sh.Cells(5, n_start) = "Тип 4"
    tpr_adress.Item("Окраска") = "'!System'!" & Sh.Range(Sh.Cells(1, n_start), Sh.Cells(5, n_start)).Address
    Set pr_adress = tpr_adress
    ReadPrSortament = True
End Function

Function ReadReinforce() As Boolean
    'Чтение сортамента
    SortamentPath = SetPath()
    nf_sort = SortamentPath & "Сортамент_арматуры.txt"
    If Len(Dir$(nf_sort)) > 0 Then
        reinforcement_specifications = ReadTxt(nf_sort)
    Else
        MsgBox ("Нет файла сортамента арматуры")
        r = LogWrite("Ошибка арматуры", "", "Нет файла сортамента арматуры")
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
    Set fso = CreateObject("scripting.filesystemobject")
    Set ts = fso.OpenTextFile(filename$, 1, True): txt$ = ts.ReadAll: ts.Close
    If def_decode Then UserForm2.decode_CB.Value = True
    If UserForm2.decode_CB.Value = True Then
        SourceCharset$ = "Windows-1251"
        DestCharset$ = "UTF-8"
        With CreateObject("ADODB.Stream")
            .Type = 2: .mode = 3
            .Charset = SourceCharset$
            .Open
            .WriteText txt$
            .Position = 0
            .Charset = DestCharset$
            txt$ = .ReadText
            .Close
        End With
    End If
    Set ts = Nothing: Set fso = Nothing
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

Function RelFName(ByVal fname As String) As String
    n_slash = InStrRev(fname, "\")
    n_len = Len(fname)
    n_dot = 4
    wt_dot = Left(fname, n_len - n_dot)
    n_len = Len(wt_dot)
    wt_path = Right(wt_dot, n_len - n_slash)
    RelFName = wt_path
End Function

Function Round_w(ByVal arg As Variant, ByVal nokrg As Variant, Optional ByVal hard_round As Boolean = False) As Variant
    If IsNumeric(arg) Then
        If arg < (1 / (10 ^ nokrg)) And hard_round = False Then
            nokrg_n = 0
            For i = nokrg To nokrg * 2 + 10
                If arg >= (1 / (10 ^ i)) And nokrg_n = 0 Then nokrg_n = i
            Next i
            nokrg = nokrg_n
        End If
        Select Case type_okrugl
            Case 1
                ost = Round(arg, nokrg) - arg
                dob = 1 / (10 ^ nokrg)
                If ost < 0 Then
                    d1 = Round(arg, nokrg)
                    d2 = Round(arg, nokrg) + dob
                    Round_w = CDbl(Round(arg, nokrg) + dob)
                Else
                    Round_w = CDbl(Round(arg, nokrg))
                End If
            Case 2
                Round_w = CDbl(Round(arg, nokrg))
            Case Else
                Round_w = CDbl(arg)
        End Select
    Else
        If arg = "" Or arg = " " Then
            Round_w = 0
        Else
            Round_w = arg
        End If
    End If
End Function

Function SetPath()
    On Error Resume Next
    form = ThisWorkbook.VBProject.VBComponents("UserForm2").Name
    isFormExistis = 0
    If Not IsEmpty(form) Then isFormExistis = CBool(Len(form))
    If isFormExistis Then
        SortamentPath = UserForm2.SortamentPath
    Else
        SortamentPath = Left$(ThisWorkbook.path, InStrRev(ThisWorkbook.path, "\")) & "sort\"
    End If
    SetPath = SortamentPath
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

Function SheetCheckName(ByVal NameSheet As String) As Boolean
    n_err = 0
    If InStr(NameSheet, "_") <> InStrRev(NameSheet, "_") Then
        MsgBox ("В имени листа может быть только один нижний пробел ('_')")
        n_err = n_err + 1
    End If
    If n_err > 0 Then
        SheetCheckName = False
        Exit Function
    End If
    SheetCheckName = True
End Function

Function SheetClear(type_out)
    n_del = 0
    For Each objWh In ThisWorkbook.Worksheets
        nm = objWh.Name
        If InStr(nm, "архикад") = 0 And Left(nm, 1) <> "|" And Left(nm, 1) <> "!" Then
            type_spec = SpecGetType(nm)
            If type_out(1) = -1 Then
                Select Case type_spec
                    Case 1, 2, 4, 5, 11, 12, 13, 14, 20
                        hh = 1
                        ThisWorkbook.Sheets(nm).Delete
                        n_del = n_del + 1
                        r = LogWrite(nm, "", "DEL")
                End Select
            Else
                For Each tdel In type_out
                    If Not IsEmpty(tdel) And tdel = type_spec Then
                        On Error Resume Next
                        ThisWorkbook.Sheets(nm).Delete
                        n_del = n_del + 1
                        r = LogWrite(nm, "", "DEL")
                    End If
                Next
            End If
        End If
    Next objWh
    SheetClear = n_del
End Function

Function SheetGetSize(ByVal objLst As Variant) As Variant
    Dim out(2)
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
    SheetGetSize = out
    Erase out
End Function

Function SheetHideAll()
    Worksheets(inx_name).Activate
    Dim sheet As Worksheet
    With ThisWorkbook
        For Each sheet In ThisWorkbook.Worksheets
            If Left(sheet.Name, 1) = "!" Then Sheets(sheet.Name).Visible = False
        Next
    End With
End Function

Function SheetImport(ByVal nm As String) As Boolean
    Set importbook = Nothing
    On Error Resume Next
    Set importbook = GetObject(nm)
    On Error GoTo 0
    If Not importbook Is Nothing Then
        listsheet = GetListOfSheet(importbook)
        For Each sheet_name In listsheet
            If SpecGetType(sheet_name) > 0 Then
                If SheetExist(sheet_name) Then
                    For n = 1 To 100
                        sn = Str(n) + " " + sheet_name
                        If Not SheetExist(sn) Then Exit For
                    Next n
                    importbook.Sheets(sheet_name).Name = sn
                    sheet_name = sn
                End If
                importbook.Sheets(sheet_name).Copy Before:=ThisWorkbook.Sheets(1)
            End If
        Next
        importbook.Close SaveChanges:=False
        SheetImport = True
    Else
        SheetImport = False
    End If
End Function

Function SheetActivate(ByVal sheetn As String)
    r = INISet()
    If sheetn = inx_name Then
        r = OutPrepare()
        r = SheetIndex()
        r = OutEnded()
    Else
        type_spec = SpecGetType(sheetn)
        If type_spec = 7 And check_on_active Then
            r = OutPrepare()
            r = ManualCheck(sheetn)
            r = OutEnded()
        End If
        If type_spec > 0 And mem_option Then r = SheetSetOption(sheetn)
    End If
End Function

Function SheetSetOption(ByVal sheetn As String)
    If IsEmpty(sheet_option) Then r = SheetReadOption()
    If IsEmpty(sheet_option.Item(sheetn & ";data")) Then r = SheetReadOption()
    tdate = sheet_option.Item(sheetn & ";data")
    UserForm2.Kzap.Text = sheet_option.Item(sheetn & ";k_zap")
    UserForm2.arm_pm_CB.Value = sheet_option.Item(sheetn & ";arm_pm")
    UserForm2.pr_pm_CB.Value = sheet_option.Item(sheetn & ";pr_pm")
    UserForm2.keep_pos_CB = sheet_option.Item(sheetn & ";keep_pos")
    UserForm2.qtyOneSubpos_CB = sheet_option.Item(sheetn & ";qtyOneSubpos")
    UserForm2.show_subpos_CB = sheet_option.Item(sheetn & ";show_subpos")
    UserForm2.ignore_subpos_CB = sheet_option.Item(sheetn & ";ignore_subpos")
    UserForm2.merge_material_CB = sheet_option.Item(sheetn & ";merge_material")
    UserForm2.otd_by_type_CB = sheet_option.Item(sheetn & ";otd_by_type")
    UserForm2.add_row_CB = sheet_option.Item(sheetn & ";add_row")
    UserForm2.ed_izm_km_CB = sheet_option.Item(sheetn & ";ed_izm_km")
    UserForm2.separate_material_CB = sheet_option.Item(sheetn & ";separate_material")
    UserForm2.show_type_CB = sheet_option.Item(sheetn & ";show_type")
    UserForm2.show_qty_spec = sheet_option.Item(sheetn & ";show_qty_spec")
    UserForm2.decode_CB = sheet_option.Item(sheetn & ";decode")
    If def_decode Then UserForm2.decode_CB.Value = True
    SheetSetOption = True
End Function

Function SheetIndex()
    r = SheetReadOption()
    If SheetExist(inx_name) Then
        ThisWorkbook.Worksheets(inx_name).Activate
    Else
        ThisWorkbook.Worksheets.Add.Name = inx_name
        ThisWorkbook.Worksheets(inx_name).Activate
    End If
    ThisWorkbook.Worksheets(inx_name).Move Before:=ThisWorkbook.Sheets(1)
    Dim sheet As Worksheet
    Dim cell As Range
    Worksheets(inx_name).Cells.Clear
    r = FormatClear()
    Worksheets(inx_name).Cells(1) = "Системные"
    Worksheets(inx_name).Cells(2) = "Скрытые"
    Worksheets(inx_name).Cells(3) = "Тип 1"
    Worksheets(inx_name).Cells(4) = "Тип 2"
    Worksheets(inx_name).Cells(5) = "Тип 3"
    Worksheets(inx_name).Cells(7) = "k_zap"
    Worksheets(inx_name).Cells(8) = "Дата"
    Worksheets(inx_name).Cells(9) = "arm_pm"
    Worksheets(inx_name).Cells(10) = "pr_pm"
    Worksheets(inx_name).Cells(11) = "keep_pos"
    Worksheets(inx_name).Cells(12) = "qtyOneSubpos"
    Worksheets(inx_name).Cells(13) = "show_subpos"
    Worksheets(inx_name).Cells(14) = "ignore_subpos"
    Worksheets(inx_name).Cells(15) = "merge_material"
    Worksheets(inx_name).Cells(16) = "otd_by_type"
    Worksheets(inx_name).Cells(17) = "add_row"
    Worksheets(inx_name).Cells(18) = "ed_izm_km"
    Worksheets(inx_name).Cells(19) = "separate_material"
    Worksheets(inx_name).Cells(20) = "show_type"
    Worksheets(inx_name).Cells(21) = "show_qty_spec"
    Worksheets(inx_name).Cells(22) = "decode"
    Dim sheetnames(): j = 0
    With ThisWorkbook
        For Each sheet In ThisWorkbook.Worksheets
            j = j + 1
            ReDim Preserve sheetnames(j)
            sheetnames(j) = sheet.Name
        Next
    End With
    sheetnames = ArraySort(sheetnames)
    For j = 1 To UBound(sheetnames)
        sheetn = sheetnames(j)
        tspec = SpecGetType(sheetn)
        Select Case tspec
            Case 1, 2, 3, 4, 5, 13, 20, 21
                With ThisWorkbook.Sheets(sheetn).Tab
                    .Color = 0
                    .TintAndShade = 0
                End With
            Case 6, 7, 9, 10, 11, 12, 14, 8
                ThisWorkbook.Worksheets(sheetn).Move After:=ThisWorkbook.Sheets(2)
                With ThisWorkbook.Sheets(sheetn).Tab
                    .ThemeColor = xlThemeColorAccent4
                    .TintAndShade = 0.4
                End With
            Case 15
                ThisWorkbook.Worksheets(sheetn).Move After:=ThisWorkbook.Sheets(2)
                With ThisWorkbook.Sheets(sheetn).Tab
                    .ThemeColor = xlThemeColorAccent5
                    .TintAndShade = 0.5
                End With
            Case Else
                With ThisWorkbook.Sheets(sheetn).Tab
                    .Color = 0
                    .TintAndShade = 1
                End With
        End Select
        If sheetn = inx_name Then
            With ThisWorkbook.Sheets(sheetn).Tab
                .Color = 5287936
                .TintAndShade = 0
            End With
        End If
    Next
    For j = 2 To UBound(sheetnames) + 1
        sheetn = sheetnames(j - 1)
        If Sheets(sheetn).Visible = 0 And Left(sheetn, 1) <> "|" And Left(sheetn, 1) <> "!" Then sheetn = "!" & sheetn
        If Left(sheetn, 1) = "|" Then
            Set cell = Worksheets(inx_name).Cells(j, 1)
            ThisWorkbook.Worksheets(inx_name).Hyperlinks.Add anchor:=cell, Address:="", SubAddress:="'" & sheetn & "'" & "!A1"
            cell.Formula = sheetn
            Sheets(sheetn).Visible = True
        Else
            If Left(sheetn, 1) = "!" Then
                Set cell = Worksheets(inx_name).Cells(j, 2)
                ThisWorkbook.Worksheets(inx_name).Hyperlinks.Add anchor:=cell, Address:="", SubAddress:="'" & sheetn & "'" & "!B2"
                cell.Formula = sheetn
                Sheets(sheetn).Visible = False
            Else
                tspec = SpecGetType(sheetn)
                Select Case tspec
                Case 6, 9, 11, 12, 14, 15, 8
                    Set cell = Worksheets(inx_name).Cells(j, 3)
                    ThisWorkbook.Worksheets(inx_name).Hyperlinks.Add anchor:=cell, Address:="", SubAddress:="'" & sheetn & "'" & "!C3"
                Case 7, 10
                    Set cell = Worksheets(inx_name).Cells(j, 5)
                    ThisWorkbook.Worksheets(inx_name).Hyperlinks.Add anchor:=cell, Address:="", SubAddress:="'" & sheetn & "'" & "!E3"
                Case 1, 2, 3, 4, 5, 13, 0, 20, 21
                    Set cell = Worksheets(inx_name).Cells(j, 4)
                    ThisWorkbook.Worksheets(inx_name).Hyperlinks.Add anchor:=cell, Address:="", SubAddress:="'" & sheetn & "'" & "!D4"
                End Select
                cell.Formula = sheetn
                If IsEmpty(sheet_option.Item(sheetn & ";k_zap")) Then
                    hh = sheet_option.Item(sheetn & ";k_zap")
                Else
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 7) = sheet_option.Item(sheetn & ";k_zap")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 8) = sheet_option.Item(sheetn & ";data")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 9) = sheet_option.Item(sheetn & ";arm_pm")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 10) = sheet_option.Item(sheetn & ";pr_pm")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 11) = sheet_option.Item(sheetn & ";keep_pos")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 12) = sheet_option.Item(sheetn & ";qtyOneSubpos")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 13) = sheet_option.Item(sheetn & ";show_subpos")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 14) = sheet_option.Item(sheetn & ";ignore_subpos")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 15) = sheet_option.Item(sheetn & ";merge_material")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 16) = sheet_option.Item(sheetn & ";otd_by_type")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 17) = sheet_option.Item(sheetn & ";add_row")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 18) = sheet_option.Item(sheetn & ";ed_izm_km")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 19) = sheet_option.Item(sheetn & ";separate_material")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 20) = sheet_option.Item(sheetn & ";show_type")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 21) = sheet_option.Item(sheetn & ";show_qty_spec")
                    ThisWorkbook.Worksheets(inx_name).Cells(j, 22) = sheet_option.Item(sheetn & ";decode")
                End If
                Sheets(sheetn).Visible = True
            End If
        End If
        Sheets(inx_name).Visible = True
    Next
    ThisWorkbook.Worksheets(inx_name).Activate
    ThisWorkbook.Worksheets(inx_name).Sort.SortFields.Clear
    ThisWorkbook.Worksheets(inx_name).Sort.SortFields.Add Key:=Range( _
        "H2:H600"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ThisWorkbook.Worksheets(inx_name).Sort
        .SetRange Range("A1:V600")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Rows("1:1").Font.Bold = True
    With Rows("1:1").Font
        .Name = "Calibri"
        .Size = 12
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
    Rows("1:1").RowHeight = 28
    Rows("2:600").RowHeight = 15
    Range("A2:N600").Rows.AutoFit
    Columns("A:V").Select
    Columns("A:E").ColumnWidth = 36
    Columns("G").ColumnWidth = 8
    Columns("H").ColumnWidth = 15
    Columns("I:V").ColumnWidth = 10
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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
    Selection.FormatConditions.Add Type:=xlTextString, String:="_спец", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399945066682943
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="!", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=ДЛСТР(СЖПРОБЕЛЫ(A1))=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.349986266670736
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="|", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="_кж", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="_км", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="_поз", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("A1").Select
End Function

Function SheetNew(ByVal NameSheet As String)
    On Error Resume Next
    If SheetExist(NameSheet) Then
        Worksheets(NameSheet).Cells.Clear
    Else
        ThisWorkbook.Worksheets.Add.Name = NameSheet
    End If
End Function

Function SheetShowAddictions()
    r = OutPrepare()
    For Each rn In Range("E1:H500")
        rn.ShowPrecedents
    Next
    r = OutEnded()
End Function

Function SheetShowAll()
    Worksheets(inx_name).Activate
    Dim sheet As Worksheet
    With ThisWorkbook
        For Each sheet In ThisWorkbook.Worksheets
            Sheets(sheet.Name).Visible = True
        Next
    End With
End Function

Function SheetReadOption()
    If IsEmpty(sheet_option) Then
        Set sheet_option = CreateObject("Scripting.Dictionary")
        sheet_option.comparemode = 1
    End If
    If SheetExist(inx_name) Then
        sheet_size = SheetGetSize(ThisWorkbook.Worksheets(inx_name))
        n_row = sheet_size(1)
        n_col = sheet_size(2)
        If n_row > 1 Then
            existssheet = ThisWorkbook.Worksheets(inx_name).Range(ThisWorkbook.Worksheets(inx_name).Cells(2, 1), ThisWorkbook.Worksheets(inx_name).Cells(n_row, n_col))
            For i = 1 To UBound(existssheet, 1)
                If IsEmpty(existssheet(i, 1)) Then
                    For j = 2 To 5
                        If Not IsEmpty(existssheet(i, j)) Then existssheet(i, 1) = existssheet(i, j)
                    Next j
                End If
                sheetn = existssheet(i, 1)
                If SpecGetType(sheetn) > 0 Then
                    If IsEmpty(existssheet(i, 7)) Then
                        r = SheetDefultOption(sheetn)
                    Else
                        sheet_option.Item(sheetn & ";k_zap") = existssheet(i, 7)
                        sheet_option.Item(sheetn & ";data") = existssheet(i, 8)
                        sheet_option.Item(sheetn & ";arm_pm") = existssheet(i, 9)
                        sheet_option.Item(sheetn & ";pr_pm") = existssheet(i, 10)
                        sheet_option.Item(sheetn & ";keep_pos") = existssheet(i, 11)
                        sheet_option.Item(sheetn & ";qtyOneSubpos") = existssheet(i, 12)
                        sheet_option.Item(sheetn & ";show_subpos") = existssheet(i, 13)
                        sheet_option.Item(sheetn & ";ignore_subpos") = existssheet(i, 14)
                        sheet_option.Item(sheetn & ";merge_material") = existssheet(i, 15)
                        sheet_option.Item(sheetn & ";otd_by_type") = existssheet(i, 16)
                        sheet_option.Item(sheetn & ";add_row") = existssheet(i, 17)
                        sheet_option.Item(sheetn & ";ed_izm_km") = existssheet(i, 18)
                        sheet_option.Item(sheetn & ";separate_material") = existssheet(i, 19)
                        sheet_option.Item(sheetn & ";show_type") = existssheet(i, 20)
                        sheet_option.Item(sheetn & ";show_qty_spec") = existssheet(i, 21)
                        sheet_option.Item(sheetn & ";decode") = existssheet(i, 22)
                    End If
                End If
            Next i
        End If
    End If
    With ThisWorkbook
        For Each sheet In ThisWorkbook.Worksheets
            If Not sheet_option.Exists(sheet.Name & ";k_zap") Then
                If SpecGetType(sheet.Name) > 0 Then r = SheetDefultOption(sheet.Name)
            End If
        Next
    End With
    
End Function

Function SheetDefultOption(ByVal sheetn As String)
    sheet_option.Item(sheetn & ";data") = "---"
    sheet_option.Item(sheetn & ";k_zap") = "1.0"
    sheet_option.Item(sheetn & ";arm_pm") = False
    sheet_option.Item(sheetn & ";pr_pm") = False
    sheet_option.Item(sheetn & ";keep_pos") = False
    sheet_option.Item(sheetn & ";qtyOneSubpos") = False
    sheet_option.Item(sheetn & ";show_subpos") = True
    sheet_option.Item(sheetn & ";ignore_subpos") = False
    sheet_option.Item(sheetn & ";merge_material") = True
    sheet_option.Item(sheetn & ";otd_by_type") = True
    sheet_option.Item(sheetn & ";add_row") = False
    sheet_option.Item(sheetn & ";ed_izm_km") = False
    sheet_option.Item(sheetn & ";separate_material") = True
    sheet_option.Item(sheetn & ";show_type") = False
    sheet_option.Item(sheetn & ";show_qty_spec") = False
    sheet_option.Item(sheetn & ";decode") = False
    SheetDefultOption = True
End Function

Function SheetWriteOption(ByVal sheetn As String)
    If IsEmpty(sheet_option) Then r = SheetReadOption()
    If IsEmpty(sheet_option.Item(sheetn & ";data")) Then r = SheetReadOption()
    tdate = Right(Str(DatePart("yyyy", Now)), 2) & Str(DatePart("m", Now)) & Str(DatePart("d", Now))
    stamp = tdate + "/" + Str(DatePart("h", Now)) + Str(DatePart("n", Now)) + Str(DatePart("s", Now))
    sheet_option.Item(sheetn & ";k_zap") = UserForm2.Kzap.Text
    sheet_option.Item(sheetn & ";data") = stamp
    sheet_option.Item(sheetn & ";arm_pm") = UserForm2.arm_pm_CB.Value
    sheet_option.Item(sheetn & ";pr_pm") = UserForm2.pr_pm_CB.Value
    sheet_option.Item(sheetn & ";keep_pos") = UserForm2.keep_pos_CB
    sheet_option.Item(sheetn & ";qtyOneSubpos") = UserForm2.qtyOneSubpos_CB
    sheet_option.Item(sheetn & ";show_subpos") = UserForm2.show_subpos_CB
    sheet_option.Item(sheetn & ";ignore_subpos") = UserForm2.ignore_subpos_CB
    sheet_option.Item(sheetn & ";merge_material") = UserForm2.merge_material_CB
    sheet_option.Item(sheetn & ";otd_by_type") = UserForm2.otd_by_type_CB
    sheet_option.Item(sheetn & ";add_row") = UserForm2.add_row_CB
    sheet_option.Item(sheetn & ";ed_izm_km") = UserForm2.ed_izm_km_CB
    sheet_option.Item(sheetn & ";separate_material") = UserForm2.separate_material_CB
    sheet_option.Item(sheetn & ";show_type") = UserForm2.show_type_CB
    sheet_option.Item(sheetn & ";show_qty_spec") = UserForm2.show_qty_spec
    sheet_option.Item(sheetn & ";decode") = UserForm2.decode_CB
    SheetWriteOption = True
End Function

Function SheetCopyOption(ByVal sheetn As String, ByVal sheetnto As String)
    If IsEmpty(sheet_option.Item(sheetn & ";data")) Then r = SheetReadOption()
    tdate = Right(Str(DatePart("yyyy", Now)), 2) & Str(DatePart("m", Now)) & Str(DatePart("d", Now))
    stamp = tdate + "/" + Str(DatePart("h", Now)) + Str(DatePart("n", Now)) + Str(DatePart("s", Now))
    sheet_option.Item(sheetn & ";k_zap") = UserForm2.Kzap.Text
    sheet_option.Item(sheetn & ";data") = stamp
    sheet_option.Item(sheetn & ";arm_pm") = UserForm2.arm_pm_CB.Value
    sheet_option.Item(sheetn & ";pr_pm") = UserForm2.pr_pm_CB.Value
    sheet_option.Item(sheetn & ";keep_pos") = UserForm2.keep_pos_CB
    sheet_option.Item(sheetn & ";qtyOneSubpos") = UserForm2.qtyOneSubpos_CB
    sheet_option.Item(sheetn & ";show_subpos") = UserForm2.show_subpos_CB
    sheet_option.Item(sheetn & ";ignore_subpos") = UserForm2.ignore_subpos_CB
    sheet_option.Item(sheetn & ";merge_material") = UserForm2.merge_material_CB
    sheet_option.Item(sheetn & ";otd_by_type") = UserForm2.otd_by_type_CB
    sheet_option.Item(sheetn & ";add_row") = UserForm2.add_row_CB
    sheet_option.Item(sheetn & ";ed_izm_km") = UserForm2.ed_izm_km_CB
    sheet_option.Item(sheetn & ";separate_material") = UserForm2.separate_material_CB
    sheet_option.Item(sheetn & ";show_type") = UserForm2.show_type_CB
    sheet_option.Item(sheetn & ";show_qty_spec") = UserForm2.show_qty_spec
    sheet_option.Item(sheetn & ";decode") = UserForm2.decode_CB
    SheetWriteOption = True
End Function

Function SpecArm(ByVal arm As Variant, ByVal n_arm As Integer, ByVal type_spec As Integer, ByVal nSubPos As Integer) As Variant
    Dim pos_out
    If UserForm2.qtyOneSubpos_CB.Value Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
    Else
        n_txt = "," & vbLf & "на все"
    End If
    If UserForm2.show_qty_spec.Value Then n_txt = "" & ",**"
    un_chsum_arm = ArrayUniqValColumn(arm, col_chksum)
    pos_chsum_arm = UBound(un_chsum_arm, 1)
    If type_spec = 1 Or UserForm2.arm_pm_CB.Value Or (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then
        'Нам нужны уникальные суммы только для диаметра и класса
        'Поэтому сформируем новый массив, где от архикадовской суммы отрежем лишнее
        For i = 1 To pos_chsum_arm
            If UserForm2.arm_pm_CB.Value Then
                If UserForm2.keep_pos_CB.Value Then
                    un_chsum_arm(i) = Split(un_chsum_arm(i), "_")(0) & Split(un_chsum_arm(i), "_")(2)
                Else
                    un_chsum_arm(i) = Split(un_chsum_arm(i), "_")(0)
                End If
            Else
                un_chsum_arm(i) = Split(un_chsum_arm(i), "_")(0) & Split(un_chsum_arm(i), "_")(2)
            End If
        Next i
        un_chsum_arm = ArrayUniqValColumn(un_chsum_arm, 1)
        pos_chsum_arm = UBound(un_chsum_arm, 1)
    End If
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    ReDim pos_out(pos_chsum_arm, n_col_spec)
    For i = 1 To pos_chsum_arm
        For j = 1 To n_arm
            If type_spec = 1 Or UserForm2.arm_pm_CB.Value Or (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then
                If type_spec = 1 Or (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then current_chksum = Split(arm(j, col_chksum), "_")(0) & Split(arm(j, col_chksum), "_")(2)
                If UserForm2.arm_pm_CB.Value And Not UserForm2.keep_pos_CB.Value Then current_chksum = Split(arm(j, col_chksum), "_")(0)
                If UserForm2.arm_pm_CB.Value And UserForm2.keep_pos_CB.Value Then current_chksum = Split(arm(j, col_chksum), "_")(0) & Split(arm(j, col_chksum), "_")(2)
            Else
                current_chksum = arm(j, col_chksum)
            End If
            chksum = un_chsum_arm(i)
            If current_chksum = chksum Then
                klass = arm(j, col_klass)
                diametr = arm(j, col_diametr)
                weight_pm = GetWeightForDiametr(diametr, klass)
                fon = arm(j, col_fon)
                If UserForm2.arm_pm_CB.Value Then fon = 1
                mp = arm(j, col_mp)
                gnut = arm(j, col_gnut)
                prim = "": If arm(j, col_gnut) And Not UserForm2.arm_pm_CB.Value Then prim = "*"
                qty = arm(j, col_qty)
                n_el = qty / nSubPos
                If k_zap_total > 1 Then n_el = (qty / nSubPos) + Round((k_zap_total - 1) * (qty / nSubPos), 0)
                length_pos = arm(j, col_length) / 1000
                Select Case type_spec
                Case 1
                    pos_out(i, 1) = arm(j, col_sub_pos) & n_txt
                    If (UserForm2.keep_pos_CB.Value And UserForm2.arm_pm_CB.Value) Or Not (UserForm2.arm_pm_CB.Value) Then
                        pos_out(i, 2) = arm(j, col_pos)
                    Else
                        pos_out(i, 2) = " "
                    End If
                    If fon Then
                        l_pos = Round_w(length_pos * k_zap_total, n_round_l) * n_el
                        If prim = "п.м." Then prim = ""
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L= п.м." & prim
                        pos_out(i, 4) = pos_out(i, 4) + l_pos
                        pos_out(i, 5) = Round_w(weight_pm, n_round_w)
                    Else
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L=" & length_pos * 1000 & "мм." & prim
                        pos_out(i, 4) = pos_out(i, 4) + n_el
                        pos_out(i, 5) = Round_w(weight_pm * length_pos * k_zap_total, n_round_w)
                    End If
                Case Else
                    If (UserForm2.keep_pos_CB.Value And UserForm2.arm_pm_CB.Value) Or Not (UserForm2.arm_pm_CB.Value) Then
                        pos_out(i, 1) = arm(j, col_pos)
                    Else
                        pos_out(i, 1) = " "
                    End If
                    pos_out(i, 2) = GetGOSTForKlass(klass)
                    If fon Then
                        l_pos = Round_w(length_pos * k_zap_total, n_round_l) * n_el
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L= п.м." & prim
                        pos_out(i, 4) = pos_out(i, 4) + l_pos
                        pos_out(i, 5) = weight_pm
                        If show_sum_prim Then pos_out(i, 6) = pos_out(i, 6) + Round_w((l_pos * weight_pm), 2)
                    Else
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L=" & length_pos * 1000 & "мм." & prim
                        pos_out(i, 4) = pos_out(i, 4) + n_el
                        pos_out(i, 5) = Round_w(weight_pm * length_pos * k_zap_total, n_round_w)
                        If show_sum_prim Then pos_out(i, 6) = pos_out(i, 6) + n_el * pos_out(i, 5)
                    End If
                End Select
            End If
        Next j
    Next i
    
    For i = 1 To UBound(pos_out, 1)
        If type_spec = 13 Then pos_out(i, 7) = t_arm
    Next
    If type_spec = 1 Then
        n_col_pos = 2
    Else
        n_col_pos = 1
    End If
    pos_out = ArraySort(pos_out, n_col_pos)
    SpecArm = pos_out
    Erase arm, pos_out
End Function

Function SpecGetType(ByVal nm As String) As Integer
    On Error Resume Next
    form = ThisWorkbook.VBProject.VBComponents("UserForm2").Name
    If IsEmpty(form) Then
        SpecGetType = 7
        Exit Function
    End If
    If Left(nm, 1) <> "!" And Left(nm, 1) <> "|" Then
        If InStr(nm, "_") > 0 Then
            type_spec = Split(nm, "_")
            Select Case type_spec(UBound(type_spec))
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
                Case "грс"
                    spec = 13
                Case "норм"
                    spec = 14
                Case "спец.И"
                    spec = 15
                Case "зап"
                    spec = 20
                Case "разб"
                    spec = 21
                Case Else
                    spec = 2
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

Function SpecIzd(ByVal izd As Variant, ByVal n_izd As Integer, ByVal type_spec As Integer, ByVal nSubPos As Integer) As Variant
    
    If UserForm2.qtyOneSubpos_CB.Value Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
    Else
        n_txt = "," & vbLf & "на все"
        nSubPos = 1
    End If
    If UserForm2.show_qty_spec.Value Then n_txt = "" & ",**"

    un_chsum_izd = ArrayUniqValColumn(izd, col_chksum)
    pos_chsum_izd = UBound(un_chsum_izd, 1)
    If type_spec = 1 Or (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then
        For i = 1 To pos_chsum_izd
            un_chsum_izd(i) = Split(un_chsum_izd(i), "_")(0) & Split(un_chsum_izd(i), "_")(2)
        Next i
        un_chsum_izd = ArrayUniqValColumn(un_chsum_izd, 1)
        pos_chsum_izd = UBound(un_chsum_izd, 1)
    End If
    
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    Dim pos_out: ReDim pos_out(pos_chsum_izd, n_col_spec)
    For i = 1 To pos_chsum_izd
        For j = 1 To n_izd
            If type_spec = 1 Or (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then
                current_chksum = Split(izd(j, col_chksum), "_")(0) & Split(izd(j, col_chksum), "_")(2)
            Else
                current_chksum = izd(j, col_chksum)
            End If
            If current_chksum = un_chsum_izd(i) Then
                qty = izd(j, col_qty)
                n_el = qty / nSubPos
                If k_zap_total > 1 Then n_el = (qty / nSubPos) + Round((k_zap_total - 1) * (qty / nSubPos), 0)
                subpos = izd(j, col_sub_pos)
                pos = izd(j, col_pos)
                
                If IsNumeric(izd(j, col_m_weight)) Then
                    Weight = Round_w(izd(j, col_m_weight), n_round_w)
                Else
                    Weight = "-"
                End If
                
                Select Case type_spec
                Case 1
                    pos_out(i, 1) = subpos & n_txt
                    pos_out(i, 2) = pos
                    pos_out(i, 4) = pos_out(i, 4) + n_el
                    If izd(j, col_m_edizm) <> "п.м." Then
                        pos_out(i, 3) = izd(j, col_m_naen) & " по " & izd(j, col_m_obozn)
                    Else
                        pos_out(i, 3) = izd(j, col_m_naen) & " по " & izd(j, col_m_obozn) & ", п.м."
                    End If
                    pos_out(i, 5) = Weight
                Case Else
                    obozn = izd(j, col_m_obozn)
                    naen = izd(j, col_m_naen)
                    pos_out(i, 1) = pos
                    pos_out(i, 2) = obozn
                    If InStr(izd(j, col_m_edizm), "#дерев") > 0 Then
                        Weight = ConvTxt2Num(Replace(izd(j, col_m_edizm), "п.м.#дерев", "")) 'теперь тут площадь сечения
                        pos_out(i, 1) = pos_out(i, 1) + "#дерев"
                        pos_out(i, 3) = naen + ", L=п.м."
                        pos_out(i, 4) = pos_out(i, 4) + n_el
                        pos_out(i, 5) = "-"
                        If IsNumeric(Weight) Then pos_out(i, 6) = Weight / 1000
                    Else
                        pos_out(i, 3) = naen
                        pos_out(i, 4) = pos_out(i, 4) + n_el
                        pos_out(i, 5) = Weight
                        If IsNumeric(Weight) And izd(j, col_m_edizm) = "" Then
                            If show_sum_prim Then pos_out(i, 6) = pos_out(i, 6) + (n_el * Weight)
                        Else
                            pos_out(i, 6) = izd(j, col_m_edizm)
                        End If
                    End If
                End Select
            End If
        Next j
    Next i
    For i = 1 To UBound(pos_out, 1)
        If type_spec = 13 Then pos_out(i, 7) = t_izd
        If type_spec <> 1 Then
            If InStr(pos_out(i, 1), "#дерев") > 0 Then
                pos_out(i, 1) = Replace(pos_out(i, 1), "#дерев", "")
                pos_out(i, 4) = Round_w(pos_out(i, 4), 0)
                pos_out(i, 6) = ConvNum2Txt(Round_w((pos_out(i, 6) * pos_out(i, 4)), 3)) + " куб.м." 'площадь сечения на длину
            End If
        End If
    Next
    If type_spec = 1 Then
        n_col_pos = 2
    Else
        n_col_pos = 1
    End If
    pos_out = ArraySort(pos_out, n_col_pos)
    SpecIzd = pos_out
    Erase izd, pos_out
End Function

Function SpecMaterial(ByVal mat As Variant, ByVal n_mat As Integer, ByVal type_spec As Integer, ByVal nSubPos As Integer) As Variant
   
    If UserForm2.qtyOneSubpos_CB.Value Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
    Else
        n_txt = "," & vbLf & "на все"
        nSubPos = 1
    End If
    If UserForm2.show_qty_spec.Value Then n_txt = ",**"
    un_pos_mat = ArrayUniqValColumn(mat, col_chksum)
    pos_mat = UBound(un_pos_mat, 1)
    
    If type_spec = 1 Or (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then
        For i = 1 To pos_mat
            un_pos_mat(i) = Split(un_pos_mat(i), "_")(0) & Split(un_pos_mat(i), "_")(2)
        Next i
        un_pos_mat = ArrayUniqValColumn(un_pos_mat, 1)
        pos_mat = UBound(un_pos_mat, 1)
    End If
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    Dim pos_out: ReDim pos_out(pos_mat, n_col_spec)
    For i = 1 To pos_mat
        For j = 1 To n_mat
            If type_spec = 1 Or (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then
                current_chksum = Split(mat(j, col_chksum), "_")(0) & Split(mat(j, col_chksum), "_")(2)
            Else
                current_chksum = mat(j, col_chksum)
            End If
            If current_chksum = un_pos_mat(i) Then
                Select Case type_spec
                Case 1
                    pos_out(i, 1) = mat(j, col_sub_pos) & n_txt
                    pos_out(i, 2) = " "
                    pos_out(i, 3) = mat(j, col_m_naen) & " по " & mat(j, col_m_obozn) & ", " & mat(j, col_m_edizm)
                    pos_out(i, 4) = pos_out(i, 4) + (mat(j, col_qty) * k_zap_total / nSubPos)
                    pos_out(i, 5) = "-"
                Case Else
                    pos_out(i, 1) = " "
                    pos_out(i, 2) = mat(j, col_m_obozn)
                    pos_out(i, 3) = mat(j, col_m_naen)
                    pos_out(i, 4) = pos_out(i, 4) + (mat(j, col_qty) * k_zap_total / nSubPos)
                    pos_out(i, 5) = "-"
                    pos_out(i, 6) = mat(j, col_m_edizm)
                End Select
            End If
        Next j
    Next i
    
    
    For i = 1 To UBound(pos_out, 1)
        If type_spec = 13 Then pos_out(i, 7) = t_mat
    Next
    If type_spec = 1 Then
        n_col_pos = 2
    Else
        n_col_pos = 1
    End If
    pos_out = ArraySort(pos_out, n_col_pos)
    SpecMaterial = pos_out
    Erase mat, un_pos_mat, pos_out
End Function

Function SpecOneSubpos(ByVal all_data As Variant, ByVal subpos As String, ByVal type_spec As Integer) As Variant
    nSubPos = GetNSubpos(subpos, type_spec)
    If Not UserForm2.qtyOneSubpos_CB.Value Then nSubPos = 1
    If (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then nSubPos = 1
    'Добавляем загаловок для сборки
    Dim pos_naen
    If UserForm2.add_row_CB.Value Then
        n_n = 2
    Else
        n_n = 1
    End If
    sb_naen = "@"
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    If type_spec = 2 Then
        ReDim pos_naen(n_n, n_col_spec)
        If subpos <> "-" Then
            naen = subpos
            If pos_data.Item("name").Count Then
                If pos_data.Item("name").Exists(subpos) Then naen = pos_data.Item("name").Item(subpos)(1)
                If UserForm2.qtyOneSubpos_CB.Value Then
                    pos_naen(n_n, 1) = naen & ", на 1 шт. (всего " & nSubPos & " шт.)"
                Else
                    pos_naen(n_n, 1) = naen & ", на все"
                End If
                If UserForm2.show_qty_spec.Value Then pos_naen(n_n, 1) = naen & ",**"
            End If
        Else
            pos_naen(n_n, 1) = " Прочие элементы "
        End If
        sb_naen = pos_naen(n_n, 1)
        pos_out = ArrayCombine(pos_out, pos_naen)
    End If
    
    n_row = UBound(all_data, 1)
    
    Dim arm(): ReDim arm(n_row, max_col)
    Dim prokat(): ReDim prokat(n_row, max_col)
    Dim mat(): ReDim mat(n_row, max_col)
    Dim izd(): ReDim izd(n_row, max_col)
    Dim subp(): ReDim subp(n_row, max_col)
    n_arm = 0: n_prokat = 0: n_mat = 0: n_izd = 0: n_subpos = 0
    For i = 1 To n_row
        сurrent_subpos = all_data(i, col_sub_pos)
        сurrent_parent = all_data(i, col_parent)
        сurrent_type_el = all_data(i, col_type_el)
        
        Select Case type_spec
            'Групповая спецификация
            'Выбираем только элементы из un_child, т.е. все второстепенные сборки для данной главной сборки
            Case 1
                usl = (сurrent_subpos = subpos) And (сurrent_type_el <> t_subpos)
            'Общая спецификация
            'Расписать каждую сборку и все элементы без сборки
            Case 2, 13
                If subpos = "-" Then
                    u1 = (сurrent_subpos = "-") 'Элементы вне сборок
                    u2 = (pos_data.Item("-").Exists(сurrent_subpos) And (сurrent_parent = "-") And (сurrent_type_el = t_subpos))   'Элементы вложенных сборок
                    usl = u1 Or u2
                Else
                    u1 = ((сurrent_parent = "-") And (сurrent_subpos = subpos) And (сurrent_type_el <> t_subpos)) 'Элементы главной сборки
                    u2 = (сurrent_parent = subpos) And (сurrent_type_el = t_subpos) 'Маркировка вложенных сборок
                    usl = (u1 Or u2)
                End If
            'Общестроительная
            'Только наименование сборки и все элементы без сборок
            Case 3
                If UserForm2.ignore_subpos_CB.Value Then
                    usl = (сurrent_type_el <> t_subpos)
                Else
                    u1 = (сurrent_subpos = "-")
                    u2 = ((сurrent_parent = "-") And (сurrent_type_el = t_subpos) And UserForm2.show_subpos_CB.Value)
                    usl = (u1 Or u2)
                End If
        End Select
        If usl Then
            Select Case сurrent_type_el
                Case t_arm
                    n_arm = n_arm + 1
                    For j = 1 To max_col
                        arm(n_arm, j) = all_data(i, j)
                    Next j
                Case t_prokat
                    n_prokat = n_prokat + 1
                    For j = 1 To max_col
                        prokat(n_prokat, j) = all_data(i, j)
                    Next j
                Case t_mat
                    n_mat = n_mat + 1
                    For j = 1 To max_col
                        mat(n_mat, j) = all_data(i, j)
                    Next j
                Case t_izd
                    n_izd = n_izd + 1
                    For j = 1 To max_col
                        izd(n_izd, j) = all_data(i, j)
                    Next j
                    If izd(n_izd, col_m_weight) = "-" Then izd(n_izd, col_m_weight) = 0
                Case t_subpos
                        n_subpos = n_subpos + 1
                        For j = 1 To max_col
                            subp(n_subpos, j) = all_data(i, j)
                        Next j
                        If pos_data.Item("weight").Item(сurrent_subpos) > 0 Then
                            subp(n_subpos, col_m_weight) = pos_data.Item("weight").Item(сurrent_subpos)
                        End If
                End Select
        End If
    Next
    
    ReDim pos_naen(n_n, n_col_spec)
    If n_subpos > 0 Then
        'subp = ArrayRedim(subp, n_subpos)
        pos_naen(n_n, 3) = " Сборочные единицы "
        If type_spec = 13 Then
            pos_naen(n_n, 7) = t_subpos
            pos_naen(1, 7) = pos_naen(n_n, 7)
        End If
        If type_spec <> 1 And UserForm2.show_type_CB.Value Then pos_out = ArrayCombine(pos_out, pos_naen)
        g = SpecSubpos(subp, n_subpos, type_spec, nSubPos)
        pos_out = ArrayCombine(pos_out, g)
    End If

    If n_izd > 0 Then
        'izd = ArrayRedim(izd, n_izd)
        pos_naen(n_n, 3) = " Изделия "
        If type_spec = 13 Then
            pos_naen(n_n, 7) = t_izd
            pos_naen(1, 7) = pos_naen(n_n, 7)
        End If
        If type_spec <> 1 And UserForm2.show_type_CB.Value Then pos_out = ArrayCombine(pos_out, pos_naen)
        pos_out = ArrayCombine(pos_out, SpecIzd(izd, n_izd, type_spec, nSubPos))
    End If
    
    If n_prokat > 0 Then
        'prokat = ArrayRedim(prokat, n_prokat)
        pos_naen(n_n, 3) = " Прокат "
        If type_spec = 13 Then
            pos_naen(n_n, 7) = t_prokat
            pos_naen(1, 7) = pos_naen(n_n, 7)
        End If
        If type_spec <> 1 And UserForm2.show_type_CB.Value Then pos_out = ArrayCombine(pos_out, pos_naen)
        pos_out = ArrayCombine(pos_out, SpecProkat(prokat, n_prokat, type_spec, nSubPos))
    End If
    
    If n_arm > 0 Then
        'arm = ArrayRedim(arm, n_arm)
        pos_naen(n_n, 3) = " Изделия арматурные "
        If type_spec = 13 Then
            pos_naen(n_n, 7) = t_arm
            pos_naen(1, 7) = pos_naen(n_n, 7)
        End If
        If type_spec <> 1 And UserForm2.show_type_CB.Value Then pos_out = ArrayCombine(pos_out, pos_naen)
        pos_out = ArrayCombine(pos_out, SpecArm(arm, n_arm, type_spec, nSubPos))
    End If
    
    If n_mat > 0 Then
        'mat = ArrayRedim(mat, n_mat)
        pos_naen(n_n, 3) = " Материалы "
        If type_spec = 13 Then
            pos_naen(n_n, 7) = t_mat
            pos_naen(1, 7) = pos_naen(n_n, 7)
        End If
        If type_spec <> 1 And UserForm2.show_type_CB.Value Then pos_out = ArrayCombine(pos_out, pos_naen)
        pos_out = ArrayCombine(pos_out, SpecMaterial(mat, n_mat, type_spec, nSubPos))
    End If
    
    If IsEmpty(pos_out) Then
        SpecOneSubpos = pos_out
    Else
        Select Case type_spec
            Case 1
                If Not UserForm2.show_type_CB.Value Then
                    pos_out = ArraySort(pos_out, 2)
                End If
                For i = 1 To UBound(pos_out, 1)
                    If IsNumeric(pos_out(i, 4)) Then
                        k = pos_out(i, 4)
                    Else
                        k = 1
                    End If
                    If IsNumeric(pos_out(i, 5)) Then
                        m = pos_out(i, 5)
                    Else
                        m = 0
                    End If
                    pos_out(1, 6) = pos_out(1, 6) + k * m
                Next i
                If UserForm2.qtyOneSubpos_CB.Value Then
                    subpos_we_group = pos_out(1, 6) / nSubPos
                    subpos_we_spec = pos_data.Item("weight").Item(subpos) / nSubPos
                Else
                    subpos_we_group = pos_out(1, 6) / nSubPos
                    subpos_we_spec = pos_data.Item("weight").Item(subpos) * GetNSubpos(subpos, type_spec)
                End If
                If Abs(subpos_we_group - subpos_we_spec) > 0.01 Then
                    r = LogWrite(lastfilespec, subpos, "Небивка массы на " & Str(subpos_we_group - subpos_we_spec) & " груп=" & Str(subpos_we_group) & ", общая=" & Str(subpos_we_spec))
                End If
                If subpos_we_group <= 0.01 Then
                    r = LogWrite(lastfilespec, subpos, "Проверьте вес " & Str(subpos_we_group))
                End If
                If subpos_we_spec <= 0.01 Then
                    r = LogWrite(lastfilespec, subpos, "Проверьте вес " & Str(subpos_we_spec))
                End If
                For i = 1 To UBound(pos_out, 1)
                    pos_out(1, 6) = Round_w(pos_out(1, 6), n_round_w)
                Next i
            Case Else
                If Not UserForm2.show_type_CB.Value Then
                    pos_out_sort = ArraySort(pos_out, 1)
                    If sb_naen = "@" Then
                        n_row = 0
                    Else
                        n_row = n_n
                    End If
                    For i = 1 To UBound(pos_out, 1)
                        If pos_out_sort(i, 1) <> "Поз." And pos_out_sort(i, 1) <> sb_naen And pos_out_sort(i, 3) <> "" And Not IsEmpty(pos_out_sort(i, 3)) Then
                            n_row = n_row + 1
                            For j = 1 To UBound(pos_out, 2)
                                pos_out(n_row, j) = pos_out_sort(i, j)
                            Next j
                        End If
                    Next i
                    If n_row <> UBound(pos_out, 1) Then pos_out = ArrayRedim(pos_out, n_row)
                End If
        End Select
        SpecOneSubpos = pos_out
        Erase pos_out
    End If
End Function

Function SpecProkat(ByVal prokat As Variant, ByVal n_prokat As Integer, ByVal type_spec As Integer, Optional ByVal nSubPos As Integer = 1) As Variant
    If UserForm2.qtyOneSubpos_CB.Value Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
    Else
        n_txt = "," & vbLf & "на все"
    End If
    If UserForm2.show_qty_spec.Value Then n_txt = "" & ",**"
    un_chsum_prokat = ArrayUniqValColumn(prokat, col_chksum)
    pos_chsum_prokat = UBound(un_chsum_prokat, 1)
    If type_spec = 1 Or UserForm2.pr_pm_CB.Value Or (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then
        For i = 1 To pos_chsum_prokat
            If UserForm2.pr_pm_CB.Value Then
                If UserForm2.keep_pos_CB.Value Then
                    un_chsum_prokat(i) = Split(un_chsum_prokat(i), "_")(0) & Split(un_chsum_prokat(i), "_")(2)
                Else
                    un_chsum_prokat(i) = Split(un_chsum_prokat(i), "_")(0)
                End If
            Else
                un_chsum_prokat(i) = Split(un_chsum_prokat(i), "_")(0) & Split(un_chsum_prokat(i), "_")(2)
            End If
        Next i
        un_chsum_prokat = ArrayUniqValColumn(un_chsum_prokat, 1)
        pos_chsum_prokat = UBound(un_chsum_prokat, 1)
    End If
    
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    ReDim pos_out(pos_chsum_prokat, n_col_spec)
    For i = 1 To pos_chsum_prokat
        For j = 1 To n_prokat
            If type_spec = 1 Or UserForm2.pr_pm_CB.Value Or (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then
                If type_spec = 1 Or (type_spec = 3 And UserForm2.ignore_subpos_CB.Value) Then current_chksum = Split(prokat(j, col_chksum), "_")(0) & Split(prokat(j, col_chksum), "_")(2)
                If UserForm2.pr_pm_CB.Value And Not UserForm2.keep_pos_CB.Value Then current_chksum = Split(prokat(j, col_chksum), "_")(0)
                If UserForm2.pr_pm_CB.Value And UserForm2.keep_pos_CB.Value Then current_chksum = Split(prokat(j, col_chksum), "_")(0) & Split(prokat(j, col_chksum), "_")(2)
            Else
                current_chksum = prokat(j, col_chksum)
            End If
            If current_chksum = un_chsum_prokat(i) Then
                name_pr = GetShortNameForGOST(prokat(j, col_pr_gost_prof))
                n_el = prokat(j, col_qty) / nSubPos
                If InStr(1, name_pr, "Лист") Then
                    L = prokat(j, col_pr_length) / 1000
                Else
                    L = Round_w(prokat(j, col_pr_length) / 1000, 3)
                End If
                If UserForm2.pr_pm_CB.Value Then
                    we = prokat(j, col_pr_weight)
                Else
                    we = prokat(j, col_pr_weight) * L
                End If
                Select Case type_spec
                    Case 1
                        If UserForm2.pr_pm_CB.Value Then
                            pos_out(i, 1) = prokat(j, col_sub_pos) & n_txt
                            If UserForm2.keep_pos_CB.Value Then
                                pos_out(i, 2) = prokat(j, col_pos)
                            Else
                                pos_out(i, 2) = " "
                            End If
                            If InStr(1, name_pr, "Лист") Then
                                naen_plate = SpecMetallPlate(prokat(j, col_pr_prof), prokat(j, col_pr_naen), L, we)
                                pos_out(i, 3) = name_pr & prokat(j, col_pr_gost_prof) & " " & naen_plate(1)
                                L = naen_plate(2)
                                we = naen_plate(3)
                            Else
                                pos_out(i, 3) = name_pr & prokat(j, col_pr_gost_prof) & " " & prokat(j, col_pr_prof) & " L = п.м."
                            End If
                            pos_out(i, 4) = pos_out(i, 4) + (L * n_el)
                            pos_out(i, 5) = Round_w(we, n_round_w)
                        Else
                            pos_out(i, 1) = prokat(j, col_sub_pos) & n_txt
                            pos_out(i, 2) = prokat(j, col_pos)
                            If InStr(1, name_pr, "Лист") Then
                                naen_plate = SpecMetallPlate(prokat(j, col_pr_prof), prokat(j, col_pr_naen), L, we)
                                pos_out(i, 3) = name_pr & prokat(j, col_pr_gost_prof) & " " & naen_plate(1)
                                L = naen_plate(2)
                                we = naen_plate(3)
                            Else
                                pos_out(i, 3) = name_pr & prokat(j, col_pr_gost_prof) & " " & prokat(j, col_pr_prof) & " L=" & L * 1000 & "мм."
                            End If
                            pos_out(i, 4) = pos_out(i, 4) + n_el
                            pos_out(i, 5) = Round_w(we, n_round_w)
                        End If
                    Case Else
                        If UserForm2.pr_pm_CB.Value Then
                            If UserForm2.keep_pos_CB.Value Then
                                pos_out(i, 1) = prokat(j, col_pos)
                            Else
                                pos_out(i, 1) = " "
                            End If
                            pos_out(i, 2) = prokat(j, col_pr_gost_prof)
                            If InStr(1, name_pr, "Лист") Then
                                naen_plate = SpecMetallPlate(prokat(j, col_pr_prof), prokat(j, col_pr_naen), L, we)
                                pos_out(i, 3) = name_pr & " " & naen_plate(1)
                                L = naen_plate(2)
                                we = naen_plate(3)
                            Else
                                pos_out(i, 3) = name_pr & " " & prokat(j, col_pr_prof) & " L = п.м."
                            End If
                            pos_out(i, 4) = pos_out(i, 4) + (L * n_el)
                            pos_out(i, 5) = we
                            If show_sum_prim Then pos_out(i, 6) = pos_out(i, 6) + Round_w(L * we, n_round_w) * n_el
                        Else
                            pos_out(i, 1) = prokat(j, col_pos)
                            pos_out(i, 2) = prokat(j, col_pr_gost_prof)
                            If InStr(1, name_pr, "Лист") Then
                                naen_plate = SpecMetallPlate(prokat(j, col_pr_prof), prokat(j, col_pr_naen), L, we)
                                pos_out(i, 3) = name_pr & " " & naen_plate(1)
                                L = naen_plate(2)
                                we = naen_plate(3)
                            Else
                                pos_out(i, 3) = name_pr & prokat(j, col_pr_prof) & " L=" & L * 1000 & "мм."
                            End If
                            pos_out(i, 4) = pos_out(i, 4) + n_el
                            pos_out(i, 5) = Round_w(we, n_round_w)
                            If show_sum_prim Then pos_out(i, 6) = pos_out(i, 6) + n_el * pos_out(i, 5)
                        End If
                End Select
            End If
        Next j
    Next i
    If type_spec = 13 Then
        For i = 1 To UBound(pos_out, 1)
            pos_out(i, 7) = t_prokat
        Next
    End If
    If type_spec = 1 Then
        n_col_pos = 2
    Else
        n_col_pos = 1
    End If
    pos_out = ArraySort(pos_out, n_col_pos)
    SpecProkat = pos_out
    Erase prokat, pos_out
End Function

Function SpecMetallPlate(ByVal prokat_prof As String, ByVal prokat_naen As String, ByVal L As Double, ByVal we As Double) As Variant
    Dim array_out: ReDim array_out(4)
    prokat_naen_t = prokat_naen
    prokat_prof = Replace(prokat_prof, " ", "")
    prokat_prof = Replace(prokat_prof, "-", "")
    prokat_prof = Trim(prokat_prof)
    prokat_naen = Replace(prokat_naen, "Лист", "")
    prokat_naen = Replace(prokat_naen, " ", "")
    prokat_naen = Replace(prokat_naen, "-", "")
    prokat_naen = Replace(prokat_naen, "X", "*")
    prokat_naen = Replace(prokat_naen, "x", "*")
    prokat_naen = Replace(prokat_naen, "Х", "*")
    prokat_naen = Replace(prokat_naen, "х", "*")
    prokat_naen = Trim(prokat_naen)
    t_list = ConvTxt2Num(prokat_prof)
    If Not IsNumeric(t_list) Or t_list < 0.0001 Then
        MsgBox ("Ошибка в имени типа профиля листа, отсутствует толщина " + prokat_prof)
        array_out(1) = "ОШИБКА ТИПА ПРОФИЛЯ ЛИСТА"
        array_out(2) = 0.001
        array_out(3) = 0.001
        array_out(4) = 0.001
        SpecMetallPlate = array_out
        Exit Function
    End If
    flag_read = False
    t_list = t_list / 1000
    abc = Split(prokat_naen, "*")
    If UBound(abc) > 0 Then
        flag_read = True
        A = 0: b = 0: t = 100000: S = 0
        For nn = 0 To UBound(abc)
            k = ConvTxt2Num(abc(nn))
            If IsNumeric(k) Then
                k = k / 1000
                If k > A Then A = k
                If k < t Then t = k
                S = S + k
            End If
        Next nn
        b = S - A - t
        b = Round(b, 3)
        A = Round(A, 3)
        t = Round(t, 3)
        prokat_prof = "--" + ConvNum2Txt(t * 1000)
        prokat_naen = "--" + ConvNum2Txt(t * 1000) + "x" + ConvNum2Txt(b * 1000) + "x" + ConvNum2Txt(A * 1000)
        we_plate_one = A * b * t * 7850
    End If
    If b < 0.000001 Or t < 0.000001 Or A < 0.000001 Then
        MsgBox ("Ошибка в имени типа профиля листа " + prokat_naen_t)
        array_out(1) = "ОШИБКА ТИПА ПРОФИЛЯ ЛИСТА"
        array_out(2) = 0.001
        array_out(3) = 0.001
        array_out(4) = 0.001
        SpecMetallPlate = array_out
        Exit Function
    End If
'    If L < 0.000001 Or we < 0.000001 Then
'        L = b * A
'        we = t * 7850
'        If Not UserForm2.pr_pm_CB.Value Then we = we * L
'    End If

    L = b * A
    we = t * 7850
    If Not UserForm2.pr_pm_CB.Value Then we = we * L
    If (UserForm2.keep_pos_CB.Value And UserForm2.pr_pm_CB.Value) Or Not UserForm2.pr_pm_CB.Value Then
        If Not UserForm2.pr_pm_CB.Value Then we = we / L
        If flag_read Then
            If Round(b * 1000, 1) = 0 Then
                l_plate = L / A
                we_plate = we / A
            Else
                l_plate = 1
                we_plate = L * we
            End If
            naen_plate = prokat_naen
            If UserForm2.pr_pm_CB.Value And Round(b * 1000, 1) = 0 Then naen_plate = prokat_naen & " L=п.м."
            If Not UserForm2.pr_pm_CB.Value And Round(b * 1000, 1) = 0 Then naen_plate = prokat_naen & " L=" & l_plate * 1000 & "мм."
        Else
            naen_plate = prokat_naen & " S=" & L & "кв.м."
            we_plate = we
            l_plate = L
        End If
    Else
        naen_plate = prokat_prof & " S=кв.м."
        we_plate = we
        l_plate = L
    End If
    array_out(1) = naen_plate
    array_out(2) = l_plate
    array_out(3) = we_plate
    array_out(4) = we_plate_one
    SpecMetallPlate = array_out
End Function

Function SpecSubpos(ByVal subp As Variant, ByVal n_subp As Integer, ByVal type_spec As Integer, ByVal nSubPos As Integer) As Variant
    If UserForm2.qtyOneSubpos_CB.Value Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
    Else
        n_txt = "," & vbLf & "на все"
    End If
    If UserForm2.show_qty_spec.Value Then n_txt = "" & ",**"
    un_chsum_subp = ArrayUniqValColumn(subp, col_chksum)
    pos_chsum_subp = UBound(un_chsum_subp, 1)
    If type_spec = 1 Then
        For i = 1 To pos_chsum_subp
            un_chsum_subp(i) = Split(un_chsum_subp(i), "_")(0) & Split(un_chsum_subp(i), "_")(2)
        Next i
        un_chsum_subp = ArrayUniqValColumn(un_chsum_subp, 1)
        pos_chsum_subp = UBound(un_chsum_subp, 1)
    End If
    
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    Dim pos_out: ReDim pos_out(pos_chsum_subp, n_col_spec)
    For i = 1 To pos_chsum_subp
        For j = 1 To n_subp
            If type_spec = 1 Then
                current_chksum = Split(subp(j, col_chksum), "_")(0) & Split(subp(j, col_chksum), "_")(2)
            Else
                current_chksum = subp(j, col_chksum)
            End If
            If current_chksum = un_chsum_subp(i) Then
                n_el = subp(j, col_qty) / nSubPos
                pos = subp(j, col_pos)
                Weight = subp(j, col_m_weight)
                Select Case type_spec
                Case 1
                    pos_out(i, 1) = subpos & n_txt
                    pos_out(i, 2) = pos
                    pos_out(i, 4) = pos_out(j, 4) + n_el
                    pos_out(i, 3) = subp(j, col_m_naen) & " по " & subp(j, col_m_obozn)
                    pos_out(i, 5) = Weight
                Case Else
                    obozn = subp(j, col_m_obozn)
                    naen = subp(j, col_m_naen)
                    If InStr(naen, "!!!") <> 0 Or InStr(obozn, "!!!") <> 0 Then
                        If pos_data.Item("name").Exists(pos) Then
                            naen = pos_data.Item("name").Item(pos)(1)
                            obozn = pos_data.Item("name").Item(pos)(2)
                        End If
                    End If
                    pos_out(i, 1) = pos
                    pos_out(i, 2) = obozn
                    pos_out(i, 3) = naen
                    pos_out(i, 4) = pos_out(i, 4) + n_el
                    pos_out(i, 5) = Weight
                    If show_sum_prim Then pos_out(i, 6) = pos_out(i, 6) + n_el * Weight
                End Select
            End If
        Next j
    Next i
    If type_spec = 13 Then
        For i = 1 To UBound(pos_out, 1)
            pos_out(i, 7) = t_subpos
        Next
    End If
    If type_spec = 1 Then
        n_col_pos = 2
    Else
        n_col_pos = 1
    End If
    pos_out = ArraySort(pos_out, n_col_pos)
    SpecSubpos = pos_out
    Erase subp, pos_out
End Function

Function Spec_AS(ByRef all_data As Variant, ByVal type_spec As Integer) As Variant
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    Dim pos_out: ReDim pos_out(1, n_col_spec)
    If IsEmpty(all_data) Then Spec_AS = Empty: Exit Function
    qty_parent = UBound(pos_data.Item("parent").keys()) + 1
    qty_child = UBound(pos_data.Item("child").keys()) + 1
    If qty_parent < 0 And qty_child < 0 And (type_spec = 2 Or type_spec = 13) Then
        r = LogWrite("Ошибка спецификации", "", "Сборки отсутвуют. Создана общестроительная спецификаця")
        MsgBox ("Сборки отсутвуют. Создана общестроительная спецификаця")
        type_spec = 3
    End If
    If type_spec = 13 And ((qty_parent <= 1) Or (qty_parent < 1 And pos_data.Exists("-"))) Then
        MsgBox ("Сборок меньше двух. Создана общестроительная спецификаця")
        r = LogWrite("Ошибка спецификации", "", "Сборок меньше двух. Создана общестроительная спецификаця")
        type_spec = 3
    End If
    Select Case type_spec
        Case 1
            pos_out(1, 1) = "Марка" & vbLf & "изделия."
            pos_out(1, 2) = "Поз." & vbLf & "дет."
            pos_out(1, 3) = "Наименование"
            pos_out(1, 4) = "Кол-во*"
            If UserForm2.qtyOneSubpos_CB.Value Then
                pos_out(1, 6) = "Масса изделия, кг."
                pos_out(1, 5) = "Масса 1 дет., кг."
            Else
                pos_out(1, 6) = "Масса изделий, кг."
                pos_out(1, 5) = "Масса, кг."
            End If
        Case 13
            end_col = 6 + qty_parent
            If pos_data.Exists("-") Then end_col = end_col + 1
            ReDim pos_out(2, end_col)
            pos_out(1, 1) = "Поз."
            pos_out(1, 2) = "Обозначение"
            pos_out(1, 3) = "Наименование"
            If UserForm2.qtyOneSubpos_CB.Value Then
                pos_out(1, 4) = "Кол-во на 1 шт."
            Else
                pos_out(1, 4) = "Кол-во на все"
            End If
            pos_out(1, end_col - 2) = "Всего"
            pos_out(1, end_col - 1) = "Масса ед., кг."
            pos_out(1, end_col) = "Примечание"
        Case Else
            pos_out(1, 1) = "Поз."
            pos_out(1, 2) = "Обозначение"
            pos_out(1, 3) = "Наименование"
            pos_out(1, 4) = "Кол-во"
            pos_out(1, 5) = "Масса ед., кг."
            pos_out(1, 6) = "Примечание"
    End Select
    pos_out_up = pos_out
    Dim ch_key As String
    ch_key = "child"
    If qty_child <= 0 And ((type_spec = 1) Or (type_spec = 2)) Then
        If qty_parent >= 0 Then
            ch_key = "parent"
        Else
            Spec_AS = Empty
            Exit Function
        End If
    End If
    
    If type_spec = 1 Then
        Dim pos_end: ReDim pos_end(1, 6)
        If UserForm2.qtyOneSubpos_CB.Value Then
            pos_end(1, 1) = Space(60) & "* расход на одно изделие"
        Else
            pos_end(1, 1) = Space(60) & "* расход на все изделия"
        End If
        For Each subpos In ArraySort(pos_data.Item(ch_key).keys(), 1)
            pos_out_onesubpos = SpecOneSubpos(all_data, subpos, type_spec)
            If delim_group_ved Then
                If UBound(pos_out, 1) > 1 Then pos_out_onesubpos = ArrayCombine(pos_out_up, pos_out_onesubpos)
                pos_out_onesubpos = ArrayCombine(pos_out_onesubpos, pos_end)
                pos_out = ArrayCombine(pos_out, pos_out_onesubpos)
            Else
                pos_out = ArrayCombine(pos_out, SpecOneSubpos(all_data, subpos, type_spec))
                pos_out = ArrayCombine(pos_out, pos_end)
            End If
        Next
    End If
    
    If type_spec = 2 Then
        If qty_parent > 0 Then
            For Each subpos In ArraySort(pos_data.Item("parent").keys(), 1)
                pos_out = ArrayCombine(pos_out, SpecOneSubpos(all_data, subpos, type_spec))
            Next
        End If
        If pos_data.Exists("-") Then pos_out = ArrayCombine(pos_out, SpecOneSubpos(all_data, "-", type_spec))
    End If
    
    If type_spec = 3 Then
        If (pos_data.Exists("-") Or (UserForm2.show_subpos_CB.Value And (UBound(pos_data.Item("parent").keys()) >= 0))) Then
            pos_out = ArrayCombine(pos_out, SpecOneSubpos(all_data, "-", type_spec))
        Else
            Spec_AS = Empty
            Exit Function
        End If
    End If
    
    If type_spec = 13 Then
        Dim n_col_sb As Integer
        n_col_sb = 4
        Dim pos_out_subpos: ReDim pos_out_subpos(UBound(all_data, 1), end_col)
        Dim pos_out_arm: ReDim pos_out_arm(UBound(all_data, 1), end_col)
        Dim pos_out_prokat: ReDim pos_out_prokat(UBound(all_data, 1), end_col)
        Dim pos_out_izd: ReDim pos_out_izd(UBound(all_data, 1), end_col)
        Dim pos_out_mat: ReDim pos_out_mat(UBound(all_data, 1), end_col)
        Dim n_row_subpos As Integer
        Dim n_row_arm As Integer
        Dim n_row_prokat As Integer
        Dim n_row_izd As Integer
        Dim n_row_mat As Integer

        For Each subpos In ArraySort(pos_data.Item("parent").keys(), 1)
            If UserForm2.qtyOneSubpos_CB.Value Then
                nSubPos = GetNSubpos(subpos, type_spec)
                pos_out(2, n_col_sb) = subpos & vbLf & "(" & nSubPos & "шт)"
            Else
                nSubPos = 1
                pos_out(2, n_col_sb) = subpos
            End If
            If UserForm2.show_qty_spec.Value Then pos_out(2, n_col_sb) = subpos & ",**"
            pos_out_tmp = SpecOneSubpos(all_data, subpos, type_spec)
            For i = 1 To UBound(pos_out_tmp, 1)
                Select Case pos_out_tmp(i, 7)
                    Case t_subpos
                        r = ins_row(pos_out_subpos, pos_out_tmp, i, n_col_sb, n_row_subpos, nSubPos)
                    Case t_arm
                        r = ins_row(pos_out_arm, pos_out_tmp, i, n_col_sb, n_row_arm, nSubPos)
                    Case t_prokat
                        r = ins_row(pos_out_prokat, pos_out_tmp, i, n_col_sb, n_row_prokat, nSubPos)
                    Case t_izd
                        r = ins_row(pos_out_izd, pos_out_tmp, i, n_col_sb, n_row_izd, nSubPos)
                    Case t_mat
                        r = ins_row(pos_out_mat, pos_out_tmp, i, n_col_sb, n_row_mat, nSubPos)
                End Select
            Next i
            n_col_sb = n_col_sb + 1
        Next
        If pos_data.Exists("-") Then
            pos_out_tmp = SpecOneSubpos(all_data, "-", type_spec)
            For i = 1 To UBound(pos_out_tmp, 1)
                Select Case pos_out_tmp(i, 7)
                    Case t_subpos
                        r = ins_row(pos_out_subpos, pos_out_tmp, i, n_col_sb, n_row_subpos, 1)
                    Case t_arm
                        r = ins_row(pos_out_arm, pos_out_tmp, i, n_col_sb, n_row_arm, 1)
                    Case t_prokat
                        r = ins_row(pos_out_prokat, pos_out_tmp, i, n_col_sb, n_row_prokat, 1)
                    Case t_izd
                        r = ins_row(pos_out_izd, pos_out_tmp, i, n_col_sb, n_row_izd, 1)
                    Case t_mat
                        r = ins_row(pos_out_mat, pos_out_tmp, i, n_col_sb, n_row_mat, 1)
                End Select
            Next i
            pos_out(2, end_col - 3) = "Прочее"
        End If
        
        pos_out_subpos = ArraySort(ArrayRedim(pos_out_subpos, n_row_subpos), 1)
        pos_out_arm = ArraySort(ArrayRedim(pos_out_arm, n_row_arm), 1)
        pos_out_prokat = ArraySort(ArrayRedim(pos_out_prokat, n_row_prokat), 1)
        pos_out_izd = ArraySort(ArrayRedim(pos_out_izd, n_row_izd), 1)
        pos_out_mat = ArraySort(ArrayRedim(pos_out_mat, n_row_mat), 1)
        
        If n_row_subpos > 0 Then pos_out = ArrayCombine(pos_out, pos_out_subpos)
        If n_row_arm > 0 Then pos_out = ArrayCombine(pos_out, pos_out_arm)
        If n_row_prokat > 0 Then pos_out = ArrayCombine(pos_out, pos_out_prokat)
        If n_row_izd > 0 Then pos_out = ArrayCombine(pos_out, pos_out_izd)
        If n_row_mat > 0 Then pos_out = ArrayCombine(pos_out, pos_out_mat)
        For i = 3 To UBound(pos_out, 1)
            If Not IsEmpty(pos_out(i, end_col - 1)) Then
                For j = 4 To end_col - 1
                    If IsEmpty(pos_out(i, j)) Then pos_out(i, j) = "-"
                Next j
            End If
            If IsNumeric(pos_out(i, end_col)) Then
                If Round_w(pos_out(i, end_col), 0) > 0 Then
                    pos_out(i, end_col) = Trim(ConvNum2Txt(pos_out(i, end_col)) & " кг.")
                    If Left(pos_out(i, end_col), 1) = "." Then pos_out(i, end_col) = "0" + pos_out(i, end_col)
                Else
                    pos_out(i, end_col) = "-"
                End If
            End If
        Next i
        
        If Not UserForm2.show_type_CB.Value Then
            pos_out_sort = ArraySort(pos_out, 1)
            n_row = 2
            For i = 1 To UBound(pos_out, 1)
                If pos_out_sort(i, 1) <> "Поз." And pos_out_sort(i, 3) <> "" Then
                    n_row = n_row + 1
                    For j = 1 To UBound(pos_out, 2)
                        pos_out(n_row, j) = pos_out_sort(i, j)
                    Next j
                End If
            Next i
        End If

    Else
        If Not IsEmpty(pos_out) Then
            For i = 2 To UBound(pos_out, 1)
                If pos_out(i, 3) <> "" Then
                    If IsNumeric(ConvTxt2Num(pos_out(i, 6))) Then
                        If Round_w(pos_out(i, 6), 0) > 0 Then
                            pos_out(i, 6) = Trim(ConvNum2Txt(pos_out(i, 6)) & " кг.")
                            If Left(pos_out(i, 6), 1) = "." Then pos_out(i, 6) = "0" + pos_out(i, 6)
                        Else
                            pos_out(i, 6) = "-"
                        End If
                        If pos_out(i, 5) = 0 Then pos_out(i, 5) = "-"
                    Else
                        For kk = 5 To 6
                            If pos_out(i, kk) = "" Or pos_out(i, kk) = " " Or pos_out(i, kk) = 0 Then pos_out(i, kk) = "-"
                        Next kk
                    End If
                End If
            Next i
        End If
    End If
    
    If Not IsEmpty(pos_out) Then
        n_col_naen = 3: n_col_pos = 1
        If type_spec = 1 Then
            n_col_naen = 2
            n_col_pos = 2
        End If
        For i = 2 To UBound(pos_out, 1)
            If (Right(Trim(UCase(pos_out(i, n_col_naen))), 1) = "*") Then
                pos_out(i, n_col_naen) = Left(pos_out(i, n_col_naen), Len(pos_out(i, n_col_naen)) - 1)
                pos_out(i, 1) = pos_out(i, 1) & "*"
            End If
        Next i
    End If
    Spec_AS = pos_out
End Function

Function Spec_AS2arr(ByVal filename As String) As Variant
    all_data = DataRead(filename & ".txt")
    data_t = ReadFile(filename & ".txt")
    type_spec = 3
    spec = SpecOneSubpos("-", type_spec)
    Spec_AS2arr = ArrayEmp2Space(spec)
End Function

Function Spec_WIN(ByRef all_data As Variant) As Variant
    all_data = ArraySelectParam(all_data, t_wind, col_type_el)
    If IsEmpty(all_data) Then Spec_WIN = Empty: Exit Function
    un_chsum = ArrayUniqValColumn(all_data, col_chksum)
    pos_chsum = UBound(un_chsum, 1)
    un_floor = ArrayUniqValColumn(all_data, col_w_nfloor)
    For i = 1 To pos_chsum
        un_chsum(i) = Split(un_chsum(i), "_")(0)
    Next i
    un_chsum = ArrayUniqValColumn(un_chsum, 1)
    pos_chsum = UBound(un_chsum, 1)
    Dim pos_zag
    If UserForm2.qtyOneSubpos_CB.Value Then
        ReDim pos_zag(2, UBound(un_floor) + 6)
        pos_zag(1, 1) = "Позиция"
        pos_zag(1, 2) = "Обозначение"
        pos_zag(1, 3) = "Наименование"
        pos_zag(1, 4) = "Количество"
        i = 4
        floor_start = i
        For Each tfloor In un_floor
            For j = 1 To UBound(all_data, 1)
                If tfloor = all_data(j, col_w_nfloor) Then
                    pos_zag(2, i) = all_data(j, col_w_floor)
                    Exit For
                End If
            Next j
            i = i + 1
        Next
        floor_end = i - 1
        pos_zag(1, i) = "Всего"
        n_col_qty = i
        i = i + 1
        pos_zag(1, i) = "Масса ед., кг"
        i = i + 1
        pos_zag(1, i) = "Примечание"
    Else
        ReDim pos_zag(1, 6)
        pos_zag(1, 1) = "Позиция"
        pos_zag(1, 2) = "Обозначение"
        pos_zag(1, 3) = "Наименование"
        pos_zag(1, 4) = "Кол-во"
        pos_zag(1, 5) = "Масса ед., кг"
        pos_zag(1, 6) = "Примечание"
        n_col_qty = 4
    End If
    Dim pos_out(): ReDim pos_out(pos_chsum, UBound(pos_zag, 2))
    pos_wind = "окно"
    pos_door = "дверь"
    n_row_out = 0
    For Each t In Array(pos_wind, pos_door)
        el_data = ArraySelectParam(all_data, t, col_pos)
        un_sub_pos_el = ArrayUniqValColumn(el_data, col_sub_pos)
        If Not IsEmpty(un_sub_pos_el) Then
            For Each sub_pos In un_sub_pos_el
                pos_dat = ArraySelectParam(all_data, sub_pos, col_sub_pos)
                'Поставим заполнение впереди
                un_pos = ArrayCombine(Array(t), ArrayUniqValColumn(pos_dat, col_pos))
                For i = 1 To UBound(un_pos)
                    If un_pos(i) = t And i > 1 Then un_pos(i) = Empty
                Next i
                For Each pos_el In un_pos
                    If Not IsEmpty(pos_el) Then n_row_out = n_row_out + 1
                    For i = 1 To UBound(pos_dat)
                        tpos = pos_dat(i, col_pos)
                        If tpos = pos_el Then
                            sub_pos = Replace(pos_dat(i, col_sub_pos), "zap", "")
                            pos = pos_dat(i, col_pos) & " "
                            If pos = t & " " Then pos = ""
                            obozn = pos_dat(i, col_w_obozn)
                            naen = pos & pos_dat(i, col_w_naen)
                            qty = pos_dat(i, col_qty)
                            Weight = pos_dat(i, col_w_weight)
                            prim = pos_dat(i, col_w_prim)
                            pos_out(n_row_out, 1) = sub_pos
                            pos_out(n_row_out, 2) = obozn
                            pos_out(n_row_out, 3) = naen
                            If UserForm2.qtyOneSubpos_CB.Value Then
                                Floor = pos_dat(i, col_w_floor)
                                For k = floor_start To floor_end
                                    If pos_zag(2, k) = Floor Then pos_out(n_row_out, k) = pos_out(n_row_out, k) + qty
                                Next k
                            End If
                            pos_out(n_row_out, n_col_qty) = pos_out(n_row_out, n_col_qty) + qty
                            pos_out(n_row_out, n_col_qty + 1) = Weight
                            If IsEmpty(pos_out(n_row_out, n_col_qty + 2)) Then pos_out(n_row_out, n_col_qty + 2) = prim
                        End If
                    Next i
                Next
            Next
        End If
    Next
    If UserForm2.qtyOneSubpos_CB.Value Then
        For k = floor_start To floor_end
            For Each deltxt In Array("План", "НА")
                pos_zag(2, k) = Replace(pos_zag(2, k), deltxt, "")
            Next
            pos_zag(2, k) = Trim(pos_zag(2, k))
            For i = 1 To n_row_out
                If IsEmpty(pos_out(i, k)) Then pos_out(i, k) = "-"
            Next i
        Next k
    End If
    Spec_WIN = ArrayCombine(pos_zag, pos_out)
End Function

Function Spec_KM(ByRef all_data As Variant) As Variant
    prokat = ArraySelectParam(all_data, t_prokat, col_type_el)
    If IsEmpty(prokat) Then
        n_prokat = 0
        MsgBox ("Прокат в файле/листе не найден")
        r = LogWrite("Ошибка спецификации", "", "Прокат в файле/листе не найден")
        Spec_KM = Empty
        Exit Function
    Else
        n_prokat = UBound(prokat, 1)
    End If

    If UserForm2.ed_izm_km_CB.Value Then
        ed_izm = "кг."
        koef = 1
        n_okr = 0
        w_format = "0"
    Else
        ed_izm = "т."
        koef = 1000
        n_okr = 2
        w_format = "0.00"
    End If

    unique_type_konstr = ArrayUniqValColumn(prokat, col_pr_type_konstr)
    n_type_konstr = UBound(unique_type_konstr)
    unique_stal = ArrayUniqValColumn(prokat, col_pr_st)
    n_stal = UBound(unique_stal)
    
    Dim pos_out(): ReDim pos_out(n_prokat * 2 + 30, n_type_konstr + 5)
    Dim weight_stal(): ReDim weight_stal(1, n_type_konstr + 5)
    weight_stal(1, 2) = "Итого"
    Dim weight_gost_prof(): ReDim weight_gost_prof(1, n_type_konstr + 5)
    weight_gost_prof(1, 1) = "Всего профиля:"
    Dim weight_total(): ReDim weight_total(1, n_type_konstr + 5)
    weight_total(1, 1) = "Всего масса металла:"

    Dim weight_stal_total(): ReDim weight_stal_total(n_stal + 1, n_type_konstr + 5)
    For i = 1 To n_stal + 1
        If i = 1 Then
            weight_stal_total(i, 1) = "В том числе по маркам:"
        Else
            weight_stal_total(i, 1) = unique_stal(i - 1)
        End If
    Next i
    row = 1
    pos_out(row, 1) = "Наименование профиля" & vbLf & "ГОСТ, ТУ"
    pos_out(row, 2) = "Наименование или марка металла" & vbLf & "ГОСТ, ТУ"
    pos_out(row, 3) = "Номер или размеры профиля, мм."
    pos_out(row, 4) = "№" & vbLf & "п.п"
    pos_out(row, n_type_konstr + 5) = "Общая" & vbLf & "масса," & ed_izm
    For i = 1 To n_type_konstr + 5
        pos_out(row + 1, i) = pos_out(row, i)
    Next i
    pos_out(row + 1, 1) = pos_out(row, 1)
    For i = 5 To n_type_konstr + 4
        pos_out(row, i) = "Масса металла, " & ed_izm
        pos_out(row + 1, i) = unique_type_konstr(i - 4)
    Next i
    row = 3
    For i = 1 To n_type_konstr + 5
        pos_out(row, i) = i
    Next i
    row = 4
    unique_gost_prof = ArrayUniqValColumn(prokat, col_pr_gost_prof)
    n_gost_prof = UBound(unique_gost_prof)
    For i = 1 To n_gost_prof
        'Все элементы с заданным типом профиля
        gost_prof = unique_gost_prof(i) 'Текущий тип профиля
        prof_stal = ArraySelectParam(prokat, gost_prof, col_pr_gost_prof) 'Выбираем все элементы с таким профилем
        unique_prof_stal = ArrayUniqValColumn(prof_stal, col_pr_st) 'Какая сталь в них используется
        n_stal = UBound(unique_prof_stal)
        For j = 1 To UBound(unique_prof_stal)
            'Все элементы с заданной сталью
            stal = unique_prof_stal(j) 'Текущая сталь
            prof = ArraySelectParam(prof_stal, stal, col_pr_st) 'Выбираем все элементы с этой сталью
            gost_stal = prof(1, col_pr_gost_st)
            unique_prof = ArrayUniqValColumn(prof, col_pr_prof) 'Какие типоразмеры профилей
            n_prof = UBound(unique_prof)
            For L = 1 To n_prof
                'Все элементы с заданным сечением
                konstr = unique_prof(L) 'Текущий типоразмер профиля
                el = ArraySelectParam(prof, konstr, col_pr_prof) 'Выбираем все элементы с этим профилем
                unique_konstr = ArrayUniqValColumn(el, col_pr_type_konstr) 'Какие типы конструкций
                n_t_konstr = UBound(unique_konstr)
                For k = 1 To n_t_konstr
                    'Все элементы с заданной конструкцией
                    type_konstr = unique_konstr(k) 'Текущий тип конструкции
                    elem_m = ArraySelectParam(el, type_konstr, col_pr_type_konstr) 'Выбираем все элементы с этим типом конструкций
                    n_el_m = UBound(elem_m, 1)
                    Weight# = 0 'Начинаем считать вес для каждого типа
                    For kk = 1 To n_el_m
                        'Вес одной строки, с учётом того, что масса дана на п.м. в кг.
                        wp = elem_m(kk, col_pr_weight) * elem_m(kk, col_qty) * elem_m(kk, col_pr_length) / 1000
                        Weight# = Weight# + wp
                    Next kk
                    'Итоговый вес для отдельного типа конструкции, в тоннах
                    'Из-за особенностей ГОСТа минимальное значение - 100 кг.
                    'Не плохой такой источник экономии
                    Weight# = Round_w(Weight# / koef, n_okr)
                    If Weight# < (1 / (10 ^ n_okr)) Then
                        wt = ConvNum2Txt(Weight#)
                        If Len(wt) > Len(w_format) Then
                            w_format_t = "0."
                            For nnul = 1 To Len(wt) - Len(w_format_t)
                                w_format_t = w_format_t + "0"
                            Next nnul
                            w_format = w_format_t
                        End If
                    End If
                    'Записываем в массив результат
                    n_konstr = GetNumberConstr(unique_type_konstr, type_konstr) + 4
                    n_stal_tot = GetNumberStal(unique_stal, stal) + 1
                    pos_out(row, 1) = GetNameForGOST(gost_prof) 'Имя
                    pos_out(row, 2) = stal & vbLf & gost_stal 'Сталь
                    pos_out(row, 3) = konstr 'Текущий типоразмер профиля
                    pos_out(row, 4) = row - 3 'Порядковый номер
                    pos_out(row, n_konstr) = pos_out(row, n_konstr) + Weight# 'Записываем вес по типу конструкции
                    pos_out(row, n_type_konstr + 5) = pos_out(row, n_type_konstr + 5) + Weight# 'Записываем вес по типоразмеру профиля
                    weight_stal(1, n_konstr) = weight_stal(1, n_konstr) + Weight# 'Записываем вес по типу стали для этого типа профиля
                    weight_gost_prof(1, n_konstr) = weight_gost_prof(1, n_konstr) + Weight# 'Записываем вес по типу профиля
                    weight_total(1, n_konstr) = weight_total(1, n_konstr) + Weight# 'Записываем общий вес
                    weight_stal_total(n_stal_tot, n_konstr) = weight_stal_total(n_stal_tot, n_konstr) + Weight# 'Записываем вес по типу стали
                Next k
                row = row + 1
            Next L
            'Ну, дальше всё понятно
            weight_stal(1, 1) = GetNameForGOST(gost_prof)
            weight_stal(1, 4) = row - 3
            For m = 1 To n_type_konstr + 4
                pos_out(row, m) = weight_stal(1, m)
                If m > 4 Then
                    pos_out(row, n_type_konstr + 5) = pos_out(row, n_type_konstr + 5) + pos_out(row, m)
                    weight_stal(1, m) = 0
                End If
            Next m
            row = row + 1
        Next j
        weight_gost_prof(1, 4) = row - 3
        For m = 1 To n_type_konstr + 4
            pos_out(row, m) = weight_gost_prof(1, m)
            If m > 4 Then
                pos_out(row, n_type_konstr + 5) = pos_out(row, n_type_konstr + 5) + pos_out(row, m)
                weight_gost_prof(1, m) = 0
            End If
        Next m
        row = row + 1
    Next i
    weight_total(1, 4) = row - 3
    For i = 1 To n_type_konstr + 4
        pos_out(row, i) = weight_total(1, i)
        If i > 4 Then
            pos_out(row, n_type_konstr + 5) = pos_out(row, n_type_konstr + 5) + pos_out(row, i)
            weight_total(1, i) = 0
        End If
    Next i
    row = row + 1
    For i = 1 To UBound(unique_stal) + 1
        weight_stal_total(i, 4) = row - 3
        For j = 1 To n_type_konstr + 4
            pos_out(row, j) = weight_stal_total(i, j)
            If j > 4 Then pos_out(row, n_type_konstr + 5) = pos_out(row, n_type_konstr + 5) + weight_stal_total(i, j)
        Next j
        row = row + 1
    Next i
    mat = ArraySelectParam(all_data, t_mat, col_type_el)
    naen_mat = Array("Покр", "Огне")
    If Not IsEmpty(mat) Then
        n_mat = UBound(mat, 1)
        un_pos_mat = ArrayUniqValColumn(mat, col_chksum)
        For i = 1 To UBound(un_pos_mat, 1)
            un_pos_mat(i) = Split(un_pos_mat(i), "_")(0)
        Next i
        un_pos_mat = ArrayUniqValColumn(un_pos_mat, 1)
        pos_mat = UBound(un_pos_mat, 1)
        
        pos_out(row, 1) = "Антикоррозийная окраска"
        row = row + 1
        For i = 1 To pos_mat
            For j = 1 To n_mat
                current_chksum = Split(mat(j, col_chksum), "_")(0)
                chksum = un_pos_mat(i)
                If current_chksum = chksum Then
                    naen = mat(j, col_m_naen)
                    obozn = mat(j, col_m_obozn): If obozn <> "" Then obozn = " по " & obozn
                    ed = mat(j, col_m_edizm)
                    qty = mat(j, col_qty)
                    usl = 0
                    For Each n In naen_mat
                        usl = usl + InStr(naen, n)
                    Next
                    If usl > 0 Then
                        pos_out(row, 1) = naen & obozn & ", " & ed
                        pos_out(row, n_type_konstr + 5) = pos_out(row, n_type_konstr + 5) + qty
                    End If
                End If
            Next j
            If pos_out(row, n_type_konstr + 5) <> 0 Then row = row + 1
        Next i
    End If
    Erase prokat, unique_gost_prof, unique_stal, prof_stal, unique_prof_stal, unique_type_konstr, prof, unique_prof, el, unique_konstr, elem_m, weight_stal, weight_gost_prof, weight_total, weight_stal_total
    pos_out = ArrayRedim(pos_out, row - 1)
    Spec_KM = pos_out
End Function

Function Spec_KZH(ByRef all_data As Variant) As Variant
    is_bet = False
    If show_bet_wkzh Then is_bet = Spec_CONC(all_data)
    Set name_subpos = pos_data.Item("name") 'Словарь с именами сборок
    un_child = ArraySort(pos_data.Item("child").keys())
    un_parent = ArraySort(pos_data.Item("parent").keys())
    If IsEmpty(un_child) Then un_child = Array()
    If IsEmpty(un_parent) Then un_parent = Array()
    'Выясняем - какие диаметры и какие классы арматуры есть для всех сборок
    'заодно отсортируем арматуру в закладных деталях и прокат
    n_row = UBound(all_data, 1)
    Dim arm_arr: ReDim arm_arr(8)
    Dim temp_arr: ReDim temp_arr(n_row, max_col)
    For i = 1 To 4
        arm_arr(i) = temp_arr
    Next i
    n_arm_a = 0: n_arm_z = 0: n_prokat_a = 0: n_prokat_z = 0
    For i = 1 To n_row
        сurrent_type_el = all_data(i, col_type_el)
        If сurrent_type_el = t_arm Or сurrent_type_el = t_prokat Then
            сurrent_subpos = all_data(i, col_sub_pos)
            naen = " "
            If name_subpos.Exists(сurrent_subpos) Then naen = name_subpos.Item(сurrent_subpos)(1)
            Select Case сurrent_type_el
                Case t_arm
                    If InStr(naen, "Заклад") = 0 Then
                        n_arm_a = n_arm_a + 1
                        For j = 1 To max_col
                            arm_arr(1)(n_arm_a, j) = all_data(i, j)
                        Next j
                        arm_arr(4 + 1) = n_arm_a
                    Else
                        n_arm_z = n_arm_z + 1
                        For j = 1 To max_col
                            arm_arr(3)(n_arm_z, j) = all_data(i, j)
                        Next j
                        arm_arr(4 + 3) = n_arm_z
                    End If
                Case t_prokat
                    If InStr(naen, "Заклад") = 0 Then
                        n_prokat_a = n_prokat_a + 1
                        For j = 1 To max_col
                            arm_arr(2)(n_prokat_a, j) = all_data(i, j)
                        Next j
                        arm_arr(4 + 2) = n_prokat_a
                    Else
                        n_prokat_z = n_prokat_z + 1
                        For j = 1 To max_col
                            arm_arr(4)(n_prokat_z, j) = all_data(i, j)
                        Next j
                        arm_arr(4 + 4) = n_prokat_z
                    End If
            End Select
        End If
    Next
    'Теперь у нас есть массив с отсортированной арматурой для всех сборок
    '1 - Арматура общая
    '2 - Прокат общий
    '3 - Арматура в закладных
    '4 - Прокат в закладных

    'Сформируем общую таблицу диаметров и классов арматуры
    n_row = 5
    If UBound(un_parent) >= 0 Then n_row = n_row + UBound(un_parent)
    If pos_data.Exists("-") Then n_row = n_row + 1
    sum_row = 0: If n_row - 5 > 1 And sum_row_wkzh = True Then sum_row = 1
    Dim pos_out: ReDim pos_out(n_row + sum_row, 1)
    For j = 1 To 4
        If Not IsEmpty(arm_arr(4 + j)) Then
            If ((j = 1) Or (j = 2)) Then
                n_type = "Изделия арматурные"
            Else
                n_type = "Изделия закладные"
            End If
            If ((j = 1) Or (j = 3)) Then
                un_klass_arm = ArrayUniqValColumn(arm_arr(j), col_klass)
                n_klass_arm = UBound(un_klass_arm, 1)
                For i = 1 To n_klass_arm
                    current_klass = un_klass_arm(i)
                    arm_current_class = ArraySelectParam(arm_arr(j), current_klass, col_klass)
                    un_diam = ArrayUniqValColumn(arm_current_class, col_diametr)
                    n_diam = UBound(un_diam, 1)
                    current_size = UBound(pos_out, 2)
                    ReDim Preserve pos_out(n_row + sum_row, current_size + n_diam + 1)
                    For k = 1 To n_diam + 1
                        pos_out(1, current_size + k) = n_type
                        pos_out(2, current_size + k) = "Арматура класса"
                        pos_out(3, current_size + k) = current_klass
                        pos_out(4, current_size + k) = GetGOSTForKlass(current_klass)
                        If k > n_diam Then
                            pos_out(5, current_size + k) = "Всего"
                        Else
                            pos_out(5, current_size + k) = symb_diam & un_diam(k)
                        End If
                    Next k
                Next i
            Else
                un_stal_pr = ArrayUniqValColumn(arm_arr(j), col_pr_st)
                n_stal_pr = UBound(un_stal_pr, 1)
                For i = 1 To n_stal_pr
                    current_slal = un_stal_pr(i)
                    pr_current_slal = ArraySelectParam(arm_arr(j), current_slal, col_pr_st)
                    stal_gost = pr_current_slal(1, col_pr_gost_st)
                    un_gost = ArrayUniqValColumn(pr_current_slal, col_pr_gost_prof)
                    n_gost = UBound(un_gost, 1)
                    For L = 1 To n_gost
                        current_gost = un_gost(L)
                        pr_current_gost = ArraySelectParam(pr_current_slal, current_gost, col_pr_gost_prof)
                        un_prof = ArrayUniqValColumn(pr_current_gost, col_pr_prof)
                        n_prof = UBound(un_prof, 1)
                        current_size = UBound(pos_out, 2)
                        ReDim Preserve pos_out(n_row + sum_row, current_size + n_prof + 1)
                        For k = 1 To n_prof + 1
                            pos_out(1, current_size + k) = n_type
                            pos_out(2, current_size + k) = "Прокат"
                            pos_out(3, current_size + k) = current_slal & vbLf & stal_gost
                            pos_out(4, current_size + k) = current_gost
                            If k > n_prof Then
                                pos_out(5, current_size + k) = "Всего"
                            Else
                                pos_out(5, current_size + k) = un_prof(k)
                            End If
                        Next k
                    Next L
                Next i
            End If
            flag = 0
            If ((n_prokat_a = 0) And (j = 1)) Then flag = 1
            If ((n_prokat_a) And (j = 2)) Then flag = 1
            If ((n_prokat_z = 0) And (j = 3)) Then flag = 1
            If ((n_prokat_z) And (j = 4)) Then flag = 1
            If flag Then
                current_size = UBound(pos_out, 2)
                ReDim Preserve pos_out(n_row + sum_row, current_size + 1)
                pos_out(1, current_size + 1) = n_type
                For kk = 2 To 5
                    pos_out(kk, current_size + 1) = "Всего"
                Next kk
            End If
        End If
    Next j
    current_size = UBound(pos_out, 2)
    ReDim Preserve pos_out(n_row + sum_row, current_size + 1)
    Dim pos_out_bet()
    ReDim pos_out_bet(n_row + sum_row, 1)
    For kk = 1 To 5
        pos_out(kk, current_size + 1) = "Всего"
    Next kk
    pos_out(1, 1) = "Марка элемента"
    current_size = current_size + 1
    'Теперь мы знаем общий размер таблицы
    'Чтобы быстро находить адрес для записи веса - сформируем словарь
    'Искать будем по комбинации ТИП(закладная/общая)+Класс+Сечение
    'Т.е. для арматуры будет "АрматураИзделия закладные_A-III(A400)_16"
    'Для проката "Прокат_Изделия закладные_С245_ГОСТ 19771-93_100x4"
    'Результатом поиска будет номер столбца для записи значения
    Set weight_index = CreateObject("Scripting.Dictionary")
    weight_index.comparemode = 1
    For k = 6 To n_row
        If (pos_data.Exists("-") And k = n_row) Or (UBound(un_parent) <= 0) Then
            subpos = "-"
            nSubPos = 1
        Else
            subpos = un_parent(k - 5)
            nSubPos = pos_data.Item("qty").Item("-_" & subpos)
            If nSubPos < 1 Then
                r = LogWrite("Ошибка спецификации", subpos, "Не определено кол-во сборок")
                MsgBox ("Не определено кол-во сборок " & subpos & ", принято 1 шт.")
                nSubPos = 1
            End If
        End If
        If UserForm2.qtyOneSubpos_CB.Value Then
            n_txt = subpos & " (" & nSubPos & " шт.)"
        Else
            nSubPos = 1
            n_txt = subpos & ", " & "на все"
        End If
        If UserForm2.show_qty_spec.Value Then n_txt = "" & ",**"
        pos_out(k, 1) = n_txt
        If subpos = "-" Then pos_out(k, 1) = "Прочие"
        weight_index.Item("row" & subpos) = k
        If is_bet = True Then
            For Each sub_bet In concrsubpos.keys()
                If InStr(sub_bet, "_") > 0 And Right(sub_bet, 4) = "_qty" And InStr(sub_bet, "bet") = 0 Then
                    subb = Split(sub_bet, "_")
                    If subb(0) = subpos Then
                        v_bet = concrsubpos.Item(sub_bet)
                        naen_bet = subb(1)
                        flag = 1
                        If IsEmpty(pos_out_bet(k, 1)) Then
                            pos_out_bet(k, 1) = v_bet
                            pos_out_bet(2, 1) = naen_bet
                            flag = 0
                        Else
                            For j = 1 To UBound(pos_out_bet, 2)
                                If pos_out_bet(2, j) = naen_bet Then
                                    pos_out_bet(k, j) = pos_out_bet(k, j) + v_bet
                                    flag = 0
                                End If
                            Next j
                        End If
                        If flag Then
                            ReDim Preserve pos_out_bet(n_row + sum_row, UBound(pos_out_bet, 2) + 1)
                            pos_out_bet(k, UBound(pos_out_bet, 2)) = v_bet
                            pos_out_bet(2, UBound(pos_out_bet, 2)) = naen_bet
                        End If
                    End If
                End If
            Next
        End If
    Next k
    
    If UBound(pos_out_bet, 2) > 1 And is_bet = True Then
        ReDim Preserve pos_out_bet(n_row + sum_row, UBound(pos_out_bet, 2) + 1)
        pos_out_bet(1, 1) = "Объём бетона, куб.м."
        pos_out_bet(1, UBound(pos_out_bet, 2)) = "Всего"
        For k = 6 To n_row
            For i = 1 To UBound(pos_out_bet, 2) - 1
                pos_out_bet(k, UBound(pos_out_bet, 2)) = pos_out_bet(k, UBound(pos_out_bet, 2)) + Round_w(pos_out_bet(k, i), n_round_wkzh)
            Next i
        Next k
    End If
    For i = 1 To current_size
        If pos_out(2, i) = "Прокат" Then
            tkey = "Прокат" & pos_out(1, i) & pos_out(3, i) & pos_out(4, i) & pos_out(5, i)
        ElseIf pos_out(2, i) = "Арматура класса" Then
            tkey = "Арматура" & pos_out(1, i) & pos_out(3, i) & pos_out(4, i) & pos_out(5, i)
        Else
            tkey = pos_out(1, i) & pos_out(2, i)
        End If
        weight_index.Item("col" & tkey) = i
    Next i
    'Теперь из ранее созданного массива будем вытаскивать элементы для каждой сборки
    For i = 1 To 4
        If ((i = 1) Or (i = 2)) Then
            n_type = "Изделия арматурные"
        Else
            n_type = "Изделия закладные"
        End If
        For j = 1 To UBound(arm_arr(i), 1)
            subpos = arm_arr(i)(j, col_sub_pos)
            tparent = arm_arr(i)(j, col_parent)
            u1 = (pos_data.Item("parent").Exists(subpos) Or pos_data.Item("parent").Exists(tparent))
            If pos_data.Exists("-") Then u2 = ((subpos = "-") Or (pos_data.Item("-").Exists(subpos) And tparent = "-"))
            If u1 Or u2 Then
                If u2 Then
                    nSubPos = 1
                    k = weight_index.Item("row" & tparent)
                End If
                If u1 Then
                    If pos_data.Item("parent").Exists(subpos) Then
                        nSubPos = pos_data.Item("qty").Item("-_" & subpos)
                        k = weight_index.Item("row" & subpos)
                    End If
                    If pos_data.Item("parent").Exists(tparent) Then
                        nSubPos = pos_data.Item("qty").Item("-_" & tparent)
                        k = weight_index.Item("row" & tparent)
                    End If
                End If
                If Not UserForm2.qtyOneSubpos_CB.Value Then nSubPos = 1
                If arm_arr(i)(j, col_type_el) = t_arm Then
                    diametr = arm_arr(i)(j, col_diametr)
                    klass = arm_arr(i)(j, col_klass)
                    qty = arm_arr(i)(j, col_qty)
                    gost = GetGOSTForKlass(klass)
                    length_pos = arm_arr(i)(j, col_length) / 1000
                    weight_pm = GetWeightForDiametr(diametr, klass)
                    w_pos = Round_w(weight_pm * length_pos * k_zap_total, n_round_w) * qty / nSubPos
                    tkeyd = "Арматура" & n_type & klass & gost & symb_diam & diametr
                    tkesum_1 = "Арматура" & n_type & klass & gost & "Всего"
                Else
                    prof = arm_arr(i)(j, col_pr_prof)
                    gost_prof = arm_arr(i)(j, col_pr_gost_prof)
                    stal = arm_arr(i)(j, col_pr_st)
                    gost_stal = arm_arr(i)(j, col_pr_gost_st)
                    qty = arm_arr(i)(j, col_qty)
                    length_pos = arm_arr(i)(j, col_pr_length) / 1000
                    weight_pm = arm_arr(i)(j, col_pr_weight)
                    name_pr = GetShortNameForGOST(arm_arr(i)(j, col_pr_gost_prof))
                    w_pos = Round_w(weight_pm * length_pos * k_zap_total, n_round_w) * qty / nSubPos
                    If InStr(1, name_pr, "Лист") Then
                        naen_plate = SpecMetallPlate(arm_arr(i)(j, col_pr_prof), arm_arr(i)(j, col_pr_naen), length_pos, weight_pm)
                        weight_pm = naen_plate(4)
                        w_pos = Round_w(naen_plate(4) * k_zap_total, n_round_w) * qty / nSubPos
                    End If
                    tkeyd = "Прокат" & n_type & stal & vbLf & gost_stal & gost_prof & prof
                    tkesum_1 = "Прокат" & n_type & stal & vbLf & gost_stal & gost_prof & "Всего"
                End If
                i_col_d = weight_index.Item("col" & tkeyd)
                i_col_s1 = weight_index.Item("col" & tkesum_1)
                i_col_s2 = weight_index.Item("col" & n_type & "Всего")
                pos_out(k, i_col_d) = pos_out(k, i_col_d) + w_pos
                pos_out(k, i_col_s1) = pos_out(k, i_col_s1) + w_pos
                pos_out(k, i_col_s2) = pos_out(k, i_col_s2) + w_pos
                pos_out(k, current_size) = pos_out(k, current_size) + w_pos
            End If
        Next j
    Next i
    If sum_row Then
        pos_out(n_row + sum_row, 1) = "Итого"
        For i = 2 To UBound(pos_out, 2)
            For j = 6 To n_row
                pos_out(n_row + sum_row, i) = pos_out(n_row + sum_row, i) + pos_out(j, i)
            Next
        Next
        If is_bet = True Then
            For i = 1 To UBound(pos_out_bet, 2)
                For j = 6 To n_row
                    pos_out_bet(n_row + sum_row, i) = pos_out_bet(n_row + sum_row, i) + pos_out_bet(j, i)
                Next
            Next
        End If
    End If
    If is_bet = True Then pos_out = ArrayTranspose(ArrayCombine(ArrayTranspose(pos_out), ArrayTranspose(pos_out_bet)))
    For i = 2 To UBound(pos_out, 2)
        For j = 6 To n_row
            If IsEmpty(pos_out(j, i)) Then
                pos_out(j, i) = "-"
            Else
                pos_out(j, i) = Round_w(pos_out(j, i), n_round_wkzh)
            End If
        Next
    Next
    Spec_KZH = pos_out
End Function

Function Spec_POL(ByRef all_data As Variant) As Variant
    out_data = all_data(1)
    rules = all_data(2)
    rules_mod = all_data(3)
    Erase all_data
    If IsEmpty(out_data) Then
        Spec_POL = Empty
        Exit Function
    End If
    isrim = 0
    Set zone = CreateObject("Scripting.Dictionary")
    zone.comparemode = 1
    un_n_zone = ArrayUniqValColumn(out_data, col_s_numb_zone)
    For Each num In un_n_zone
        perim_total = 0
        perim_hole = 0
        free_len = 0
        wall_len = 0
        door_len = 0
        n_wall = 0
        If IsNumeric(num) Then num = CStr(num)
        zone_el = ArraySelectParam(out_data, num, col_s_numb_zone, "ЗОНА", col_s_type)
        If Not IsEmpty(zone_el) Then
            If UBound(zone_el, 1) > 1 Then MsgBox ("Зоны с одинаковыми именами считаются не правильно - " + num)
            perim_total = zone_el(1, col_s_perim_zone)
            perim_hole = zone_el(1, col_s_perimhole_zone)
            free_len = zone_el(1, col_s_freelen_zone)
            wall = ArraySelectParam(out_data, num, col_s_numb_zone, "СТЕНА", col_s_type)
            If Not IsEmpty(wall) Then
                For i = 1 To UBound(wall, 1)
                    door_len = door_len + wall(i, col_s_doorlen_zone)
                    wall_len = wall_len + wall(i, col_s_walllen_zone)
                Next i
            End If
        End If
        zone.Item(num + ";perim_total") = perim_total
        zone.Item(num + ";perim_hole") = perim_hole
        zone.Item(num + ";free_len") = free_len
        zone.Item(num + ";wall_len") = wall_len
        zone.Item(num + ";door_len") = door_len
    Next
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
        un_pol(i) = ConvNum2Txt(un_pol(i))
    Next i
    For j = 1 To n_type_pol
        type_pol = un_pol(j)
        t_pol = ArraySelectParam(pol, type_pol, col_s_type_pol)
        t_un_zone = ArrayUniqValColumn(t_pol, col_s_numb_zone)
        pol_area = 0
        pol_perim_el = 0
        For i = 1 To UBound(t_pol, 1)
            pol_area = pol_area + t_pol(i, col_s_area_pol)
            pol_perim_el = pol_perim_el + t_pol(i, col_s_perim_pol)
        Next i
        perim_total = 0
        perim_hole = 0
        free_len = 0
        wall_len = 0
        door_len = 0
        For Each num In t_un_zone
            If IsNumeric(num) Then num = CStr(num)
            perim_total = zone.Item(num + ";perim_total") + perim_total
            perim_hole = zone.Item(num + ";perim_hole") + perim_hole
            free_len = zone.Item(num + ";free_len") + free_len
            wall_len = zone.Item(num + ";wall_len") + wall_len
            door_len = zone.Item(num + ";door_len") + door_len
        Next
        Select Case type_perim
            Case 1 'Из аксессуаров
                pol_perim = pol_perim_el
            Case 2 'Из длины стен
                pol_perim = wall_len
            Case 3 'Из периметра зоны
                pol_perim = perim_total
        End Select
        If del_dor_perim Then pol_perim = pol_perim - door_len
        If del_freelen_perim Then pol_perim = pol_perim - free_len
        If add_holes_perim Then pol_perim = pol_perim + perim_hole
        t_zone = ""
        For i = 1 To UBound(t_un_zone, 1) - 1
            t_un_zone(i) = Replace(t_un_zone(i), ",", ".")
            t_zone = t_zone + t_un_zone(i) + ", "
        Next i
        t_zone = t_zone + t_un_zone(i)
        pos_out(j, 1) = type_pol
        pos_out(j, 2) = t_zone
        pos_out(j, 3) = Round_w(pol_area * k_zap_total, n_round_area)
        pos_out(j, 4) = Round_w(pol_perim * k_zap_total / 1000, n_round_area)
        If InStr(type_pol, "I") Or InStr(type_pol, "V") Or InStr(type_pol, "X") Then isrim = isrim + 1
    Next j
    'TODO Добавить сортировку римских цифр
    Spec_POL = pos_out
End Function

Function SetKzap()
    tt = ConvTxt2Num(UserForm2.Kzap.Text)
    If IsNumeric(tt) Then
        If tt > 1 And tt < 2 Then
            k_zap_total = tt
        Else
            k_zap_total = 1
        End If
    End If
End Function

Function Spec_Select(ByVal lastfilespec As String, ByVal suffix As String, Optional quiet As Boolean = False) As String
    r = INISet()
    If SpecGetType(lastfilespec) = 7 Then
        nm = Split(lastfilespec, "_")(0) & suffix
    Else
        nm = lastfilespec & suffix
    End If
    r = SheetWriteOption(nm)
    r = SetKzap()
    If Not SheetCheckName(nm) Then
        r = LogWrite(lastfilespec, suffix, "Ошибка имени листа или файла")
        If Not (quiet) Then MsgBox ("Данные отсутвуют")
        Exit Function
    End If
    type_spec = SpecGetType(nm)
    Select Case type_spec
        Case 10
            If Not (quiet) Then MsgBox ("Перейдите на лист _вед и повторите")
            Exit Function
        Case 11
            all_data = VedRead(nm)
        Case 12
            all_data = VedReadPol(nm)
        Case Else
            If IsEmpty(pr_adress) Then r = ReadPrSortament()
            all_data = DataRead(lastfilespec)
    End Select
    If IsEmpty(all_data) Then
        r = LogWrite(lastfilespec, suffix, "Данные отсутвуют")
        If Not (quiet) Then MsgBox ("Данные отсутвуют")
        Exit Function
    End If
    'Ищем файл или лист _разб для разбивки на части
    flag_split = False
    If SheetExist(Split(lastfilespec, "_")(0) & "_разб") Then
        split_data = VedSplitSheet(Split(lastfilespec, "_")(0))
        flag_split = True
    Else
        listFile = GetListFile("*.txt")
        file = ArraySelectParam(listFile, Split(lastfilespec, "_")(0) & "_разб", 1)
        If Not IsEmpty(file) Then
            split_data = VedSplitFile(Split(lastfilespec, "_")(0))
            flag_split = True
        End If
    End If
    pos_out_all = Empty
    Dim pos_zag()
    Select Case type_spec
        Case 1, 2, 3, 13
            If Not (quiet) Then MsgBox ("Коэффицент запаса для объёма, площади и длин " & ConvNum2Txt(k_zap_total))
            pos_out = Spec_AS(all_data, type_spec)
        Case 4
            pos_out = Spec_KM(all_data)
        Case 5
            If Not (quiet) Then MsgBox ("Коэффицент запаса для веса " & ConvNum2Txt(k_zap_total))
            pos_out = Spec_KZH(all_data)
        Case 11
            If Not (quiet) Then MsgBox ("Коэффицент запаса площади отделки -" & ConvNum2Txt(k_zap_total))
            'Проверка возможности разделения на типы (если они заданы)
            If UserForm2.otd_by_type_CB.Value Then
                zone_el = ArraySelectParam(all_data(1), "ЗОНА", col_s_type)
                flag = Empty
                If Not IsEmpty(zone_el) Then
                    For jj = LBound(zone_el, 1) To UBound(zone_el, 1)
                        is_type_otd = zone_el(1, col_s_type_otd)
                        If is_type_otd = 0 Or is_type_otd = "" Then
                            flag = zone_el(1, col_s_numb_zone)
                            jj = UBound(zone_el, 1)
                        End If
                    Next jj
                End If
                If Not IsEmpty(flag) Then
                    r = LogWrite(lastfilespec, suffix, "Тип отделки не задан: " & flag)
                    If Not (quiet) Then MsgBox ("Тип отделки в помещении " & flag & " не задан. Вывожу без типов отделки.")
                End If
            Else
                flag = 1
            End If
            If flag_split Then
                all_data = VedSplitData(all_data, split_data, Split(lastfilespec, "_")(0), suffix)
                For nsheet = 1 To UBound(all_data, 1)
                    nm = all_data(nsheet, 1)
                    nm_data = all_data(nsheet, 2)
                    If IsEmpty(flag) Then
                        pos_out = Spec_VED_GR(nm_data)
                    Else
                        pos_out = Spec_VED(nm_data)
                    End If
                    If delim_by_sheet Then
                        Spec_Select = Spec_OUT(pos_out, nm, suffix, quiet)
                        r = VedWriteLog(nm)
                    Else
                        ReDim pos_zag(1, UBound(pos_out, 2))
                        pos_zag(1, 2) = "@@@" & nm
                        pos_out = ArrayCombine(pos_zag, pos_out)
                        pos_out_all = ArrayCombine(pos_out_all, pos_out)
                    End If
                Next nsheet
            Else
                If IsEmpty(flag) Then
                    pos_out = Spec_VED_GR(all_data)
                Else
                    pos_out = Spec_VED(all_data)
                End If
            End If
        Case 12
            If Not (quiet) Then MsgBox ("Коэффицент запаса площади пола -" & ConvNum2Txt(k_zap_total))
            If flag_split Then
                all_data = VedSplitData(all_data, split_data, Split(lastfilespec, "_")(0), suffix)
                For nsheet = 1 To UBound(all_data, 1)
                    nm = all_data(nsheet, 1)
                    nm_data = all_data(nsheet, 2)
                    pos_out = Spec_POL(nm_data)
                    If delim_by_sheet Then
                        Spec_Select = Spec_OUT(pos_out, nm, suffix, quiet)
                    Else
                        ReDim pos_zag(1, UBound(pos_out, 2))
                        pos_zag(1, 1) = "@@@" & nm
                        pos_out = ArrayCombine(pos_zag, pos_out)
                        pos_out_all = ArrayCombine(pos_out_all, pos_out)
                    End If
                Next nsheet
            Else
                pos_out = Spec_POL(all_data)
            End If
        Case 14
            pos_out = Spec_NRM(all_data)
        Case 20
            pos_out = Spec_WIN(all_data)
    End Select
    If Not IsEmpty(pos_out_all) Then pos_out = pos_out_all
    If flag_split = False Or (delim_by_sheet = False And flag_split = True) Then Spec_Select = Spec_OUT(pos_out, nm, suffix, quiet)
End Function

Function VedAddAreaGR(ByVal area As Double, ByVal mat_fin As String, ByVal type_constr As String, ByVal type_name As String, ByVal mat_draft As String, ByRef rules_mod As Variant, ByRef materials_by_type As Variant, Optional ByVal num As String) As Integer
    If area < 0.001 Then
        VedAddAreaGR = 0
        Exit Function
    End If
    flag_fin = 1
    flag_draft = 1
    'Если есть черновая отделка - запишем её
    If Len(mat_draft) > 1 Then
        'Если в названии черновой отделки стоит = - чистовая отделка не нужна
        If InStr(mat_draft, "=") > 0 Then
            mat_draft = Trim(Left(mat_draft, Len(mat_draft) - 1))
            flag_fin = 0
        End If
    Else
        flag_draft = 0
    End If
    num = Replace(num, ",", ".")
    'Если в названии чистовой отделки стоит --- или УНИВЕРСАЛЬНЫЙ - чистовая отделка не нужна
    If InStr(mat_fin, "--") > 0 Or InStr(mat_fin, "УНИВЕРСАЛЬНОЕ") > 0 Then flag_fin = 0
    If flag_draft Then
        'Черновая отделка с учётом исключений
        all_name_mat = Split(Replace(VedModMat(Replace(mat_fin, fin_str, ""), mat_draft, rules_mod), "@", ";"), ";")
        For Each mat In all_name_mat
            mat = Trim(mat)
            materials_by_type.Item(type_name + type_constr + mat) = materials_by_type.Item(type_name + type_constr + mat) + area
            If InStr(type_constr, ";pot;") > 0 And zonenum_pot Then
                If materials_by_type.Exists(type_name + ";pot_num" + mat) Then
                    materials_by_type.Item(type_name + ";pot_num" + mat) = materials_by_type.Item(type_name + ";pot_num" + mat) + "; " + num
                Else
                    materials_by_type.Item(type_name + ";pot_num" + mat) = num
                End If
            End If
        Next
    End If
    If flag_fin Then
        'Чистовая отделка
        all_name_mat = Split(Replace(mat_fin, "@", ";"), ";")
        For ni = 0 To UBound(all_name_mat)
            mat = Trim(all_name_mat(ni))
            materials_by_type.Item(type_name + type_constr + mat) = materials_by_type.Item(type_name + type_constr + mat) + area
            If InStr(type_constr, ";pot;") > 0 And zonenum_pot Then
                If materials_by_type.Exists(type_name + ";pot_num" + mat) Then
                    materials_by_type.Item(type_name + ";pot_num" + mat) = materials_by_type.Item(type_name + ";pot_num" + mat) + "; " + num
                Else
                    materials_by_type.Item(type_name + ";pot_num" + mat) = num
                End If
            End If
        Next ni
    End If
    VedAddAreaGR = flag_draft + flag_fin
End Function
Function Spec_OUT(ByRef pos_out As Variant, ByVal nm As String, ByVal suffix As String, ByVal quiet As Boolean) As String
    If IsEmpty(pos_out) Then
        r = LogWrite(nm, suffix, "Данные отсутвуют")
        If Not (quiet) Then MsgBox ("Данные отсутвуют")
    End If
    If SheetExist(nm) Then
        Worksheets(nm).Activate
        Worksheets(nm).Cells.Clear
    Else
        ThisWorkbook.Worksheets.Add.Name = nm
    End If
    r = SheetWriteOption(nm)
    Worksheets(nm).Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    r = FormatTable(nm, pos_out)
    r = FormatTable(nm)
    r = LogWrite(nm, suffix, "ОК")
    If inx_on_new And Not (quiet) Then
        r = SheetIndex()
        Worksheets(nm).Activate
    End If
    Spec_OUT = nm
End Function
Function VedAddArea(ByRef zone As Variant, ByRef materials As Variant, ByVal mat_draft As String, ByVal mat_fin As String, ByVal num As String, ByVal area_mat As Double, ByVal rules_mod As Variant, Optional ByVal perim As Double = 0, Optional ByVal h_pan As Double = 0) As Integer
    If UserForm2.separate_material_CB.Value Then
        razd = ";"
        mat_fin = Trim(mat_fin)
    Else
        razd = "&"
        mat_fin = " " + mat_fin
    End If
    mat_fin = Replace(mat_fin, "@", ";a@")
    If Trim(mat_fin) = "0" Then mat_fin = "---"
    mat_draft = VedModMat(Replace(mat_fin, fin_str, ""), mat_draft, rules_mod)
    mat_draft = Trim(mat_draft)
    mat_draft = "b@" & Replace(mat_draft, razd, ";b@")
    mat_draft = Replace(mat_draft, "@ ", "@")
    If InStr(mat_draft, "=") > 0 Then
        name_mat = Array(Trim(Left(mat_draft, Len(mat_draft) - 1)))
    Else
        If InStr(mat_fin, "--") > 0 And razd = "&" Then
            If isErrorNoFin Then
                name_mat = Split((mat_draft & ";a@" & "НЕТ ОТДЕЛКИ"), razd)
            Else
                name_mat = Split((mat_draft), razd)
                
                
            End If
        Else
            name_mat = Split((mat_draft & ";a@" & mat_fin), razd)
        End If
    End If
    flag = 0
    If perim > 0.01 Then
        If zone.Exists(num + "perim;") Then
            zone.Item(num + "perim;") = zone.Item(num + "perim;") + perim
        Else
            zone.Item(num + "perim;") = perim
        End If
    End If
    If area_mat > 0.01 Then
        For Each mat In name_mat
            mat = Trim(mat)
            naen_mat = Trim(Replace(Replace(mat, "b@", ""), "a@", ""))
            If Left(naen_mat, 1) <> "" Then naen_mat = StrConv(Left(naen_mat, 1), vbUpperCase) + Right(naen_mat, Len(naen_mat) - 1)
            If Left(naen_mat, 1) = ";" Then naen_mat = Trim(Right(naen_mat, Len(naen_mat) - 1))
            If naen_mat <> "" And Not IsEmpty(naen_mat) And InStr(naen_mat, "--") = 0 And InStr(naen_mat, "УНИВЕРСАЛЬНОЕ") = 0 Then
                If Not zone.Exists(num) Then
                    Set mat_collect = CreateObject("Scripting.Dictionary")
                    mat_collect.Item(mat) = 1
                    Set zone.Item(num) = mat_collect
                    flag = flag + 1
                Else
                    If Not zone.Item(num).Exists(mat) Then
                        zone.Item(num).Item(mat) = 1
                        flag = flag + 1
                    End If
                End If
                
                If zone.Exists(num + "n;" + mat) Then
                    zone.Item(num + "a;" + mat) = zone.Item(num + "a;" + mat) + area_mat
                Else
                    zone.Item(num + "a;" + mat) = area_mat
                    zone.Item(num + "n;" + mat) = naen_mat
                End If
                
                If h_pan > 0.01 Then
                    If Not zone.Exists(num + "h;" + mat) Then
                        zone.Item(num + "h;" + mat) = h_pan
                    End If
                End If
                If InStr(num, ";pot") > 0 Then
                    mat = mat + ";pot"
                    naen_mat = "Потолки: " + naen_mat
                End If
                If materials.Exists(mat) Then
                    materials.Item(mat + ";a") = materials.Item(mat + ";a") + area_mat
                Else
                    materials.Item(mat) = naen_mat
                    materials.Item(mat + ";a") = materials.Item(mat + ";a") + area_mat
                End If
            End If
        Next
    End If
    VedAddArea = flag
End Function

Function Spec_CONC(ByRef all_data As Variant) As Boolean
    all_bet = ArraySelectParam_2(all_data, t_mat, col_type_el, "?етон?", col_m_naen)
    If IsEmpty(concrsubpos) Then Set concrsubpos = CreateObject("Scripting.Dictionary")
    flag = False
    concrsubpos.Item("bet_qty") = 0
    For Each subpos In pos_data.Item("parent").keys()
        all_bet_subpos = ArraySelectParam_2(all_bet, subpos, col_sub_pos, "?етон?", col_m_naen)
        concrsubpos.Item(subpos & "_bet_qty") = 0
        If Not IsEmpty(all_bet_subpos) Then
            nSubPos = GetNSubpos(subpos, 1)
            n_mat = UBound(all_bet_subpos, 1)
            spec_subpos = SpecMaterial(all_bet_subpos, n_mat, 2, nSubPos)
            For j = 1 To UBound(spec_subpos, 1)
                bet = Trim(Replace(spec_subpos(j, 3), "B", "В"))
                qty = spec_subpos(j, 4)
                concrsubpos.Item(subpos & "_" & bet & "_qty") = qty
                concrsubpos.Item(subpos & "_bet_qty") = concrsubpos.Item(subpos & "_bet_qty") + qty
                concrsubpos.Item("bet_qty") = concrsubpos.Item("bet_qty") + qty
            Next j
            concrsubpos.Item(subpos & "_bet") = ArrayUniqValColumn(spec_subpos, 3)
            flag = True
        End If
    Next
    Spec_CONC = flag
End Function

Function Spec_NRM(ByRef all_data As Variant) As Variant
    UserForm2.qtyOneSubpos_CB.Value = False
    pos_out_arm = Spec_KZH(all_data)
    n_col_arm = UBound(pos_out_arm, 2)
    For i = 1 To UBound(pos_out_arm, 2)
        If InStr(pos_out_arm(2, i), "етон") > 0 Then n_col_arm = n_col_arm - 1
    Next
    If UBound(pos_out_arm, 2) - n_col_arm > 1 Then n_col_arm = n_col_arm - 1
    r = Spec_CONC(all_data)
    If r = False Then
        MsgBox ("В спецификациях не найден бетон")
        Spec_NRM = Empty
        Exit Function
    End If
    subpos = pos_data.Item("parent").keys()
    Dim pos_out_norm: ReDim pos_out_norm(UBound(subpos, 1) + 3, 5)
    n_out = 1
    pos_out_norm(n_out, 1) = "Поз."
    pos_out_norm(n_out, 2) = "Марка бетона"
    pos_out_norm(n_out, 3) = "Объём" & vbLf & "бетона, куб.м."
    pos_out_norm(n_out, 4) = "Вес арматуры, кг."
    pos_out_norm(n_out, 5) = "Расход, кг/куб.м"
    sum_bet = 0: sum_arm = 0
    For Each subpos In pos_data.Item("parent").keys()
        v_bet = 0
        v_arm = 0
        naen_bet = ""
        If concrsubpos.Exists(subpos & "_bet_qty") Then
            If concrsubpos.Exists(subpos & "@бет") Then
                bet_ank = concrsubpos.Item(subpos & "@бет")
                For Each sub_bet In concrsubpos.keys()
                    If InStr(sub_bet, "_") > 0 And Right(sub_bet, 4) = "_qty" And InStr(sub_bet, "bet") = 0 Then
                        subb = Split(sub_bet, "_")
                        If subb(0) = subpos And InStr(subb(1), bet_ank) > 0 Then
                            v_bet = v_bet + concrsubpos.Item(sub_bet)
                            naen_bet = naen_bet & vbLf & subb(1)
                        End If
                    End If
                Next
            Else
                v_bet = concrsubpos.Item(subpos & "_bet_qty")
                For Each nbet In concrsubpos.Item(subpos & "_bet")
                    naen_bet = naen_bet & vbLf & nbet
                Next
            End If
        End If
        v_bet = Round(v_bet, 0)
        If v_bet > 0 Then
            For k = 1 To UBound(pos_out_arm, 1)
                If Left(pos_out_arm(k, 1), Len(subpos)) = subpos Then v_arm = pos_out_arm(k, n_col_arm)
            Next k
        End If
        v_arm = Round(v_arm, 0)
        If v_bet > 0 And v_arm > 0 Then
            n_out = n_out + 1
            pos_out_norm(n_out, 1) = subpos
            pos_out_norm(n_out, 2) = naen_bet
            pos_out_norm(n_out, 3) = v_bet
            pos_out_norm(n_out, 4) = v_arm
            pos_out_norm(n_out, 5) = Round(v_arm / v_bet, 0)
            sum_bet = sum_bet + v_bet
            sum_arm = sum_arm + v_arm
        End If
    Next
    n_out = n_out + 1
    pos_out_norm(n_out, 1) = "Итого"
    pos_out_norm(n_out, 2) = " "
    pos_out_norm(n_out, 3) = sum_bet
    pos_out_norm(n_out, 4) = sum_arm
    pos_out_norm(n_out, 5) = Round(sum_arm / sum_bet, 0)
    diff = concrsubpos.Item("bet_qty") - sum_bet
    If n_out <> UBound(pos_out_norm, 1) Then pos_out_norm = ArrayRedim(pos_out_norm, n_out)
    Spec_NRM = pos_out_norm
End Function

Function Spec_VED_GR(ByRef all_data As Variant) As Variant
    out_data = all_data(1)
    rules = all_data(2)
    rules_mod = all_data(3)
    Erase all_data
    If IsEmpty(out_data) Or IsEmpty(rules) Then
        Spec_VED_GR = Empty
        Exit Function
    End If
    is_column = False
    is_pan = False
    tot_area_wall = 0
    tot_area_column = 0
    tot_perim_zone = 0
    tot_area_wall_zone = 0
    tot_area_pan = 0
    Set materials_by_type = CreateObject("Scripting.Dictionary")
    Set materials = CreateObject("Scripting.Dictionary")
    materials_by_type.comparemode = 1
    materials.comparemode = 1
    spec_type = 1 'И лестницы, и полы
    If UBound(out_data, 2) < col_s_tipverh_l Then spec_type = 2 'Без лестниц
    If UBound(out_data, 2) < col_s_type_el Then spec_type = 3 'Только зоны
    '------- Предварительная выборка элементов --------------------
    zones_el_all = ArraySelectParam(out_data, "ЗОНА", col_s_type)
    walls_el_all = ArraySelectParam(out_data, "СТЕНА", col_s_type)
    pots_el_all = ArraySelectParam(out_data, "Потолок", col_s_type_el)
    pols_el_all = ArraySelectParam(out_data, "Пол", col_s_type_el)
    '--------------------------------------------------------------
    un_type_otd = ArrayUniqValColumn(zones_el_all, col_s_type_otd)
    materials_by_type.Item("list") = un_type_otd
    For Each type_name In un_type_otd
        If InStr(type_name, "Без отделки") = 0 Then
            If IsNumeric(type_name) Then type_name = CStr(type_name)
            zone_bytype_el = ArraySelectParam(zones_el_all, type_name, col_s_type_otd) 'Список зон с этим типом отделки
            un_n_zone_type = ArrayUniqValColumn(zone_bytype_el, col_s_numb_zone) 'Список номеров зон
            materials_by_type.Item(type_name + ";zone") = un_n_zone_type
            '------- Предварительная выборка элементов --------------------
            zones_el = ArraySelectParam_2(zones_el_all, un_n_zone_type, col_s_numb_zone)
            walls_el = ArraySelectParam_2(walls_el_all, un_n_zone_type, col_s_numb_zone)
            pots_el = ArraySelectParam_2(pots_el_all, un_n_zone_type, col_s_numb_zone)
            pols_el = ArraySelectParam_2(pols_el_all, un_n_zone_type, col_s_numb_zone)
            '--------------------------------------------------------------
            For Each num In un_n_zone_type
                ' Теперь для каждой зоны с этим типом отделки считаем всё что можем
                is_error = 0
                If IsNumeric(num) Then num = CStr(num)
                zone_el = ArraySelectParam(zone_bytype_el, num, col_s_numb_zone)
                If Not IsEmpty(zone_el) Then
                    ' --- Финишная отделка для данного типа ----
                    fin_pot = fin_str + Replace(zone_el(1, col_s_mpot_zone), "@", "; ")
                    fin_pan = fin_str + Replace(zone_el(1, col_s_mpan_zone), "@", "; ")
                    fin_wall = fin_str + Replace(zone_el(1, col_s_mwall_zone), "@", "; ")
                    fin_column = fin_str + Replace(zone_el(1, col_s_mcolumn_zone), "@", "; ")
                    ' ---
                    ' Если в имени финишного материала стоят знаки <>
                    ' То все колонны переходят в стены
                    ' Отделка панелей для них не выполняется
                    ' ---
                    If InStr(fin_column, "<>") > 0 Then
                        column_is_wall = True
                        fin_column = Replace(fin_column, "<>", "")
                    Else
                        column_is_wall = False
                    End If
                    ' ------------------------------------------
                    area_total = zone_el(1, col_s_area_zone)
                    perim_total = zone_el(1, col_s_perim_zone) / 1000
                    perim_hole = zone_el(1, col_s_perimhole_zone) / 1000
                    h_zone = zone_el(1, col_s_h_zone) / 1000
                    free_len = zone_el(1, col_s_freelen_zone) / 1000
                    h_pan = zone_el(1, col_s_hpan_zone) / 1000
                    If h_pan < 0.01 And h_pan > 0 Then
                        r = LogWrite("Проверьте высоту панелей, должна быть в мм. - " + CStr(h_pan), "Ошибка", num)
                        is_error = is_error + 1
                        h_pan = h_pan * 1000
                    End If
                    If h_pan > 0.01 Then is_pan = True
                    If UBound(zone_el, 1) > 1 Then
                        r = LogWrite("Одинаковых зон - " + CStr(UBound(zone_el, 1)), "Ошибка", num)
                        is_error = is_error + 1
                        For nzone = 2 To UBound(zone_el, 1)
                            area_total = area_total + zone_el(nzone, col_s_area_zone)
                            perim_total = perim_total + zone_el(nzone, col_s_perim_zone) / 1000
                            perim_hole = perim_hole + zone_el(nzone, col_s_perimhole_zone) / 1000
                            free_len = free_len + zone_el(nzone, col_s_freelen_zone) / 1000
                        Next nzone
                    End If
                    tot_area_wall_zone = tot_area_wall_zone + (perim_total - free_len) * h_zone
                    tot_area_zone = tot_area_zone + (perim_total - free_len) * h_zone
                    tot_perim_zone = tot_perim_zone + perim_total
                    ' --- Длины стен и дверей ---
                    wall = ArraySelectParam(walls_el, num, col_s_numb_zone)
                    wall_len = 0
                    door_len = 0
                    If Not IsEmpty(wall) Then
                        For i = 1 To UBound(wall, 1)
                            door_len = door_len + wall(i, col_s_doorlen_zone) / 1000
                            wall_len = wall_len + wall(i, col_s_walllen_zone) / 1000
                        Next i
                    End If
                    ' -----------------------
                    
                    '----------------------
                    '        КОЛОННЫ
                    '----------------------
                    'Добавить возможность выбора наличия отверстий
                    column_perim_in_wall = perim_total - wall_len - free_len + perim_hole * (hole_in_zone = True)
                    column_perim = perim_hole
                    column_perim_total = column_perim + column_perim_in_wall
                    column_pan_area = column_perim_total * h_pan
                    column_area = column_perim_total * (h_zone - h_pan)
                    tot_area_column = tot_area_column + column_area
                    tot_area_pan = tot_area_pan + column_pan_area
                    colm = VedNameMat("Колонны", "ЖБ", rules)
                    If column_is_wall Then
                        'Стены
                        r = VedAddAreaGR(column_area + column_pan_area, fin_column, ";wall;", type_name, "", rules_mod, materials_by_type, num)
                    Else
                        If column_area > 0.01 Then is_column = True
                        If column_pan_area > 0.01 Then is_pan = True
                        colm = VedNameMat("Колонны", "ЖБ", rules)
                        'Колонны
                        r = VedAddAreaGR(column_area, fin_column, ";column;", type_name, colm, rules_mod, materials_by_type, num)
                        'Панели на колоннах
                        r = VedAddAreaGR(column_pan_area, fin_pan, ";pan;", type_name, colm, rules_mod, materials_by_type, num)
                    End If
                    '----------------------
                    '        СТЕНЫ
                    '----------------------
                    wall_area_zone = 0
                    un_wall = ArrayUniqValColumn(wall, col_s_mat_wall)
                    If Not IsEmpty(un_wall) Then
                        For Each w In un_wall
                            wall_len = 0
                            wall_area = 0
                            wall_c_len = 0
                            wall_c_area = 0
                            pan_area = 0
                            pan_c_area = 0
                            twall_area = 0
                            tpan_area = 0
                            wall_by_key = ArraySelectParam(wall, w, col_s_mat_wall)
                            For i = 1 To UBound(wall_by_key, 1)
                                twall_area = wall_by_key(i, col_s_area_wall)
                                If twall_area > 0 Then
                                    tdoor_len = wall_by_key(i, col_s_doorlen_zone) / 1000
                                    twall_len = wall_by_key(i, col_s_walllen_zone) / 1000
                                    If twall_len > tdoor_len Then th_wall = twall_area / (twall_len - tdoor_len)
                                    If th_wall > h_pan Then
                                        tpan_area = (twall_len - tdoor_len) * h_pan
                                        twall_area = twall_area - tpan_area
                                    Else
                                        If twall_len > tdoor_len Then
                                            tpan_area = twall_area
                                            twall_area = 0
                                            r = LogWrite("Панели на всю высоту стен? " + CStr(h_pan), "", num)
                                        Else
                                            tpan_area = 0
                                            twall_area = twall_area
                                            r = LogWrite("Стена полностью скрыта дверью? " + CStr(h_pan), "", num)
                                        End If
                                    End If
                                    If InStr(wall_by_key(i, col_s_layer), "Колонн") > 0 Then
                                        wall_c_area = wall_c_area + twall_area
                                        pan_c_area = pan_c_area + tpan_area
                                    Else
                                        wall_area = wall_area + twall_area
                                        pan_area = pan_area + tpan_area
                                    End If
                                End If
                            Next i
                            wall_area_zone = wall_area_zone + wall_area + wall_c_area + pan_c_area + pan_area
                            tot_area_wall = tot_area_wall + wall_area
                            tot_area_column = tot_area_column + wall_c_area
                            tot_area_pan = tot_area_pan + pan_area + pan_c_area
                            'Стены
                            r = VedAddAreaGR(wall_area, fin_wall, ";wall;", type_name, w, rules_mod, materials_by_type, num)
                            'Колонны, смоделированные стенами
                            r = VedAddAreaGR(wall_c_area, fin_column, ";column;", type_name, w, rules_mod, materials_by_type, num)
                            'Панели
                            r = VedAddAreaGR(pan_area, fin_pan, ";pan;", type_name, w, rules_mod, materials_by_type, num)
                            'Панели на колоннах, смоделированных стенами
                            r = VedAddAreaGR(pan_c_area, fin_pan, ";pan;", type_name, w, rules_mod, materials_by_type, num)
                        Next
                    End If
                    If h_pan > 0.001 And ((pan_c_area > 0.001) Or (column_pan_area > 0.001) Or (pan_area > 0.001)) Then
                        materials_by_type.Item(type_name + ";panh") = CStr(h_pan)
                    End If
                    If wall_area_zone < 0.1 Then
                        r = LogWrite("Почти нет стен" & num, "Ошибка", wall_area_zone)
                        is_error = is_error + 1
                    End If
                    
                    '----------------------
                    '        ПОТОЛКИ
                    '----------------------
                    diff_area_pot = 0
                    area_total_pot = 0
                    If spec_type < 3 Then
                        pot = ArraySelectParam(pots_el, num, col_s_numb_zone, "Потолок")
                        un_pot = ArrayUniqValColumn(pot, col_s_type_pol)
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
                                r = VedAddAreaGR(pot_area, fin_pot, ";pot;", type_name, p, rules_mod, materials_by_type, num)
                            Next
                            materials_by_type.Item(type_name + ";pot_perim;") = materials_by_type.Item(type_name + ";pot_perim;") + pot_perim
                            diff_area_pot = area_total - area_total_pot
                            diff_area_pot = Round(diff_area_pot, 4)
                            'Если разница площади и подвесного потолка больше 1-го квадрата - добавим финишной отделки на разницу.
'                            If diff_area_pot > 1 Then
'                                r = VedAddAreaGR(diff_area_pot, fin_pot, ";pot;", type_name, "", rules_mod, materials_by_type, num)
'                                r = LogWrite("Добавлена окраска" & num, "Ошибка", diff_area_pot)
'                                is_error = is_error + 1
'                            End If
                            If Abs(diff_area_pot) > 1 Then
                                r = LogWrite("Разница площади помещения и потолка в помещении " & num, "Ошибка", diff_area_pot)
                                is_error = is_error + 1
                            End If
                        Else
                            r = VedAddAreaGR(area_total, fin_pot, ";pot;", type_name, "", rules_mod, materials_by_type, num)
                            materials_by_type.Item(type_name + ";pot_perim;") = materials_by_type.Item(type_name + ";pot_perim;") + perim_total
                        End If
                    Else
                        r = VedAddAreaGR(area_total, fin_pot, ";pot;", type_name, "", rules_mod, materials_by_type, num)
                        materials_by_type.Item(type_name + ";pot_perim;") = materials_by_type.Item(type_name + ";pot_perim;") + perim_total
                    End If
                    '----------------------
                    '        ПОЛЫ
                    '----------------------
                    area_total_pol = 0
                    diff_area_pol = 0
                    If spec_type < 3 Then
                        pol = ArraySelectParam(pols_el, num, col_s_numb_zone)
                        If Not IsEmpty(pol) Then
                            n_pol = UBound(pol, 1)
                            For i = 1 To n_pol
                                area_total_pol = area_total_pol + pol(i, col_s_area_pol)
                            Next
                            diff_area_pol = area_total - area_total_pol
                            diff_area_pol = Round(diff_area_pol, 4)
                            If Abs(diff_area_pol) > 1 Then
                                r = LogWrite("Разница площади помещения и пола в помещении " & num, "Ошибка", diff_area_pol)
                                is_error = is_error + 1
                            End If
                        End If
                    End If
                    If is_error > 0 Then
                        zone_error.Item(num + "_err") = zone_error.Item(num + "_err") + is_error
                        zone_error.Item(num + "_h_zone") = h_zone
                        zone_error.Item(num + "_h_pan") = h_pan
                        zone_error.Item(num + "_area_total") = area_total
                        zone_error.Item(num + "_area_total_pot") = area_total_pot
                        If Abs(diff_area_pot) > 1 Then
                            zone_error.Item(num + "_pot_diff") = diff_area_pot
                        Else
                            zone_error.Item(num + "_pot_diff") = ""
                        End If
                        If Abs(diff_area_pol) > 1 Then
                            zone_error.Item(num + "_pol_diff") = diff_area_pol
                        Else
                            zone_error.Item(num + "_pol_diff") = ""
                        End If
                        zone_error.Item(num + "_area_total_pol") = area_total_pol
                        zone_error.Item(num + "_column_area") = column_area
                        zone_error.Item(num + "_wall_area_zone") = wall_area_zone
                        is_error = 0
                    End If
                Else
                    MsgBox ("Номер зоны в элементе записан не правильно - " + num)
                    Spec_VED_GR = Empty
                    Exit Function
                End If
            Next
            Dim all_mat_pot(): ReDim all_mat_pot(1): npot = 0: len_find_pot = Len(type_name + ";pot;")
            Dim all_mat_wall(): ReDim all_mat_wall(1): nwall = 0: len_find_wall = Len(type_name + ";wall;")
            Dim all_mat_column(): ReDim all_mat_column(1): ncolumn = 0: len_find_column = Len(type_name + ";column;")
            Dim all_mat_pan(): ReDim all_mat_pan(1): npan = 0: len_find_pan = Len(type_name + ";pan;")
            For Each mat In materials_by_type.keys()
                If InStr(mat, type_name) > 0 Then
                    len_mat = Len(mat)
                    If (Left(mat, len_find_pot) = type_name + ";pot;") Then
                        npot = npot + 1
                        ReDim Preserve all_mat_pot(npot)
                        all_mat_pot(npot) = Right(mat, len_mat - len_find_pot)
                    End If
    
                    If (Left(mat, len_find_wall) = type_name + ";wall;") Then
                        nwall = nwall + 1
                        ReDim Preserve all_mat_wall(nwall)
                        all_mat_wall(nwall) = Right(mat, len_mat - len_find_wall)
                    End If
    
                    If (Left(mat, len_find_column) = type_name + ";column;") Then
                        ncolumn = ncolumn + 1
                        ReDim Preserve all_mat_column(ncolumn)
                        all_mat_column(ncolumn) = Right(mat, len_mat - len_find_column)
                    End If
    
                    If (Left(mat, len_find_pan) = type_name + ";pan;") Then
                        npan = npan + 1
                        ReDim Preserve all_mat_pan(npan)
                        all_mat_pan(npan) = Right(mat, len_mat - len_find_pan)
                    End If
                End If
            Next
            If npot > 0 Then
                materials_by_type.Item(type_name + ";pot") = ArrayUniqValColumn(all_mat_pot, 1)
'                If zonenum_pot Then
'                    For Each mat In all_mat_pot
'                        mat_find = Array("?" + Replace(mat, fin_str, "") + "?")
'                        zone_by_mat = ArraySelectParam_2(zones_el, mat_find, col_s_mpot_zone)
'                        pot_by_mat = ArraySelectParam_2(pots_el, mat_find, col_s_type_pol)
'                        un_num_zone = ArrayUniqValColumn(zone_by_mat, col_s_numb_zone)
'                        un_num_pot = ArrayUniqValColumn(pot_by_mat, col_s_numb_zone)
'                        un_num = ArrayUniqValColumn(ArrayCombine(un_num_zone, un_num_pot), 1)
'                        materials_by_type.Item(type_name + ";pot_num" + mat) = un_num
'                        hh = 1
'                    Next mat
'                End If
            Else
                materials_by_type.Item(type_name + ";pot") = Empty
            End If
            If nwall > 0 Then
                materials_by_type.Item(type_name + ";wall") = ArrayUniqValColumn(all_mat_wall, 1)
            Else
                materials_by_type.Item(type_name + ";wall") = Empty
            End If
            If ncolumn > 0 Then
                materials_by_type.Item(type_name + ";column") = ArrayUniqValColumn(all_mat_column, 1)
            Else
                materials_by_type.Item(type_name + ";column") = Empty
            End If
            If npan > 0 Then
                materials_by_type.Item(type_name + ";pan") = ArrayUniqValColumn(all_mat_pan, 1)
            Else
                materials_by_type.Item(type_name + ";pan") = Empty
            End If
        End If
    Next
    n_col_out = 7
    If is_pan Then n_col_out = n_col_out + 3
    If is_column Then n_col_out = n_col_out + 2
    Dim pos_out: ReDim pos_out(3400, n_col_out)
    pos_out(1, 1) = "Тип"
    pos_out(1, 2) = "Номера помещений"
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
    un_type_otd = materials_by_type.Item("list")
    sum_pot = 0
    sum_wall = 0
    sum_column = 0
    sum_pan = 0
    For Each type_name In un_type_otd
        If InStr(type_name, "Без отделки") = 0 Then
            n_row_p = n_row
            n_row_w = n_row
            n_row_c = n_row
            n_row_pan = n_row
            pos_out(n_row, 1) = type_name
    'ПОТОЛКИ
            mat_list = materials_by_type.Item(type_name + ";pot")
            If Not IsEmpty(mat_list) Then
                For Each mat In mat_list
                    area = Round_w(materials_by_type.Item(type_name + ";pot;" + mat) * k_zap_total, n_round_area)
                    If area > 0.001 Then
                        If InStr(mat, fin_str) > 0 Then
                            sum_pot = sum_pot + area
                            materials.Item("@Потолок: " + mat) = materials.Item("@Потолок: " + mat) + area
                        Else
                            materials.Item("Потолок: " + mat) = materials.Item("Потолок: " + mat) + area
                        End If
                        num_zone = ""
                        If Not IsEmpty(materials_by_type.Item(type_name + ";pot_num" + mat)) And zonenum_pot Then
                            num_zone = materials_by_type.Item(type_name + ";pot_num" + mat)
                            pos_out(n_row_p, 2) = num_zone
                        End If
                        pos_out(n_row_p, 3) = Replace(mat, fin_str, "")
                        pos_out(n_row_p, 4) = area
                        n_row_p = n_row_p + 1
                    End If
                Next
            Else
                pos_out(n_row_p, 3) = "-"
                pos_out(n_row_p, 4) = "-"
                n_row_p = n_row_p + 1
            End If
            If zonenum_pot = False Or IsEmpty(mat_list) Then
                num_zone = ""
                For Each num In materials_by_type.Item(type_name + ";zone")
                    If IsNumeric(num) Then num = CStr(num)
                    num = Replace(num, ",", ".")
                    If num_zone = "" Then
                        num_zone = num
                    Else
                        num_zone = num_zone + ", " + num
                    End If
                Next
                pos_out(n_row, 2) = num_zone
            End If
    'СТЕНЫ
            mat_list = materials_by_type.Item(type_name + ";wall")
            If Not IsEmpty(mat_list) Then
                For Each mat In mat_list
                    area = materials_by_type.Item(type_name + ";wall;" + mat)
                    If area > 0.001 Then
                        If InStr(mat, fin_str) > 0 Then sum_wall = sum_wall + Round_w(area * k_zap_total, n_round_area)
                        materials.Item(mat) = materials.Item(mat) + Round_w(area * k_zap_total, n_round_area)
                        pos_out(n_row_w, 5) = Replace(mat, fin_str, "")
                        pos_out(n_row_w, 6) = Round_w(area * k_zap_total, n_round_area)
                        n_row_w = n_row_w + 1
                    End If
                Next
            Else
                pos_out(n_row_w, 5) = "-"
                pos_out(n_row_w, 6) = "-"
                n_row_w = n_row_w + 1
            End If
    'КОЛОННЫ
            If is_column Then
                mat_list = materials_by_type.Item(type_name + ";column")
                If Not IsEmpty(mat_list) Then
                    For Each mat In mat_list
                        area = materials_by_type.Item(type_name + ";column;" + mat)
                        If area > 0.001 Then
                            If InStr(mat, fin_str) > 0 Then sum_column = sum_column + Round_w(area * k_zap_total, n_round_area)
                            materials.Item(mat) = materials.Item(mat) + Round_w(area * k_zap_total, n_round_area)
                            pos_out(n_row_c, 7) = Replace(mat, fin_str, "")
                            pos_out(n_row_c, 8) = Round_w(area * k_zap_total, n_round_area)
                            n_row_c = n_row_c + 1
                        End If
                    Next
                                Else
                                        pos_out(n_row_c, 7) = "-"
                                        pos_out(n_row_c, 8) = "-"
                                        n_row_c = n_row_c + 1
                                End If
            End If
    'ПАНЕЛИ
            If is_pan Then
                mat_list = materials_by_type.Item(type_name + ";pan")
                If Not IsEmpty(mat_list) Then
                    For Each mat In mat_list
                        area = materials_by_type.Item(type_name + ";pan;" + mat)
                        If area > 0.001 Then
                            If InStr(mat, fin_str) > 0 Then sum_pan = sum_pan + Round_w(area * k_zap_total, n_round_area)
                            materials.Item(mat) = materials.Item(mat) + Round_w(area * k_zap_total, n_round_area)
                            pos_out(n_row_pan, col_end + 1) = Replace(mat, fin_str, "")
                            pos_out(n_row_pan, col_end + 2) = Round_w(area * k_zap_total, n_round_area)
                            pos_out(n_row_pan, col_end + 3) = materials_by_type.Item(type_name + ";panh")
                            n_row_pan = n_row_pan + 1
                        End If
                    Next
                Else
                    pos_out(n_row_pan, col_end + 1) = "-"
                    pos_out(n_row_pan, col_end + 2) = "-"
                    pos_out(n_row_pan, col_end + 3) = "-"
                    n_row_pan = n_row_pan + 1
                End If
            End If
            If show_perim Then pos_out(n_row, n_col_out) = "Lперим=" + CStr(Round_w(materials_by_type.Item(type_name + ";pot_perim;") * k_zap_total, n_round_area)) + "п.м."
            n_row = Application.WorksheetFunction.Max(n_row_p, n_row_w, n_row_c, n_row_pan)
        End If
    Next
    If (show_surf_area Or show_mat_area) And delim_by_sheet = True Then n_row = n_row + 1
    If show_surf_area And delim_by_sheet = True Then
        nn_col = 5
        If Not show_mat_area Then nn_col = 1
        pos_out(n_row, nn_col) = "Общяя площадь поверхности, кв.м."
        pos_out(n_row + 1, nn_col) = "Потолки"
        pos_out(n_row + 2, nn_col) = "Стены(за вычетом панелей)"
        pos_out(n_row + 3, nn_col) = "Колонны"
        pos_out(n_row + 4, nn_col) = "Низ стен/колонн"
        pos_out(n_row + 1, nn_col + 3) = sum_pot
        pos_out(n_row + 2, nn_col + 3) = sum_wall
        pos_out(n_row + 3, nn_col + 3) = sum_column
        pos_out(n_row + 4, nn_col + 3) = sum_pan
    End If
    If show_mat_area And delim_by_sheet = True Then
        pos_out(n_row, 1) = "Общяя площадь отделки, кв.м."
        For Each type_mat In Array("@Потолок: ", "Потолок: ", fin_str, "@@@")
            len_type_mat = Len(type_mat)
            For Each mat In materials.keys()
                If Len(mat) > 2 And (Left(mat, len_type_mat) = type_mat Or (type_mat = "@@@" And InStr(mat, "@") = 0 And InStr(mat, "Потолок: ") = 0 And InStr(mat, fin_str) = 0)) Then
                    n_row = n_row + 1
                    pos_out(n_row, 1) = Replace(Replace(mat, "@", ""), fin_str, "")
                    pos_out(n_row, 4) = materials.Item(mat)
                End If
            Next
        Next
    End If
    If Not show_mat_area And Not show_mat_area Then n_row = n_row + 1
    pos_out = ArrayRedim(pos_out, n_row)
    r = LogWrite("Ведомость отделки", "ИТОГ", "'====")
    r = LogWrite("Ведомость отделки", "Потолки", CStr(sum_pot))
    r = LogWrite("Ведомость отделки", "Стены", CStr(sum_wall))
    r = LogWrite("Ведомость отделки", "Колонны", CStr(sum_column))
    r = LogWrite("Ведомость отделки", "Панели", CStr(sum_pan))
    r = LogWrite("Ведомость отделки", "КОНЕЦ", "'====")
    Spec_VED_GR = pos_out
End Function

Function Spec_VED(ByRef all_data As Variant) As Variant
    out_data = all_data(1)
    rules = all_data(2)
    rules_mod = all_data(3)
    Erase all_data
    If IsEmpty(out_data) Or IsEmpty(rules) Then
        Spec_VED = Empty
        Exit Function
    End If
    Set zone = CreateObject("Scripting.Dictionary")
    Set materials = CreateObject("Scripting.Dictionary")
    zone.comparemode = 1
    materials.comparemode = 1
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
        un_n_zone(i) = ConvNum2Txt(un_n_zone(i))
    Next i
    spec_type = 1 'И лестницы, и полы
    If UBound(out_data, 2) < col_s_tipverh_l Then spec_type = 2 'Без лестниц
    If UBound(out_data, 2) < col_s_type_el Then spec_type = 3 'Только зоны
    For Each num In un_n_zone
        n_row_pot = 0
        n_row_w = 0
        n_row_pn = 0
        n_row_c = 0
        If IsNumeric(num) Then num = CStr(num)
        zone_el = ArraySelectParam(out_data, num, col_s_numb_zone, "ЗОНА", col_s_type)
        If Not IsEmpty(zone_el) Then
            zone.Item(num + ";name") = zone_el(1, col_s_name_zone)
            area_total = zone_el(1, col_s_area_zone)
            perim_total = zone_el(1, col_s_perim_zone) / 1000
            perim_hole = zone_el(1, col_s_perimhole_zone) / 1000
            h_zone = zone_el(1, col_s_h_zone) / 1000
            free_len = zone_el(1, col_s_freelen_zone) / 1000
            h_pan = zone_el(1, col_s_hpan_zone) / 1000
            If h_pan < 0.01 And h_pan > 0 Then
                r = LogWrite("Проверьте высоту панелей, должна быть в мм. - " + CStr(h_pan), "Ошибка", num)
                h_pan = h_pan * 1000
            End If
            If UBound(zone_el, 1) > 1 Then
                r = LogWrite("Одинаковых зон - " + CStr(UBound(zone_el, 1)), "Ошибка", num)
                For nzone = 2 To UBound(zone_el, 1)
                    area_total = area_total + zone_el(nzone, col_s_area_zone)
                    perim_total = perim_total + zone_el(nzone, col_s_perim_zone) / 1000
                    perim_hole = perim_hole + zone_el(nzone, col_s_perimhole_zone) / 1000
                    free_len = free_len + zone_el(nzone, col_s_freelen_zone) / 1000
                Next nzone
            End If
            fin_pot = fin_str + zone_el(1, col_s_mpot_zone)
            fin_pan = fin_str + zone_el(1, col_s_mpan_zone)
            fin_wall = fin_str + zone_el(1, col_s_mwall_zone)
            fin_column = fin_str + zone_el(1, col_s_mcolumn_zone)
            wall = ArraySelectParam(out_data, num, col_s_numb_zone, "СТЕНА", col_s_type)
            wall_len = 0
            door_len = 0
            If Not IsEmpty(wall) Then
                For i = 1 To UBound(wall, 1)
                    door_len = door_len + wall(i, col_s_doorlen_zone) / 1000
                    wall_len = wall_len + wall(i, col_s_walllen_zone) / 1000
                Next i
            End If
            '----------------------
            '        КОЛОННЫ
            '----------------------
            'Добавить возможность выбора наличия отверстий
            column_perim_in_wall = perim_total - wall_len - free_len + perim_hole * (hole_in_zone = True)
            column_perim = perim_hole
            column_perim_total = column_perim + column_perim_in_wall
            column_pan_area = column_perim_total * h_pan
            column_pan_area = Round(column_pan_area * 10, 0) / 10
            column_area = column_perim_total * (h_zone - h_pan)
            column_area = Round(column_area * 10, 0) / 10
            name_mat_column = "ЖБ"
            colmn = VedNameMat("Колонны", name_mat_column, rules)
            If column_area > 0.1 Then
                n_row_c = n_row_c + VedAddArea(zone, materials, colmn, fin_column, num + ";c", column_area, rules_mod)
                is_column = True
            End If
            If column_pan_area > 0.1 Then n_row_pn = n_row_pn + VedAddArea(zone, materials, colmn, fin_pan, num + ";pn", column_pan_area, rules_mod, 0, h_pan)
            '----------------------
            '        СТЕНЫ
            '----------------------
            un_wall = ArrayUniqValColumn(wall, col_s_mat_wall)
            If Not IsEmpty(un_wall) Then
                For Each w In un_wall
                    wall_by_key = ArraySelectParam(wall, w, col_s_mat_wall)
                    wall_len = 0
                    wall_area = 0
                    wall_c_len = 0
                    wall_c_area = 0
                    pan_area = 0
                    pan_c_area = 0
                    twall_area = 0
                    tpan_area = 0
                    For i = 1 To UBound(wall_by_key, 1)
                        twall_area = wall_by_key(i, col_s_area_wall)
                        If twall_area > 0 Then
                            tdoor_len = wall_by_key(i, col_s_doorlen_zone) / 1000
                            twall_len = wall_by_key(i, col_s_walllen_zone) / 1000
                            If twall_len > tdoor_len Then th_wall = twall_area / (twall_len - tdoor_len)
                            If th_wall > h_pan Then
                                tpan_area = (twall_len - tdoor_len) * h_pan
                                twall_area = twall_area - tpan_area
                            Else
                                If twall_len > tdoor_len Then
                                    tpan_area = twall_area
                                    twall_area = 0
                                    r = LogWrite("Панели на всю высоту стен? " + CStr(Round(th_wall, 2)), "", num)
                                Else
                                    tpan_area = 0
                                    twall_area = twall_area
                                    r = LogWrite("Стена полностью скрыта дверью? " + CStr(Round(twall_len, 2)), "", num)
                                End If
                            End If
                            If InStr(wall_by_key(i, col_s_layer), "Колонн") > 0 Then
                                wall_c_area = wall_c_area + twall_area
                                pan_c_area = pan_c_area + tpan_area
                            Else
                                wall_area = wall_area + twall_area
                                pan_area = pan_area + tpan_area
                            End If
                        End If
                    Next i
                    result = VedAddArea(zone, materials, w, fin_column, num + ";c", wall_c_area, rules_mod)
                    n_row_c = n_row_c + result
                    result = VedAddArea(zone, materials, w, fin_wall, num + ";w", wall_area, rules_mod)
                    n_row_w = n_row_w + result
                    result = VedAddArea(zone, materials, w, fin_pan, num + ";pn", pan_area + pan_c_area, rules_mod, 0, h_pan)
                    n_row_pn = n_row_pn + result
                Next
            Else
                result = VedAddArea(zone, materials, "", fin_wall, num + ";w", wall_area, rules_mod)
                n_row_w = n_row_w + result
            End If
            If wall_c_area > 0.1 Then is_column = True
            If h_pan > 0 Then is_pan = True

            '----------------------
            '        ПОТОЛКИ
            '----------------------
            If spec_type < 3 Then
                area_total_pot = 0
                pot = ArraySelectParam(out_data, num, col_s_numb_zone, "Потолок", col_s_type_el)
                un_pot = ArrayUniqValColumn(pot, col_s_type_pol)
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
                        n_row_pot = n_row_pot + VedAddArea(zone, materials, p, fin_pot, num + ";pot", pot_area, rules_mod, pot_perim)
                    Next
                    diff_area_pot = area_total - area_total_pot
                    If Abs(diff_area_pot) > 1 Then
                        r = LogWrite("Разница площади помещения " & area_total & " и потолка " & area_total_pot & " в помещении " & num, "Ошибка", diff_area_pot)
                    End If
                Else
                    n_row_pot = n_row_pot + VedAddArea(zone, materials, "", fin_pot, num + ";pot", area_total, rules_mod, perim_total)
                End If
            Else
                n_row_pot = n_row_pot + VedAddArea(zone, materials, "", fin_pot, num + ";pot", area_total, rules_mod, perim_total)
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
            n_row_tot = n_row_tot + Application.WorksheetFunction.Max(n_row_pot, n_row_w, n_row_pn, n_row_c)
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
    n_un_mat = (materials.Count / 2)
    If (n_un_mat - Int(n_un_mat)) <> 0 Then MsgBox ("Ошибка записи в словарь")
    Dim pos_out: ReDim pos_out(3 + n_row_tot + n_un_mat + 60, n_col_out)
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
    n_row = 2
    sum_pot = 0
    sum_wall = 0
    sum_column = 0
    sum_pan = 0
    For Each num In un_n_zone
        n_row_p = n_row
        n_row_w = n_row
        n_row_c = n_row
        n_row_pan = n_row
        pos_out(n_row + 1, 1) = "'" + Replace(num, ",", ".")
        pos_out(n_row + 1, 2) = zone.Item(num + ";name")
        pos_out(n_row + 1, n_col_out) = "Lперим=" + CStr(Round_w(zone.Item(num + ";potperim;") * k_zap_total, n_round_area)) + "п.м."
        '-- ПОТОЛКИ ---
        If Not zone.Exists(num + ";pot") Then
            pot = Empty
        Else
            pot = ArraySort(zone.Item(num + ";pot").keys())
        End If
        If Not IsEmpty(pot) Then
            For Each p In pot
                n_row_p = n_row_p + 1
                mat = zone.Item(num + ";potn;" + p)
                area = Round_w(zone.Item(num + ";pota;" + p) * k_zap_total, n_round_area)
                If InStr(mat, fin_str) > 0 Then sum_pot = sum_pot + area
                pos_out(n_row_p, 3) = Replace(mat, fin_str, "")
                pos_out(n_row_p, 4) = area
                summ_area_pot = summ_area_pot + pos_out(n_row_p, 4)
            Next p
        Else
            n_row_p = n_row_p + 1
            pos_out(n_row_p, 3) = "-"
            pos_out(n_row_p, 4) = "-"
        End If
        '-- СТЕНЫ ---
        If Not zone.Exists(num + ";w") Then
            wall = Empty
        Else
            wall = ArraySort(zone.Item(num + ";w").keys())
        End If
        If Not IsEmpty(wall) Then
            For Each w In wall
                n_row_w = n_row_w + 1
                mat = zone.Item(num + ";wn;" + w)
                area = Round_w(zone.Item(num + ";wa;" + w) * k_zap_total, n_round_area)
                If InStr(mat, fin_str) > 0 Then sum_wall = sum_wall + area
                pos_out(n_row_w, 5) = Replace(mat, fin_str, "")
                pos_out(n_row_w, 6) = area
            Next w
        Else
            n_row_w = n_row_w + 1
            pos_out(n_row_w, 5) = "-"
            pos_out(n_row_w, 6) = "-"
        End If
         '-- КОЛОННЫ ---
        If is_column Then
            If Not zone.Exists(num + ";c") Then
                colum = Empty
            Else
                colum = ArraySort(zone.Item(num + ";c").keys())
            End If
            If Not IsEmpty(colum) Then
                For Each c In colum
                    n_row_c = n_row_c + 1
                    mat = zone.Item(num + ";cn;" + c)
                    area = Round_w(zone.Item(num + ";ca;" + c) * k_zap_total, n_round_area)
                    If InStr(mat, fin_str) > 0 Then sum_column = sum_column + area
                    pos_out(n_row_c, 7) = Replace(mat, fin_str, "")
                    pos_out(n_row_c, 8) = area
                Next c
            Else
                n_row_c = n_row_c + 1
                pos_out(n_row_c, 7) = "-"
                pos_out(n_row_c, 8) = "-"
            End If
        End If
         '-- ПАНЕЛИ ---
        If is_pan Then
            If zone.Exists(num + ";pn") Then
                panel = ArraySort(zone.Item(num + ";pn").keys())
                For Each p In panel
                    n_row_pan = n_row_pan + 1
                    mat = zone.Item(num + ";pnn;" + p)
                    area = Round_w(zone.Item(num + ";pna;" + p) * k_zap_total, n_round_area)
                    If InStr(mat, fin_str) > 0 Then sum_pan = sum_pan + area
                    pos_out(n_row_pan, col_end + 1) = Replace(mat, fin_str, "")
                    pos_out(n_row_pan, col_end + 2) = area
                    pos_out(n_row_pan, col_end + 3) = zone.Item(num + ";pnh;" + p)
                Next p
            Else
                n_row_pan = n_row_pan + 1
                pos_out(n_row_pan, col_end + 1) = "-"
                pos_out(n_row_pan, col_end + 2) = "-"
                pos_out(n_row_pan, col_end + 3) = "-"
            End If
        End If
        n_row = Application.WorksheetFunction.Max(n_row_p, n_row_w, n_row_c, n_row_pan)
    Next
    n_row = n_row + 1
    If show_surf_area Then
        pos_out(n_row, 5) = "Общяя площадь поверхности, кв.м."
        pos_out(n_row + 1, 5) = "Потолки"
        pos_out(n_row + 2, 5) = "Стены(за вычетом панелей)"
        pos_out(n_row + 3, 5) = "Колонны"
        pos_out(n_row + 4, 5) = "Панели"
        pos_out(n_row + 1, 8) = sum_pot
        pos_out(n_row + 2, 8) = sum_wall
        pos_out(n_row + 3, 8) = sum_column
        pos_out(n_row + 4, 8) = sum_pan
    End If
    If show_mat_area Then
        pos_out(n_row, 1) = "Общяя площадь отделки, кв.м."
        material_all = ArraySort(materials.keys())
        For Each mat In material_all
            If (Right(mat, 2) <> ";a") Then
                n_row = n_row + 1
                pos_out(n_row, 1) = Replace(materials.Item(mat), fin_str, "")
                pos_out(n_row, 4) = Round_w(materials.Item(mat + ";a") * k_zap_total, n_round_area)
            End If
        Next
    End If
    pos_out = ArrayRedim(pos_out, n_row)
    r = LogWrite("Ведомость отделки", "ИТОГ", "'====")
    r = LogWrite("Ведомость отделки", "Потолки", CStr(sum_pot))
    r = LogWrite("Ведомость отделки", "Стены", CStr(sum_wall))
    r = LogWrite("Ведомость отделки", "Колонны", CStr(sum_column))
    r = LogWrite("Ведомость отделки", "Панели", CStr(sum_pan))
    r = LogWrite("Ведомость отделки", "КОНЕЦ", "'====")
    Spec_VED = pos_out
End Function


Function VedAddRules(ByVal nm As String, ByVal add_rule As Variant) As Boolean
    nm_rule = ""
    nm = Split(nm, "_")(0)
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
        lsize = SheetGetSize(rule_sheet)
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
        VedAddRules = True
    Else
        VedAddRules = False
        r = VedNewListRules(nm)
        MsgBox ("Не найден лист с правилами отделки (оканчивается на '_правила')")
    End If
End Function

Function VedGetRules(ByVal nm As String) As Variant
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
        lsize = SheetGetSize(rule_sheet)
        n_row = lsize(1)
        n_col = lsize(2)
        If n_row = 1 Then n_row = 2
        Set Data_out = rule_sheet.Range(rule_sheet.Cells(1, 1), rule_sheet.Cells(n_row, n_col))
        Worksheets(nm_rule).Activate
        r = FormatClear()
        r = FormatSpec_Rule(Data_out)
        Dim rules: ReDim rules(n_row - 1, 3)
        Dim rules_mod: ReDim rules_mod(n_row - 1, 3)
        n_rules = 0
        n_rules_mod = 0
        For i = 2 To n_row
            If Not IsEmpty(Data_out(i, 1)) And Len(Data_out(i, 1)) > 1 And Left(Data_out(i, 1), 2) <> "!!" Then
                If InStr(Data_out(i, 1), "Исключить") = 0 And InStr(Data_out(i, 1), "Добавить") = 0 Then
                    n_rules = n_rules + 1
                    rules(n_rules, 1) = Data_out(i, 1)
                    rules(n_rules, 2) = Data_out(i, 2)
                    rules(n_rules, 3) = Data_out(i, 3)
                Else
                    n_rules_mod = n_rules_mod + 1
                    rules_mod(n_rules_mod, 1) = Trim(Data_out(i, 1))
                    rules_mod(n_rules_mod, 1) = Replace(rules_mod(n_rules_mod, 1), "Исключить", "-")
                    rules_mod(n_rules_mod, 1) = Replace(rules_mod(n_rules_mod, 1), "Добавить", "+")
                    rules_mod(n_rules_mod, 2) = Trim(Data_out(i, 2))
                    rules_mod(n_rules_mod, 3) = Trim(Data_out(i, 3))
                End If
            End If
        Next i
        rules = ArrayRedim(rules, n_rules)
        rules_mod = ArrayRedim(rules_mod, n_rules_mod)
        VedGetRules = Array(rules, rules_mod)
        Erase rules
    Else
        VedGetRules = Array(Empty, Empty)
        r = VedNewListRules(nm)
        MsgBox ("Создан лист с правилами.")
    End If
End Function

Function VedModMat(ByVal fin_material As String, ByVal all_material As String, ByRef rules_mod As Variant) As String
    If Not IsEmpty(rules_mod) Then
        For i = 1 To UBound(rules_mod, 1)
            If Trim(fin_material) = Trim(rules_mod(i, 2)) Then
                If rules_mod(i, 1) = "-" Then
                    arr_mat = Split(all_material, ";")
                    arr_mod = Split(rules_mod(i, 3), ";")
                    all_material_out = ""
                    For Each mat In arr_mat
                        mat = Trim(mat)
                        flag_in = True
                        For Each modd In arr_mod
                            modd = Trim(modd)
                            If mat = modd Then flag_in = False
                        Next modd
                        If flag_in = True Then
                            If all_material_out = "" Then
                                all_material_out = mat
                            Else
                                all_material_out = all_material_out & ";" & mat
                            End If
                        End If
                    Next mat
                    all_material = Trim(all_material_out)
                End If
                'all_material = Replace(all_material, rules_mod(i, 3), "")
                If rules_mod(i, 1) = "+" Then all_material = all_material + ";" + rules_mod(i, 3)
            End If
        Next i
        all_material = Replace(all_material, "; ;", ";")
        all_material = Replace(all_material, ";;", ";")
        all_material = Trim(all_material)
        If all_material = ";" Then all_material = ""
    End If
    VedModMat = all_material
End Function

Function VedNameMat(ByVal layer As String, ByVal material As String, ByRef rules As Variant) As String
    name_m = ""
    flag = 0
    For i = 1 To UBound(rules, 1)
        m = rules(i, 1)
        L = rules(i, 2)
        If (layer = L Or layer = "") And m = material Then
            name_m = rules(i, 3)
            flag = flag + 1
        End If
    Next i
    If flag < 1 Then
        For i = 1 To UBound(rules, 1)
            m = rules(i, 1)
            L = rules(i, 2)
            If (layer = L Or layer = "") And InStr(material, m) > 0 Then
                name_m = rules(i, 3)
                flag = flag + 1
            End If
        Next i
    End If
    If flag = 1 Then
        If InStr(name_m, "ез отделк") > 0 Then
            name_m = Replace(name_m, "; БЕЗ ОТДЕЛКИ", "")
            name_m = Replace(name_m, "без отделки", "")
            name_m = Trim(name_m)
            If Right(name_m, 1) = ";" Then name_m = Trim(Left(name_m, Len(name_m) - 1))
            name_m = name_m + "="
        End If
        VedNameMat = name_m
    Else
        VedNameMat = material + ";" + layer + ";ОШИБКА"
        If flag > 1 Then
            MsgBox ("Несколько правил для одного материала - " + material + " слой" + layer)
        End If
    End If
End Function

Function VedNewListRules(ByVal nm As String) As Boolean
    ThisWorkbook.Worksheets.Add.Name = nm & "_правила"
    Worksheets(nm & "_правила").Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Worksheets(nm & "_правила").Activate
    Cells(1, 1).Value = "Имя многослойной конструкции (целиком или часть имени, строки с !! не учитываются)"
    Cells(1, 2).Value = "Слой"
    Cells(1, 3).Value = "Черновая отделка (разделитель ';')"
    
    Cells(2, 1).Value = "!!Исключить"
    Cells(2, 2).Value = "Облицовка керам. плиткой"
    Cells(2, 3).Value = "Шпатлёвка"
    
    Cells(3, 1).Value = "!!Добавить"
    Cells(3, 2).Value = "Облицовка керам. плиткой"
    Cells(3, 3).Value = "Обработка бетоноконтактом"
    
    Cells(5, 1).Value = "ЖБ"
    Cells(5, 2).Value = "Колонны"
    Cells(5, 3).Value = "Затирка; Шпатлёвка"
    
    Cells(6, 1).Value = "!!П1"
    Cells(6, 2).Value = "Потолок"
    Cells(6, 3).Value = "Армстронг; Без отделки"
    
    Columns("A:A").ColumnWidth = 50
    Columns("B:B").ColumnWidth = 30
    Columns("C:C").ColumnWidth = 60
    Rows("1:1").EntireRow.AutoFit
End Function

Function VedRead(ByVal lastfilespec As String) As Variant
    lastfilespec = Left(lastfilespec, Len(lastfilespec) - Len("_вед"))
    out_data = ReadFile(lastfilespec & ".txt")
    If Not DataIsOtd(out_data) Then
        MsgBox ("Неверный формат файла")
        VedRead = Empty
        Exit Function
    End If
    rules = VedGetRules(lastfilespec)(1)
    rules_mod = VedGetRules(lastfilespec)(2)
    Set add_rule = CreateObject("Scripting.Dictionary")
    Set zone_error = CreateObject("Scripting.Dictionary")
    Set zone_num = CreateObject("Scripting.Dictionary")
    add_rule.comparemode = 1
    If IsEmpty(rules) Or IsEmpty(out_data) Then
        VedRead = Empty
        Exit Function
    End If
    n_row_a = UBound(out_data, 1) - 2
    n_col_a = UBound(out_data, 2)
    n_zone = 999999
    For i = 1 To n_row_a
       out_data(i, col_s_type_otd) = ConvNum2Txt(out_data(i, col_s_type_otd))
        If out_data(i, col_s_numb_zone) = 0 Or out_data(i, col_s_numb_zone) = "" Then
            out_data(i, col_s_numb_zone) = n_zone
        Else
            n_zone = ConvNum2Txt(out_data(i, col_s_numb_zone))
            If zone_num.Exists(n_zone) Then
                'Если такая зона уже есть - добавим их количество. Но вообще зон с одинаковым номером быть не дОлжно.
                zone_num.Item(n_zone) = zone_num.Item(n_zone) + 1
                n_zone = n_zone + "@@" + CStr(zone_num.Item(n_zone))
            Else
                zone_num.Item(n_zone) = 1
            End If
            out_data(i, col_s_numb_zone) = n_zone
        End If
        If n_col_a > col_s_mun_zone Then ' Если есть лестницы
            out_data(i, col_s_tipverh_l) = ConvNum2Txt(out_data(i, col_s_tipverh_l))
            out_data(i, col_s_tipniz_l) = ConvNum2Txt(out_data(i, col_s_tipniz_l))
            out_data(i, col_s_tippl_l) = ConvNum2Txt(out_data(i, col_s_tippl_l))
            out_data(i, col_s_tipl_l) = ConvNum2Txt(out_data(i, col_s_tipl_l))
        End If
        If out_data(i, col_s_type) = "СТЕНА" Then
            If out_data(i, col_s_area_wall) > 0 Then
                layer = out_data(i, col_s_layer)
                material = out_data(i, col_s_mat_wall)
                name_mat = VedNameMat(layer, material, rules)
                If InStr(name_mat, "ОШИБКА") > 0 Then
                    If Not add_rule.Exists(name_mat) Then add_rule.Item(name_mat) = name_mat
                    out_data(i, col_s_mat_wall) = "ОШИБКА"
                Else
                    out_data(i, col_s_mat_wall) = name_mat
                End If
            End If
        End If
        If n_col_a > col_s_layer Then 'Если есть пол или потолок
            out_data(i, col_s_type_pol) = ConvNum2Txt(out_data(i, col_s_type_pol))
            If out_data(i, col_s_type) = "ОБЪЕКТ" And out_data(i, col_s_type_el) = "Потолок" Then
                layer = "Потолок"
                material = out_data(i, col_s_type_pol)
                name_mat = VedNameMat(layer, material, rules)
                If InStr(name_mat, "ОШИБКА") > 0 Then
                    If Not add_rule.Exists(name_mat) Then add_rule.Item(name_mat) = name_mat
                    out_data(i, col_s_type_pol) = "ОШИБКА"
                Else
                    out_data(i, col_s_type_pol) = name_mat
                End If
            End If
            If out_data(i, col_s_type) = "ОБЪЕКТ" Then
                out_data(i, col_s_n_mun_zone) = ConvNum2Txt(out_data(i, col_s_n_mun_zone))
                If out_data(i, col_s_n_mun_zone) <> "" And out_data(i, col_s_n_mun_zone) <> out_data(i, col_s_numb_zone) Then
                    If out_data(i, col_s_mun_zone) = 1 Then
                        out_data(i, col_s_numb_zone) = out_data(i, col_s_n_mun_zone)
                    Else
                        r = LogWrite("Проверьте пол/потолок номер помещения " & out_data(i, col_s_numb_zone) & " или " & out_data(i, col_s_n_mun_zone), "Ошибка", num)
                    End If
                End If
                If out_data(i, col_s_type_el) = "Потолок" Then zone_error.Item(out_data(i, col_s_numb_zone) + "_pot_qty") = zone_error.Item(out_data(i, col_s_numb_zone) + "_pot_qty") + 1
                If out_data(i, col_s_type_el) = "Пол" Then zone_error.Item(out_data(i, col_s_numb_zone) + "_pol_qty") = zone_error.Item(out_data(i, col_s_numb_zone) + "_pol_qty") + 1
            End If
        End If
        For j = 1 To n_col_a
            If out_data(i, j) = "" Then out_data(i, j) = 0
        Next j
        If is_error > 0 Then
            zone_error.Item(num + "_err") = is_error
        End If
    Next i
    n_err = 0
    For Each zonerr In zone_error.keys()
        nqty = zone_error.Item(zonerr)
        If InStr(zonerr, "_qty") > 0 And nqty > 1 Then
            n_err = n_err + 1
            num = Split(zonerr, "_")(0)
            zone_error.Item(num + "_err") = zone_error.Item(num + "_err") + 1
        End If
    Next
    Dim pos_out: ReDim pos_out(3)
    If add_rule.Count = 0 Then
        pos_out(1) = out_data
        pos_out(2) = rules
        pos_out(3) = rules_mod
    Else
        r = VedAddRules(lastfilespec, add_rule.keys)
        pos_out(1) = Empty
        pos_out(2) = Empty
        pos_out(3) = Empty
    End If
    VedRead = pos_out
End Function

Function VedReadPol(ByVal lastfilespec As String) As Variant
    lastfilespec = Split(lastfilespec, "_")(0)
    out_data = ReadFile(lastfilespec & ".txt")
    If IsEmpty(out_data) Then
        VedReadPol = Empty
        Exit Function
    End If
    If Not DataIsOtd(out_data) Then
        MsgBox ("Неверный формат файла")
        VedReadPol = Empty
        Exit Function
    End If
    n_row_a = UBound(out_data, 1) - 1
    n_col_a = UBound(out_data, 2)
    If n_col_a <= col_s_layer Then
        VedReadPol = Empty
        Exit Function
    End If
    Dim add_pol: ReDim add_pol(1, 1)
    If UBound(out_data, 2) >= col_s_tipverh_l Then
        ReDim add_pol(col_s_areapl_l, n_row_a)
        n_add = 0
    End If
    n_zone = 999999
    For i = 1 To n_row_a
        If out_data(i, col_s_numb_zone) = 0 Then
            out_data(i, col_s_numb_zone) = n_zone
        Else
            n_zone = ConvNum2Txt(out_data(i, col_s_numb_zone))
            out_data(i, col_s_numb_zone) = n_zone
        End If
        If out_data(i, col_s_numb_zone) = 0 Then
            out_data(i, col_s_numb_zone) = n_zone
        Else
            n_zone = ConvNum2Txt(out_data(i, col_s_numb_zone))
            out_data(i, col_s_numb_zone) = n_zone
        End If
        If out_data(i, col_s_type) = "ОБЪЕКТ" Then
            out_data(i, col_s_n_mun_zone) = ConvNum2Txt(out_data(i, col_s_n_mun_zone))
            If out_data(i, col_s_n_mun_zone) <> "" And out_data(i, col_s_n_mun_zone) <> out_data(i, col_s_numb_zone) Then
                If out_data(i, col_s_mun_zone) = 1 Then
                    out_data(i, col_s_numb_zone) = out_data(i, col_s_n_mun_zone)
                Else
                    r = LogWrite("Проверьте пол/потолок номер помещения " & out_data(i, col_s_numb_zone) & " или " & out_data(i, col_s_n_mun_zone), "Ошибка", num)
                End If
            End If
        End If
        If n_col_a >= col_s_type_el Then out_data(i, col_s_type_pol) = ConvNum2Txt(out_data(i, col_s_type_pol))
        If n_col_a >= col_s_tipverh_l Then
            out_data(i, col_s_tipverh_l) = ConvNum2Txt(out_data(i, col_s_tipverh_l))
            out_data(i, col_s_tipniz_l) = ConvNum2Txt(out_data(i, col_s_tipniz_l))
            out_data(i, col_s_tippl_l) = ConvNum2Txt(out_data(i, col_s_tippl_l))
            out_data(i, col_s_tipl_l) = ConvNum2Txt(out_data(i, col_s_tipl_l))
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
    If n_add > 0 Then
        ReDim Preserve add_pol(col_s_areapl_l, n_add)
        add_pol = ArrayTranspose(add_pol)
        out_data = ArrayCombine(out_data, add_pol)
    End If
    VedReadPol = Array(out_data, Empty, Empty)
End Function

Function VedWriteLog(ByVal nm As String)
    If Debug_mode = False Then Exit Function
    nm_log = Right(nm, 24) & "_log"
    If SheetExist(nm_log) Then
        Worksheets(nm_log).Activate
        Worksheets(nm_log).Cells.Clear
    Else
        ThisWorkbook.Worksheets.Add.Name = nm_log
    End If
    Set name_col = CreateObject("Scripting.Dictionary")
    name_col.Item("01@err") = "Кол-во ошибок"
    name_col.Item("04@area_total") = "Площадь"
    name_col.Item("05@area_total_pot") = "Потолка"
    name_col.Item("06@pot_diff") = "Разница"
    name_col.Item("07@area_total_pol") = "Пола"
    name_col.Item("08@pol_diff") = "Разница"
    name_col.Item("09@column_area") = "Пл.колонн"
    name_col.Item("11@wall_area_zone") = "Пл.стен"
    name_col.Item("14@pot_qty") = "Nпотолков"
    name_col.Item("15@pol_qty") = "Nполов"
    un_col = ArraySort(name_col.keys(), 1)
    Dim zone_errornum: ReDim zone_errornum(1): n_row = 1
    For Each zoneerr In zone_error.keys()
        If InStr(zoneerr, "_err") > 0 Then
            zone_errornum(n_row) = Split(zoneerr, "_")(0)
            n_row = n_row + 1
            ReDim Preserve zone_errornum(n_row)
        End If
    Next
    n_col = UBound(un_col, 1) + 1
    Dim pos_out: ReDim pos_out(n_row, n_col)
    For i = 2 To n_row
        pos_out(i, 1) = "'" + Replace(ConvNum2Txt(zone_errornum(i - 1)), ",", ".")
    Next i
    For j = 2 To n_col
        pos_out(1, j) = name_col.Item(un_col(j - 1))
    Next
    For i = 2 To n_row
        num = zone_errornum(i - 1)
        For j = 2 To n_col
            nkey = Split(un_col(j - 1), "@")(1)
            pos_out(i, j) = zone_error.Item(num + "_" + nkey)
            If InStr(nkey, "_qty") > 0 And zone_error.Item(num + "_" + nkey) = 0 Then pos_out(i, j) = 0
            If InStr(nkey, "_qty") > 0 And zone_error.Item(num + "_" + nkey) = 1 Then pos_out(i, j) = ""
        Next j
    Next i
    Set Sh = Application.ThisWorkbook.Sheets(nm_log)
    Sh.Range(Sh.Cells(2, 1), Sh.Cells(n_row + 1, n_col)) = pos_out
    Set Data_out = Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col))
End Function

Function VedSplitData(ByVal all_data As Variant, ByVal split_data As Variant, ByVal lastfilespec As Variant, ByVal suffix As String) As Variant
    n_split = UBound(split_data, 1)
    Dim out_data: ReDim out_data(n_split, 2)
    raw_data = all_data(1)
    rules = all_data(2)
    rules_mod = all_data(3)
    zones_el_all = Empty
    For i = 1 To n_split
        nm = Right(lastfilespec & "-" & split_data(i, 1) & suffix, 31)
        If split_data(i, 3) <> col_s_numb_zone Then
            If IsEmpty(zones_el_all) Then zones_el_all = ArraySelectParam(raw_data, "ЗОНА", col_s_type)
            un_zone = ArrayUniqValColumn(ArraySelectParam_2(zones_el_all, split_data(i, 2), split_data(i, 3)), col_s_numb_zone)
            split_data(i, 2) = un_zone
            split_data(i, 3) = col_s_numb_zone
        End If
        zones = ArraySelectParam_2(raw_data, split_data(i, 2), split_data(i, 3))
        out_data(i, 1) = nm
        out_data(i, 2) = Array(zones, rules, rules_mod)
    Next i
    VedSplitData = out_data
    Erase all_data, split_data
End Function

Function VedSplitSheet(ByVal lastfilespec As String)
    Set split_sheet = Application.ThisWorkbook.Sheets(Split(lastfilespec, "_")(0) & "_разб")
    r = FormatTable(Split(lastfilespec, "_")(0) & "_разб")
    sheet_size = SheetGetSize(split_sheet)
    raw_data = split_sheet.Range(split_sheet.Cells(2, 1), split_sheet.Cells(sheet_size(1), 3))
    n_split = UBound(raw_data, 1)
    n_row = 0
    Dim split_data: ReDim split_data(n_split, 3)
    For i = 1 To n_split
        If Not IsEmpty(raw_data(i, 1)) Then
            nm = raw_data(i, 1)
            num_zone = Split(raw_data(i, 2), ";")
            n_col_param = CInt(raw_data(i, 3))
            If n_col_param <= 0 Or n_col_param > col_s_type_otd Then n_col_param = 1
            If Not IsEmpty(num_zone) Then
                num_zone = ArrayUniqValColumn(num_zone, 1)
                For j = LBound(num_zone) To UBound(num_zone)
                    If IsNumeric(num_zone(j)) Then num = CStr(num_zone(j))
                    num_zone(j) = Trim(Trim(num_zone(j)))
                Next
                n_row = n_row + 1
                split_data(n_row, 1) = nm
                split_data(n_row, 2) = num_zone
                split_data(n_row, 3) = n_col_param
            End If
        End If
    Next i
    split_data = ArrayRedim(split_data, n_row)
    VedSplitSheet = split_data
End Function

Function VedSplitFile(ByVal lastfilespec As String)
    raw_data = ReadTxt(ThisWorkbook.path & "\import\" & lastfilespec & "_разб" & ".txt", 1, vbTab, vbNewLine)
    sheet_name = ArrayUniqValColumn(raw_data, 1)
    n_split = UBound(sheet_name, 1)
    Dim split_data: ReDim split_data(n_split, 3)
    For i = 1 To n_split
        If Not IsEmpty(sheet_name(i)) Then
            nm = sheet_name(i)
            For Each del_txt In Array("План", "Кровля", "на", "отм.", "отметке", "отметка", "этаж", "  ")
                nm = Replace(nm, del_txt, "")
            Next
            nm = Trim(Trim(nm)) 'Безусловное удаление пробелов
            num_zone = ArrayUniqValColumn(ArraySelectParam(raw_data, sheet_name(i), 1), 2)
            If Not IsEmpty(num_zone) Then
                For j = 1 To UBound(num_zone)
                    If IsNumeric(num_zone(j)) Then num = CStr(num_zone(j))
                    num_zone(j) = Trim(Trim(num_zone(j)))
                Next
                n_row = n_row + 1
                split_data(n_row, 1) = nm
                split_data(n_row, 2) = num_zone
                split_data(n_row, 3) = col_s_numb_zone
            End If
        End If
    Next i
    split_data = ArrayRedim(split_data, n_row)
    VedSplitFile = split_data
End Function

