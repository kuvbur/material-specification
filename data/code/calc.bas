Attribute VB_Name = "calc"
Option Compare Text
Option Base 1
Public Const macro_version As String = "4.12"
'-------------------------------------------------------
'Типы элементов (столбец col_type_el)
Public Const t_arm As Long = 10
Public Const t_prokat As Long = 20
Public Const t_mat As Long = 30
Public Const t_mat_spc As Long = 35
Public Const t_izd As Long = 40
Public Const t_subpos As Long = 45
Public Const t_else As Long = 50
Public Const t_wind As Long = 60
Public Const t_perem_m As Long = 70
Public Const t_perem As Long = 71
Public Const t_error As Long = -1 'Ошибка распознавания типов
Public Const t_sys As Long = -10 'Вспомогательный тип
Public Const t_syserror As Long = -20 'Строки с ошибками
'Столбцы общие
Public Const col_marka As Long = 1
Public Const col_sub_pos As Long = 2
Public Const col_type_el As Long = 3
Public Const col_pos As Long = 4
Public Const col_qty As Long = 8
Public Const col_chksum As Long = 12
Public Const col_parent As Long = 15
Public Const col_nfloor As Long = 16
Public Const col_floor As Long = 17
Public Const col_param As Long = 18
'Столбцы арматуры (t_arm)
Public Const col_klass As Long = 5
Public Const col_diametr As Long = 6
Public Const col_length As Long = 7
Public Const col_fon As Long = 9
Public Const col_mp As Long = 10
Public Const col_gnut As Long = 11
'Столбцы проката (t_prokat)
Public Const col_pr_type_konstr As Long = 5
Public Const col_pr_gost_st As Long = 6
Public Const col_pr_st As Long = 7
Public Const col_pr_gost_prof As Long = 9
Public Const col_pr_prof As Long = 10
Public Const col_pr_length As Long = 11
Public Const col_pr_weight As Long = 13
Public Const col_pr_naen As Long = 14
'Столбцы материалов и изделий (t_izd, t_mat, t_subpos)
Public Const col_m_obozn As Long = 5
Public Const col_m_naen As Long = 6
Public Const col_m_weight As Long = 7
Public Const col_m_edizm As Long = 9
'Столбцы окон, дверей
Public Const col_w_obozn As Long = 5
Public Const col_w_naen As Long = 6
Public Const col_w_weight As Long = 7
Public Const col_w_prim As Long = 9
Public Const col_w_guid As Long = 11
'Общее количество столбцов во входном массиве
Public Const max_col As Long = 19
'-------------------------------------------------------
'Описание таблицы с именами сборок (суффикс "_поз")
Public Const col_add_pos As Long = 1
Public Const col_add_obozn As Long = 2
Public Const col_add_naen As Long = 3
Public Const col_add_qty As Long = 4
Public Const col_add_prim As Long = 5
'-------------------------------------------------------
'Описание файла с отделкой
Public Const col_s_numb_zone As Long = 1
Public Const col_s_name_zone As Long = 2
Public Const col_s_area_zone As Long = 3
Public Const col_s_perim_zone As Long = 4
Public Const col_s_perimhole_zone As Long = 5
Public Const col_s_h_zone As Long = 6
Public Const col_s_freelen_zone As Long = 7
Public Const col_s_walllen_zone As Long = 8
Public Const col_s_doorlen_zone As Long = 9
Public Const col_s_hpan_zone As Long = 10
Public Const col_s_mpot_zone As Long = 11
Public Const col_s_mpan_zone As Long = 12
Public Const col_s_mwall_zone As Long = 13
Public Const col_s_mcolumn_zone As Long = 14
Public Const col_s_area_wall As Long = 15
Public Const col_s_type As Long = 16
Public Const col_s_mat_wall As Long = 17
Public Const col_s_type_otd As Long = 18
Public Const col_s_layer As Long = 19
Public Const max_col_type_1_1 As Long = 19
Public Const col_s_type_el_1 As Long = 20
Public Const col_s_type_pol_1 As Long = 21
Public Const col_s_area_pol_1 As Long = 22
Public Const col_s_perim_pol_1 As Long = 23
Public Const col_s_n_mun_zone_1 As Long = 24
Public Const col_s_mun_zone_1 As Long = 25
Public Const max_col_type_2_1 As Long = 25
Public Const col_s_tipverh_l_1 As Long = 26
Public Const col_s_tipl_l_1 As Long = 27
Public Const col_s_tipniz_l_1 As Long = 28
Public Const col_s_tippl_l_1 As Long = 29
Public Const col_s_areaverh_l_1 As Long = 30
Public Const col_s_areal_l_1 As Long = 31
Public Const col_s_areaniz_l_1 As Long = 32
Public Const col_s_areapl_l_1 As Long = 33
Public Const max_col_type_3_1 As Long = 33
Public Const max_s_col_1 As Long = 33
'Описание файла с отделкой v2
Public Const col_s_type_pot_zone As Long = 20
Public Const col_s_type_pol_zone As Long = 21
Public Const col_s_h_pot_zone As Long = 22
Public Const col_s_mwall_up_zone As Long = 23
Public Const col_s_param_zone As Long = 24
Public Const col_s_h_wall As Long = 25
Public Const max_col_type_1_2 As Long = 25
Public Const col_s_type_el_2 As Long = 26
Public Const col_s_type_pol_2 As Long = 21
Public Const col_s_area_pol_2 As Long = 27
Public Const col_s_perim_pol_2 As Long = 28
Public Const col_s_n_mun_zone_2 As Long = 29
Public Const col_s_mun_zone_2 As Long = 30
Public Const max_col_type_2_2 As Long = 30
Public Const col_s_tipverh_l_2 As Long = 31
Public Const col_s_tipl_l_2 As Long = 32
Public Const col_s_tipniz_l_2 As Long = 33
Public Const col_s_tippl_l_2 As Long = 34
Public Const col_s_areaverh_l_2 As Long = 35
Public Const col_s_areal_l_2 As Long = 36
Public Const col_s_areaniz_l_2 As Long = 37
Public Const col_s_areapl_l_2 As Long = 38
Public Const max_col_type_3_2 As Long = 38
Public Const max_s_col_2 As Long = 38
'Столбцы с изменяющимеся номерами, в зависимости от версии
Public col_s_type_el As Long
Public col_s_type_pol As Long
Public col_s_area_pol As Long
Public col_s_perim_pol As Long
Public col_s_n_mun_zone As Long
Public col_s_mun_zone As Long
Public col_s_tipverh_l As Long
Public col_s_tipl_l As Long
Public col_s_tipniz_l As Long
Public col_s_tippl_l As Long
Public col_s_areaverh_l As Long
Public col_s_areal_l As Long
Public col_s_areaniz_l As Long
Public col_s_areapl_l As Long
Public max_col_type_1 As Long
Public max_col_type_2 As Long
Public max_col_type_3 As Long
Public max_s_col As Long

Public fin_str As String
Public fin_str_sec As String
Public zone_error As Variant
'-------------------------------------------------------
'Описание файла сортамента
Public Const col_gost_spec As Long = 1
Public Const col_klass_spec As Long = 2
Public Const col_diametr_spec As Long = 3
Public Const col_area_spec As Long = 4
Public Const col_weight_spec As Long = 5
'-------------------------------------------------------
'Столбцы ручной спецификации (суффикс "_спец")
'Общие
Public Const col_man_subpos As Long = 1
Public Const col_man_pos As Long = 2
Public Const col_man_obozn As Long = 3
Public Const col_man_naen As Long = 4
Public Const col_man_qty As Long = 5
Public Const col_man_weight As Long = 6
Public Const col_man_prim As Long = 7
Public Const col_man_komment As Long = 18
Public Const col_man_ank As Long = 19
Public Const col_man_nahl As Long = 20
Public Const col_man_dgib As Long = 21
'Арматура
Public Const col_man_length As Long = 8
Public Const col_man_diametr As Long = 9
Public Const col_man_klass As Long = 10
'Прокат
Public Const col_man_pr_length As Long = 11
Public Const col_man_pr_gost_pr As Long = 12
Public Const col_man_pr_prof As Long = 13
Public Const col_man_pr_type As Long = 14
Public Const col_man_pr_st As Long = 15
Public Const col_man_pr_okr As Long = 16
Public Const col_man_pr_ogn As Long = 17
Public Const max_col_man As Long = col_man_dgib

'-------------------------------------------------------
Public symb_diam As String 'Символ диаметра в спецификацию
Public material_ed_izm As Variant
Public gost2fklass
Public name_gost
Public reinforcement_specifications As Variant
 'Разные массивы
Public pr_adress As Variant
Public wbk As Workbook
Public swap_gost As Variant
Public k_zap_total As Double
Public w_format As String 'Формат вывода в техничку
Public pos_data As Variant
Public floor_txt_arr As Variant
Public sheet_option_inx As Variant
Public concrsubpos As Variant
Public otd_version As Long 'Версия файла с отделкой
Public spec_version As Long 'Версия файла со спецификацией
Public Const log_sheet_name As String = "|Лог|"
Public type_el_name As Variant
Public aIniLines As Variant
Public error_ini As String
'-------------------------------------------------------
'-----Переменные, читаемые из INI-----------------------
'-------------------------------------------------------
    'Тип округления
    ' 1 - округление в большую сторону
    ' 2 - округление стандартным round
    ' 3 - округление отключено
Public isINIset As Boolean
Public type_okrugl As Long
Public n_round_l As Long 'Длина
Public n_round_w As Long 'Вес
Public n_round_wkzh As Long 'Вес в ведомости расхода стали
Public n_round_mat As Long 'Вес в ведомости расхода стали
Public sum_row_wkzh As Boolean
Public show_bet_wkzh As Boolean
Public show_sum_prim As Boolean
Public del_dor_perim As Boolean
Public type_perim As Long
Public del_freelen_perim As Boolean
Public add_holes_perim As Boolean
Public show_mat_area As Boolean
Public show_surf_area As Boolean
Public show_perim As Boolean
Public zonenum_pot As Boolean
Public delim_by_sheet As Boolean
Public delim_group_ved As Boolean
Public delim_zone_fin As Boolean
Public ignore_zap_material As Boolean
Public n_round_area As Long 'Площадь в ведомость отделки
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
Public lenght_ed_arm As Long 'Максимальная длина стержня арматуры
Public hard_round_km As Boolean

Public checktxt_on_load As Boolean 'Проверка текстовых файлов при подгрузке
Public clear_bet_name As Boolean 'Удаляем пояснения к марке бетона, сдаланные в скобках
Public zap_only_mp As Boolean 'Запас только для п.м арматуры и материала
Public functime_data As Variant
Public log_sheet As Variant
Public this_sheet_option As Variant 'хранилище значений форм
Public set_sheet_option As Variant 'хранилище значений листа
Public spec_type_suffix As Variant
Public defult_values_ini As Variant 'Значения по умолчанию
Public fontname As String

Function dprint(ByVal msg As String, Optional ByVal type_msg As Integer = 0)
    If type_msg = 1 Then msg = "ERROR      !!!!! " & msg
    If Debug_mode Then Debug.Print msg
End Function

Function functime(ByVal namefunc As String, ByVal tfunctime As Double) As Double
    If IsEmpty(functime_data) Then
        Set functime_data = CreateObject("Scripting.Dictionary")
    End If
    If Not functime_data.Exists(namefunc) Then functime_data.Item(namefunc) = Array(0#, 0#)
    Dim tt As Double
    tt = Timer - tfunctime
    jj = functime_data.Item(namefunc)
    jj(1) = jj(1) + tt
    jj(2) = jj(2) + 1
    functime_data.Item(namefunc) = jj
    functime = Timer
End Function

Function print_functime()
    If Not IsEmpty(functime_data) Then
        Dim arr(): ReDim arr(functime_data.Count + 1, 3)
        n = 0
        For Each varKey In functime_data.keys
            n = n + 1
            arr(n, 1) = varKey
            arr(n, 2) = functime_data.Item(varKey)(2)
            arr(n, 3) = Round(functime_data.Item(varKey)(1), 4)
        Next
        arr = ArraySort(arr, 3)
        r = dprint("---------- Общее время -------------")
        For i = 1 To n
            If arr(i, 3) > 0 Then dprint (arr(i, 1) & " x " & arr(i, 2) & " - " & arr(i, 3))
        Next i
        r = dprint("------------------------------------")
        functime_data.RemoveAll
    Else
        r = dprint("Таймеров нет")
    End If
End Function

Function INISet()
    mtype = ModeType()
    If mtype Then Exit Function
    Set defult_values_ini = CreateObject("Scripting.Dictionary")
    defult_values_ini.Item("type_okrugl") = 1
    defult_values_ini.Item("n_round_l") = 2
    defult_values_ini.Item("n_round_w") = 2
    defult_values_ini.Item("n_round_wkzh") = 1
    defult_values_ini.Item("n_round_mat") = 1
    defult_values_ini.Item("ignore_pos") = "!!"
    defult_values_ini.Item("subpos_delim") = "'"
    defult_values_ini.Item("n_round_area") = 1
    defult_values_ini.Item("hole_in_zone") = False
    defult_values_ini.Item("isErrorNoFin") = True
    defult_values_ini.Item("delim_zone_fin") = False
    defult_values_ini.Item("izd_sheet_name") = "Изделия"
    defult_values_ini.Item("inx_name") = "|Содержание|"
    defult_values_ini.Item("mem_option") = True
    defult_values_ini.Item("inx_on_new") = True
    defult_values_ini.Item("check_on_active") = True
    defult_values_ini.Item("def_decode") = False
    defult_values_ini.Item("Debug_mode") = False
    defult_values_ini.Item("check_version") = True
    defult_values_ini.Item("del_dor_perim") = False
    defult_values_ini.Item("type_perim") = 1
    defult_values_ini.Item("del_freelen_perim") = False
    defult_values_ini.Item("add_holes_perim") = False
    defult_values_ini.Item("show_mat_area") = True
    defult_values_ini.Item("show_surf_area") = True
    defult_values_ini.Item("show_perim") = True
    defult_values_ini.Item("zonenum_pot") = False
    defult_values_ini.Item("delim_by_sheet") = False
    defult_values_ini.Item("sum_row_wkzh") = True
    defult_values_ini.Item("show_bet_wkzh") = False
    defult_values_ini.Item("delim_group_ved") = False
    defult_values_ini.Item("show_sum_prim") = True
    defult_values_ini.Item("lenght_ed_arm") = 11700
    defult_values_ini.Item("hard_round_km") = True
    defult_values_ini.Item("ignore_zap_material") = False
    defult_values_ini.Item("clear_bet_name") = False
    defult_values_ini.Item("zap_only_mp") = False
    defult_values_ini.Item("checktxt_on_load") = False
    defult_values_ini.Item("fontname") = "ISOCPEUR"
    sIniFile = UserForm2.CodePath & "setting.ini"
    If Not CBool(Len(Dir$(sIniFile))) Then r = Download_Settings()
    If CBool(Len(Dir$(sIniFile))) Then
        aIniLines = INIReadFile(sIniFile)    'Read the file into memory
    Else
        aIniLines = Empty
    End If
    error_ini = vbNullString
    type_okrugl = INIReadKeyVal("РАСЧЁТЫ", "type_okrugl")
    n_round_l = INIReadKeyVal("РАСЧЁТЫ", "n_round_l")
    n_round_w = INIReadKeyVal("РАСЧЁТЫ", "n_round_w")
    n_round_wkzh = INIReadKeyVal("РАСЧЁТЫ", "n_round_wkzh")
    n_round_mat = INIReadKeyVal("РАСЧЁТЫ", "n_round_mat")
    ignore_pos = INIReadKeyVal("РАСЧЁТЫ", "ignore_pos")
    subpos_delim = INIReadKeyVal("РАСЧЁТЫ", "subpos_delim")
    n_round_area = INIReadKeyVal("ОТДЕЛКА", "n_round_area")
    hole_in_zone = INIReadKeyVal("ОТДЕЛКА", "hole_in_zone")
    isErrorNoFin = INIReadKeyVal("ОТДЕЛКА", "isErrorNoFin")
    delim_zone_fin = INIReadKeyVal("ОТДЕЛКА", "delim_zone_fin")
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
    lenght_ed_arm = INIReadKeyVal("РАСЧЁТЫ", "lenght_ed_arm")
    hard_round_km = INIReadKeyVal("РАСЧЁТЫ", "hard_round_km")
    ignore_zap_material = INIReadKeyVal("РАСЧЁТЫ", "ignore_zap_material")
    clear_bet_name = INIReadKeyVal("РАСЧЁТЫ", "clear_bet_name")
    zap_only_mp = INIReadKeyVal("РАСЧЁТЫ", "zap_only_mp")
    checktxt_on_load = INIReadKeyVal("ЛИСТЫ", "checktxt_on_load")
    izd_sheet_name = izd_sheet_name + "_спец.И"
    fontname = INIReadKeyVal("ЛИСТЫ", "fontname")
    '----Принудительное включение
    delim_by_sheet = True
    isINIset = True
    fin_str = "!!AA_"
    fin_str_sec = "!!BB_"
    If lenght_ed_arm < 1 Then lenght_ed_arm = 11700
    material_ed_izm = Array("кв.м.", "куб.м.", "кг.", "л.")
    '----Наименования типов элементов
    Set type_el_name = CreateObject("Scripting.Dictionary")
    type_el_name.Item(t_prokat) = " Прокат "
    type_el_name.Item(t_arm) = " Изделия арматурные "
    type_el_name.Item(t_mat) = " Материалы "
    type_el_name.Item(t_izd) = " Изделия "
    type_el_name.Item(t_subpos) = " Сборочные единицы "
    type_el_name.Item(t_wind) = " Заполнения проёмов "
    type_el_name.Item(t_perem) = " Сборочные единицы "
    type_el_name.Item(t_perem_m) = " Заполнения проёмов "
    r = SetWorkbook()
    If Not SheetExist(log_sheet_name) Then r = LogNewSheet(log_sheet_name)
    Set log_sheet = wbk.Sheets(log_sheet_name)
    If IsEmpty(spec_type_suffix) Then
        Set spec_type_suffix = CreateObject("Scripting.Dictionary")
        spec_type_suffix.comparemode = 1
        spec_type_suffix.Item("гр") = 1
        spec_type_suffix.Item("об") = 2
        spec_type_suffix.Item("км") = 4
        spec_type_suffix.Item("кж") = 5
        spec_type_suffix.Item("спец") = 7
        spec_type_suffix.Item("поз") = 9
        spec_type_suffix.Item("правила") = 10
        spec_type_suffix.Item("вед") = 11
        spec_type_suffix.Item("экспл") = 12
        spec_type_suffix.Item("грс") = 13
        spec_type_suffix.Item("норм") = 14
        spec_type_suffix.Item("спец.И") = 15
        spec_type_suffix.Item("зап") = 20
        spec_type_suffix.Item("разб") = 21
        spec_type_suffix.Item("мат") = 22
        spec_type_suffix.Item("autocad") = 23
        spec_type_suffix.Item("archicad") = 24
        spec_type_suffix.Item("раскр") = 25
    End If
End Function

Function INIReadKeyVal(ByVal sSection As String, ByVal sKey As String) As Variant
    If IsEmpty(aIniLines) Then
        If defult_values_ini.Exists(sKey) Then
            tval = defult_values_ini.Item(sKey)
            t = INIWriteKeyVal(sSection, sKey, tval)
            INIReadKeyVal = tval
        Else
            error_ini = error_ini & "Значение по умолчанию не задано " & sKey & vbLf
            r = dprint(error_ini, 1)
            INIReadKeyVal = Empty
        End If
    End If
    bSectionExists = False
    bKeyExists = False
    tval = Empty
    For i = LBound(aIniLines) To UBound(aIniLines)
        sLine = aIniLines(i)
        If bSectionExists = True And Left$(sLine, 1) = "[" And Right$(sLine, 1) = "]" Then
            Exit For    'Start of a new section
        End If
        If sLine = "[" & sSection & "]" Then
            bSectionExists = True
        End If
        If bSectionExists = True Then
            If Len(sLine) > Len(sKey) Then
                If Left$(sLine, Len(sKey) + 1) = sKey & "=" Then
                    bKeyExists = True
                    tval = Mid$(sLine, InStr(sLine, "=") + 1)
                End If
            End If
        End If
    Next i
    If bSectionExists = False Or bKeyExists = False Then
        If bSectionExists = False And bKeyExists = False Then
            error_ini = error_ini & "Не найден параметр " & sKey & " в разделе " & sSection & vbLf
        Else
            If bSectionExists = False Then error_ini = error_ini & "Не найден раздел " & sSection & vbLf
            If bKeyExists = False Then error_ini = error_ini & "Не найден параметр " & sKey & vbLf
        End If
        If defult_values_ini.Exists(sKey) Then
            tval = defult_values_ini.Item(sKey)
            t = INIWriteKeyVal(sSection, sKey, tval)
            INIReadKeyVal = tval
        Else
            error_ini = error_ini & "Значение по умолчанию не задано " & sKey & vbLf
            INIReadKeyVal = Empty
        End If
        r = dprint(error_ini, 1)
    Else
        If InStr(tval, "#") > 0 Then tval = Trim$(Split(tval, "#")(0))
        INIReadKeyVal = tval
    End If
End Function

Function INIWriteKeyVal(ByVal sSection As String, ByVal sKey As String, ByVal sValue As String) As Boolean
    sIniFile = UserForm2.CodePath & "setting.ini"
    sIniFileContent = vbNullString
    bSectionExists = False
    bKeyExists = False
    If IsEmpty(aIniLines) Then aIniLines = INIReadFile(sIniFile)
    sIniFileContent = vbNullString    'Reset it
    sIniFileContent_before = vbNullString
    sIniFileContent_after = vbNullString
    For i = LBound(aIniLines) To UBound(aIniLines)    'Loop through each line
        sNewLine = vbNullString
        sLine = Trim$(aIniLines(i))
        If bSectionExists Then
            sIniFileContent_after = sIniFileContent_after & sLine & vbCrLf
        Else
            sIniFileContent_before = sIniFileContent_before & sLine & vbCrLf
        End If
        If sLine = "[" & sSection & "]" Then bSectionExists = True
    Next i
    sNewSection = "[" & sSection & "]" & vbCrLf
    sNewValue = sKey & "=" & sValue & " # Значение по умолчанию" & vbCrLf
    If bSectionExists Then
        sIniFileContent = sIniFileContent_before & sNewValue & sIniFileContent_after
    Else
        sIniFileContent = sIniFileContent_before & sNewSection & sNewValue
    End If
    r = ExportSaveTXTfile(sIniFile, sIniFileContent)
    Ini_WriteKeyVal = True
End Function

Function INIReadFile(ByVal strFile As String) As Variant
    On Error Resume Next
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.OpenTextFile(strFile$, 1, True): sFile$ = ts.ReadAll: ts.Close
    sFile = Replace(sFile, " = ", "=")
    sFile = Replace(sFile, "= ", "=")
    sFile = Replace(sFile, " =", "=")
    taIniLines = Split(sFile, vbCrLf)
    If UBound(taIniLines) < 1 Then taIniLines = Split(sFile, vbLf)
    For i = 0 To UBound(taIniLines)
        taIniLines(i) = Trim$(taIniLines(i))
    Next i
    INIReadFile = taIniLines
End Function

Function ArrayCol(ByVal array_in As Variant, ByVal col As Long) As Variant
    If IsEmpty(array_in) Then ArrayCol = Empty: Exit Function
    If ArrayIsSecondDim(array_in) = False Then ArrayCol = array_in: Exit Function
Dim tfunctime As Double
tfunctime = Timer
    n = UBound(array_in, 1)
    Dim out(): ReDim out(n)
    For i = 1 To n
        out(i) = array_in(i, col)
    Next i
    ArrayCol = out
r = functime("ArrayCol", tfunctime)
End Function

Function ArrayDelElement(ByVal array_in As Variant, ByVal param1 As Variant, Optional ByVal n_col1 As Long) As Variant
    Dim arrout
    If IsEmpty(array_in) Then
        ArrayDelElement = Empty
        Exit Function
    End If
    If Not IsArray(param1) Then
        param1 = Array(param1)
    End If
Dim tfunctime As Double
tfunctime = Timer
    If ArrayIsSecondDim(array_in) Then
        n_row = UBound(array_in, 1)
        n_param = UBound(array_in, 2)
        n_row_s = 0
        If n_col1 > n_param Then
            ArrayDelElement = Empty
            Exit Function
        End If
        ReDim arrout(n_row, n_param)
        For j = 1 To n_row
            flag1 = 0
            For Each tparam1 In param1
                If array_in(j, n_col1) = tparam1 Then
                    flag1 = 1
                Else
                    If InStr(tparam1, "?") > 0 Then
                        tpar = array_in(j, n_col1)
                        If IsNumeric(tpar) Then tparam1 = ConvNum2Txt(tpar)
                        If Right$(tparam1, 1) = "?" And Left$(tparam1, 1) = "?" Then
                            tparam1 = Trim$(Replace(tparam1, "?", vbNullString))
                            If InStr(tpar, tparam1) > 0 Then flag1 = 1
                        End If
                        If Left$(tparam1, 1) = "?" Then
                            tparam1 = Trim$(Replace(tparam1, "?", vbNullString))
                            If Right$(tpar, Len(tparam1)) = tparam1 Then flag1 = 1
                        End If
                        If Right$(tparam1, 1) = "?" Then
                            tparam1 = Trim$(Replace(tparam1, "?", vbNullString))
                            If Left$(tpar, Len(tparam1)) = tparam1 Then flag1 = 1
                        End If
                    End If
                End If
                If flag1 = 1 Then Exit For
            Next
            If flag1 = 0 Then 'Если все согласны
                n_row_s = n_row_s + 1
                For k = 1 To n_param
                    arrout(n_row_s, k) = array_in(j, k)
                Next k
            End If
        Next j
        If n_param > 0 And n_row_s > 0 Then
            ArrayDelElement = ArrayRedim(arrout, n_row_s)
r = functime("ArrayDelElement", tfunctime)
            Exit Function
        Else
            ArrayDelElement = Empty
r = functime("ArrayDelElement", tfunctime)
            Exit Function
        End If
    Else
        n_row = UBound(array_in)
        n_row_s = 0
        ReDim arrout(n_row)
        For j = 1 To n_row
            flag1 = 0
            For Each tparam1 In param1
                If array_in(j) = tparam1 Then
                    flag1 = 1
                Else
                    If InStr(tparam1, "?") > 0 Then
                        tpar = array_in(j)
                        If IsNumeric(tpar) Then tparam1 = ConvNum2Txt(tpar)
                        If Right$(tparam1, 1) = "?" And Left$(tparam1, 1) = "?" Then
                            tparam1 = Trim$(Replace(tparam1, "?", vbNullString))
                            If InStr(tpar, tparam1) > 0 Then flag1 = 1
                        End If
                        If Left$(tparam1, 1) = "?" Then
                            tparam1 = Trim$(Replace(tparam1, "?", vbNullString))
                            If Right$(tpar, Len(tparam1)) = tparam1 Then flag1 = 1
                        End If
                        If Right$(tparam1, 1) = "?" Then
                            tparam1 = Trim$(Replace(tparam1, "?", vbNullString))
                            If Left$(tpar, Len(tparam1)) = tparam1 Then flag1 = 1
                        End If
                    End If
                End If
                If flag1 = 1 Then Exit For
            Next
            If flag1 = 0 Then
                n_row_s = n_row_s + 1
                arrout(n_row_s) = array_in(j)
            End If
        Next j
        If n_row_s > 0 Then
            ReDim Preserve arrout(n_row_s)
r = functime("ArrayDelElement", tfunctime)
            ArrayDelElement = arrout
            Exit Function
        Else
             ArrayDelElement = Empty
r = functime("ArrayDelElement", tfunctime)
            Exit Function
        End If
    End If
End Function
    
Function ArrayCombine(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    isarr1 = IsArray(arr1)
    isarr2 = IsArray(arr2)
    If Not isarr1 And isarr2 Then ArrayCombine = arr2: Exit Function
    If Not isarr2 And isarr1 Then ArrayCombine = arr1: Exit Function
    If Not isarr2 And Not isarr1 Then ArrayCombine = Empty: Exit Function
    On Error Resume Next: Err.Clear
    If Err.Number = 9 Then ArrayCombine = Empty: Exit Function
    n_rec_row = 1: n_rec_col = 1
Dim tfunctime As Double
tfunctime = Timer
    If ArrayIsSecondDim(arr1) And ArrayIsSecondDim(arr2) Then
        n_row = (UBound(arr1, 1) - LBound(arr1, 1) + 1) + (UBound(arr2, 1) - LBound(arr2, 1) + 1)
        n_col = (UBound(arr1, 2) - LBound(arr1, 2) + 1)
        ReDim arr(n_row, n_col)
        For i = LBound(arr1, 1) To UBound(arr1, 1)
            n_rec_col = 1
            For j = LBound(arr1, 2) To UBound(arr1, 2)
                arr(n_rec_row, n_rec_col) = arr1(i, j)
                n_rec_col = n_rec_col + 1
            Next j
            n_rec_row = n_rec_row + 1
        Next i
        For i = LBound(arr2, 1) To UBound(arr2, 1)
            n_rec_col = 1
            For j = LBound(arr2, 2) To UBound(arr2, 2)
                arr(n_rec_row, n_rec_col) = arr2(i, j)
                n_rec_col = n_rec_col + 1
            Next j
            n_rec_row = n_rec_row + 1
        Next i
        ArrayCombine = arr
    Else
        If ArrayIsSecondDim(arr1) Then ArrayCombine = Empty: Exit Function
        If ArrayIsSecondDim(arr2) Then ArrayCombine = Empty: Exit Function
        n_row = (UBound(arr1) - LBound(arr1) + 1) + (UBound(arr2) - LBound(arr2) + 1)
        ReDim arr_(n_row)
        For i = LBound(arr1) To UBound(arr1)
            arr_(n_rec_row) = arr1(i)
            n_rec_row = n_rec_row + 1
        Next i
        For i = LBound(arr2) To UBound(arr2)
            arr_(n_rec_row) = arr2(i)
            n_rec_row = n_rec_row + 1
        Next i
        ArrayCombine = arr_
    End If
r = functime("ArrayCombine", tfunctime)
End Function

Function ArrayEmp2Space(ByVal array_in As Variant) As Variant
Dim tfunctime As Double
tfunctime = Timer
    If Not (IsEmpty(array_in)) Then
        seconddim = ArrayIsSecondDim(array_in)
        If Not (seconddim) Then
            For i = 1 To UBound(array_in, 1)
                If IsNumeric(array_in(i)) Then
                    If array_in(i) = 0 Then array_in(i) = " "
                    'If type_okrugl > 2 Then array_in(i) = Round(array_in(i), 4)
                    array_in(i) = "'" + CStr(array_in(i))
                Else
                    If Right$(array_in(i), 2) = "--" Then array_in(i) = " "
                End If
            Next
        Else
            For i = 1 To UBound(array_in, 1)
                For j = 1 To UBound(array_in, 2)
                    If IsNumeric(array_in(i, j)) Then
                        If array_in(i, j) = 0 And Not IsEmpty(array_in(i, j)) Then
                            array_in(i, j) = " "
                        End If
                        If type_okrugl > 2 Then array_in(i, j) = Round(array_in(i, j), 4)
                        array_in(i, j) = "'" + CStr(array_in(i, j))
                    Else
                        If Right$(array_in(i, j), 2) = "--" Then array_in(i, j) = " "
                    End If
                Next
            Next
        End If
    End If
    ArrayEmp2Space = array_in
r = functime("ArrayEmp2Space", tfunctime)
End Function

Function ArrayGetRowIndex(ByVal array_in As Variant, ByVal param As Variant, Optional ByVal n_col As Long) As Long
    index = Empty
    If IsEmpty(array_in) Then
        ArrayGetRowIndex = index
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
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
r = functime("ArrayGetRowIndex", tfunctime)
End Function

Function ArrayIsSecondDim(ByVal array_in As Variant) As Boolean
    If IsEmpty(array_in) Or Not IsArray(array_in) Then
        ArrayIsSecondDim = False
        Exit Function
    Else
Dim tfunctime As Double
tfunctime = Timer
        temp = 0
        On Error Resume Next
        n = UBound(array_in)
        For i = 1 To 60
            On Error Resume Next
            tmp = tmp + UBound(array_in, i)
        Next
        If tmp > n Then
            ArrayIsSecondDim = True
        Else
            ArrayIsSecondDim = False
        End If
    End If
tfunctime = functime("ArrayIsSecondDim", tfunctime)
End Function

Function ArrayIsEmpty(parArray As Variant) As Boolean
    ArrayIsEmpty = True
    On Error Resume Next
    ArrayIsEmpty = Not (LBound(parArray) <= UBound(parArray))
End Function

Function ArrayRedim(ByVal array_in As Variant, ByVal n_row As Long) As Variant
    If IsEmpty(array_in) Then ArrayRedim = Empty: Exit Function
    If n_row < 1 Then ArrayRedim = Empty: Exit Function
    If ArrayIsSecondDim(array_in) = False Then
        ReDim Preserve array_in(n_row)
        ArrayRedim = array_in
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
    n_col = UBound(array_in, 2)
    Dim arr(): ReDim arr(n_row, n_col)
    For i = 1 To n_row
        For j = 1 To n_col
            arr(i, j) = array_in(i, j)
        Next j
    Next i
    ArrayRedim = arr
r = functime("ArrayRedim", tfunctime)
End Function

Function ArrayRow(ByVal array_in As Variant, ByVal row As Long, Optional ByVal seconddim As Boolean = False) As Variant
    If IsEmpty(array_in) Then ArrayRow = Empty: Exit Function
    If ArrayIsSecondDim(array_in) = False Then ArrayRow = array_in: Exit Function
    If UBound(array_in, 1) < row Then ArrayRow = Empty: Exit Function
    If row < LBound(array_in, 1) Then ArrayRow = Empty: Exit Function
Dim tfunctime As Double
tfunctime = Timer
    n = UBound(array_in, 2)
    Dim out()
    If seconddim Then
        ReDim out(1, n)
        For i = 1 To n
            out(1, i) = array_in(row, i)
        Next i
    Else
        ReDim out(n)
        For i = 1 To n
            out(i) = array_in(row, i)
        Next i
    End If
    ArrayRow = out
r = functime("ArrayRow", tfunctime)
End Function

Function ArraySelectParam(ByVal array_in As Variant, ByVal param1 As Variant, Optional ByVal n_col1 As Long, Optional ByVal param2 As Variant, Optional ByVal n_col2 As Long) As Variant
    Dim arrout
    If IsEmpty(array_in) Then
        ArraySelectParam = Empty
        Exit Function
    End If
    If IsArray(param1) Or IsArray(param2) Then
        ArraySelectParam = ArraySelectParam_2(array_in, param1, n_col1, param2, n_col2)
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
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
            ArraySelectParam = ArrayRedim(arrout, n_row_s)
r = functime("ArraySelectParam", tfunctime)
            Exit Function
        Else
            ArraySelectParam = Empty
r = functime("ArraySelectParam", tfunctime)
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
r = functime("ArraySelectParam", tfunctime)
            Exit Function
        Else
            ArraySelectParam = Empty
r = functime("ArraySelectParam", tfunctime)
            Exit Function
        End If
    End If
End Function
Function ArraySelectParam_2(ByVal array_in As Variant, ByVal param1 As Variant, Optional ByVal n_col1 As Long, Optional ByVal param2 As Variant, Optional ByVal n_col2 As Long) As Variant
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
Dim tfunctime As Double
tfunctime = Timer
    Dim n_row As Integer
    Dim n_row_s As Integer
    Dim n_param As Integer
    If ArrayIsSecondDim(array_in) Then
        n_row = UBound(array_in, 1)
        n_param = UBound(array_in, 2)
        n_row_s = 0
        If n_col1 > n_param Then
            ArraySelectParam_2 = Empty
            Exit Function
        End If
        If n_row = 0 Or n_param = 0 Then
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
                        rparam = (StrComp(Right$(tparam1, 1), "?") = 0)
                        lParam = (StrComp(Left$(tparam1, 1), "?") = 0)
                        tparam1 = Trim$(Replace(tparam1, "?", vbNullString))
                        tpar = array_in(j, n_col1)
                        'If IsNumeric(tpar) Then tparam1 = ConvNum2Txt(tpar)
                        If rparam And lParam Then
                            If InStr(tpar, tparam1) > 0 Then flag1 = 1
                        Else
                            If lParam Then
                                If Right$(tpar, Len(tparam1)) = tparam1 Then flag1 = 1
                            Else
                                If rparam Then
                                    If Left$(tpar, Len(tparam1)) = tparam1 Then flag1 = 1
                                End If
                            End If
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
                                rparam = (StrComp(Right$(tparam2, 1), "?") = 0)
                                lParam = (StrComp(Left$(tparam2, 1), "?") = 0)
                                tparam2 = Trim$(Replace(tparam2, "?", vbNullString))
                                tpar = array_in(j, n_col2)
                                'If IsNumeric(tpar) Then tparam1 = ConvNum2Txt(tpar)
                                If rparam And lParam Then
                                    If InStr(tpar, tparam2) > 0 Then flag2 = 1
                                Else
                                    If lParam Then
                                        If Right$(tpar, Len(tparam2)) = tparam2 Then flag2 = 1
                                    Else
                                        If rparam Then
                                            If Left$(tpar, Len(tparam2)) = tparam1 Then flag2 = 1
                                        End If
                                    End If
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
            ArraySelectParam_2 = ArrayRedim(arrout, n_row_s)
r = functime("ArraySelectParam2", tfunctime)
            Exit Function
        Else
            ArraySelectParam_2 = Empty
r = functime("ArraySelectParam2", tfunctime)
            Exit Function
        End If
    Else
        n_row = UBound(array_in, 1)
        If n_row = 0 Then
            ArraySelectParam_2 = Empty
            Exit Function
        End If
        n_row_s = 0
        ReDim arrout(n_row)
        For j = 1 To n_row
            flag1 = 0 'Не записывать ни в коем случае
            For Each tparam1 In param1
                If array_in(j) = tparam1 Then
                    flag1 = 1 'Конечно, записывать
                Else
                    If InStr(tparam1, "?") > 0 Then
                        tpar = array_in(j)
                        If IsNumeric(tpar) Then tparam1 = ConvNum2Txt(tpar)
                        If Right$(tparam1, 1) = "?" And Left$(tparam1, 1) = "?" Then
                            tparam1 = Trim$(Replace(tparam1, "?", vbNullString))
                            If InStr(tpar, tparam1) > 0 Then flag1 = 1
                        End If
                        If Left$(tparam1, 1) = "?" Then
                            tparam1 = Trim$(Replace(tparam1, "?", vbNullString))
                            If Right$(tpar, Len(tparam1)) = tparam1 Then flag1 = 1
                        End If
                        If Right$(tparam1, 1) = "?" Then
                            tparam1 = Trim$(Replace(tparam1, "?", vbNullString))
                            If Left$(tpar, Len(tparam1)) = tparam1 Then flag1 = 1
                        End If
                    End If
                End If
                If flag1 = 1 Then Exit For
            Next
            If flag1 = 1 Then
                n_row_s = n_row_s + 1
                arrout(n_row_s) = array_in(j)
            End If
        Next j
        If n_row_s > 0 Then
            ReDim Preserve arrout(n_row_s)
            ArraySelectParam_2 = arrout
r = functime("ArraySelectParam2", tfunctime)
            Exit Function
        Else
            ArraySelectParam_2 = Empty
r = functime("ArraySelectParam2", tfunctime)
            Exit Function
        End If
    End If
End Function
Function ArraySort_2(ByVal array_in As Variant, ByVal nCol_arr As Variant, Optional ByVal type_sort As Long = 0) As Variant
    If IsEmpty(array_in) Then
        ArraySort_2 = Empty
        Exit Function
    End If
    If Not ArrayIsSecondDim(array_in) Then
        ArraySort_2 = Empty
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
    n_row = UBound(array_in, 1)
    n_col = UBound(array_in, 2)
    For i = LBound(nCol_arr) To UBound(nCol_arr)
        If nCol_arr(i) > n_col Then
            ArraySort_2 = Empty
            Exit Function
        End If
    Next i
    'Не мудрствуя лукаво - сформируем ключ из содержимого столбцов
    'К каждому такому ключу пропишем номера строк - просто текстом
    Set unic_dict = CreateObject("Scripting.Dictionary")
    For i = LBound(array_in) To UBound(array_in)
        str_key = vbNullString
        For j = LBound(nCol_arr) To UBound(nCol_arr)
            str_key = str_key & array_in(i, nCol_arr(j))
        Next j
        If unic_dict.Exists(str_key) Then
            unic_dict.Item(str_key) = unic_dict.Item(str_key) + ";" + CStr(i)
        Else
            unic_dict.Item(str_key) = CStr(i)
        End If
    Next i
    sort_key = ArraySort(unic_dict.keys, 1, type_sort)
    Dim array_out As Variant
    ReDim array_out(n_row, n_col)
    n_row_out = 0
    For Each str_key In sort_key
        'Формируем массив с номерами строй для данного ключа
        inx_row_arr = unic_dict.Item(str_key)
        If InStr(inx_row_arr, ";") = 0 Then
            inx_row_arr = Array(inx_row_arr)
        Else
            inx_row_arr = Split(inx_row_arr, ";")
        End If
        'Итак, у нас есть массив с номерами строк. Будем вытаскивать их в итоговый массив
        For j = LBound(inx_row_arr) To UBound(inx_row_arr)
            inx_row = CInt(inx_row_arr(j))
            n_row_out = n_row_out + 1
            For k = 1 To n_col
                array_out(n_row_out, k) = array_in(inx_row, k)
            Next k
        Next j
    Next
    If n_row_out <> n_row Then
        array_out = array_in
        r = dprint("Ошибка сортировки", 1)
    End If
''   при ошибках - вернуть как было
'    If n_col1 > n_col Or n_col2 > n_col Then
'        ArraySort_2 = Empty
'        Exit Function
'    End If
'    Dim array_out As Variant
'    sort_key = ArrayUniqValColumn(array_in, nCol1, type_sort)
'    For Each stkey In sort_key
'        array_by_key = ArraySelectParam(array_in, stkey, nCol1)
'        array_by_key = ArraySort(array_by_key, nCol2, type_sort)
'        array_out = ArrayCombine(array_out, array_by_key)
'    Next
    ArraySort_2 = array_out
r = functime("ArraySort_2", tfunctime)
End Function
Function ArraySort(ByVal array_in As Variant, Optional ByVal nCol As Long = 1, Optional ByVal type_sort As Long = 0) As Variant
'    type_sort = 0 - разбивка массива на числа и текст, с последующей сортировкой
'    type_sort = 1 - сортировка всего массива как чисел
'    type_sort = 2 - сортировка всего массива как текст
    If IsEmpty(array_in) Then
        ArraySort = Empty
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
    Dim array_in_str As Variant
    Dim array_in_num As Variant
    If ArrayIsSecondDim(array_in) Then
        If type_sort = 0 Then
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
        End If
        If type_sort = 1 Then array_in = ArraySortNum(array_in, nCol)
        If type_sort = 2 Then array_in = ArraySortABC(array_in, nCol)
    Else
        n_row = UBound(array_in)
        If LBound(array_in) = 0 Then n_row = n_row + 1
        If n_row <= 0 Then
            ArraySort = Empty
            Exit Function
        End If
        If type_sort = 0 Then
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
        If type_sort = 1 Then array_in = ArraySortNum(array_in, nCol)
        If type_sort = 2 Then array_in = ArraySortABC(array_in, nCol)
    End If
r = functime("ArraySort", tfunctime)
    ArraySort = array_in
End Function

Function ArraySortABC(ByVal array_in As Variant, Optional ByVal nCol As Long = 1) As Variant
    If ArrayIsSecondDim(array_in) Then
        r = ArraySortText_TwoDim(array_in, LBound(array_in, 1), UBound(array_in, 1), nCol)
    Else
        r = ArraySortText_OneDim(array_in, LBound(array_in), UBound(array_in))
    End If
    ArraySortABC = array_in
End Function
Function ArraySortNum(ByVal array_in As Variant, Optional ByVal nCol As Long = 1) As Variant
    If ArrayIsSecondDim(array_in) Then
        r = ArraySortNum_TwoDim(array_in, LBound(array_in, 1), UBound(array_in, 1), nCol)
    Else
        r = ArraySortNum_OneDim(array_in, LBound(array_in), UBound(array_in))
    End If
    ArraySortNum = array_in
End Function
Function ArraySortNum_OneDim(ByRef vArray As Variant, ByVal inLow As Long, ByVal inHi As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long
  tmpLow = inLow
  tmpHi = inHi
  pivot = vArray((inLow + inHi) \ 2)
  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend
     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend
     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend
  If (inLow < tmpHi) Then r = ArraySortNum_OneDim(vArray, inLow, tmpHi)
  If (tmpLow < inHi) Then r = ArraySortNum_OneDim(vArray, tmpLow, inHi)
End Function
Function ArraySortNum_TwoDim(ByRef vArray As Variant, ByVal inLow As Long, ByVal inHi As Long, ByVal nCol As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long
  tmpLow = inLow
  tmpHi = inHi
  pivot = vArray((inLow + inHi) \ 2, nCol)
  While (tmpLow <= tmpHi)
     While (vArray(tmpLow, nCol) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend
     While (pivot < vArray(tmpHi, nCol) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend
     If (tmpLow <= tmpHi) Then
        For k = LBound(vArray, 2) To UBound(vArray, 2)
            tmpSwap = vArray(tmpLow, k)
            vArray(tmpLow, k) = vArray(tmpHi, k)
            vArray(tmpHi, k) = tmpSwap
        Next k
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend
  If (inLow < tmpHi) Then r = ArraySortNum_TwoDim(vArray, inLow, tmpHi, nCol)
  If (tmpLow < inHi) Then r = ArraySortNum_TwoDim(vArray, tmpLow, inHi, nCol)
End Function

Function ArraySortText_TwoDim(ByRef strArray As Variant, ByVal intBottom As Long, ByVal intTop As Long, ByVal nCol As Long)
    Dim strPivot As String, strTemp As String
    Dim intBottomTemp As Long, intTopTemp As Long
    intBottomTemp = intBottom
    intTopTemp = intTop
    strPivot = strArray((intBottom + intTop) \ 2, nCol)
    Do While (intBottomTemp <= intTopTemp)
        Do While (CompareNaturalNum(strArray(intBottomTemp, nCol), strPivot) < 0 And intBottomTemp < intTop)
            intBottomTemp = intBottomTemp + 1
        Loop
        Do While (CompareNaturalNum(strPivot, strArray(intTopTemp, nCol)) < 0 And intTopTemp > intBottom) '
            intTopTemp = intTopTemp - 1
        Loop
        If intBottomTemp < intTopTemp Then
            For k = LBound(strArray, 2) To UBound(strArray, 2)
                strTemp = strArray(intBottomTemp, k)
                strArray(intBottomTemp, k) = strArray(intTopTemp, k)
                strArray(intTopTemp, k) = strTemp
            Next k
        End If
        If intBottomTemp <= intTopTemp Then
            intBottomTemp = intBottomTemp + 1
            intTopTemp = intTopTemp - 1
        End If
    Loop
    If (intBottom < intTopTemp) Then r = ArraySortText_TwoDim(strArray, intBottom, intTopTemp, nCol)
    If (intBottomTemp < intTop) Then r = ArraySortText_TwoDim(strArray, intBottomTemp, intTop, nCol)
End Function

Function ArraySortText_OneDim(ByRef strArray As Variant, ByVal intBottom As Long, ByVal intTop As Long)
    Dim strPivot As String, strTemp As String
    Dim intBottomTemp As Long, intTopTemp As Long
    intBottomTemp = intBottom
    intTopTemp = intTop
    tt = (intBottom + intTop) \ 2
    strPivot = strArray(tt)
    Do While (intBottomTemp <= intTopTemp)
        Do While (CompareNaturalNum(strArray(intBottomTemp), strPivot) < 0 And intBottomTemp < intTop)
            intBottomTemp = intBottomTemp + 1
        Loop
        Do While (CompareNaturalNum(strPivot, strArray(intTopTemp)) < 0 And intTopTemp > intBottom) '
            intTopTemp = intTopTemp - 1
        Loop
        If intBottomTemp < intTopTemp Then
            strTemp = strArray(intBottomTemp)
            strArray(intBottomTemp) = strArray(intTopTemp)
            strArray(intTopTemp) = strTemp
        End If
        If intBottomTemp <= intTopTemp Then
            intBottomTemp = intBottomTemp + 1
            intTopTemp = intTopTemp - 1
        End If
    Loop
    If (intBottom < intTopTemp) Then r = ArraySortText_OneDim(strArray, intBottom, intTopTemp)
    If (intBottomTemp < intTop) Then r = ArraySortText_OneDim(strArray, intBottomTemp, intTop)
End Function
Function CompareNaturalNum(string1 As Variant, string2 As Variant) As Long
Dim tfunctime As Double
tfunctime = Timer
    Dim n1 As Double, n2 As Double
    Dim iPosOrig1 As Long, iPosOrig2 As Long
    Dim iPos1 As Long, iPos2 As Long
    Dim nOffset1 As Long, nOffset2 As Long
    If Not (IsNull(string1) Or IsNull(string2)) Then
        iPos1 = 1
        iPos2 = 1
        Do While iPos1 <= Len(string1)
            If iPos2 > Len(string2) Then
                CompareNaturalNum = 1
                Exit Function
            End If
            If isDigit(string1, iPos1) Then
                If Not isDigit(string2, iPos2) Then
                    CompareNaturalNum = -1
                    Exit Function
                End If
                iPosOrig1 = iPos1
                iPosOrig2 = iPos2
                Do While isDigit(string1, iPos1)
                    iPos1 = iPos1 + 1
                Loop

                Do While isDigit(string2, iPos2)
                    iPos2 = iPos2 + 1
                Loop

                nOffset1 = (iPos1 - iPosOrig1)
                nOffset2 = (iPos2 - iPosOrig2)
                tt1_ = Mid$(string1, iPosOrig1, nOffset1)
                tt2_ = Mid$(string2, iPosOrig2, nOffset2)
                n1 = val(tt1_)
                n2 = val(tt2_)

                If (n1 < n2) Then
                    CompareNaturalNum = -1
tfunctime = functime("CompareNaturalNum", tfunctime)
                    Exit Function
                ElseIf (n1 > n2) Then
                    CompareNaturalNum = 1
tfunctime = functime("CompareNaturalNum", tfunctime)
                    Exit Function
                End If

                ' front padded zeros (put 01 before 1)
                If (n1 = n2) Then
                    If (nOffset1 > nOffset2) Then
                        CompareNaturalNum = -1
tfunctime = functime("CompareNaturalNum", tfunctime)
                        Exit Function
                    ElseIf (nOffset1 < nOffset2) Then
                        CompareNaturalNum = 1
tfunctime = functime("CompareNaturalNum", tfunctime)
                        Exit Function
                    End If
                End If
            ElseIf isDigit(string2, iPos2) Then
                CompareNaturalNum = 1
tfunctime = functime("CompareNaturalNum", tfunctime)
                Exit Function
            Else
                If (Mid$(string1, iPos1, 1) < Mid$(string2, iPos2, 1)) Then
                    CompareNaturalNum = -1
tfunctime = functime("CompareNaturalNum", tfunctime)
                    Exit Function
                ElseIf (Mid$(string1, iPos1, 1) > Mid$(string2, iPos2, 1)) Then
                    CompareNaturalNum = 1
tfunctime = functime("CompareNaturalNum", tfunctime)
                    Exit Function
                End If

                iPos1 = iPos1 + 1
                iPos2 = iPos2 + 1
            End If
        Loop
        If Len(string2) > Len(string1) Then
            CompareNaturalNum = -1
tfunctime = functime("CompareNaturalNum", tfunctime)
            Exit Function
        End If
    Else
        If IsNull(string1) And Not IsNull(string2) Then
            CompareNaturalNum = -1
tfunctime = functime("CompareNaturalNum", tfunctime)
            Exit Function
        ElseIf IsNull(string1) And IsNull(string2) Then
            CompareNaturalNum = 0
tfunctime = functime("CompareNaturalNum", tfunctime)
            Exit Function
        ElseIf Not IsNull(string1) And IsNull(string2) Then
            CompareNaturalNum = 1
tfunctime = functime("CompareNaturalNum", tfunctime)
            Exit Function
        End If
    End If
tfunctime = functime("CompareNaturalNum_", tfunctime)
End Function
Function isDigit(ByVal str As String, pos As Long) As Boolean
    Dim iCode As Long
    If pos <= Len(str) Then
        iCode = Asc(Mid$(str, pos, 1))
        If iCode >= 48 And iCode <= 57 Then isDigit = True
    End If
End Function
Function ArrayTranspose(ByVal array_in As Variant) As Variant
    If IsEmpty(array_in) Then
        ArrayTranspose = Empty
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
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
r = functime("ArrayTranspore", tfunctime)
End Function

Function ArrayUniqValColumn(ByVal array_in As Variant, Optional ByVal cols As Long = 1, Optional ByVal type_sort As Long = 0) As Variant
    If IsEmpty(array_in) Or Not IsArray(array_in) Or ArrayIsEmpty(array_in) Then
        ArrayUniqValColumn = Empty
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
    Set unic_dict = CreateObject("Scripting.Dictionary")
    flag_twodim = 0
    If ArrayIsSecondDim(array_in) Then flag_twodim = 1
    If flag_twodim Then
        n_row = UBound(array_in, 1)
        n_start = LBound(array_in, 1)
    Else
        n_row = UBound(array_in)
        n_start = LBound(array_in)
    End If
    For i = n_start To n_row
        If flag_twodim Then
            var = array_in(i, cols)
        Else
            var = array_in(i)
        End If
        If Not unic_dict.Exists(var) Then
            flag = 1
            If IsEmpty(var) Then flag = 0
            If Not IsNumeric(var) Then
                If Len(var) = 0 Then flag = 0
                If var = " " Then flag = 0
            End If
            If flag Then unic_dict.Item(var) = var
        End If
    Next i
    If unic_dict.Count > 0 Then
        Dim array_out()
        ReDim array_out(unic_dict.Count)
        n_row = 0
        For Each k In unic_dict.keys
            n_row = n_row + 1
            array_out(n_row) = k
        Next
        array_out = ArraySort(array_out, 1, type_sort)
        ArrayUniqValColumn = array_out
    Else
        ArrayUniqValColumn = Empty
    End If
r = functime("ArrayUniqValColumn", tfunctime)
End Function
Function ArrayHasElement(ByVal array_in As Variant, ByVal elem As Variant, Optional ByVal cols As Long = 1) As Boolean
    If IsEmpty(array_in) Or Not IsArray(array_in) Or ArrayIsEmpty(array_in) Then
        ArrayHasElement = False
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
    If ArrayIsSecondDim(array_in) Then
        If cols = 0 Then cols = 1
        For i = 1 To UBound(array_in, 1)
            If array_in(i, cols) = elem Then
                ArrayHasElement = True
r = functime("ArrayHasElement", tfunctime)
                Exit Function
            End If
        Next
    Else
        For i = LBound(array_in) To UBound(array_in)
            If array_in(i) = elem Then
                ArrayHasElement = True
r = functime("ArrayHasElement", tfunctime)
                Exit Function
            End If
        Next
    End If
r = functime("ArrayHasElement", tfunctime)
End Function

Function ControlSumEl(ByVal array_in As Variant) As String
    Dim param
    isel = 0
Dim tfunctime As Double
tfunctime = Timer
    If ArrayIsSecondDim(array_in) Then
        Dim t: ReDim t(UBound(array_in, 2))
        For i = 1 To UBound(array_in, 2)
            t(i) = array_in(1, i)
        Next i
        array_in = t
    End If
    'marka = array_in(col_marka)
    subpos = array_in(col_sub_pos)
    type_el = array_in(col_type_el)
    pos = array_in(col_pos)
    qty = array_in(col_qty)
    chksum = array_in(col_chksum)
    sparent = array_in(col_parent)
    nfloor = 0
    t_floor = array_in(col_floor)
    If sparent = 0 Then sparent = "-"
    If subpos = 0 Then subpos = "-"
    If pos = 0 Then pos = "-"
    If type_el = t_arm Then
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
            Else
                param(9) = "l" + ConvNum2Txt(Int(Length))
                param(10) = "f0"
                param(11) = 0
            End If
            If gnut Then
                param(12) = "g3"
            Else
                param(12) = "g0"
            End If
        End If
        If type_el = t_prokat Then
            isel = 1
            type_konstr = array_in(col_pr_type_konstr)
            gost_st = array_in(col_pr_gost_st)
            st = array_in(col_pr_st)
            gost_prof = array_in(col_pr_gost_prof)
            prof = array_in(col_pr_prof)
            naen = array_in(col_pr_naen)
            Length = array_in(col_pr_length)
            'Weight = array_in(col_pr_weight)
            'Если в примечании стоит п.м. - выводим в п.м.
            pm = False
            If InStr(array_in(col_pr_naen), "@@") > 0 Then
                prim = Split(array_in(col_pr_naen), "@@")(1)
                If InStr(prim, "п.м.") > 0 Then pm = True
            End If
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
            If pm Then
                param(11) = "lpm"
            Else
                If Not IsNumeric(Length) Then
                    r = LogWrite("Неизвестная длина проката - " & pos & suboos, naen, "ОШИБКА")
                Else
                    param(11) = "l" + ConvNum2Txt(Int(Length) * 1000)
                End If
                If Not IsNumeric(naen) Then param(11) = param(11) + naen
            End If
        End If
        If type_el = t_mat Then
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
        End If
        If type_el = t_izd Or type_el = t_perem Then
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
        End If
        If type_el = t_subpos Or type_el = t_perem_m Then
            isel = 1
            obozn = array_in(col_m_obozn)
            naen = array_in(col_m_naen)
            Weight = array_in(col_m_weight)
            edizm = array_in(col_m_edizm)
            sparam = array_in(col_param)
            
            ReDim param(7)
            param(1) = subpos
            param(2) = "_"
            param(3) = subpos
            param(4) = sparent
            param(5) = "_"
            param(6) = naen
            param(7) = sparam
        End If
        If type_el = t_wind Then
            isel = 1
            obozn = array_in(col_w_obozn)
            naen = array_in(col_w_naen)

            ReDim param(8)
            param(1) = pos
            param(2) = subpos
            param(3) = "_"
            param(4) = t_floor
            param(5) = nfloor
            param(6) = "_"
            param(7) = obozn
            param(8) = naen
        End If
    If Not IsEmpty(this_sheet_option) Then
        If spec_version > 1 And this_sheet_option.Item("qty_one_floor") And Not IsEmpty(param) Then
            n_param = UBound(param, 1)
            ReDim Preserve param(n_param + 3)
            param(n_param + 1) = "_"
            param(n_param + 2) = 0
            param(n_param + 3) = array_in(col_floor)
        End If
    End If
    control_sum = vbNullString
    If isel Then
        For i = 1 To UBound(param, 1)
            var = param(i)
            If IsNumeric(var) Then param(i) = CStr(var)
        Next i
        control_sum = Join(param, vbNullString)
        d_txt = Array(" ", "--", "x", "х", "-", " ")
        For i = 1 To UBound(d_txt)
            If InStr(control_sum, d_txt(i)) Then
                control_sum = Trim$(Replace(control_sum, d_txt(i), vbNullString))
            End If
        Next
    End If
tfunctime = functime("ControlSumEl", tfunctime)
    ControlSumEl = control_sum
End Function

Function ConvNum2Txt(ByVal var As Variant, Optional ByVal n_end As Long, Optional ByVal force_zero As Boolean = False) As String
Dim tfunctime As Double
tfunctime = Timer
    txt = vbNullString
    If IsNumeric(var) Then
        If var = 0 Then
            txt = vbNullString
        Else
            txt = Trim$(CStr(var))
            If Left$(txt, 1) = "." Or Left$(txt, 1) = "," Then txt = "0" + txt
        End If
        txt = Replace(txt, ".", ",")
        If force_zero And InStr(txt, ",") <= 0 Then txt = txt + ",0"
        If n_end > 0 And InStr(txt, ",") > 0 Then
            var = Split(txt, ",")
            n_zero = n_end - Len(var(1))
            If n_zero > 0 Then
                For i = 1 To n_zero
                    txt = txt + "0"
                Next i
            End If
        End If
    Else
        txt = var
    End If
    ConvNum2Txt = txt
r = functime("ConvNum2Txt", tfunctime)
End Function

Function ConvTxt2Num(ByVal x As Variant) As Variant
Dim tfunctime As Double
tfunctime = Timer
    If IsError(x) Then
        ConvTxt2Num = x
        Exit Function
    End If
    If IsNumeric(x) Then
        out = CDbl(x)
    Else
        x_tmp = x
        x = Replace(x, " ", vbNullString)
        x = Replace(x, ".", ",")
        x = Replace(x, "'", vbNullString)
        x = Trim(x)
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
r = functime("ConvTxt2Num", tfunctime)
End Function
Function ConvNum2Otm(ByVal x As Variant) As Variant
Dim tfunctime As Double
tfunctime = Timer
    If Not IsNumeric(x) Then x = ConvTxt2Num(x)
    If IsNumeric(x) Then
        If Abs(x) < 0.001 Then
            x = "'0.000"
        Else
            txt = Format$(x, "#,###0.000")
            txt = Replace(txt, ",", ".")
            sgn_txt = "+": If x < 0 Then sgn_txt = "-"
            x = "'" + sgn_txt + txt
        End If
    End If
    ConvNum2Otm = x
r = functime("ConvNum2Otm", tfunctime)
End Function

Function DataAddNullSubpos(ByVal array_in As Variant) As Variant
    'TODO переделать под новую систему
    'Если в массиве есть элементы, состоящие в сборках, но маркировки сборок (t_subpos) нет - добавляет строки маркировок сборок
    If IsEmpty(array_in) Then
        DataAddNullSubpos = Array(Empty, Empty)
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
    Dim add_subpos
    Dim out_subpos
    Set name_subpos = DataNameSubpos(exist_subpos) 'Получим для них имена
    arr_subpos = ArrayUniqValColumn(array_in, col_sub_pos)
    If IsEmpty(arr_subpos) Then
        DataAddNullSubpos = Array(Empty, Empty)
        Exit Function
    End If
    add_txt = Empty
    For Each current_subpos In arr_subpos
        If current_subpos <> "-" Then
            'Проверяем - есть ли маркировка для главных сборок
            seach_subpos = ArraySelectParam_2(array_in, current_subpos, col_sub_pos, Array(t_subpos, t_perem_m, t_wind), col_type_el)
            If IsEmpty(seach_subpos) Then
                If name_subpos.Exists(current_subpos) Then
                    naen = name_subpos(current_subpos)(1)
                    obozn = name_subpos(current_subpos)(2)
                Else
                    naen = current_subpos
                    obozn = "!!!"
                End If
                nfloor = 1
                If spec_version > 1 And this_sheet_option.Item("qty_one_floor") Then  'Учтём кол-во этажей
                    el_in_subpos = ArraySelectParam(array_in, current_subpos, col_sub_pos)
                    un_floor = ArrayUniqValColumn(el_in_subpos, col_floor)
                    If Not IsEmpty(un_floor) Then nfloor = UBound(un_floor, 1)
                End If
                ReDim add_subpos(nfloor, max_col)
                For i = 1 To nfloor
                    If this_sheet_option.Item("qty_one_floor") Then
                        tadd_txt = current_subpos + " на отм. " + CStr(un_floor(i))
                    Else
                        tadd_txt = current_subpos
                    End If
                    If IsEmpty(add_txt) Then
                        add_txt = tadd_txt
                    Else
                        add_txt = tadd_txt + ";" + add_txt
                    End If
                    add_subpos(i, col_sub_pos) = current_subpos
                    add_subpos(i, col_type_el) = t_subpos
                    add_subpos(i, col_pos) = current_subpos
                    add_subpos(i, col_m_naen) = Replace(naen, "@", subpos_delim)
                    add_subpos(i, col_m_obozn) = obozn
                    add_subpos(i, col_qty) = 1
                    If spec_version > 1 And this_sheet_option.Item("qty_one_floor") Then
                        t_floor = un_floor(i)
                        el_in_floor = ArraySelectParam(el_in_subpos, t_floor, col_floor)
                        add_subpos(i, col_nfloor) = 0
                        add_subpos(i, col_floor) = el_in_floor(1, col_floor)
                        add_subpos(i, col_param) = el_in_floor(1, col_param)
                    End If
                    add_subpos(i, col_chksum) = ControlSumEl(ArrayRow(add_subpos, i))
                Next i
                out_subpos = ArrayCombine(out_subpos, add_subpos)
            End If
        End If
    Next
    out_subpos = DataCheck(out_subpos)
    DataAddNullSubpos = Array(out_subpos, add_txt)
r = functime("DataAddNullSubpos", tfunctime)
End Function

Function CheckGost(ByVal obozn As String) As String
    If IsEmpty(obozn) Then obozn = ""
    If Len(obozn) > 0 Then
        new_obozn = ""
        If swap_gost.Exists(obozn) Then new_obozn = swap_gost.Item(obozn)
        If IsEmpty(new_obozn) Then new_obozn = ""
        If Len(new_obozn) > 0 Then obozn = new_obozn
    End If
    CheckGost = obozn
End Function

Function DataCheck(ByVal array_in As Variant) As Variant
    If IsEmpty(array_in) Then DataCheck = Empty: Exit Function
    If IsEmpty(pr_adress) Then r = ReadPrSortament()
Dim tfunctime As Double
tfunctime = Timer
    n_col = UBound(array_in, 2)
    n_row_in = UBound(array_in, 1)
    n_ingore = 0
    n_error = 0
    Dim out_data: ReDim out_data(n_row_in, n_col): n_row = 0
    For i = 1 To n_row_in
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
        ignore_flag = False
        If InStr(array_in(i, col_sub_pos), ignore_pos) Then ignore_flag = True
        If InStr(array_in(i, col_parent), ignore_pos) Then ignore_flag = True
        If InStr(array_in(i, col_marka), ignore_pos) Then ignore_flag = True
        If ignore_flag = False And Not IsEmpty(this_sheet_option) Then
            subpos_filter = True
            subpos_filter = DataFilter(array_in(i, col_sub_pos), this_sheet_option.Item("arr_subpos_add"), this_sheet_option.Item("arr_subpos_del"))
            If subpos_filter = True And Len(array_in(i, col_parent)) > 0 And array_in(i, col_parent) <> "-" Then subpos_filter = DataFilter(array_in(i, col_parent), this_sheet_option.Item("arr_subpos_add"), this_sheet_option.Item("arr_subpos_del"))
            If subpos_filter = False Then ignore_flag = True
        End If
        If type_el > 0 And ignore_flag = False Then
            If type_el = t_arm Then
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
            End If
            If type_el = t_prokat Then
                'Сразу отфильтруем по типу элемента КМ
                constr_filter = DataFilter(array_in(i, col_pr_type_konstr), this_sheet_option.Item("arr_typeKM_add"), this_sheet_option.Item("arr_typeKM_del"))
                If constr_filter Then
                    name_pr = GetShortNameForGOST(array_in(i, col_pr_gost_prof))
                    If InStr(1, name_pr, "Лист") > 0 Then
                        naen_plate = SpecMetallPlate(array_in(i, col_pr_prof), array_in(i, col_pr_naen), array_in(i, col_pr_length), array_in(i, col_pr_weight), array_in(i, col_chksum))
                        array_in(i, col_pr_length) = naen_plate(2) * 1000
                        array_in(i, col_pr_weight) = naen_plate(3)
                        If naen_plate(7) <= 0 Then ignore_flag = True
                    End If
                    If Not IsNumeric(array_in(i, col_pr_weight)) Then
                        array_in(i, col_pr_weight) = 0.01
                        r = LogWrite("Нулевая масса элемента проката", CStr(array_in(i, col_sub_pos)) + " " + CStr(array_in(i, col_pos)), array_in(i, col_pr_naen))
                        n_error = n_error + 1
                    End If
                    If Not IsNumeric(array_in(i, col_pr_length)) Then
                        r = LogWrite("Непонятная длина проката", CStr(array_in(i, col_sub_pos)) + " " + CStr(array_in(i, col_pos)), array_in(i, col_pr_naen))
                        n_error = n_error + 1
                    End If
                    array_in(i, col_pr_gost_st) = pr_adress.Item(array_in(i, col_pr_st))
                    array_in(i, col_pr_gost_st) = CheckGost(array_in(i, col_pr_gost_st))
                    array_in(i, col_pr_gost_prof) = CheckGost(array_in(i, col_pr_gost_prof))
                Else
                    ignore_flag = True
                End If
            End If
            If type_el = t_mat_spc Then
                If InStr(array_in(i, col_marka), subpos_delim) Then
                    array_in(i, col_sub_pos) = Split(array_in(i, col_marka), subpos_delim)(1)
                Else
                    array_in(i, col_sub_pos) = array_in(i, col_marka)
                End If
                array_in(i, col_pos) = Empty
                array_in(i, col_type_el) = t_mat
                array_in(i, col_m_weight) = "-"
            End If
            If type_el = t_wind Then
                If IsNumeric(array_in(i, col_w_weight)) = False Then array_in(i, col_w_weight) = 0
                array_in(i, col_w_naen) = Replace(array_in(i, col_w_naen), "  ", " ")
                array_in(i, col_w_naen) = Replace(array_in(i, col_w_naen), "  ", " ")
                array_in(i, col_w_naen) = Replace(array_in(i, col_w_naen), "  ", " ")
            End If
            If type_el <> t_arm And type_el <> t_prokat Then array_in(i, col_m_obozn) = CheckGost(array_in(i, col_m_obozn))
            If type_el = t_mat Then
' TODO дописать в архи деление материалов по типам конструкций
                obozn = array_in(i, col_m_obozn)
                type_konstr = GetZoneParam(array_in(i, col_param), "tk")
                If Not IsEmpty(type_konstr) Then
                    constr_filter = DataFilter(type_konstr, this_sheet_option.Item("arr_typeKM_add"), this_sheet_option.Item("arr_typeKM_del"))
                    If constr_filter = False Then ignore_flag = True
                End If
            End If
            If Len(array_in(i, col_sub_pos)) = 0 Then array_in(i, col_sub_pos) = "-"
            If array_in(i, col_sub_pos) = " " Then array_in(i, col_sub_pos) = "-"
            If array_in(i, col_sub_pos) = 0 Then array_in(i, col_sub_pos) = "-"
            If array_in(i, col_sub_pos) = "-" Then array_in(i, col_parent) = "-"
            If IsEmpty(array_in(i, col_parent)) Then array_in(i, col_parent) = "-"
            array_in(i, col_sub_pos) = Replace(array_in(i, col_sub_pos), "@", subpos_delim)
            array_in(i, col_parent) = Replace(array_in(i, col_parent), "@", subpos_delim)
            array_in(i, col_marka) = Replace(array_in(i, col_marka), "@", subpos_delim)
            array_in(i, col_pos) = Replace(array_in(i, col_pos), "@", subpos_delim)
            array_in(i, col_pos) = Replace(array_in(i, col_pos), ",", ".")
        End If
        If ignore_flag Then
            n_ingore = n_ingore + 1
        Else
            'Вычисление и проверка контрольных сумм
            array_in(i, col_chksum) = ControlSumEl(ArrayRow(array_in, i))
            n_row = n_row + 1
            For j = 1 To n_col
                If IsNumeric(array_in(i, j)) Then
                    out_data(n_row, j) = array_in(i, j)
                Else
                    out_data(n_row, j) = Trim$(Replace(array_in(i, j), "  ", " "))
                End If
            Next j
        End If
    Next i
    If n_ingore > 0 Then r = LogWrite("Пропущено элементов ", vbNullString, n_ingore)
    If n_error > 0 Then
        MsgBox ("Ошибка в данных файла, см. лист " + log_sheet_name)
        DataCheck = Empty
    End If
    If n_row Then
        out_data = ArrayRedim(out_data, n_row)
        DataCheck = out_data
    Else
        DataCheck = Empty
    End If
tfunctime = functime("DataCheck", tfunctime)
End Function

Function DataIsOtd(ByVal array_in As Variant) As Boolean
    If IsEmpty(array_in) Then
        DataIsOtd = False
        Exit Function
    End If
    n_col = UBound(array_in, 2)
    otd_version = 0
    If array_in(1, col_s_type) = "ЗОНА" And (n_col = max_col_type_1_1 Or n_col = max_col_type_2_1 Or n_col = max_col_type_3_1) Then otd_version = 1
    If array_in(1, col_s_type) = "ЗОНА" And (n_col = max_col_type_1_2 Or n_col = max_col_type_2_2 Or n_col = max_col_type_3_2) Then otd_version = 2
    If otd_version = 2 Then
        If array_in(1, col_s_type_pol_zone) = "Не задан" And array_in(1, col_s_type_pot_zone) = 0 And n_col = max_col_type_2_1 Then otd_version = 1
    End If
    If otd_version = 1 Then
        col_s_type_el = col_s_type_el_1
        col_s_type_pol = col_s_type_pol_1
        col_s_area_pol = col_s_area_pol_1
        col_s_perim_pol = col_s_perim_pol_1
        col_s_n_mun_zone = col_s_n_mun_zone_1
        col_s_mun_zone = col_s_mun_zone_1
        col_s_tipverh_l = col_s_tipverh_l_1
        col_s_tipl_l = col_s_tipl_l_1
        col_s_tipniz_l = col_s_tipniz_l_1
        col_s_tippl_l = col_s_tippl_l_1
        col_s_areaverh_l = col_s_areaverh_l_1
        col_s_areal_l = col_s_areal_l_1
        col_s_areaniz_l = col_s_areaniz_l_1
        col_s_areapl_l = col_s_areapl_l_1
        max_s_col = max_s_col_1
        max_col_type_1 = max_col_type_1_1
        max_col_type_2 = max_col_type_2_1
        max_col_type_3 = max_col_type_3_1
    End If
    If otd_version = 2 Then
        col_s_type_el = col_s_type_el_2
        col_s_type_pol = col_s_type_pol_2
        col_s_area_pol = col_s_area_pol_2
        col_s_perim_pol = col_s_perim_pol_2
        col_s_n_mun_zone = col_s_n_mun_zone_2
        col_s_mun_zone = col_s_mun_zone_2
        col_s_tipverh_l = col_s_tipverh_l_2
        col_s_tipl_l = col_s_tipl_l_2
        col_s_tipniz_l = col_s_tipniz_l_2
        col_s_tippl_l = col_s_tippl_l_2
        col_s_areaverh_l = col_s_areaverh_l_2
        col_s_areal_l = col_s_areal_l_2
        col_s_areaniz_l = col_s_areaniz_l_2
        col_s_areapl_l = col_s_areapl_l_2
        max_s_col = max_s_col_2
        max_col_type_1 = max_col_type_1_2
        max_col_type_2 = max_col_type_2_2
        max_col_type_3 = max_col_type_3_2
    End If
    If otd_version > 0 Then
        DataIsOtd = True
    Else
        DataIsOtd = False
    End If
End Function

Function DataIsShort(ByVal array_in As Variant) As Boolean
    If IsEmpty(array_in) Then
        DataIsShort = False
        Exit Function
    End If
'Если номер столбца с типами элементов отличается от col_type_el - то первый столбец, скорее всего - количество элементов
'    colum = 0
'    n_row = Int(UBound(array_in, 1) / 2) + 1
'    For j = col_type_el To col_type_el + 1
'        n = 0
'        For i = 1 To n_row
'            If type_el_name.exists(array_in(i, j)) Then
'                n = n + 1
'            End If
'        Next i
'        If n > 0 And colum = 0 Then colum = j
'    Next j
'    res = False
'    If colum <> col_type_el Then res = True
    res = True
    DataIsShort = res
End Function

Function DataIsSpec(ByVal array_in As Variant) As Boolean
Dim tfunctime As Double
tfunctime = Timer
    If IsEmpty(array_in) Then
        DataIsSpec = False
        Exit Function
    End If
    n_row = Int(UBound(array_in, 1) / 2) + 1
    n = 0
    For i = 1 To n_row
        If type_el_name.Exists(array_in(i, col_type_el)) Then n = n + 1
    Next i
    If n > 0 Then DataIsSpec = True Else DataIsSpec = False
tfunctime = functime("DataIsSpec", tfunctime)
End Function

Function DataGetVersion(ByRef array_in As Variant) As Long
Dim tfunctime As Double
tfunctime = Timer
    s_version = -1
    If IsEmpty(array_in) Then
        DataGetVersion = s_version
        Exit Function
    End If
    If InStr(array_in(1, col_qty), "%%") = 0 And InStr(array_in(1, col_qty), "v") = 0 Then
        s_version = 1
    Else
        ttxt = Left$(array_in(1, col_qty), InStr(array_in(1, col_qty), "%%") - 1)
        s_version = Int(Right$(ttxt, Len(ttxt) - 1))
        If Not IsEmpty(array_in(1, 19)) Then s_version = 3
    End If
    DataGetVersion = s_version
tfunctime = functime("DataGetVersion", tfunctime)
End Function

Function DataConvertVersion(ByVal array_in As Variant) As Variant
Dim tfunctime As Double
tfunctime = Timer
    add_row = 0: If DataIsShort(array_in) Then add_row = 1
    n_row = UBound(array_in, 1)
    Dim pos_out(): ReDim pos_out(n_row, max_col + add_row)
    
    'Обработаем данные по окнам/дверям с длинным именем
    If spec_version = 3 Then
        wind_1 = ArraySelectParam(array_in, 61, 4)
        If Not IsEmpty(wind_1) Then
            wind_2 = ArraySelectParam(array_in, 62, 4)
            If Not IsEmpty(wind_2) Then
                For i = 1 To UBound(wind_1, 1)
                    GUID_w = wind_1(i, 15)
                        par_ = wind_1(i, 2)
                        pos_ = wind_1(i, 3)
                        qty_ = wind_1(i, 1)
                    wind = ArraySelectParam(wind_2, GUID_w, 7)
                    If Not IsEmpty(wind) Then
                        naen = wind(1, 5)
                        par = wind(1, 2)
                        pos = wind(1, 3)
                        qty = wind(1, 1)
                        If par = par_ And pos_ = pos And qty_ = qty Then
                            wind_1(i, 10) = naen
                            wind_1(i, 15) = wind_1(i, 15) + par + pos + str(qty)
                        Else
                            MsgBox ("Проверь окна/двери с длинным названием")
                        End If
                    End If
                    hh = 1
                Next i
                For i = 1 To n_row
                    If array_in(i, 4) = 61 Then
                        GUID_w = array_in(i, 15)
                        par_ = array_in(i, 2)
                        pos_ = array_in(i, 3)
                        qty_ = array_in(i, 1)
                        find_str = GUID_w + par + pos + str(qty)
                        wind = ArraySelectParam(wind_1, find_str, 15)
                        If Not IsEmpty(wind) Then
                            array_in(i, 4) = 60
                            array_in(i, 10) = wind(1, 10)
                        End If
                    End If
                Next i
            End If
        End If
    End If

    For i = 1 To n_row
        If spec_version = 2 Then
            For j = 1 To col_pos + add_row
                pos_out(i, j) = array_in(i, j)
            Next j
            For j = col_pos + add_row + 1 To UBound(array_in, 2) - 3
                pos_out(i, j) = array_in(i, j + 3)
            Next j
            pos_out(i, col_nfloor + add_row) = 0 'array_in(i, col_pos + 1 + add_row)
            pos_out(i, col_floor + add_row) = array_in(i, col_pos + 2 + add_row)
            pos_out(i, col_param + add_row) = array_in(i, col_pos + 3 + add_row)
        End If
        If spec_version = 3 Then
            For j = 1 To col_pos + add_row
                pos_out(i, j) = array_in(i, j)
            Next j
            For j = col_pos + add_row + 1 To UBound(array_in, 2) - 3
                pos_out(i, j) = array_in(i, j + 3)
            Next j
            pos_out(i, col_param + add_row) = array_in(i, col_pos + 3 + add_row)
            pos_out(i, col_nfloor + add_row) = 0
            inx_col_floor = UBound(array_in, 2)
            For j = UBound(array_in, 2) To 1 Step -1
                If Not IsEmpty(array_in(i, j)) Then
                    inx_col_floor = j
                    Exit For
                End If
            Next j
            pos_out(i, col_floor + add_row) = array_in(i, inx_col_floor)
        End If
    Next i
    DataConvertVersion = pos_out
tfunctime = functime("DataConvertVersion", tfunctime)
End Function

Function DataIsWall(ByVal nm As String) As Variant
    array_in = ReadTxt(ThisWorkbook.path & "\import\" & nm, 1, vbTab, vbNewLine)
    n_row = UBound(array_in, 1)
    Dim pos_out(): ReDim pos_out(n_row - 1, max_col)
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
            t_start = Trim$(Mid$(naen, 1, p_start))
            t_end = Trim$(Mid$(naen, p_end, Len(naen)))
            obozn = Trim$(Mid$(naen, p_start + 2, p_end - p_start - 3))
            naen = t_start & " " & t_end
        End If
        t_sl = array_in(i, 3)
        If t_sl > 0.1 And InStr(naen, "t=") = 0 Then naen = naen & " t=" & ConvNum2Txt(t_sl) & "мм."
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
Dim tfunctime As Double
tfunctime = Timer
    Set name_subpos = CreateObject("Scripting.Dictionary")

    sheet = "Имена сборок_поз"
    If SheetExist(sheet) Then
        array_in = ReadPos(sheet)
        all_subpos_in_sheet = ArraySelectParam(array_in, t_subpos, col_type_el)
        If Not IsEmpty(all_subpos_in_sheet) Then all_subpos = ArrayCombine(all_subpos, all_subpos_in_sheet)
    End If
    
    nm = ThisWorkbook.ActiveSheet.Name
    type_sheet = SpecGetType(nm)
    If Not IsEmpty(type_sheet) And type_sheet <> 3 Then
        sheet = Trim$(Split(nm, "_")(0)) & "_поз"
    Else
        sheet = nm & "_поз"
    End If
    If SheetExist(sheet) Then
        array_in = ReadPos(sheet)
        all_subpos_in_sheet = ArraySelectParam(array_in, t_subpos, col_type_el)
        If Not IsEmpty(all_subpos_in_sheet) Then all_subpos = ArrayCombine(all_subpos, all_subpos_in_sheet)
    End If

    If Not IsEmpty(sub_pos_arr) Then all_subpos = ArrayCombine(all_subpos, sub_pos_arr)
    If Not IsEmpty(all_subpos) Then
        For i = 1 To UBound(all_subpos, 1)
            subpos = all_subpos(i, col_sub_pos)
            naen = all_subpos(i, col_m_naen)
            obozn = all_subpos(i, col_m_obozn)
            qty = all_subpos(i, col_qty)
            If name_subpos.Exists(subpos) Then
                tqty = name_subpos.Item(subpos)(3) + qty
                name_subpos.Item(subpos) = Array(naen, obozn, tqty)
            Else
                name_subpos.Item(subpos) = Array(naen, obozn, qty)
            End If
        Next i
    End If
    
    Set DataNameSubpos = name_subpos
tfunctime = functime("DataNameSubpos", tfunctime)
End Function

Function DataRead(ByVal nm As String) As Variant
Dim tfunctime As Double
tfunctime = Timer
    r = SetWorkbook()
    errread = 0
    If InStr(nm, "_") > 0 Then
        nsfile = Split(nm, "_")(0)
    Else
        nsfile = nm
    End If
    out_data_auto = DataReadFile(nsfile)
    out_data_mat = Empty
    un_subpos = ArrayUniqValColumn(out_data, col_sub_pos)
    If StrComp(nsfile, "Сводная") <> 0 Then out_data_mat = DataReadAutoMat(nsfile, un_subpos)
    type_spec = SpecGetType(nm)
    isReadFromSheet = False
    Select Case type_spec
        Case 7
            'Читаем с листа
            out_data = ManualSpec(nm)
            isReadFromSheet = True
        Case Else
            'Поищем листы с суффиксом "_спец"
            nsht = Trim$(nsfile) & "_спец"
            sheet = Empty
            listsheet = GetListOfSheet(wbk)
            For i = 1 To UBound(listsheet)
                If nsht = Trim$(listsheet(i)) Then sheet = nsht
            Next i
            If IsEmpty(sheet) And IsEmpty(out_data_auto) And IsEmpty(out_data_mat) Then
                'Нет ни файла, ни листа.
                errread = 1
            End If
            If Not IsEmpty(sheet) Then
                'Читаем с листа
                r = ManualCheck(nsht)
                out_data = ManualSpec(nsht)
                isReadFromSheet = True
            Else
                out_data = out_data_auto
            End If
    End Select
    If Not IsEmpty(out_data) And errread = 0 Then
        spec_version = DataGetVersion(out_data)
        If spec_version > 1 Then out_data = DataConvertVersion(out_data)
        If DataIsShort(out_data) And isReadFromSheet = False Then out_data = DataShort(out_data)
    Else
        spec_version = 3
    End If
    If Not IsEmpty(out_data_mat) Then
        out_data = ArrayCombine(out_data, out_data_mat)
    End If
    If Not IsEmpty(out_data_auto) And isReadFromSheet And StrComp(nsfile, "Сводная") <> 0 Then
        spec_version = DataGetVersion(out_data_auto)
        If spec_version > 1 Then
            out_data_auto = DataConvertVersion(out_data_auto)
        Else
            'Отключаем неиспользуемое
            this_sheet_option.Item("qty_one_floor") = False
        End If
        If DataIsShort(out_data_auto) Then out_data_auto = DataShort(out_data_auto)
        out_data = ArrayCombine(out_data, out_data_auto)
    End If
    If Not DataIsSpec(out_data) And type_spec <> 7 Or errread Then
        MsgBox ("Неверный формат файла")
        r = LogWrite(nm, vbNullString, "Неверный формат файла")
        DataRead = Empty
        Exit Function
    End If
    out_data = DataPrepare(out_data)
    DataRead = out_data
tfunctime = functime("DataRead", tfunctime)
End Function

Function DataReadFile(ByVal nsfile As String) As Variant
    'Проверим - есть ли такой файл
    listFile = GetListFile("*.txt")
    File = ArraySelectParam(listFile, nsfile, 1)
    If IsEmpty(File) Then
        out_data = Empty
    Else
        out_data = ReadFile(File(1, 1) & ".txt")
    End If
    DataReadFile = out_data
End Function


Function DataPrepare(ByVal out_data As Variant) As Variant
    out_data = DataCheck(out_data) 'Проверяем и корректируем
    If IsEmpty(out_data) Then DataPrepare = Empty: Exit Function
Dim tfunctime As Double
tfunctime = Timer
    Set pos_data = Nothing
    Set pos_data = CreateObject("Scripting.Dictionary")
    pos_data.comparemode = 1
    nfloor = 1
    floor_txt = "all_floor"
    add_subpos_txt = vbNullString
    add_subpos = DataAddNullSubpos(out_data) 'Добавляем объявления сборок для всех элементов
    
    add_subpos_txt = add_subpos_txt + ";" + add_subpos(2)
    If Not IsEmpty(add_subpos(1)) Then out_data = ArrayCombine(add_subpos(1), out_data)
    out_data = DataSumByControlSum(out_data) 'Объединяем все позиции с одинаковой контрольной суммой
    Set pos_data.Item(floor_txt) = DataUniqParent(ArraySelectParam_2(out_data, Array(t_subpos, t_perem_m), col_type_el))
    Set pos_data.Item(floor_txt).Item("weight") = DataWeightSubpos(out_data, floor_txt)
    If Not IsEmpty(ArraySelectParam(out_data, "-", col_sub_pos)) And this_sheet_option.Item("only_subpos") = False Then
        If pos_data.Item(floor_txt).Exists("-") Then
            pos_data.Item(floor_txt).Item("-").Item("-") = 1
        Else
            Set dfirst = CreateObject("Scripting.Dictionary")
            dfirst.Item("-") = 1
            Set pos_data.Item(floor_txt).Item("-") = dfirst
        End If
    End If
    If spec_version > 1 And this_sheet_option.Item("qty_one_floor") Then 'Учтём кол-во этажей
        out_data_allfloor = out_data
        out_data = Empty
        un_floor = ArrayUniqValColumn(out_data_allfloor, col_floor)
        nfloor = UBound(un_floor, 1)
        ReDim floor_txt_arr(nfloor, 3)
        For inxfloor = 1 To nfloor
            t_floor = un_floor(inxfloor)
            floor_txt = "%%" + CStr(t_floor) + "%" + CStr(inxfloor)
            floor_txt_arr(inxfloor, 1) = 0
            floor_txt_arr(inxfloor, 2) = t_floor
            floor_txt_arr(inxfloor, 3) = floor_txt
        Next inxfloor
        For inxfloor = 1 To nfloor
            t_floor = floor_txt_arr(inxfloor, 2)
            floor_txt = floor_txt_arr(inxfloor, 3)
            'Выбираем элементы с этим этажом
            out_data_floor = ArraySelectParam_2(out_data_allfloor, t_floor, col_floor)
            'На каждом этаже должна быть объявлена сборка
            add_subpos = DataAddNullSubpos(out_data_floor)
            add_subpos_txt = add_subpos_txt + ";" + add_subpos(2)
            If Not IsEmpty(add_subpos(1)) Then out_data_floor = ArrayCombine(add_subpos(1), out_data_floor)
            out_data_floor = DataSumByControlSum(out_data_floor) 'Объединяем все позиции с одинаковой контрольной суммой
            Set pos_data.Item(floor_txt) = DataUniqParent(ArraySelectParam_2(out_data_floor, Array(t_subpos, t_perem_m), col_type_el))
            Set pos_data.Item(floor_txt).Item("weight") = DataWeightSubpos(out_data_floor, floor_txt)
            If Not IsEmpty(ArraySelectParam(out_data_floor, "-", col_sub_pos)) And this_sheet_option.Item("only_subpos") = False Then
                If pos_data.Item(floor_txt).Exists("-") Then
                    pos_data.Item(floor_txt).Item("-").Item("-") = 1
                Else
                    Set dfirst = CreateObject("Scripting.Dictionary")
                    dfirst.Item("-") = 1
                    Set pos_data.Item(floor_txt).Item("-") = dfirst
                End If
            End If
            If inxfloor > 1 Then
                out_data = ArrayCombine(out_data, out_data_floor)
            Else
                out_data = out_data_floor
            End If
        Next inxfloor
    End If
'    add_subpos_txt = Trim(add_subpos_txt)
'    If Len(add_subpos_txt) > 2 Then
'        add_subpos_txt = Split(add_subpos_txt, ";")
'        add_subpos_txt = ArrayUniqValColumn(add_subpos_txt)
'        add_subpos_txt = Join(add_subpos_txt, vbCrLf)
'        MsgBox ("Добавлена недостающая маркировка сборок " & vbCrLf & add_subpos_txt)
'    End If
    DataPrepare = out_data
tfunctime = functime("DataPrepare", tfunctime)
End Function

Function DataShort(ByVal array_in As Variant) As Variant
    If IsEmpty(array_in) Then
        DataShort = Empty
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
    'Домножаем количество элементов на число в первом столбце
    rows_array_in = UBound(array_in, 1)
    cols_array_in = UBound(array_in, 2)
    ReDim out(rows_array_in, cols_array_in)
    n_row = 0
    n_error = 0
    For i = 1 To rows_array_in
        If IsNumeric(array_in(i, 1)) And IsNumeric(array_in(i, col_qty + 1)) And type_el_name.Exists(array_in(i, col_type_el + 1)) Then
            n_row = n_row + 1
            For j = 2 To cols_array_in
                out(n_row, j - 1) = array_in(i, j)
            Next j
            qty = array_in(i, 1)
            out(n_row, col_qty) = out(n_row, col_qty) * array_in(i, 1)
        Else
            n_error = n_error + 1
            If type_el_name.Exists(array_in(i, col_type_el + 1)) Then
                r = LogWrite(array_in(i, 2) & array_in(i, 6) & array_in(i, 10), array_in(i, 5), "ДЛИНА")
            Else
                kk = 1
            End If
        End If
    Next i
tfunctime = functime("DataShort", tfunctime)
    If n_error > 2 Then
        MsgBox ("Пропущено строк -" & CStr(n_error))
    End If
    ReDim Preserve out(rows_array_in, max_col)
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
Dim tfunctime As Double
tfunctime = Timer
    For Each t_el In Array(t_arm, t_prokat, t_mat, t_izd, t_subpos, t_wind, t_perem, t_perem_m)
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
tfunctime = functime("DataSumByControlSum", tfunctime)
End Function

Function IsEqFloor(ByVal arr1, ByVal arr2) As Boolean
    iseq = True
    If spec_version > 1 And this_sheet_option.Item("qty_one_floor") Then
    
    End If
    IsEqFloor = iseq
End Function

Function DataUniqParent(ByVal sub_pos_arr As Variant) As Variant
    'Возвращает словарь с главными сборками и входящими в них вложенными
    Set dparent = CreateObject("Scripting.Dictionary")
    Set dchild = CreateObject("Scripting.Dictionary")
    Set dqty = CreateObject("Scripting.Dictionary")
    Set dfirst = CreateObject("Scripting.Dictionary")
    Set out = CreateObject("Scripting.Dictionary")
    out.comparemode = 1
    dparent.comparemode = 1
    dchild.comparemode = 1
    dqty.comparemode = 1
    dfirst.comparemode = 1
    If Not IsEmpty(sub_pos_arr) Then
Dim tfunctime As Double
tfunctime = Timer
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
    If dfirst.Count > 0 And this_sheet_option.Item("only_subpos") = False Then Set out.Item("-") = dfirst
    Set out.Item("name") = DataNameSubpos(sub_pos_arr)
    Set DataUniqParent = out
tfunctime = functime("DataUniqParent", tfunctime)
End Function

Function DataWeightSubpos(ByVal array_in As Variant, ByVal floor_txt As String) As Variant
    Set dweight = CreateObject("Scripting.Dictionary")
    dweight.comparemode = 1
    Dim tweight As Double
    If (UBound(pos_data.Item(floor_txt).Item("parent").keys()) < 0) Then
        Set DataWeightSubpos = dweight
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
    'Общий вес всех элементов сборки
    For i = 1 To UBound(array_in, 1)
        subpos = array_in(i, col_sub_pos)
        type_el = array_in(i, col_type_el)
        If (subpos <> "-") Then
            tweight = 0
            Select Case type_el
                Case t_arm
                    klass = array_in(i, col_klass)
                    diametr = array_in(i, col_diametr)
                    weight_pm = GetWeightForDiametr(diametr, klass)
                    length_pos = array_in(i, col_length) / 1000
                    qty = array_in(i, col_qty)
                    If (qty = 0) Or IsEmpty(qty) Then qty = 1
                    fon = array_in(i, col_fon)
'If k_zap_total > 1 Then qty = qty + Round((k_zap_total - 1) * qty, 0)
                    If fon Or this_sheet_option.Item("arm_pm") Then
                        length_pos = Round_w(length_pos * k_zap_total, n_round_l)
                        tweight = length_pos * weight_pm * qty
                    Else
                        If zap_only_mp Then
                            tweight = Round_w(weight_pm * length_pos, n_round_w) * qty
                        Else
                            tweight = Round_w(weight_pm * length_pos * k_zap_total, n_round_w) * qty
                        End If
                    End If
                Case t_prokat
                    qty = array_in(i, col_qty)
                    If (qty = 0) Or IsEmpty(qty) Then qty = 1
                    name_pr = GetShortNameForGOST(array_in(i, col_pr_gost_prof))
                    If InStr(1, name_pr, "Лист") Then
                        naen_plate = SpecMetallPlate(array_in(i, col_pr_prof), array_in(i, col_pr_naen), array_in(i, col_pr_length), array_in(i, col_pr_weight), array_in(i, col_chksum))
                        weight_pm = naen_plate(4)
                        length_pos = naen_plate(2)
                    Else
                        length_pos = Round_w(array_in(i, col_pr_length) / 1000, 3)
                        weight_pm = array_in(i, col_pr_weight) * length_pos
                    End If
                    pm = False: If InStr(array_in(i, col_chksum), "lpm") > 0 Then pm = True
                    If this_sheet_option.Item("pr_pm") Or pm Then
                        tweight = Round_w(weight_pm * k_zap_total, n_round_w) * qty
                    Else
                        If zap_only_mp Then
                            tweight = Round_w(weight_pm, n_round_w) * qty
                        Else
                            tweight = Round_w(weight_pm * k_zap_total, n_round_w) * qty
                        End If
                    End If
                Case t_izd
                    qty = array_in(i, col_qty)
                    If (qty = 0) Or IsEmpty(qty) Then qty = 1
                    If Not IsNumeric(array_in(i, col_m_weight)) Then
                        tweight = 0
                    Else
                        tweight = Round_w(array_in(i, col_m_weight), n_round_w) * qty
                    End If
            End Select
            If tweight Then dweight.Item(subpos) = dweight.Item(subpos) + tweight
        End If
    Next
    'Делим на количество вхождений, чтоб получить массу одной шт.
    For Each subpos In dweight.keys()
        If pos_data.Item(floor_txt).Item("child").Exists(subpos) Then
            nSubPos = pos_data.Item(floor_txt).Item("qty").Item("all" & subpos)
        Else
            nSubPos = pos_data.Item(floor_txt).Item("qty").Item("-_" & subpos)
        End If
        If nSubPos < 1 Then
            MsgBox ("Не определено кол-во сборок " & subpos & ", принято 1 шт.")
            r = LogWrite(subpos, vbNullString, "Не определено кол-во сборок")
            nSubPos = 1
        End If
        w = (dweight.Item(subpos) / nSubPos)
        dweight.Item(subpos) = Round_w(w, n_round_w)
    Next
    'Для сборок первого уровня учтём вхождения сборок второго уровня
    For Each subpos In pos_data.Item(floor_txt).Item("parent").keys()
        For Each tchild In pos_data.Item(floor_txt).Item("child").keys()
            If pos_data.Item(floor_txt).Item("qty").Exists(subpos & "_" & tchild) Then
                qty = pos_data.Item(floor_txt).Item("qty").Item(subpos & "_" & tchild) / pos_data.Item(floor_txt).Item("qty").Item("-_" & subpos)
                tweight = dweight.Item(tchild)
                dweight.Item(subpos) = dweight.Item(subpos) + qty * tweight
            End If
        Next
    Next
    Set DataWeightSubpos = dweight
tfunctime = functime("DataWeightSubpos", tfunctime)
End Function

Function SetWorkbook() As Boolean
    If IsEmpty(wbk) Or wbk Is Nothing Then
        Set wbk = ThisWorkbook
        SetWorkbook = True
    Else
        SetWorkbook = False
    End If
End Function


Function DebugOut(ByVal pos_out As Variant, Optional ByVal module_name As String) As Boolean
    sh_name = "DEBUG"
    If SheetExist(sh_name) Then
        Set Sh = wbk.Sheets("DEBUG")
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
    buffer$ = vbNullString
    For i = 1 To UBound(arr, 1)
        txt = vbNullString
        For j = 1 To UBound(arr, 2)
            If arr(i, j) <> vbNullString Then txt = txt & ColumnsSeparator$ & arr(i, j)
        Next j
        If txt <> vbNullString Then
            Range2CSV = Range2CSV & Mid$(txt, Len(ColumnsSeparator$) + 1) & RowsSeparator$
            If Len(Range2CSV) > 50000 Then buffer$ = buffer$ & Range2CSV: Range2CSV = vbNullString
        End If
    Next i
    CSVtext$ = buffer$ & Range2CSV
    ExportList2CSV = ExportSaveTXTfile(CSVfilename$, CSVtext$)
End Function

Function ExportArray2CSV(ByVal arr As Variant, ByVal CSVfilename As String, Optional ByVal ColumnsSeparator$ = ";", Optional ByVal RowsSeparator$ = vbNewLine) As String
    buffer$ = vbNullString
    For i = 1 To UBound(arr, 1)
        txt = vbNullString
        For j = 1 To UBound(arr, 2)
            If arr(i, j) <> vbNullString Then txt = txt & ColumnsSeparator$ & arr(i, j)
        Next j
        If txt <> vbNullString Then
            Range2CSV = Range2CSV & Mid$(txt, Len(ColumnsSeparator$) + 1) & RowsSeparator$
            If Len(Range2CSV) > 50000 Then buffer$ = buffer$ & Range2CSV: Range2CSV = vbNullString
        End If
    Next i
    CSVtext$ = buffer$ & Range2CSV
    ExportArray2CSV = ExportSaveTXTfile(CSVfilename$, CSVtext$)
End Function

Function ExportSaveTXTfile(ByVal filename As String, ByVal txt As String) As Boolean
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.CreateTextFile(filename, True)
    ts.Write txt: ts.Close
    ExportSaveTXTfile = Err = 0
    Set ts = Nothing: Set FSO = Nothing
End Function

Function ExportAttribut(ByVal nm As String) As Boolean
    '-------------------------------------------------------
    out_data_raw = Empty
    Dim out_data_diff
    Set data_out = wbk.Sheets(nm)
    n_row = SheetGetSize(data_out)(1)
    col = max_col_man
    spec = data_out.Range(data_out.Cells(1, 1), data_out.Cells(n_row, max_col_man))
    block_data = ArraySelectParam_2(spec, "АВТОКАД_?", col_man_naen)
    If IsEmpty(block_data) Then
        MsgBox ("Данные для выгрузки отсутвуют")
        ExportAttribut = False
        Exit Function
    End If
    For i = 1 To UBound(block_data, 1)
        If InStr(block_data(i, col_man_subpos), "!") = 0 And InStr(block_data(i, col_man_pos), "!") = 0 Then
            arr = Split(block_data(i, col_man_naen), "_")
            block_data(i, col_man_pr_length) = arr(1)
            block_data(i, col_man_pr_gost_pr) = arr(2)
        End If
    Next i
    subpos_arr = ArrayUniqValColumn(block_data, col_man_subpos)
    For Each subpos In subpos_arr
        If InStr(subpos, "!") = 0 Then
            block_data_subpos = ArraySelectParam_2(block_data, subpos, col_man_subpos)
            coll = GetListFile(subpos + "_autocad.txt")
            For i = 1 To UBound(coll, 1)
                snm = Split(coll(i, 1), "_")(0)
                If snm = subpos Then
                    out_data_sheet = ReadFile(coll(i, 1) + ".txt", 1, vbTab, vbNewLine)
                    out_data_raw = ArrayCombine(out_data_raw, out_data_sheet)
                End If
            Next i
            n_row = UBound(out_data_raw, 1)
            max_col_acad = UBound(out_data_raw, 2)
            'Описание файла c извлечением данных из автокада
            col_acad_handle = 1
            col_acad_blockname = 2
            'Поиск подходящих столбцов по имени
            col_acad_qty = -1 'Кол-во стержней в блоке
            col_acad_diametr = -1
            col_acad_length = -1
            col_acad_pos = -1
            For i = 1 To max_col_acad
                If Trim$(UCase$(out_data_raw(1, i))) = "КОЛИЧЕСТВО" Then col_acad_qty = i
                If Trim$(UCase$(out_data_raw(1, i))) = "ДИАМЕТР" Then col_acad_diametr = i
                If Trim$(UCase$(out_data_raw(1, i))) = "ПОЗИЦИЯ" Then col_acad_pos = i
                If Trim$(UCase$(out_data_raw(1, i))) = "ДЛИНА_СТЕРЖНЯ" Then col_acad_length = i
            Next i
            ReDim out_data_diff(n_row, 4)
            n_diff = 0
            n_diff = n_diff + 1
            out_data_diff(n_diff, 1) = "HANDLE"
            out_data_diff(n_diff, 2) = "BLOCKNAME"
            out_data_diff(n_diff, 3) = "ПОЗИЦИЯ"
            out_data_diff(n_diff, 4) = "ДИАМЕТР"
            For i = 1 To n_row
                block = ArraySelectParam_2(block_data_subpos, out_data_raw(i, col_acad_handle), col_man_pr_length, out_data_raw(i, col_acad_blockname), col_man_pr_gost_pr)
                If Not IsEmpty(block) Then
                    If out_data_raw(i, col_acad_diametr) <> block(1, col_man_diametr) Or out_data_raw(i, col_acad_pos) <> block(1, col_man_pos) Then
                        n_diff = n_diff + 1
                        out_data_diff(n_diff, 1) = out_data_raw(i, col_acad_handle)
                        out_data_diff(n_diff, 2) = out_data_raw(i, col_acad_blockname)
                        out_data_diff(n_diff, 3) = block(1, col_man_pos)
                        out_data_diff(n_diff, 4) = block(1, col_man_diametr)
                    End If
                End If
            Next i
            CSVfilename$ = ThisWorkbook.path & "\list\Autocad_" & subpos & ".txt"
            If n_diff > 0 Then r = ExportArray2CSV(out_data_diff, CSVfilename, vbTab, vbNewLine)
        End If
    Next
    ExportAttribut = True
End Function
Function ExportSetPageBreaks(ByRef Sh As Variant, ByVal h_list As Double, Optional ByVal n_first As Long, Optional ByVal page_delim As String) As Boolean
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
    If SpecGetType(nm) > 0 And mem_option Then r = OptionSheetSet(nm)

    If type_spec = 12 Then
        r = FormatSpec_Pol(data_out)
        type_spec = 0
    End If
    If type_spec <> 7 And type_spec > 0 And Len(nm) > 1 Then
        Set Sh = wbk.Sheets(nm)
        lsize = SheetGetSize(Sh)
        n_row = lsize(1)
        n_col = lsize(2)
        Set data_out = Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col))
        If type_spec = 3 And Right$(nm, 3) = "зап" Then
            r = FormatSpec_Perem(data_out, n_row)
        End If
        'Вытащим коэфф. запаса
        k_zap_total_t = ""
        If InStr(data_out.Cells(n_row, 1).Value, "k=") > 0 Then
            k_zap_total_t = Replace(Sh.Cells(n_row, 1).Value, "k=", vbNullString)
            k_zap_total_t = ConvTxt2Num(k_zap_total_t)
            k_zap_total_t = ConvNum2Txt(k_zap_total_t * 10)
        Else
            If IsEmpty(this_sheet_option) Then
                r = OptionSheetSet(nm)
                Set this_sheet_option = OptionGetForm(nm)
            End If
            r = SetKzap()
            k_zap_total_t = ConvNum2Txt(k_zap_total * 10)
        End If
        filename$ = ThisWorkbook.path & "\list\Спец_" & nm & "_" & k_zap_total_t & ".pdf"
        If Dir(filename) <> vbNullString Then
            If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\list\old\") Then
                MkDir (ThisWorkbook.path & "\list\old\")
            End If
            tdate = Right$(str(DatePart("yyyy", Now)), 2) & str(DatePart("m", Now)) & str(DatePart("d", Now))
            stamp = "=" + tdate + "=" + str(DatePart("h", Now)) + str(DatePart("n", Now)) + str(DatePart("s", Now))
            stamp = Replace(stamp, " ", vbNullString)
            Set fs = CreateObject("Scripting.FileSystemObject")
            fs.CopyFile filename, ThisWorkbook.path & "\list\old\Спец_" & nm & ConvNum2Txt(k_zap_total * 10) & stamp & ".pdf"
            Set fs = Nothing
        End If
        type_print = 0
        If SpecGetType(nm) = 11 Then type_print = 1
        If delim_group_ved And type_spec = 1 Then
            r = ExportSetPageBreaks(Sh, 420, 2, "Марка" & vbLf & "изделия.")
        Else
            If GetHeightSheet(Sh) > 420 Then r = ExportSetPageBreaks(Sh, 420, 2)
        End If
        If type_spec <> 1 Then
            Sh.ResetAllPageBreaks
            delim_group_ved = False
        End If
        r = ExportSheet2Pdf(data_out, filename, type_print)
        r = LogWrite(filename, "PDF", "ОК")
    End If
End Function

Function ExportSheet2Pdf(ByVal data_out As Range, ByVal filename As String, Optional ByVal type_print As Long = 0) As Boolean
    data_out.Select
    On Error Resume Next
    'Application.PrintCommunication = False
    ActiveSheet.PageSetup.PrintArea = data_out.Address
    Select Case type_print
        Case 0
            With ActiveSheet.PageSetup
                .LeftHeader = vbNullString
                .CenterHeader = vbNullString
                .RightHeader = vbNullString
                .LeftFooter = vbNullString
                .CenterFooter = vbNullString
                .RightFooter = vbNullString
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
                .EvenPage.LeftHeader.Text = vbNullString
                .EvenPage.CenterHeader.Text = vbNullString
                .EvenPage.RightHeader.Text = vbNullString
                .EvenPage.LeftFooter.Text = vbNullString
                .EvenPage.CenterFooter.Text = vbNullString
                .EvenPage.RightFooter.Text = vbNullString
                .FirstPage.LeftHeader.Text = vbNullString
                .FirstPage.CenterHeader.Text = vbNullString
                .FirstPage.RightHeader.Text = vbNullString
                .FirstPage.LeftFooter.Text = vbNullString
                .FirstPage.CenterFooter.Text = vbNullString
                .FirstPage.RightFooter.Text = vbNullString
            End With
        Case 1
            With ActiveSheet.PageSetup
                .LeftHeader = vbNullString
                .CenterHeader = vbNullString
                .RightHeader = vbNullString
                .LeftFooter = vbNullString
                .CenterFooter = vbNullString
                .RightFooter = vbNullString
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
                .EvenPage.LeftHeader.Text = vbNullString
                .EvenPage.CenterHeader.Text = vbNullString
                .EvenPage.RightHeader.Text = vbNullString
                .EvenPage.LeftFooter.Text = vbNullString
                .EvenPage.CenterFooter.Text = vbNullString
                .EvenPage.RightFooter.Text = vbNullString
                .FirstPage.LeftHeader.Text = vbNullString
                .FirstPage.CenterHeader.Text = vbNullString
                .FirstPage.RightHeader.Text = vbNullString
                .FirstPage.LeftFooter.Text = vbNullString
                .FirstPage.CenterFooter.Text = vbNullString
                .FirstPage.RightFooter.Text = vbNullString
            End With
    End Select
    On Error Resume Next
    'Application.PrintCommunication = True
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=filename$, Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ExportSheet2Pdf = True
End Function

Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal mask As String = vbNullString, Optional ByVal SearchDeep As Long = 2) As Collection
    Set FilenamesCollection = New Collection
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetAllFileNamesUsingFSO FolderPath, mask, FSO, FilenamesCollection, SearchDeep
    Set FSO = Nothing: Application.StatusBar = False
End Function

Function FormatClear(ByRef Sh As Variant) As Boolean
Dim tfunctime As Double
tfunctime = Timer
    With Sh.Cells
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
tfunctime = functime("FormatClear", tfunctime)
End Function

Function FormatFont(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean
Dim tfunctime As Double
tfunctime = Timer
    arr_bold = Array("шт.)", ", на ", "Элементы на отм.")
    For Each txt In arr_bold
        data_out.FormatConditions.Add Type:=xlTextString, String:=txt, TextOperator:=xlContains
        data_out.FormatConditions(data_out.FormatConditions.Count).SetFirstPriority
        data_out.FormatConditions(1).Font.Bold = True
    Next
    
    arr_underline = type_el_name.items
    For Each txt In arr_underline
        data_out.FormatConditions.Add Type:=xlTextString, String:=txt, TextOperator:=xlContains
        data_out.FormatConditions(data_out.FormatConditions.Count).SetFirstPriority
        data_out.FormatConditions(1).Font.Underline = xlUnderlineStyleSingle
    Next
    
    arr_warning = Array("!!!", "ИЗ ФАЙЛА", "С ЛИСТА")
    For Each txt In arr_warning
        data_out.FormatConditions.Add Type:=xlTextString, String:=txt, TextOperator:=xlContains
        data_out.FormatConditions(data_out.FormatConditions.Count).SetFirstPriority
        data_out.FormatConditions(1).Font.Color = -16751204
    Next

    With data_out.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    With data_out.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    data_out.Borders(xlDiagonalDown).LineStyle = xlNone
    data_out.Borders(xlDiagonalUp).LineStyle = xlNone
    With data_out.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With data_out.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With data_out.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With data_out.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With data_out.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With data_out.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    data_out.Borders(xlDiagonalDown).LineStyle = xlNone
    data_out.Borders(xlDiagonalUp).LineStyle = xlNone
    With data_out.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With data_out.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With data_out.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With data_out.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With data_out.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With data_out.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With data_out.Font
        .Name = fontname
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
                n = data_out.Cells(i, j)
                On Error Resume Next
                If IsNumeric(data_out.Cells(i, j)) And data_out.Cells(i, j) <> 0 Then
                    With data_out.Cells(i, j)
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
                    With data_out.Cells(i, j)
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
        With data_out
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
tfunctime = functime("FormatFont", tfunctime)
End Function

Function FormatManual(ByVal nm As String) As Boolean
    'Наведение красоты на листе с ручной спецификацией
    If IsEmpty(pr_adress) Then r = ReadPrSortament()
    Set data_out = wbk.Sheets(nm)
    size_sh = SheetGetSize(data_out)
    nrow = size_sh(1) + 30
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
    data_out.Cells.UnMerge
    
    data_out.Cells(1, col_man_subpos) = "Марка" & vbLf & "элемента"
    data_out.Cells(1, col_man_pos) = "Поз."
    data_out.Cells(1, col_man_obozn) = "Обозначение"
    data_out.Cells(1, col_man_naen) = "Наименование"
    data_out.Cells(1, col_man_qty) = "Кол-во" & vbLf & "на один элемент"
    data_out.Cells(1, col_man_weight) = "Масса, кг"
    data_out.Cells(1, col_man_prim) = "Примечание" & vbLf & "(на лист)"
    data_out.Cells(1, col_man_komment) = "Комментарий"
    
    data_out.Cells(1, col_man_length) = "Арматура"
    data_out.Cells(2, col_man_length) = "Длина, мм"
    data_out.Cells(2, col_man_diametr) = "Диаметр"
    data_out.Cells(2, col_man_klass) = "Класс"
    
    data_out.Cells(1, col_man_pr_length) = "Прокат"
    data_out.Cells(2, col_man_pr_length) = "Длина" & vbLf & "(площадь кв.мм для пластин), мм"
    data_out.Cells(2, col_man_pr_gost_pr) = "ГОСТ профиля"
    data_out.Cells(2, col_man_pr_prof) = "Профиль"
    data_out.Cells(2, col_man_pr_type) = "Тип конструкции"
    data_out.Cells(2, col_man_pr_st) = "Сталь"
    data_out.Cells(2, col_man_pr_okr) = "Окраска"
    data_out.Cells(2, col_man_pr_ogn) = "Огнезащита"
    
    data_out.Cells(1, col_man_ank) = "Всё в мм"
    data_out.Cells(2, col_man_ank) = "Анкеровка"
    data_out.Cells(2, col_man_nahl) = "Нахлёст"
    data_out.Cells(2, col_man_dgib) = "Радиус оправки"
    
    data_out.Range("A1:A2").Merge
    data_out.Range("B1:B2").Merge
    data_out.Range("C1:C2").Merge
    data_out.Range("D1:D2").Merge
    data_out.Range("E1:E2").Merge
    data_out.Range("F1:F2").Merge
    data_out.Range("G1:G2").Merge
    data_out.Range("H1:J1").Merge
    data_out.Range("K1:Q1").Merge
    data_out.Range("R1:R2").Merge
    data_out.Range("S1:U1").Merge
    
    data_out.Cells(1, col_man_subpos).ColumnWidth = 8
    data_out.Cells(1, col_man_pos).ColumnWidth = 8
    data_out.Cells(1, col_man_obozn).ColumnWidth = 25
    data_out.Cells(1, col_man_naen).ColumnWidth = 25
    data_out.Cells(1, col_man_qty).ColumnWidth = 8
    data_out.Cells(1, col_man_weight).ColumnWidth = 8
    data_out.Cells(1, col_man_prim).ColumnWidth = 15
    data_out.Cells(2, col_man_length).ColumnWidth = 10
    data_out.Cells(2, col_man_diametr).ColumnWidth = 10
    data_out.Cells(2, col_man_klass).ColumnWidth = 10
    data_out.Cells(1, col_man_komment).ColumnWidth = 15
    
    data_out.Cells(1, col_man_pr_length).ColumnWidth = 15
    data_out.Cells(2, col_man_pr_length).ColumnWidth = 15
    data_out.Cells(2, col_man_pr_gost_pr).ColumnWidth = 34
    data_out.Cells(2, col_man_pr_prof).ColumnWidth = 11
    data_out.Cells(2, col_man_pr_type).ColumnWidth = 15
    data_out.Cells(2, col_man_pr_st).ColumnWidth = 8
    data_out.Cells(2, col_man_pr_okr).ColumnWidth = 8
    data_out.Cells(2, col_man_pr_ogn).ColumnWidth = 8
    
    data_out.Cells(2, col_man_ank).ColumnWidth = 8
    data_out.Cells(2, col_man_nahl).ColumnWidth = 8
    data_out.Cells(2, col_man_dgib).ColumnWidth = 8
    
    data_out.Range(r_all).FormatConditions.Add Type:=xlExpression, Formula1:="=ЕОШИБКА(A1)"
    data_out.Range(r_all).FormatConditions(data_out.Range(r_all).FormatConditions.Count).SetFirstPriority
    With data_out.Range(r_all).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10040319
        .TintAndShade = 0
    End With
    data_out.Range(r_all).FormatConditions(1).StopIfTrue = False
'
'    'Создаём столбец с марками элементов и добавим раскрывающийся список
'    Set name_subpos = DataNameSubpos(Empty)
'    If name_subpos.Count > 0 Then
'        un_pos = name_subpos.Keys()
'        If Not IsEmpty(un_pos) Then
'            istart = max_col_man + 1
'            iend = UBound(un_pos, 1)
'            'Data_out.range(Data_out.Cells(1, istart), Data_out.Cells((iEnd + 3) * 500, istart)).ClearContents
'            For i = 1 To iend
'                Data_out.Range(Data_out.Cells(i, istart), Data_out.Cells(i, istart)) = un_pos(i)
'            Next
'            With Range(r_subpos).Validation
'                .Delete
'                .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
'                .IgnoreBlank = True
'                .InCellDropdown = True
'                .ShowInput = True
'                .ShowError = False
'            End With
'            With Data_out.Range(Data_out.Cells(1, istart), Data_out.Cells(iend, istart)).Font
'                .Name = "Calibri"
'                .Size = 8
'                .Strikethrough = False
'                .Superscript = False
'                .Subscript = False
'                .OutlineFont = False
'                .Shadow = False
'                .Underline = xlUnderlineStyleNone
'                .ThemeColor = xlThemeColorLight1
'                .TintAndShade = 0
'                .ThemeFont = xlThemeFontMinor
'            End With
'            With Data_out.Range(Data_out.Cells(1, istart), Data_out.Cells(iend, istart)).Interior
'                .Pattern = xlSolid
'                .PatternColorIndex = xlAutomatic
'                .ThemeColor = xlThemeColorDark1
'                .TintAndShade = 0
'                .PatternTintAndShade = 0
'            End With
'            With Range(r_subpos).Validation
'                .Delete
'                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & Data_out.Range(Data_out.Cells(1, istart), Data_out.Cells(iend, istart)).Address
'                .IgnoreBlank = True
'                .InCellDropdown = True
'                .InputTitle = vbNullString
'                .ErrorTitle = vbNullString
'                .InputMessage = vbNullString
'                .ErrorMessage = vbNullString
'                .ShowInput = True
'                .ShowError = False
'            End With
'        End If
'    End If
    
    With data_out.Range(r_prim).Validation
        .Delete
        On Error Resume Next
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertWarning, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("Примечания")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .ErrorTitle = vbNullString
        .InputMessage = vbNullString
        .ErrorMessage = vbNullString
        .ShowInput = True
        .ShowError = True
    End With
    
    With data_out.Range(r_klass).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("Классы")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .ErrorTitle = vbNullString
        .InputMessage = vbNullString
        .ErrorMessage = vbNullString
        .ShowInput = True
        .ShowError = True
    End With

    With data_out.Range(r_pr_gost_pr).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("ГОСТпрокат")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .ErrorTitle = vbNullString
        .InputMessage = vbNullString
        .ErrorMessage = vbNullString
        .ShowInput = True
        .ShowError = True
    End With
    
    With data_out.Range(r_pr_st).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("Марки стали")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .ErrorTitle = vbNullString
        .InputMessage = vbNullString
        .ErrorMessage = vbNullString
        .ShowInput = True
        .ShowError = True
    End With
    
    
    With data_out.Range(r_pr_okr).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & pr_adress.Item("Окраска")
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = vbNullString
        .ErrorTitle = vbNullString
        .InputMessage = vbNullString
        .ErrorMessage = vbNullString
        .ShowInput = True
        .ShowError = True
    End With
    has_prof = False
    For i = 1 To nrow + 30
        gost = data_out.Cells(i, col_man_pr_gost_pr).Value
        addr = pr_adress.Item(gost)
        If Not IsEmpty(addr) And Not IsEmpty(gost) Then
            With data_out.Cells(i, col_man_pr_prof).Validation
                            .Delete
                            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & addr(1)
                            .IgnoreBlank = True
                            .InCellDropdown = True
                            .InputTitle = vbNullString
                            .ErrorTitle = vbNullString
                            .InputMessage = vbNullString
                            .ErrorMessage = vbNullString
                            .ShowInput = True
                            .ShowError = True
            End With
        End If
        
        klass = data_out.Cells(i, col_man_klass).Value
        addr = pr_adress.Item(klass)
        If Not IsEmpty(addr) And Not IsEmpty(klass) Then
            With data_out.Cells(i, col_man_diametr).Validation
                            .Delete
                            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & addr
                            .IgnoreBlank = True
                            .InCellDropdown = True
                            .InputTitle = vbNullString
                            .ErrorTitle = vbNullString
                            .InputMessage = vbNullString
                            .ErrorMessage = vbNullString
                            .ShowInput = True
                            .ShowError = True
            End With
        End If
        If Not IsEmpty(data_out.Cells(i, col_man_pr_gost_pr).Value) And i > 2 Then has_prof = True
    Next i
    data_out.Range(r_all).Rows.AutoFit
    If has_prof = False Then
        data_out.Columns("K:Q").Group
        data_out.Columns("K:Q").EntireColumn.Hidden = True
    End If
    FormatManual = True
End Function

Function FormatManuallitera(ByVal col As Long) As String
    If col > 0 Then
        litera = Split(Cells(1, col).Address, "$")(1)
    Else
        litera = "A"
    End If
    FormatManuallitera = litera
End Function

Function FormatManualrange(ByVal col As Long, ByVal nrow As Long) As String
    litera = FormatManuallitera(col)
    out = litera & "3:" & litera & Trim$(str(nrow))
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

Function FormatRowPrint(ByRef data_out As Range, ByVal n_row As Long)
    Application.PrintCommunication = False
    With wbk.Sheets(data_out.Parent.Name).PageSetup
        .PrintTitleRows = "$1:$" + CStr(n_row)
        .PrintTitleColumns = vbNullString
    End With
    Application.PrintCommunication = True
End Function

Function FormatSpec_AS(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean
    start_row = 2
    If this_sheet_option.Item("qty_one_floor") And spec_version > 1 Then
        n_emp = 0
        For i = 4 To n_col - 4
            If Len(data_out.Cells(start_row + 1, i)) < 2 Then
                data_out.Columns(i).Delete
                n_emp = n_emp + 1
            End If
        Next i
        n_col = n_col - n_emp
        For i = 1 To 3
            data_out.Range(data_out.Cells(start_row, i), data_out.Cells(start_row + 1, i)).Merge
        Next i
        data_out.Range(data_out.Cells(start_row, 4), data_out.Cells(start_row, n_col - 3)).Merge
        For i = n_col - 2 To n_col
            data_out.Range(data_out.Cells(start_row, i), data_out.Cells(start_row + 1, i)).Merge
        Next i
        n_qty = 3
    Else
        n_qty = 6
    End If
    n_naen = 3
    For i = start_row + 1 To n_row
        If InStr(data_out(i, 1), ", на ") > 0 Or InStr(data_out(i, 1), ",**") > 0 Then
            data_out(i, 1) = Replace(data_out(i, 1), ",**", vbNullString)
            If InStr(data_out(i, 1), "_") > 0 Then data_out(i, 1) = Split(data_out(i, 1), "_")(1)
            If this_sheet_option.Item("qty_one_floor") And spec_version > 1 Then
                 If InStr(data_out(i, 1), ", на ") > 0 Then data_out(i, 1) = Split(data_out(i, 1), ", на ")(0)
            End If
            data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, n_qty)).Merge
        End If
        If IsNumeric(Application.Match(data_out.Cells(i, n_naen), type_el_name.items, 0)) Then
            data_out.Cells(i, 1).Value = data_out.Cells(i, n_naen).Value
            data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, n_qty)).Merge
        End If
        If InStr(data_out(i, 1), " Прочие") > 0 Then data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, n_qty)).Merge
    Next i
    
    If this_sheet_option.Item("merge_material") Then
        n_c = 2
        start_row_t = start_row + 1
        r = FormatSpec_merge(data_out.Range(data_out.Cells(start_row_t, 1), data_out.Cells(n_row, n_col)), n_c, False)
    End If
    
    s1 = 15
    s2 = 50
    s3 = 60
    s4 = 15
    s5 = 20
    s6 = 25
    sall = s1 + s2 + s3 + s4 + s5
    koeff = (sall / 209) * 100
    dblPoints = Application.CentimetersToPoints(1)
    r = FormatFont(data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(n_row, n_col)), n_row, n_col)
    For i = start_row + 1 To n_row
        If Not IsNumeric(Application.Match(data_out.Cells(i, 1), type_el_name.items, 0)) Then
            If data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, n_qty)).MergeCells Then
                data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, n_qty)).Font.Bold = True
            End If
        End If
    Next i
    data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(n_row, n_col)).Rows.AutoFit
    If data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(start_row, n_col)).RowHeight < dblPoints * 0.8 Then
        data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(start_row, n_col)).RowHeight = dblPoints * 0.8
    End If
    data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(start_row, 1)).ColumnWidth = (s1 / sall) * koeff
    data_out.Range(data_out.Cells(start_row, 2), data_out.Cells(start_row, 2)).ColumnWidth = (s2 / sall) * koeff
    data_out.Range(data_out.Cells(start_row, 3), data_out.Cells(start_row, 3)).ColumnWidth = (s3 / sall) * koeff
    data_out.Range(data_out.Cells(start_row, 4), data_out.Cells(start_row, 4)).ColumnWidth = (s4 / sall) * koeff
    data_out.Range(data_out.Cells(start_row, 5), data_out.Cells(start_row, 5)).ColumnWidth = (s5 / sall) * koeff
    data_out.Range(data_out.Cells(start_row, 6), data_out.Cells(start_row, 6)).ColumnWidth = (s6 / sall) * koeff
End Function

Function FormatSpec_ASGR(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean
    start_row = 2
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
    n_emp = 0
    For i = 4 To n_col - 4
        tcval = data_out.Cells(start_row + 1, i).Value
        If Len(tcval) < 2 Then
            data_out.Columns(i).Delete
            n_emp = n_emp + 1
        End If
    Next i
    n_col = n_col - n_emp
    r = FormatFont(data_out, n_row, n_col)
    For i = start_row + 2 To n_row
        flag = 1
        For j = 2 To n_col
            tcval = data_out.Cells(i, j).Value
            tcval = Replace(tcval, "'", vbNullString)
            tcval = Replace(tcval, "-", vbNullString)
            tcval = Trim$(tcval)
            If Len(tcval) > 0 Then flag = 0
        Next j
        If IsNumeric(Application.Match(Cells(i, 3), type_el_name.items, 0)) Then
            Cells(i, 1).Value = Cells(i, 3).Value
            Range(Cells(i, 1), Cells(i, 3)).Merge
            flag = 0
        End If
        If flag = 1 Then
            Range(data_out.Cells(i, 1), Cells(i, n_col)).Merge
            If Not (IsNumeric(Application.Match(Cells(i, 1), type_el_name.items, 0))) Then Range(data_out.Cells(i, 1), Cells(i, n_col)).Font.Bold = True
        End If
    Next i
    For i = 1 To 3
        Range(data_out.Cells(start_row, i), Cells(start_row + 1, i)).Merge
    Next i
    Range(data_out.Cells(start_row, 4), Cells(start_row, n_col - 3)).Merge
    For i = n_col - 2 To n_col
        Range(data_out.Cells(start_row, i), Cells(start_row + 1, i)).Merge
    Next i
    If this_sheet_option.Item("merge_material") Then
        n_c = 2
        start_row_t = start_row + 2
        n_start = start_row_t
        n_end = start_row
        temp_1 = data_out.Cells(n_start, n_c).MergeArea.Cells(1, 1).Value
        For i = start_row_t To n_row
            temp_2 = data_out.Cells(i, n_c).MergeArea.Cells(1, 1).Value
            If temp_1 <> temp_2 And temp_2 <> Empty Then
                temp_1 = temp_2
                If n_end > n_start Then Range(data_out.Cells(n_start, n_c), data_out.Cells(n_end, n_c)).Merge
                n_start = i
            Else
                n_end = i
            End If
            If i = n_row And temp_1 = temp_2 And temp_2 <> Empty Then
                n_end = i
                Range(data_out.Cells(n_start, n_c), data_out.Cells(n_end, n_c)).Merge
            End If
        Next i
    End If
    Range(data_out.Cells(start_row, 1), data_out.Cells(n_row, n_col)).Rows.AutoFit
    Range(data_out.Cells(start_row, 1), data_out.Cells(start_row, 1)).ColumnWidth = (s1 / sall) * koeff
    Range(data_out.Cells(start_row, 2), data_out.Cells(start_row, 2)).ColumnWidth = (s2 / sall) * koeff
    Range(data_out.Cells(start_row, 3), data_out.Cells(start_row, 3)).ColumnWidth = (s3 / sall) * koeff
    For i = 4 To n_col - 2
        Range(data_out.Cells(start_row, i), data_out.Cells(start_row, i)).ColumnWidth = (ssb / sall) * koeff
    Next i
    Range(data_out.Cells(start_row, n_col - 1), data_out.Cells(start_row, n_col - 1)).ColumnWidth = (s5 / sall) * koeff
    Range(data_out.Cells(start_row, n_col), data_out.Cells(start_row, n_col)).ColumnWidth = (s6 / sall) * koeff
End Function

Function FormatSpec_Fas(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean
    If n_col < 5 Or n_row < 2 Then
        If n_col < 5 Then n_col = 5
        If n_row < 2 Then n_row = 2
        Set data_out = Range(data_out.Cells(1, 1), data_out.Cells(n_row, n_col))
    End If
    data_out.Cells(1, 1) = "Поз." & vbLf & "отделки"
    data_out.Cells(1, 2) = "Наименование" & vbLf & "элементов фасада"
    data_out.Cells(1, 3) = "Наименование материала отделки"
    data_out.Cells(1, 4) = "Наименование и номер эталона цвета или образец колера"
    data_out.Cells(1, 5) = "Примечание"
    
    s1 = 20
    s2 = 45
    s3 = 65
    s4 = 30
    s5 = 25
    sall = s1 + s2 + s3 + s4 + s5
    koeff = (sall / 207.5) * 100
    dblPoints = Application.CentimetersToPoints(1)
    
    Range(data_out.Cells(1, 1), data_out.Cells(n_row, n_col)).Rows.AutoFit
    If Range(data_out.Cells(1, 1), data_out.Cells(1, n_col)).RowHeight < dblPoints * 1.5 Then
        Range(data_out.Cells(1, 1), data_out.Cells(1, n_col)).RowHeight = dblPoints * 1.5
    End If
    Range(data_out.Cells(1, 1), data_out.Cells(1, 1)).ColumnWidth = (s1 / sall) * koeff
    Range(data_out.Cells(1, 2), data_out.Cells(1, 2)).ColumnWidth = (s2 / sall) * koeff
    Range(data_out.Cells(1, 3), data_out.Cells(1, 3)).ColumnWidth = (s3 / sall) * koeff
    Range(data_out.Cells(1, 4), data_out.Cells(1, 4)).ColumnWidth = (s4 / sall) * koeff
    Range(data_out.Cells(1, 5), data_out.Cells(1, 5)).ColumnWidth = (s5 / sall) * koeff
    r = FormatFont(data_out, n_row, n_col)
End Function

Function FormatSpec_GR(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean
    start_cell = 1
    start_row = 2
        For j = start_row To n_row - 1
            data_out(j, 1) = Replace(data_out(j, 1), ",**", vbNullString)
            If (data_out(j - 1, 1).Value <> data_out(j, 1).Value) Then
                EndCell = j - 1
                data_out.Range(data_out.Cells(start_cell, 1), data_out.Cells(EndCell, 1)).Merge
                data_out.Range(data_out.Cells(start_cell, 6), data_out.Cells(EndCell, 6)).Merge
                start_cell = j
            End If
            If j = n_row - 1 Then
                EndCell = j
                data_out.Range(data_out.Cells(start_cell, 1), data_out.Cells(EndCell, 1)).Merge
                data_out.Range(data_out.Cells(start_cell, 6), data_out.Cells(EndCell, 6)).Merge
            End If
            If InStr(data_out(j, 1), "* расход на ") > 0 Then data_out.Range(data_out.Cells(j, 1), data_out.Cells(j, 6)).Merge
        Next j
    data_out.Range(data_out.Cells(n_row, 1), data_out.Cells(n_row, 6)).Merge
    koeff = (185 / 208) * 100
    
'    r = FormatFont(data_out, n_row, n_col)
    r = FormatFont(data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(n_row, n_col)), n_row, n_col)
    
    dblPoints = Application.CentimetersToPoints(1)
    data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(n_row, n_col)).Rows.AutoFit
    
    If data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(start_row, n_col)).RowHeight < dblPoints * 1.5 Then
        data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(start_row, n_col)).RowHeight = dblPoints * 1.5
    End If
    data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(n_row, 1)).Columns.AutoFit
    data_out.Range(data_out.Cells(start_row, 2), data_out.Cells(start_row, 2)).ColumnWidth = 0.07 * koeff
    data_out.Range(data_out.Cells(start_row, 3), data_out.Cells(start_row, 3)).ColumnWidth = 0.5 * koeff
    data_out.Range(data_out.Cells(start_row, 4), data_out.Cells(start_row, 4)).ColumnWidth = 0.07 * koeff
    data_out.Range(data_out.Cells(start_row, 5), data_out.Cells(start_row, 5)).ColumnWidth = 0.1 * koeff
    data_out.Range(data_out.Cells(start_row, 6), data_out.Cells(start_row, 6)).ColumnWidth = 0.1 * koeff
End Function

Function FormatSpec_KM(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean
    start_cell = 0
    start_row = 2
    For i = 1 To 2
        If start_cell = 0 Then start_cell = 1
            For j = start_row To n_row
                If (data_out.Cells(j - 1, i) <> data_out.Cells(j, i)) Then
                    EndCell = j - 1
                    Range(data_out.Cells(start_cell, i), data_out.Cells(EndCell, i)).Merge
                    start_cell = j
                End If
            Next j
        start_cell = 0
    Next i
    For i = start_row + 3 To n_row
        k = 0
        For j = 2 To 3
            tcval = data_out.Cells(i, j).Value
            tcval = Replace(tcval, "'", vbNullString)
            tcval = Replace(tcval, "-", vbNullString)
            tcval = Trim$(tcval)
            If Len(tcval) = 0 Then k = k + 1
        Next j
        tcval = data_out.Cells(i, 1).Value
        tcval = Replace(tcval, "'", vbNullString)
        tcval = Replace(tcval, "-", vbNullString)
        tcval = Trim$(tcval)
        If k = 2 Then Range(data_out.Cells(i, 1), data_out.Cells(i, 3)).Merge
        If data_out.Cells(i, 2).Value = "Итого" Then Range(data_out.Cells(i, 2), data_out.Cells(i, 3)).Merge
        If tcval = "Всего масса металла:" Then r_obsh = i
        If tcval = "Антикоррозийная окраска" Then
            data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, n_col)).Merge
            r_okr = i
        End If
        If i > r_okr And r_okr <> 0 Then data_out.Range(data_out.Cells(i, 4), data_out.Cells(i, n_col - 1)).Merge
    Next i
    data_out.Range(data_out.Cells(start_row, 3), data_out.Cells(start_row + 1, 3)).Merge
    data_out.Range(data_out.Cells(start_row, 4), data_out.Cells(start_row + 1, 4)).Merge
    data_out.Range(data_out.Cells(start_row, 5), data_out.Cells(start_row, n_col - 1)).Merge
    data_out.Range(data_out.Cells(start_row, n_col), data_out.Cells(start_row + 1, n_col)).Merge
    
    r = FormatFont(data_out, n_row, n_col)
    If Not IsEmpty(r_okr) Then
        data_out.Range(data_out.Cells(start_row + 3, 5), data_out.Cells(r_okr, n_col)).NumberFormat = w_format
        data_out.Range(data_out.Cells(r_okr, 5), data_out.Cells(n_row, n_col)).NumberFormat = "0.00"
    End If
    
    dblPoints = Application.CentimetersToPoints(1)
    data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(start_row + 1, n_col)).RowHeight = dblPoints * 1.5
    data_out.Range(data_out.Cells(start_row + 2, 1), data_out.Cells(start_row + 2, n_col)).RowHeight = dblPoints * 0.4
    data_out.Range(data_out.Cells(start_row + 3, 1), data_out.Cells(n_row, n_col)).Rows.AutoFit
    koeff = 5
    data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(start_row + 1, 3)).ColumnWidth = 3 * koeff
    data_out.Range(data_out.Cells(start_row, 4), data_out.Cells(start_row + 1, 4)).ColumnWidth = 1 * koeff
    data_out.Range(data_out.Cells(start_row, 5), data_out.Cells(start_row + 1, n_col - 1)).ColumnWidth = 1.5 * koeff
    data_out.Range(data_out.Cells(start_row, n_col), data_out.Cells(start_row + 1, n_col)).ColumnWidth = 2.5 * koeff

    Set MyRange = data_out.Range(data_out.Cells(r_obsh, n_col), data_out.Cells(r_obsh, n_col))
    MyRange.Font.Bold = True
    With MyRange.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
End Function

Function FormatSpec_KZH(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean
    start_row = 2
    If IsEmpty(data_out.Cells(start_row, 1).Value) Then Exit Function
    'Удаляем пустые строки
    n_del = 0
    For i = start_row + 5 To n_row
        data_out(i, 1).Value = Replace(data_out(i, 1).Value, ",**", vbNullString)
        flag = 1
        For j = 2 To n_col
            tcval = data_out.Cells(i, j).Value
            tcval = Replace(tcval, "'", vbNullString)
            tcval = Replace(tcval, "-", vbNullString)
            tcval = Trim$(tcval)
            If Len(tcval) > 0 Then
                flag = 0
                j = n_col
            End If
        Next j
        If flag Then
            data_out.Rows(i).Delete Shift:=xlUp
            n_del = n_del + 1
        End If
    Next i
    n_row = n_row - n_del
    If n_row = start_row + 6 Then
        data_out.Rows(start_row + 6).Delete Shift:=xlUp
        n_row = start_row + 5
    End If

    n_col_bet = 0 'Начало столбцов с бетоном
    'Объединение по колонкам
    For j = 1 To n_col
        flag_merge = False
        For i = start_row To start_row + 3
            end_val = data_out.Cells(i, j).Value
            If InStr(end_val, "Всего") > 0 Then flag_merge = True
            If InStr(end_val, "Марка элемента") > 0 Then flag_merge = True
            If InStr(end_val, "бетон") > 0 And InStr(end_val, "Объём") = 0 Then
                flag_merge = True
                n_col_bet = j
            End If
            If flag_merge Then
                data_out.Range(data_out.Cells(i, j), data_out.Cells(start_row + 4, j)).Merge
                i = start_row + 3
            End If
        Next i
    Next j
    'Объединение по строкам
    For i = start_row To start_row + 3
        start_cell = 2
        start_val = data_out.Cells(i, start_cell).Value
        For j = 2 To n_col
            end_val = data_out.Cells(i, j).Value
            If end_val <> start_val Then
                end_cell = j - 1
                If end_cell <> start_cell And data_out.Cells(i, end_cell).MergeCells = False And data_out.Cells(i, start_cell).MergeCells = False Then
                    data_out.Range(data_out.Cells(i, start_cell), data_out.Cells(i, end_cell)).Merge
                End If
                start_cell = j
                start_val = data_out.Cells(i, start_cell).Value
            End If
        Next j
    Next i

    r = FormatRowHigh(0.8, data_out)
    r = FormatColWidth(1.5, data_out)
    r = FormatColWidth(3, data_out.Columns(1))
    r = FormatFont(data_out.Range(data_out.Cells(start_row, 1), data_out.Cells(start_row + 4, n_col)), 5, n_col)
    r = FormatFont(data_out.Range(data_out.Cells(start_row + 5, 1), data_out.Cells(n_row, n_col)), n_row - 6, n_col)
    With data_out.Range(data_out.Cells(start_row + 5, 2), data_out.Cells(n_row, n_col))
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

    Range(data_out.Cells(n_row, 1), data_out.Cells(n_row, n_col)).Font.Bold = True
    With data_out.Cells(n_row, n_col).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
'    If n_col_bet > 0 Then
'        r = FormatFont(data_out.Range(data_out.Cells(start_row, n_col + 1), data_out.Cells(n_row, n_col + n_col_bet)), n_row, n_col + n_col_bet)
'        With data_out.Cells(n_row, n_col + n_col_bet).Interior
'            .Pattern = xlSolid
'            .PatternColorIndex = xlAutomatic
'            .ThemeColor = xlThemeColorAccent2
'            .TintAndShade = 0.599993896298105
'            .PatternTintAndShade = 0
'        End With
'        For i = n_col + 1 To n_col + n_col_bet - 1
'            If InStr(data_out(start_row + 1, i), "етон") > 0 Then
'                With data_out.Range(data_out.Cells(start_row + 1, i), data_out.Cells(start_row + 4, i))
'                    .HorizontalAlignment = xlCenter
'                    .VerticalAlignment = xlCenter
'                    .WrapText = True
'                    .Orientation = 90
'                    .AddIndent = False
'                    .IndentLevel = 0
'                    .ShrinkToFit = False
'                    .ReadingOrder = xlContext
'                    .MergeCells = True
'                End With
'            End If
'        Next i
'    End If
End Function

Function FormatSpec_RSK(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean
    Range(data_out.Cells(1, 1), data_out.Cells(n_row, 1)).ColumnWidth = 10
    Range(data_out.Cells(1, 2), data_out.Cells(n_row, 2)).ColumnWidth = 8
    Range(data_out.Cells(1, 3), data_out.Cells(n_row, 3)).ColumnWidth = 10
    For i = 4 To n_col - 1
        Range(data_out.Cells(1, i), data_out.Cells(n_row, i)).ColumnWidth = 8
    Next i
    Range(data_out.Cells(1, n_col), data_out.Cells(n_row, n_col)).ColumnWidth = 30
    r = FormatFont(data_out, n_row, n_col)
    For n_c = 1 To 2
        start_row = 2
        n_start = start_row
        n_end = start_row
        temp_1 = data_out.Cells(n_start, n_c).MergeArea.Cells(1, 1).Value
        For i = start_row To n_row
            temp_2 = data_out.Cells(i, n_c).MergeArea.Cells(1, 1).Value
            If temp_1 <> temp_2 And temp_2 <> Empty Then
                temp_1 = temp_2
                If n_end > n_start Then Range(data_out.Cells(n_start, n_c), data_out.Cells(n_end, n_c)).Merge
                n_start = i
            Else
                n_end = i
            End If
            If i = n_row And temp_1 = temp_2 And temp_2 <> Empty Then
                n_end = i
                Range(data_out.Cells(n_start, n_c), data_out.Cells(n_end, n_c)).Merge
            End If
        Next i
    Next n_c
    Range("A1:H1").Select
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
    Selection.Merge
    Rows("1:1").RowHeight = 24
    Range(data_out.Cells(1, 1), data_out.Cells(n_row, n_col)).Rows.AutoFit
End Function

Function FormatSpec_NRM(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean

    r = FormatFont(data_out, n_row, n_col)
    Range(data_out.Cells(1, 1), data_out.Cells(n_row, n_col)).Rows.AutoFit
    
    Range(data_out.Cells(1, 1), data_out.Cells(n_row, 1)).ColumnWidth = 15
    Range(data_out.Cells(1, 2), data_out.Cells(n_row, 2)).ColumnWidth = 25
    Range(data_out.Cells(1, 3), data_out.Cells(n_row, 5)).ColumnWidth = 15

    Range(data_out.Cells(n_row, 1), data_out.Cells(n_row, n_col)).Font.Bold = True
    Range(data_out.Cells(1, 1), data_out.Cells(1, n_col)).Font.Bold = True
    
    data_out.Range(data_out.Cells(2, 1), data_out.Cells(n_row, n_col)).Select
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
    data_out.Cells(n_row, n_col).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
   
    data_out.Range(data_out.Cells(2, 5), data_out.Cells(n_row - 1, 5)).Select
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
 
Function FormatSpec_Pol(ByVal data_out As Range) As Boolean
    CSVfilename$ = ThisWorkbook.path & "\list\Спец_" & ThisWorkbook.ActiveSheet.Name & ".txt"
    n = ExportList2CSV(data_out, CSVfilename$)
    MsgBox ("Данные о полах записаны в файл" & vbLf & "\list\Спец_" & ThisWorkbook.ActiveSheet.Name & ".txt")
    FormatSpec_Pol = True
End Function

Function FormatSpec_Perem(ByVal data_out As Variant, ByVal n_row As Long) As Boolean
    istart = 1
    For i = 1 To 4
        If Len(data_out(i, 1)) > 0 And InStr(data_out(i, 1), "Поз.") = 0 And istart = 1 Then istart = i
    Next i
    Dim pos_out
    ReDim pos_out(n_row - istart + 1, 3)
    For i = istart To n_row
        pos = CStr(data_out(i, 1).Value)
        naen = CStr(data_out(i, 3).Value)
        obozn = CStr(data_out(i, 2).Value)
        If Len(pos) < 1 Then pos = "   "
        If Len(naen) < 1 Then naen = "   "
        If Len(obozn) < 1 Then obozn = "-"
        pos_out(i - istart + 1, 1) = pos
        pos_out(i - istart + 1, 2) = naen
        pos_out(i - istart + 1, 3) = obozn
    Next i
    CSVfilename$ = ThisWorkbook.path & "\list\Поз_" & ThisWorkbook.ActiveSheet.Name & ".txt"
    n = ExportArray2CSV(pos_out, CSVfilename$)
    MsgBox ("Данные о позициях перемычек записаны в файл" & vbLf & "\list\Поз_" & ThisWorkbook.ActiveSheet.Name & ".txt")
    FormatSpec_Perem = True
End Function


Function FormatSpec_Split(ByVal data_out As Range) As Boolean
    data_out.Range("A1").FormulaR1C1 = "Имя листа"
    data_out.Range("B1").FormulaR1C1 = "Список значений параметров зоны"
    data_out.Range("C1").FormulaR1C1 = "Номер столбца параметров"
    With data_out
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
    With data_out.Font
        .Name = fontname
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
    data_out.Borders(xlDiagonalDown).LineStyle = xlNone
    data_out.Borders(xlDiagonalUp).LineStyle = xlNone
    data_out.Borders(xlEdgeLeft).LineStyle = xlNone
    data_out.Borders(xlEdgeTop).LineStyle = xlNone
    data_out.Borders(xlEdgeBottom).LineStyle = xlNone
    data_out.Borders(xlEdgeRight).LineStyle = xlNone
    With data_out.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With data_out.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    data_out.Borders(xlDiagonalDown).LineStyle = xlNone
    data_out.Borders(xlDiagonalUp).LineStyle = xlNone
    With data_out.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With data_out.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With data_out.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With data_out.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With data_out.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With data_out.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    data_out.Columns("B:B").ColumnWidth = 35
    data_out.Columns("C:C").ColumnWidth = 11.57
    data_out.Cells.Rows.AutoFit
End Function

Function FormatSpec_Rule(ByVal data_out As Range) As Boolean
    data_out.Select
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
        .Name = fontname
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
    Selection.FormatConditions.Add Type:=xlTextString, String:="Стены-разделители зон", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = vbYellow
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

Function FormatSpec_merge(ByRef data_out As Range, ByVal n_col As Long, Optional ByVal is_empty_eq As Boolean, Optional ByVal n_col_arr As Variant) As Boolean
    n_row = data_out.Rows.Count
    n_start = 1
    n_end = 1
    temp_1 = data_out.Cells(n_start, n_col).MergeArea.Cells(1, 1).Value
    temp_1 = Replace(temp_1, "'", vbNullString)
    temp_1 = Replace(temp_1, "-", vbNullString)
    temp_1 = Trim$(temp_1)
    For i = 1 To n_row
        temp_2 = data_out.Cells(i, n_col).MergeArea.Cells(1, 1).Value
        temp_2 = Replace(temp_2, "'", vbNullString)
        temp_2 = Replace(temp_2, "-", vbNullString)
        temp_2 = Trim$(temp_2)
        If is_empty_eq And (IsEmpty(temp_2) Or Len(temp_2) = 0) Then temp_1 = temp_2
        If temp_1 <> temp_2 Or i = n_row Then
            If i = n_row And temp_1 = temp_2 Then
                n_end = i
            Else
                temp_1 = temp_2
            End If
            If n_end > n_start Then
                Range(data_out.Cells(n_start, n_col), data_out.Cells(n_end, n_col)).Merge
                If Not IsMissing(n_col_arr) Then
                    For k = LBound(n_col_arr) To UBound(n_col_arr)
                        data_out.Range(data_out.Cells(n_start, n_col_arr(k)), data_out.Cells(n_end, n_col_arr(k))).Merge
                    Next k
                End If
            End If
            n_start = i
        Else
            n_end = i
        End If
    Next i
End Function

Function FormatSpec_Ved(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean
    s_mat = 5
    s_ar = 1.5
    s1 = 1
    s2 = 5
    sp = 3
    Cells.UnMerge
    Cells.NumberFormat = "@"
    Range(data_out.Cells(1, 1), data_out.Cells(2, 1)).Merge
    Range(data_out.Cells(1, 2), data_out.Cells(2, 2)).Merge
    Range(data_out.Cells(1, 3), data_out.Cells(1, n_col - 1)).Merge
    Range(data_out.Cells(1, n_col), data_out.Cells(2, n_col)).Merge
    For i = 1 To n_row
        If InStr(data_out.Cells(i, 1), "Общяя площадь") > 0 Or InStr(data_out.Cells(i, 5), "Общяя площадь") > 0 Then
            n_all = n_row
            n_row = i - 1
            n_start_all = i
        End If
    Next i
    If n_all = Empty Then
        n_all = n_row
        n_start_all = n_all
    End If
    If this_sheet_option.Item("otd_by_type") Then
        If zonenum_pot = False Then
            n_cst = 3
        Else
            n_cst = 2
        End If
    Else
        n_cst = 3
    End If
    n_start = 3
    n_end = 3
    temp_1 = data_out.Cells(n_start, 1).MergeArea.Cells(1, 1).Value
    temp_1 = Replace(temp_1, "'", vbNullString)
    temp_1 = Replace(temp_1, "-", vbNullString)
    temp_1 = Trim$(temp_1)
    For i = 3 To n_row
    'Идём по номерам помещений или типам отделки
        temp_2 = data_out.Cells(i, 1).MergeArea.Cells(1, 1).Value
        temp_2 = Replace(temp_2, "'", vbNullString)
        temp_2 = Replace(temp_2, "-", vbNullString)
        temp_2 = Trim$(temp_2)
        If IsEmpty(temp_2) Or Len(temp_2) = 0 Then temp_2 = temp_1
        If temp_1 <> temp_2 Or i = n_row Then
            If i = n_row Then n_end = i
            temp_1 = temp_2
            If n_end > n_start Then
                Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, 1)).Merge
                If zonenum_pot = False Or Not this_sheet_option.Item("otd_by_type") Then Range(data_out.Cells(n_start, 2), data_out.Cells(n_end, 2)).Merge
                Range(data_out.Cells(n_start, n_col), data_out.Cells(n_end, n_col)).Merge
                For n_c = n_cst To n_col - 1

                    n_start_ = n_start
                    n_end_ = n_start
                    temp_1 = data_out.Cells(n_start_, n_c).MergeArea.Cells(1, 1).Value
                    temp_1 = Replace(temp_1, "'", vbNullString)
                    temp_1 = Replace(temp_1, "-", vbNullString)
                    temp_1 = Trim$(temp_1)
                    For i_ = n_start To n_end
                        temp_2 = data_out.Cells(i_, n_c).MergeArea.Cells(1, 1).Value
                        temp_2 = Replace(temp_2, "'", vbNullString)
                        temp_2 = Replace(temp_2, "-", vbNullString)
                        temp_2 = Trim$(temp_2)
                        If Len(temp_2) = 0 And i_ < n_end Then
                            n_end_ = i_
                        Else
                            If i_ = n_end And Len(temp_2) = 0 Then n_end_ = i_
                            If n_end_ > n_start_ Then Range(data_out.Cells(n_start_, n_c), data_out.Cells(n_end_, n_c)).Merge
                            n_start_ = i_
                        End If

'                        If Len(temp_2) = 0 Then temp_1 = temp_2
'                        If temp_1 <> temp_2 Or Len(temp_2) = 0 Or i_ = n_end Then
'                            If i_ = n_end And temp_1 = temp_2 Then
'                                n_end_ = i_
'                            Else
'                                temp_1 = temp_2
'                            End If
'                            'If n_end_ > n_start_ Then Range(data_out.Cells(n_start_, n_c), data_out.Cells(n_end_, n_c)).Merge
'                            n_start_ = i_
'                        Else
'                            n_end_ = i_
'                        End If
                    Next i_

                    
'                    r = FormatSpec_merge(data_out.Range(data_out.Cells(n_start, n_cst), data_out.Cells(n_end, n_col - 1)), n_c, True)
                Next n_c
                inside_weight = xlThin
            Else
                inside_weight = xlMedium
            End If
            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlMedium
            End With
            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = xlThin
            End With
            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .ColorIndex = 0
                .TintAndShade = 0
                .Weight = inside_weight
            End With
            n_start = i
        Else
            n_end = i
        End If
    Next i
    
    
    
'
'    For i = 3 To n_row
'        temp = data_out.Cells(i, 1).MergeArea.Cells(1, 1).Value
'        If temp = Empty Or temp = "-" Then n_end = i
'        If temp <> Empty And temp <> "-" Then
'            If n_end > n_start Then
'                Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, 1)).Merge
'                If zonenum_pot = False Then Range(data_out.Cells(n_start, 2), data_out.Cells(n_end, 2)).Merge
'                Range(data_out.Cells(n_start, n_col), data_out.Cells(n_end, n_col)).Merge
'                With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeLeft)
'                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
'                    .Weight = xlMedium
'                End With
'                With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeTop)
'                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
'                    .Weight = xlMedium
'                End With
'                With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeBottom)
'                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
'                    .Weight = xlMedium
'                End With
'                With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeRight)
'                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
'                    .Weight = xlMedium
'                End With
'                With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlInsideVertical)
'                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
'                    .Weight = xlThin
'                End With
'                With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlInsideHorizontal)
'                    .LineStyle = xlContinuous
'                    .ColorIndex = 0
'                    .TintAndShade = 0
'                    .Weight = xlThin
'                End With
'            End If
'            n_start = i
'        End If
'        If i = n_row And temp = Empty Or temp = "-" Then
'            n_end = i
'            Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, 1)).Merge
'            If zonenum_pot = False Then Range(data_out.Cells(n_start, 2), data_out.Cells(n_end, 2)).Merge
'            Range(data_out.Cells(n_start, n_col), data_out.Cells(n_end, n_col)).Merge
'            Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlDiagonalDown).LineStyle = xlNone
'            Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlDiagonalUp).LineStyle = xlNone
'
'            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeLeft)
'                .LineStyle = xlContinuous
'                .ColorIndex = 0
'                .TintAndShade = 0
'                .Weight = xlMedium
'            End With
'            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeTop)
'                .LineStyle = xlContinuous
'                .ColorIndex = 0
'                .TintAndShade = 0
'                .Weight = xlMedium
'            End With
'            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeBottom)
'                .LineStyle = xlContinuous
'                .ColorIndex = 0
'                .TintAndShade = 0
'                .Weight = xlMedium
'            End With
'            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlEdgeRight)
'                .LineStyle = xlContinuous
'                .ColorIndex = 0
'                .TintAndShade = 0
'                .Weight = xlMedium
'            End With
'            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlInsideVertical)
'                .LineStyle = xlContinuous
'                .ColorIndex = 0
'                .TintAndShade = 0
'                .Weight = xlThin
'            End With
'            With Range(data_out.Cells(n_start, 1), data_out.Cells(n_end, n_col)).Borders(xlInsideHorizontal)
'                .LineStyle = xlContinuous
'                .ColorIndex = 0
'                .TintAndShade = 0
'                .Weight = xlThin
'            End With
'        End If
'    Next i
'    For n_c = n_cst To n_col - 1
'        n_start = 3
'        n_end = 3
'        For i = 3 To n_row
'            temp = data_out.Cells(i, n_c).MergeArea.Cells(1, 1).Value
'            If temp = Empty Then n_end = i
'            If temp <> Empty Then
'                If n_end > n_start Then Range(data_out.Cells(n_start, n_c), data_out.Cells(n_end, n_c)).Merge
'                n_start = i
'            End If
'            If i = n_row And temp = Empty Then
'                n_end = i
'                Range(data_out.Cells(n_start, n_c), data_out.Cells(n_end, n_c)).Merge
'            End If
'        Next i
'    Next n_c

    If this_sheet_option.Item("otd_by_type") Then
        For n_c = 3 To n_col - 1
            If InStr(data_out.Cells(2, n_c), "Высота") > 0 Then
                temp_1 = data_out.Cells(n_start, n_c).MergeArea.Cells(1, 1).Value
                n_start = 3
                n_end = 3
                For i = 3 To n_row
                    temp_2 = data_out.Cells(i, n_c).MergeArea.Cells(1, 1).Value
                    If temp_1 <> temp_2 And temp_2 <> Empty Then
                        temp_1 = temp_2
                        If n_end > n_start Then Range(data_out.Cells(n_start, n_c), data_out.Cells(n_end, n_c)).Merge
                        n_start = i
                    Else
                        n_end = i
                    End If
                    If i = n_row And temp_1 = temp_2 And temp_2 <> Empty Then
                        n_end = i
                        Range(data_out.Cells(n_start, n_c), data_out.Cells(n_end, n_c)).Merge
                    End If
                Next i
            End If
        Next n_c
    End If
    
    If show_mat_area Then
        Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_start_all, 4)).Merge
        For i = n_start_all + 1 To n_all
            If Len(data_out.Cells(i, 4).Value) > 1 Then
                Range(data_out.Cells(i, 1), data_out.Cells(i, 3)).Merge
            Else
                Range(data_out.Cells(i, 1), data_out.Cells(i, 4)).Merge
            End If
        Next i
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlDiagonalDown).LineStyle = xlNone
        Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlDiagonalUp).LineStyle = xlNone
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlDiagonalDown).LineStyle = xlNone
        Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlDiagonalUp).LineStyle = xlNone
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, 1), data_out.Cells(n_all, 4)).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End If
    
    If show_surf_area Then
        nn_col = 5
        If Not show_mat_area Then nn_col = 1
        Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_start_all, nn_col + 3)).Merge
        For i = n_start_all + 1 To n_all
            If Len(data_out.Cells(i, nn_col).Value) > 1 Then
                Range(data_out.Cells(i, nn_col), data_out.Cells(i, nn_col + 2)).Merge
                n_surf = i
            End If
        Next i
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Interior
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = -0.249977111117893
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlDiagonalDown).LineStyle = xlNone
        Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlDiagonalUp).LineStyle = xlNone
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlDiagonalDown).LineStyle = xlNone
        Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlDiagonalUp).LineStyle = xlNone
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(n_start_all, nn_col), data_out.Cells(n_surf, nn_col + 3)).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
    End If
    
    
    With Range(data_out.Cells(1, 1), data_out.Cells(n_start_all - 1, n_col)).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        With Range(data_out.Cells(1, 1), data_out.Cells(n_start_all - 1, n_col)).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(1, 1), data_out.Cells(2, n_col)).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(1, 1), data_out.Cells(2, n_col)).Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(1, 1), data_out.Cells(2, n_col)).Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlMedium
        End With
        With Range(data_out.Cells(1, 1), data_out.Cells(2, n_col)).Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With Range(data_out.Cells(1, 1), data_out.Cells(2, n_col)).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .Weight = xlThin
        End With
        
    With data_out.Font
        .Name = fontname
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
    
    With data_out
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    
     With Range(data_out.Cells(1, 1), data_out.Cells(2, n_col)).Font
        .Name = fontname
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
    

    dblPoints = Application.CentimetersToPoints(1)
    
    r = FormatRowHigh(0.5, data_out.Rows(1))
    r = FormatRowHigh(0.8, Range(data_out.Cells(2, 1), data_out.Cells(n_row, n_col)))
    
    r = FormatColWidth(s1, data_out.Columns(1))
    r = FormatColWidth(s2, data_out.Columns(2))
    r = FormatColWidth(s_mat, data_out.Columns(3))
    r = FormatColWidth(s_ar, data_out.Columns(4))
    r = FormatColWidth(s_mat, data_out.Columns(5))
    r = FormatColWidth(s_ar, data_out.Columns(6))

    If data_out.Cells(2, 7).Value = "Колонн" Then
        r = FormatColWidth(s_mat, data_out.Columns(7))
        r = FormatColWidth(s_ar, data_out.Columns(8))
        If data_out.Cells(2, 9).Value = "Низа стен/колонн" Then
            r = FormatColWidth(s_mat, data_out.Columns(9))
            r = FormatColWidth(s_ar, data_out.Columns(10))
            r = FormatColWidth(s_ar, data_out.Columns(11))
        End If
    Else
        If data_out.Cells(2, 7).Value = "Низа стен/колонн" Then
            r = FormatColWidth(s_mat, data_out.Columns(7))
            r = FormatColWidth(s_ar, data_out.Columns(8))
            r = FormatColWidth(s_ar, data_out.Columns(9))
        End If
    End If
    For n_c = 1 To n_col
        g = data_out.Cells(2, n_c).Value
        If InStr(g, "Площадь") Or InStr(g, "Высота") Then
            data_out.Cells(2, n_c).Orientation = 90
            Range(data_out.Cells(3, n_c), data_out.Cells(n_row, n_c)).Font.Size = 11
            Range(data_out.Cells(3, n_c), data_out.Cells(n_row, n_c)).ShrinkToFit = True
        End If
        g = data_out.Cells(1, n_c).Value
        If InStr(g, "Тип") And this_sheet_option.Item("otd_by_type") Then Range(data_out.Cells(1, n_c), data_out.Cells(n_start_all - 1, n_c)).Orientation = 90
        If InStr(g, "Номер") And Not this_sheet_option.Item("otd_by_type") Then data_out.Cells(1, n_c).Orientation = 90
    Next n_c
    r = FormatColWidth(sp, data_out.Columns(n_col))
    data_out.FormatConditions.Add Type:=xlTextString, String:="НЕТ ОТДЕЛКИ", TextOperator:=xlContains
    With data_out.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With data_out.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    FormatSpec_Ved = True
End Function

Function FormatSpec_VOR(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean

End Function


Function FormatSpec_WIN(ByVal data_out As Range, ByVal n_row As Long, ByVal n_col As Long) As Boolean
    s1 = 1.5
    s2 = 5.5
    s3 = 6.5
    sqty = 2
    sprim = 2.5
    data_out.NumberFormat = "@"
    If by_floor Then
        start_row = 3
    Else
        start_row = 2
    End If
    n_col_m = 1: If this_sheet_option.Item("merge_material") Then n_col_m = 2
    For n_c = 1 To n_col_m
        n_start = start_row
        n_end = start_row
        temp_1 = data_out.Cells(n_start, n_c).MergeArea.Cells(1, 1).Value
        For i = start_row To n_row
            temp_2 = data_out.Cells(i, n_c).MergeArea.Cells(1, 1).Value
            If temp_1 <> temp_2 And temp_2 <> Empty Then
                temp_1 = temp_2
                If n_end > n_start Then Range(data_out.Cells(n_start, n_c), data_out.Cells(n_end, n_c)).Merge
                n_start = i
            Else
                n_end = i
            End If
            If i = n_row And temp_1 = temp_2 And temp_2 <> Empty Then
                n_end = i
                Range(data_out.Cells(n_start, n_c), data_out.Cells(n_end, n_c)).Merge
            End If
        Next i
    Next n_c
    r = FormatColWidth(s1, data_out.Columns(1))
    r = FormatColWidth(s2, data_out.Columns(2))
    r = FormatColWidth(s3, data_out.Columns(3))
    r = FormatColWidth(sprim, Range(data_out.Cells(1, n_col - 1), data_out.Cells(n_row, n_col)))
    If this_sheet_option.Item("qty_one_floor") Then
        For i = 1 To 3
            Range(data_out.Cells(1, i), data_out.Cells(2, i)).Merge
        Next i
        Range(data_out.Cells(1, 4), data_out.Cells(1, n_col - 3)).Merge
        For i = n_col - 2 To n_col
            Range(data_out.Cells(1, i), data_out.Cells(2, i)).Merge
        Next i
        r = FormatRowHigh(0.5, data_out.Rows(1))
        r = FormatRowHigh(0.8, data_out.Rows(2))
        r = FormatRowHigh(0.8, Range(data_out.Cells(3, n_col), data_out.Cells(n_row, n_col)))
        r = FormatColWidth(sqty, Range(data_out.Cells(1, 4), data_out.Cells(n_row, n_col - 2)))
        r = FormatRowPrint(data_out, 2)
    Else
        r = FormatRowHigh(1.5, data_out.Rows(1))
        r = FormatRowHigh(0.8, Range(data_out.Cells(2, n_col), data_out.Cells(n_row, n_col)))
        r = FormatColWidth(sqty, Range(data_out.Cells(1, 4), data_out.Cells(n_row, n_col - 2)))
        r = FormatRowPrint(data_out, 1)
    End If
    r = FormatFont(data_out, n_row, n_col)
End Function

Function FormatTable(ByVal nm As String, Optional ByVal pos_out As Variant, Optional ByVal str_kzap As String = Empty) As Boolean
Dim tfunctime As Double
tfunctime = Timer

    If IsEmpty(this_sheet_option) Then
        r = OptionSheetSet(nm)
        Set this_sheet_option = OptionGetForm(nm)
        r = SetKzap()
    End If

    Set Sh = wbk.Sheets(nm)
    If IsError(pos_out) Or IsEmpty(pos_out) Then
        lsize = SheetGetSize(Sh)
        n_row = lsize(1)
        If InStr(Sh.Cells(n_row, 1).Value, "k=") > 0 Then
            str_kzap = Sh.Cells(n_row, 1).Value
            Sh.Rows(n_row).Delete
            n_row = n_row - 1
        End If
        n_col = lsize(2)
        Set data_out = Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col))
    Else
        n_row = UBound(pos_out, 1)
        n_col = UBound(pos_out, 2)
        pos_out = ArrayEmp2Space(pos_out)
        Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col)) = pos_out
        Set data_out = Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col))
    End If
    type_spec = SpecGetType(Sh.Name)
    If type_spec <> 7 Then r = FormatClear(Sh)
    Select Case type_spec
        Case 1
            r = FormatSpec_GR(data_out, n_row, n_col)
        Case 2, 3
            If this_sheet_option.Item("qty_one_floor") And spec_version > 1 Then
                r = FormatSpec_ASGR(data_out, n_row, n_col)
            Else
                r = FormatSpec_AS(data_out, n_row, n_col)
            End If
        Case 4
            r = FormatSpec_KM(data_out, n_row, n_col)
        Case 5
            r = FormatSpec_KZH(data_out, n_row, n_col)
        Case 6
            r = FormatSpec_AS(data_out, n_row, n_col)
        Case 8
            r = FormatSpec_Fas(data_out, n_row, n_col)
        Case 11
            r = FormatSpec_Ved(data_out, n_row, n_col)
        Case 12
            r = FormatSpec_Pol(data_out)
        Case 13
            r = FormatSpec_ASGR(data_out, n_row, n_col)
        Case 14
            r = FormatSpec_NRM(data_out, n_row, n_col)
        Case 20
            r = FormatSpec_WIN(data_out, n_row, n_col)
        Case 21
            r = FormatSpec_Split(data_out)
        Case 25
            r = FormatSpec_RSK(data_out, n_row, n_col)
    End Select
    
    If this_sheet_option.Item("title_on") = True Then
        dblPoints = Application.CentimetersToPoints(1)
        Sh.Range(Sh.Cells(1, 1), Sh.Cells(1, n_col)).Merge
        If Not IsEmpty(Sh.Cells(1, 1).Value) Then
            With Sh.Range(Sh.Cells(1, 1), Sh.Cells(1, n_col))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = True
                .ReadingOrder = xlContext
            End With
            With Sh.Range(Sh.Cells(1, 1), Sh.Cells(1, n_col)).Font
                .Name = fontname
                .FontStyle = "обычный"
                .Size = 14
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
            Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col)).Rows.AutoFit
    '        If sh.Range(sh.Cells(1, 1), sh.Cells(1, n_col)).RowHeight < dblPoints * 1.5 Then
    '            sh.Range(sh.Cells(1, 1), sh.Cells(1, n_col)).RowHeight = dblPoints * 1.5
    '        End If
        Else
            Sh.Range(Sh.Cells(1, 1), Sh.Cells(1, n_col)).RowHeight = dblPoints * 0.01
        End If
        If Not IsEmpty(str_kzap) Then
            Sh.Cells(n_row + 1, 1).Value = str_kzap
            With Sh.Cells(n_row + 1, 1).Font
                .Name = fontname
                .Size = 3
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
            With Sh.Cells(n_row + 1, 1).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Sh.Range(Sh.Cells(n_row + 1, 1), Sh.Cells(n_row + 1, n_col)).RowHeight = dblPoints * 0.1
        End If
    End If
    FormatTable = True
tfunctime = functime("FormatTable", tfunctime)
End Function

Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal mask As String, ByRef FSO, ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
    On Error Resume Next: Set curfold = FSO.GetFolder(FolderPath)
    If Not curfold Is Nothing Then
        For Each fil In curfold.Files
            If fil.Name Like "*" & mask Then FileNamesColl.Add fil.path
        Next
        SearchDeep = SearchDeep - 1
        If SearchDeep Then
           For Each sfol In curfold.SubFolders
               GetAllFileNamesUsingFSO sfol.path, mask, FSO, FileNamesColl, SearchDeep
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
Dim tfunctime As Double
tfunctime = Timer
    If IsEmpty(gost2fklass) Then r = ReadReinforce()
    klass = Replace(klass, "А", "A")
    klass = Replace(klass, "С", "C")
    gost = gost2fklass.Item(klass)
    If Len(swap_gost.Item(gost)) > 0 Then gost = swap_gost.Item(gost)
    GetGOSTForKlass = gost
tfunctime = functime("GetGOSTForKlass", tfunctime)
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

Function GetZoneParam(ByVal param_zone As String, ByVal param_name As String) As Variant
Dim tfunctime As Double
tfunctime = Timer
    GetZoneParam = Empty
    If InStr(param_zone, "@" + param_name + "=") > 0 Then
        For Each el In Split(param_zone, "@")
            If InStr(el, "=") > 0 Then
                name_param = Split(el, "=")(0)
                If name_param = param_name Then
                    value_param = ConvTxt2Num(Split(el, "=")(1))
                    GetZoneParam = value_param
                    Exit Function
                End If
            End If
        Next
    End If
tfunctime = functime("GetZoneParam", tfunctime)
End Function

Function GetClassBeton(ByVal txt As String) As String
Dim tfunctime As Double
tfunctime = Timer
    class = Empty
    If InStr(txt, "етон") > 0 Then
        txt = Trim$(Replace(txt, "  ", " "))
        wrd = Split(txt, " ")
        For Each w In wrd
            w = Trim$(Replace(Replace(w, "B", vbNullString), "В", vbNullString))
            If IsNumeric(ConvTxt2Num(w)) Then
                class = w
            End If
        Next
        txt = Trim$(Replace(Replace(txt, "B", vbNullString), "В", vbNullString))
    End If
    GetClassBeton = class
tfunctime = functime("GetClassBeton", tfunctime)
End Function

Function GetListFile(ByRef mask As String) As Variant
    path = ThisWorkbook.path & "\import"
    Set coll = FilenamesCollection(path, mask)
    If coll.Count <= 0 Then
        GetListFile = Empty
        Exit Function
    End If
    Dim out(): ReDim out(coll.Count, 2)
    i = 0
    For Each Fl In coll
        i = i + 1
        fname = RelFName(Fl)
        out(i, 1) = fname
        out(i, 2) = Fl
    Next
    out = ArraySort(out, 1)
    GetListFile = out
End Function

Function GetFileName(ByVal nm As String) As String
    nm = Replace(nm, "/", "\")
    f = Split(nm, "\")
    fname$ = f(UBound(f))
    GetFileName = fname
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
End Function

Function GetNameForGOST(ByVal gost As String) As String
Dim tfunctime As Double
tfunctime = Timer
    If IsEmpty(name_gost) Then r = ReadMetall()
    If Len(swap_gost.Item(gost)) > 0 Then gost = swap_gost.Item(gost)
    For i = 1 To UBound(name_gost, 1)
        If name_gost(i, 1) = gost Then
            GetNameForGOST = name_gost(i, 2) & vbLf & gost
            Exit Function
        End If
    Next
    GetNameForGOST = gost
tfunctime = functime("GetNameForGOST", tfunctime)
End Function

Function GetNSubpos(ByVal subpos As String, ByVal type_spec As Long, ByVal floor_txt As String) As Long
    'Получаем количество сборок с именем = subpos
    Dim nSubPos As Long
    If subpos <> "-" Then
        If type_spec = 1 Then
            nSubPos = pos_data.Item(floor_txt).Item("qty").Item("all" & subpos)
            If nSubPos = 0 Then nSubPos = pos_data.Item(floor_txt).Item("qty").Item("-_" & subpos)
        Else
            nSubPos = pos_data.Item(floor_txt).Item("qty").Item("-_" & subpos)
        End If
        If nSubPos < 1 Then
            MsgBox ("Не определено кол-во сборок " & subpos & ", принято 1 шт.")
            r = LogWrite(subpos, vbNullString, "Не определено кол-во сборок")
            nSubPos = 1
        End If
    Else
        nSubPos = 1
    End If
    GetNSubpos = nSubPos
End Function

Function GetNumberConstr(ByVal unique_type_konstr As Variant, ByVal konstr As String) As Long
Dim tfunctime As Double
tfunctime = Timer
    For i = 1 To UBound(unique_type_konstr)
        If unique_type_konstr(i) = konstr Then
            GetNumberConstr = i
        End If
    Next i
tfunctime = functime("GetNumberConstr", tfunctime)
End Function

Function GetNumberStal(ByVal unique_stal As Variant, ByVal stal As String) As Long
Dim tfunctime As Double
tfunctime = Timer
    For i = 1 To UBound(unique_stal)
        If unique_stal(i) = stal Then
            GetNumberStal = i
        End If
    Next i
tfunctime = functime("GetNumberStal", tfunctime)
End Function

Function GetSheetOfBook(ByRef objCloseBook As Variant, ByVal sName As String) As Worksheet
    Set GetSheetOfBook = objCloseBook.Sheets(sName)
End Function

Function GetShortNameForGOST(ByVal gost As String) As String
Dim tfunctime As Double
tfunctime = Timer
    If IsEmpty(name_gost) Then r = ReadMetall()
    If Len(swap_gost.Item(gost)) > 0 Then gost = swap_gost.Item(gost)
    For i = 1 To UBound(name_gost, 1)
        If name_gost(i, 1) = gost Then
            GetShortNameForGOST = " " & name_gost(i, 3) & " "
            Exit Function
        End If
    Next
tfunctime = functime("GetShortNameForGOST", tfunctime)
End Function

Function GetWeightForDiametr(ByVal diametr As Long, ByVal klass As String) As Double
Dim tfunctime As Double
tfunctime = Timer
    If IsEmpty(reinforcement_specifications) Then r = ReadReinforce()
    klass = Replace(klass, "А", "A")
    klass = Replace(klass, "С", "C")
    For i = 1 To UBound(reinforcement_specifications, 1)
        diametr_r = reinforcement_specifications(i, col_diametr_spec)
        klass_r = reinforcement_specifications(i, col_klass_spec)
        If klass_r = klass And diametr_r = diametr Then
            GetWeightForDiametr = CDbl(reinforcement_specifications(i, col_weight_spec))
tfunctime = functime("GetWeightForDiametr", tfunctime)
            Exit Function
        End If
    Next
    r = LogWrite("Ошибка арматуры", vbNullString, "Отсутвует вес для " & diametr & " " & klass)
    GetWeightForDiametr = -1
End Function

Private Function ins_row(ByRef arr_out As Variant, ByRef arr_tmp As Variant, ByVal i As Long, ByVal n_col_sb As Long, ByRef n_row_ex As Long, ByVal nSubPos As Long) As Boolean
Dim tfunctime As Double
tfunctime = Timer
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
tfunctime = functime("ins_row", tfunctime)
End Function

Function LogNewSheet(ByVal log_sheet_name As String)
    ThisWorkbook.Worksheets.Add.Name = log_sheet_name
    Set log_sheet = wbk.Sheets(log_sheet_name)
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
    i = i + 1: log_sheet.Cells(1, i) = 0
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
Dim tfunctime As Double
tfunctime = Timer
    Dim out_log(): ReDim out_log(1, 15)
    i = 0
    i = i + 1: out_log(1, i) = Now
    i = i + 1: out_log(1, i) = Environ$("computername") & "-" & Environ$("username")
    i = i + 1: out_log(1, i) = sheet_name
    i = i + 1: out_log(1, i) = suffix
    i = i + 1: out_log(1, i) = rezult
    i = i + 1: out_log(1, i) = macro_version
    i = i + 1: out_log(1, i) = common_version
    i = i + 1: out_log(1, i) = UserForm2.form_ver.Caption
    If Not IsEmpty(this_sheet_option) Then
        i = i + 1: out_log(1, i) = k_zap_total
        i = i + 1: out_log(1, i) = this_sheet_option.Item("arm_pm")
        i = i + 1: out_log(1, i) = this_sheet_option.Item("pr_pm")
        i = i + 1: out_log(1, i) = this_sheet_option.Item("keep_pos")
        i = i + 1: out_log(1, i) = this_sheet_option.Item("qty_one_subpos")
        i = i + 1: out_log(1, i) = this_sheet_option.Item("show_subpos")
        i = i + 1: out_log(1, i) = this_sheet_option.Item("ignore_subpos")
    End If
    j = log_sheet.Cells(1, 16).Value2 + 1
    log_sheet.Cells(1, 16).Value2 = j
    log_sheet.Range(log_sheet.Cells(j, 1), log_sheet.Cells(j, 15)) = out_log
    log_sheet.Visible = False
tfunctime = functime("LogWrite", tfunctime)
End Function

Function DataReadAutoMat(ByVal nm As String, ByRef un_subpos As Variant) As Variant
    '-------------------------------------------------------
    'Описание файла ИК архикада с компонентами
    col_archimat_marka = 1 'Полный ID
    col_archimat_sub_pos = 2 'Свойство - принадлежит к элементу
    col_archimat_type_el = 3 'Тип элемента
    col_archimat_pos = 4 'Упрощённый ID
    col_archimat_naen = 5 'Имя штриховки
    col_archimat_volume = 6 'Объём слоя
    col_archimat_area = 7 'Площадь слоя
    col_archimat_thickness = 8 'Толщина слоя
    col_archimat_edizm = 9 'Расход, ед. изм
    col_archimat_density = 10 'Плотность
    col_archimat_floor = 11 'Имя собственного этажа
    max_archimat_acad = 11
    eps_mat = 0.000001
Dim tfunctime As Double
tfunctime = Timer
    If InStr(nm, "_") > 0 Then nm = Split(nm, "_")(0)
    nm = LCase(nm)
    out_data_raw = Empty
    del_no_subpos = False 'Удалять материалы, для которых не задана сборка
    'Выберем все файлы с окончанием "_мат"
    coll = GetListFile("*_мат.txt")
    If IsEmpty(coll) Then
        DataReadAutoMat = Empty
        Exit Function
    End If
    
    ' Сначала поищем по имени листа
    For i = 1 To UBound(coll, 1)
        snm = ThisWorkbook.path & "\import\" & coll(i, 1)
        If LCase(coll(i, 1)) = nm + "_мат" Then
            tdate = FileDateTime(coll(i, 2))
            short_fname = coll(i, 1)
            out_data_sheet = ReadTxt(snm + ".txt", 1, vbTab, vbNewLine, False)
            out_data_raw = ArrayCombine(out_data_raw, out_data_sheet)
        End If
    Next i
    ' Затем поищем по содержанию имени листа
    If IsEmpty(out_data_raw) Then
        del_no_subpos = True
        nm_ = Trim(nm)
        For i = 1 To UBound(coll, 1)
            snm = ThisWorkbook.path & "\import\" & coll(i, 1)
            If InStr(LCase(coll(i, 1)), nm_) > 0 Then
                tdate = FileDateTime(coll(i, 2))
                short_fname = coll(i, 1)
                out_data_sheet = ReadTxt(snm + ".txt", 1, vbTab, vbNewLine, False)
                out_data_raw = ArrayCombine(out_data_raw, out_data_sheet)
            End If
        Next i
    End If
    
    ' Теперь отрежем цифры в листе и повторим
    If IsEmpty(out_data_raw) Then
        For k = 0 To 9
            nm_ = Replace(nm_, CStr(k), "")
        Next k
        For i = 1 To UBound(coll, 1)
            snm = ThisWorkbook.path & "\import\" & coll(i, 1)
            If InStr(LCase(coll(i, 1)), nm_) > 0 Then
                tdate = FileDateTime(coll(i, 2))
                short_fname = coll(i, 1)
                out_data_sheet = ReadTxt(snm + ".txt", 1, vbTab, vbNewLine, False)
                out_data_raw = ArrayCombine(out_data_raw, out_data_sheet)
            End If
        Next i
    End If
    
    ' Ну и просто поищем по раздела - АР, КЖ, КМ
    If IsEmpty(out_data_raw) Then
        For i = 1 To UBound(coll, 1)
            snm = ThisWorkbook.path & "\import\" & coll(i, 1)
            If LCase(coll(i, 1)) = "кж_мат" Or LCase(coll(i, 1)) = "кр_мат" Or LCase(coll(i, 1)) = "ар_мат" Or LCase(coll(i, 1)) = "ас_мат" Then
                tdate = FileDateTime(coll(i, 2))
                short_fname = coll(i, 1)
                out_data_sheet = ReadTxt(snm + ".txt", 1, vbTab, vbNewLine, False)
                out_data_raw = ArrayCombine(out_data_raw, out_data_sheet)
            End If
        Next i
    End If

    If IsEmpty(out_data_raw) Then
        DataReadAutoMat = Empty
        Exit Function
    End If
    
    'Поставим все позиуии использованных сборок в нижний регистр
    If del_no_subpos Then
        For i = LBound(un_subpos) To UBound(un_subpos)
            un_subpos(i) = Trim(LCase(un_subpos(i)))
        Next i
    End If
    n_row = UBound(out_data_raw, 1)
    Dim out_data: ReDim out_data(n_row, max_col)
    n_row_out = 0
    'Для строк без толщины попробьуем её вычислить
    'просмотрим остальные материалы и посмотрим толщину слоёв с таким же названием
    Set thickness_mat = CreateObject("Scripting.Dictionary")
    For i = 1 To n_row
        tthickness = out_data_raw(i, col_archimat_thickness)
        tvolume = out_data_raw(i, col_archimat_volume)
        If IsNumeric(tvolume) And Not IsNumeric(tthickness) Then
            tnaen = out_data_raw(i, col_archimat_naen)
            For j = 1 To n_row
                tthickness_j = out_data_raw(j, col_archimat_thickness)
                tnaen_j = out_data_raw(j, col_archimat_naen)
                If out_data_raw(j, col_archimat_naen) = tnaen And IsNumeric(tthickness_j) Then
                    If Not thickness_mat.Exists(tnaen) Then
                        thickness_mat.Item(tnaen) = tthickness_j
                    Else
                        thickness_mat.Item(tnaen) = Application.WorksheetFunction.Max(tthickness_j, thickness_mat.Item(tnaen))
                    End If
                End If
            Next j
        End If
    Next i
    For i = 1 To n_row
        marka = Trim(CStr(out_data_raw(i, col_archimat_marka)))
        sub_pos = Trim(CStr(out_data_raw(i, col_archimat_sub_pos)))
        pos = Trim(CStr(out_data_raw(i, col_archimat_pos)))
        If pos = sub_pos Then pos = ""
        If pos = "---" Then pos = ""
        If sub_pos = "---" Then sub_pos = ""
        If marka = "---" Then marka = ""
        tfloor = CStr(out_data_raw(i, col_archimat_floor))
        tarea = ConvTxt2Num(out_data_raw(i, col_archimat_area))
        tthickness = ConvTxt2Num(out_data_raw(i, col_archimat_thickness))
        thickness_in_param = 0
        tvolume = ConvTxt2Num(out_data_raw(i, col_archimat_volume))
        tnaen = Trim(CStr(out_data_raw(i, col_archimat_naen)))
        tedizm = Trim(CStr(out_data_raw(i, col_archimat_edizm)))
        flag_add = 1
        kzap_mat = 1
        trate = 0
        trate_edizm = vbNullString
        trate_edizm_raw = vbNullString
        trate_edizm_fist = vbNullString
        tprim = vbNullString
        edizm = "куб.м."
        If InStr(tedizm, "=") > 0 And flag_add Then
            'Смотрим - что за параметры прилетели
            arr = Split(tedizm, ";")
            For Each tel In arr
                If Len(tel) > 1 Then
                    arr2 = Split(tel, "=")
                    tvalue = Trim$(arr2(1))
                    name_param = Trim$(LCase$(arr2(0)))
                    If InStr(name_param, "прим") > 0 Then tprim = tvalue
                    tvalue = Replace(LCase$(tvalue), " ", vbNullString)
                    name_param = Replace(LCase$(name_param), " ", vbNullString)
                    name_param = Replace(name_param, ",", vbNullString)
                    name_param = Replace(name_param, ".", vbNullString)
                    name_param = Replace(name_param, "\", vbNullString)
                    name_param = Replace(name_param, "/", vbNullString)
                    If InStr(name_param, "кзап") > 0 Then
                        kzap_mat = ConvTxt2Num(tvalue)
                        If IsNumeric(kzap_mat) Then
                            If kzap_mat < 1 Then kzap_mat = 1
                            If kzap_mat >= 2 Then kzap_mat = 1.5
                        Else
                            kzap_mat = 1
                        End If
                    End If
                    If ignore_zap_material Then kzap_mat = 1
                    If InStr(name_param, "едизм") > 0 Then edizm = tvalue
                    If InStr(name_param, "расход") > 0 Then
                        tvalue = Replace(tvalue, ":", "\")
                        tvalue = Replace(tvalue, "/", "\")
                        If InStr(tvalue, "\") > 0 Then
                            trate_edizm_raw = tvalue
                            arr3 = Split(tvalue, "\")
                            trate_edizm_fist = arr3(0)
                            trate_edizm = Replace(arr3(1), ",", vbNullString)
                            trate_edizm = Replace(trate_edizm, ".", vbNullString)
                            trate = Replace(arr3(0), "л", vbNullString)
                            trate = Replace(trate, "кг", vbNullString)
                            trate = Replace(trate, "шт", vbNullString)
                            trate = Replace(trate, "ед", vbNullString)
                            trate = Replace(trate, "пм", vbNullString)
                            trate = ConvTxt2Num(trate)
                            If Not IsNumeric(trate) Then
                                trate = 0
                                trate_edizm = vbNullString
                            End If
                        End If
                    End If
                    If InStr(name_param, "t") > 0 Or InStr(name_param, "толщ") > 0 Then
                        'Толщина нам нужна в мм. По умолчанию, если ничего не задано, предполагаем что это мм.
                        edizm_thickness_in_param = 1
                        If InStr(tvalue, "мм") > 0 Then
                            edizm_thickness_in_param = 1
                        Else
                            If InStr(tvalue, "см") > 0 Then
                                edizm_thickness_in_param = 10
                            Else
                                If InStr(tvalue, "м") > 0 Then
                                    edizm_thickness_in_param = 1000
                                Else
                                    ' Если ничего не подошло
                                    edizm_thickness_in_param = 1
                                End If
                            End If
                        End If
                        tvalue = Replace(tvalue, "м", vbNullString)
                        tvalue = Replace(tvalue, "c", vbNullString)
                        tvalue = ConvTxt2Num(tvalue)
                        If IsNumeric(tvalue) Then thickness_in_param = tvalue * edizm_thickness_in_param
                    End If
                End If
            Next
        End If
        ' Если тощина не определена - посмотрим что прописано в параметрах материала
        If Not IsNumeric(tthickness) And thickness_in_param > 0 Then tthickness = thickness_in_param
        ' Если тощина не определена - попробуем поискать среди встречавшихся ранее слоёв с тем же наименованием
        If Not IsNumeric(tthickness) And thickness_mat.Exists(tnaen) Then tthickness = thickness_mat.Item(tnaen)
        'Очистим от скверны
        edizm_purge = Replace(edizm, " ", vbNullString)
        edizm_purge = Replace(edizm_purge, ",", vbNullString)
        edizm_purge = Replace(edizm_purge, ".", vbNullString)
        edizm_purge = Replace(edizm_purge, "\", vbNullString)
        edizm_purge = Replace(edizm_purge, "/", vbNullString)
        edizm_purge = Replace(edizm_purge, Chr$(34), vbNullString)
        If IsNumeric(tarea) Then
            If tarea < eps_mat Then tarea = vbNullString
        End If
        If IsNumeric(tvolume) Then
            If tvolume < eps_mat Then tvolume = vbNullString
        End If
        If IsNumeric(tthickness) Then
            If tthickness < eps_mat Then tthickness = vbNullString
        End If
        If Not IsNumeric(tarea) And Not IsNumeric(tvolume) And Not IsNumeric(tthickness) Then
            flag_add = 0
        Else
            If IsNumeric(tarea) And IsNumeric(tvolume) And IsNumeric(tthickness) Then
                flag_add = 1
            Else
                If IsNumeric(tvolume) And IsNumeric(tarea) And Not IsNumeric(tthickness) Then tthickness = (tvolume / tarea) * 1000
                If Not IsNumeric(tvolume) And IsNumeric(tarea) And IsNumeric(tthickness) Then tvolume = tarea * (tthickness / 1000)
                If IsNumeric(tvolume) And Not IsNumeric(tarea) And IsNumeric(tthickness) Then tarea = tvolume / (tthickness / 1000)
                If IsNumeric(tarea) And IsNumeric(tvolume) And IsNumeric(tthickness) Then
                    flag_add = 1
                Else
                    'А так ли нам важно всё это знать?
                    Select Case edizm_purge
                        Case "кубм", "мкуб", "м3"
                            'А мы выдадим в кв.м.
                            If Not IsNumeric(tvolume) And IsNumeric(tarea) Then
                                edizm_purge = "квм"
                                flag_add = 1
                            End If
                        Case "квм", "мкв", "м2"
                            'А мы выдадим в куб.м.
                            If Not IsNumeric(tarea) And IsNumeric(tvolume) Then
                                edizm_purge = "кубм"
                                flag_add = 1
                            End If
                    End Select
                    If Not IsNumeric(tvolume) Then tvolume = 0.001
                    If Not IsNumeric(tarea) Then tarea = 0.001
                    If Not IsNumeric(tthickness) Then tthickness = 1
                End If
            End If
        End If
        If flag_add Then
            tthickness = tthickness / 1000
            tnaen = Replace(tnaen, "  ", vbNullString)
            tnaen = Replace(tnaen, "\n", " ")
            tnaen = Replace(tnaen, "/n", " ")
            tnaen = Replace(tnaen, "( ", "(")
            tnaen = Replace(tnaen, " )", ")")
            tdensity = CStr(out_data_raw(i, col_archimat_density))
            If InStr(tdensity, "кг") > 0 Then
                tdensity = ConvTxt2Num(Split(tdensity, "кг")(0))
                If Not IsNumeric(tdensity) Then tdensity = -1
            Else
                tdensity = -1
            End If
            'Выделяем ГОСТ, ТУ, СТО для графы Обозначение
            obozn = " "
            p_start = 0
            For Each ttype_norm In Array("(ТУ", "(СТО", "(ГОСТ", "(Серия")
                p_start = InStr(1, tnaen, ttype_norm, vbBinaryCompare)
                If p_start > 0 Then
                    type_norm = ttype_norm
                    Exit For
                End If
            Next
            If p_start > 0 Then
                arr = Split(tnaen, type_norm)
                obozn = Trim$(arr(1))
                obozn = Left$(obozn, InStr(obozn, ")"))
                obozn = type_norm + " " + Trim$(obozn)
                naen = Replace(tnaen, "по " + obozn, vbNullString)
                naen = Replace(naen, "(" + obozn + ")", vbNullString)
                naen = Replace(naen, obozn, vbNullString)
                naen = Replace(naen, "  ", " ")
                naen = Replace(naen, " ,", ",")
                obozn = Replace(obozn, "(", vbNullString)
                obozn = Replace(obozn, ")", vbNullString)
                obozn = Replace(obozn, "по", vbNullString)
                obozn = Trim$(obozn)
                naen = Trim$(naen)
            Else
                naen = Trim$(tnaen)
            End If
            If Len(swap_gost.Item(obozn)) > 0 Then obozn = swap_gost.Item(obozn)
            If InStr(naen, "по уклону") > 0 And InStr(naen, " от") > 0 And InStr(naen, " до") > 0 Then
                naen = Replace(naen, " до", " до " + CStr(Int(tthickness * 1000)) + "мм.")
            End If
            array_add_th = Array("ТЕХНОРУФ", "кладка", "азобетон", "CARBON PROF", "ТЕХНОВЕНТ")
            flag_add_th = 1
            For jj = 1 To UBound(array_add_th)
                If InStr(naen, array_add_th(jj)) > 0 And flag_add_th And InStr(naen, "t=") = 0 Then
                    naen = naen + ", t=" + ConvNum2Txt(tthickness * 1000) + "мм."
                    flag_add_th = 0
                End If
            Next jj
            'Если расход не задан - просто выберем по единицам измерения нужный параметр
            tqty = tvolume
            If trate = 0 Then
                Select Case edizm_purge
                    Case "л", "кг", "пм", "шт", "ед"
                        'Если не задан расход - то и выдать ничего не можем. Выдадим в кубах 8)
                        edizm = "куб.м."
                        tqty = tvolume
                    Case "кубм", "мкуб", "м3"
                        tqty = tvolume
                    Case "квм", "мкв", "м2"
                        tqty = tarea
                        If tthickness > 0.005 And InStr(naen, "t=") = 0 Then naen = naen + ", t=" + ConvNum2Txt(tthickness * 1000) + "мм."
                End Select
                trate = 1
            Else
                If InStr(trate_edizm_fist, edizm_purge) = 0 Then
                    'Единицы в расходе и в выдаче не совпадают. Повод задуматься
                    trate = 1
                End If
                Select Case trate_edizm
                    Case "кубм", "мкуб", "м3"
                        tqty = tvolume
                    Case "квм", "мкв", "м2"
                        tqty = tarea
                End Select
                naen = naen + ", расход " + trate_edizm_raw
            End If
            Select Case edizm_purge
                Case "л"
                    edizm = "л."
                    qty = trate * tqty * kzap_mat
                    Weight = 0
                Case "кг"
                    edizm = "кг."
                    qty = trate * tqty * kzap_mat
                    Weight = 0
                Case "пм"
                    edizm = "п.м."
                    qty = trate * tqty * kzap_mat
                    Weight = 0
                Case "кубм", "мкуб", "м3"
                    edizm = "куб.м."
                    qty = tqty * kzap_mat
                    Weight = tdensity
                Case "квм", "мкв", "м2"
                    edizm = "кв.м."
                    qty = tqty * kzap_mat
                    Weight = tdensity * tthickness
                Case "шт", "ед"
                    edizm = "шт."
                    qty = Round_w(trate * tqty * kzap_mat, 0)
                    Weight = tdensity * tvolume
            End Select
            If Len(tprim) > 0 Then naen = naen + ", " + tprim
        Else
            flag_add = 0
        End If
        'Если прочитали из файла с именем, отличным от имяни листа - то берём только материалы со сборок, встречающихся на листе.
        If del_no_subpos Then
            'Берём только элементы, для которых задана сборка
            If Len(sub_pos) > 0 Then
                'Выделяем среди них только элементы, расположенные в использованных на листе сборках
                is_subpose_use = ArrayHasElement(un_subpos, LCase(sub_pos))
                If Not is_subpose_use Then flag_add = 0
            Else
                flag_add = 0
            End If
        End If
        If flag_add = 1 Then
            n_row_out = n_row_out + 1
            out_data(n_row_out, col_marka) = marka
            out_data(n_row_out, col_sub_pos) = sub_pos
            out_data(n_row_out, col_type_el) = t_mat
            out_data(n_row_out, col_pos) = pos
            out_data(n_row_out, col_qty) = qty
            out_data(n_row_out, col_nfloor) = 0
            out_data(n_row_out, col_floor) = tfloor
            out_data(n_row_out, col_m_obozn) = obozn
            out_data(n_row_out, col_m_naen) = naen
            out_data(n_row_out, col_m_weight) = Weight
            If Weight > 0 Then
                hh = 1
            End If
            out_data(n_row_out, col_m_edizm) = edizm
        End If
    Next i
    If n_row_out <> n_row Then out_data = ArrayRedim(out_data, n_row_out)
r = functime("DataReadAutoMat", tfunctime)
    DataReadAutoMat = out_data
End Function

Function DataReadAutoArm(ByVal subpos As String) As Variant
    '-------------------------------------------------------
    'Описание файла c извлечением данных из автокада
    col_acad_pos = 1
    col_acad_qty_all = 2  'Кол-во блоков
    col_acad_qty = 3 'Кол-во стержней в блоке
    col_acad_diametr = 4
    col_acad_length = 5
    col_acad_layer = 6
    col_acad_klass = 7
    col_acad_fon = 8
    col_acad_mp = 9
    col_acad_gnut = 10
    max_col_acad = 6
    
    out_data_raw = Empty
    path = ThisWorkbook.path & "\import"
    Set coll = FilenamesCollection(path, subpos + ".xls")
    For Each snm In coll
        If Not IsEmpty(snm) Then
            tdate = FileDateTime(snm)
            fname = snm
            full_fname = Split(fname, "\")
            n_path = UBound(full_fname)
            short_fname = full_fname(n_path)
        End If
        If Split(short_fname, ".")(0) = subpos Then
            Set spec_book = GetObject(snm)
            For Each nm In GetListOfSheet(spec_book)
                Set spec_sheet = spec_book.Sheets(nm)
                n = SheetGetSize(spec_sheet)
                out_data_sheet = spec_sheet.Range(spec_sheet.Cells(2, 1), spec_sheet.Cells(n(1), n(2)))
                out_data_raw = ArrayCombine(out_data_raw, out_data_sheet)
            Next
            spec_book.Close True
        End If
    Next
    If IsEmpty(out_data_raw) Then
        DataReadAutoArm = Empty
        Exit Function
    End If
    n_row = UBound(out_data_raw, 1)
    Dim out_data: ReDim out_data(n_row, col_man_klass)
    Dim out_data_up: ReDim out_data_up(1, col_man_klass)
    For j = 1 To UBound(out_data_up, 2)
        out_data_up(1, j) = Empty
    Next j
    short_fname = " " + full_fname(n_path) + ".." + full_fname(n_path - 3) + ".." + full_fname(n_path - 4)
    out_data_up(1, col_man_subpos) = subpos
    out_data_up(1, col_man_pos) = "!!!"
    out_data_up(1, col_man_obozn) = "'" + str(tdate)
    out_data_up(1, col_man_naen) = "АВТОКАД_" + short_fname
    n_row_out = 0
    For i = 1 To n_row
        pos = out_data_raw(i, col_acad_pos)
        qty_all = ConvTxt2Num(out_data_raw(i, col_acad_qty_all))
        qty_in_one = ConvTxt2Num(out_data_raw(i, col_acad_qty))
        diametr = ConvTxt2Num(out_data_raw(i, col_acad_diametr))
        Length = ConvTxt2Num(out_data_raw(i, col_acad_length))
        If IsNumeric(qty_all) And IsNumeric(qty_in_one) And IsNumeric(diametr) And IsNumeric(Length) Then
            For j = 1 To UBound(out_data, 2)
                out_data(i, j) = Empty
            Next j
            n_row_out = n_row_out + 1
            qty = qty_all * qty_in_one
            out_data(n_row_out, col_man_subpos) = subpos
            out_data(n_row_out, col_man_naen) = "АВТОКАД_" + Trim$(out_data_raw(i, col_acad_layer))
            out_data(n_row_out, col_man_qty) = qty
            out_data(n_row_out, col_man_pos) = Trim$(out_data_raw(i, col_acad_pos))
            out_data(n_row_out, col_man_diametr) = diametr
            If Length > lenght_ed_arm Then
                out_data(n_row_out, col_man_length) = Length
                out_data(n_row_out, col_man_prim) = "п.м."
            Else
                out_data(n_row_out, col_man_length) = Length
            End If
            out_data(n_row_out, col_man_klass) = "A500C"
        End If
    Next i
    If n_row_out <> n_row Then out_data = ArrayRedim(out_data, n_row_out)
    out_data = ArrayCombine(out_data_up, out_data)
    DataReadAutoArm = out_data
End Function

Function DataReadAutoArm_2way(ByVal subpos As String) As Variant
    '-------------------------------------------------------
    'Описание файла c извлечением данных из автокада
Dim tfunctime As Double
tfunctime = Timer
    col_acad_handle = 1
    col_acad_blockname = 2
    out_data_raw = Empty
    coll = GetListFile(subpos + "_autocad.txt")
    If IsEmpty(coll) Then
        DataReadAutoArm_2way = Empty
        Exit Function
    End If
    For i = 1 To UBound(coll, 1)
        snm = ThisWorkbook.path & "\import\" & coll(i, 1)
        If coll(i, 1) = subpos + "_autocad" Then
            tdate = FileDateTime(coll(i, 2))
            short_fname = coll(i, 1)
            out_data_sheet = ReadTxt(snm + ".txt", 1, vbTab, vbNewLine, False)
            out_data_raw = ArrayCombine(out_data_raw, out_data_sheet)
        End If
    Next i
    If IsEmpty(out_data_raw) Then
        DataReadAutoArm_2way = Empty
        Exit Function
    End If
    n_row = UBound(out_data_raw, 1)
    max_col_acad = UBound(out_data_raw, 2)
    'Поиск подходящих столбцов по имени
    col_acad_qty = -1 'Кол-во стержней в блоке
    col_acad_diametr = -1
    col_acad_length = -1
    col_acad_pos = -1
    col_acad_gnut = -1
    col_acad_mp = -1
    col_acad_klass = -1
    For i = 1 To max_col_acad
        If Trim$(UCase$(out_data_raw(1, i))) = "ПОЗИЦИЯ" Then col_acad_pos = i
        If Trim$(UCase$(out_data_raw(1, i))) = "КОЛИЧЕСТВО" Then col_acad_qty = i
        If Trim$(UCase$(out_data_raw(1, i))) = "ДИАМЕТР" Then col_acad_diametr = i
        If Trim$(UCase$(out_data_raw(1, i))) = "ДЛИНА_СТЕРЖНЯ" Then col_acad_length = i
        'Добавка
        If Trim$(UCase$(out_data_raw(1, i))) = "ГНУТИК" Then col_acad_gnut = i
        If Trim$(UCase$(out_data_raw(1, i))) = "ПОГОНАЖ" Then col_acad_mp = i
        If Trim$(UCase$(out_data_raw(1, i))) = "КЛАСС" Then col_acad_klass = i
    Next i
    If col_acad_pos < 0 Or col_acad_qty < 0 Or col_acad_diametr < 0 Or col_acad_length < 0 Then
        DataReadAutoArm_2way = Empty
        Exit Function
    End If
    n_add = 0
    For i = 2 To n_row
        If InStr(out_data_raw(i, col_acad_pos), "@") Then
            n_add = n_add + UBound(Split(out_data_raw(i, col_acad_pos), "@"))
        End If
    Next i
    Dim out_data: ReDim out_data(n_row + n_add, col_man_klass)
    Dim out_data_up: ReDim out_data_up(1, col_man_klass)
    For j = 1 To UBound(out_data_up, 2)
        out_data_up(1, j) = Empty
    Next j
    out_data_up(1, col_man_subpos) = subpos
    out_data_up(1, col_man_pos) = "!!!"
    out_data_up(1, col_man_obozn) = "'" + str(tdate)
    out_data_up(1, col_man_naen) = "АВТОКАД_" + short_fname

    def_klass = "A500C"
    Set summ_arm = CreateObject("Scripting.Dictionary")
    summ_arm.comparemode = 1
    Set data_arm = CreateObject("Scripting.Dictionary")
    data_arm.comparemode = 1
    n_row_out = 0
    For i = 2 To n_row
        Handle = CStr(out_data_raw(i, col_acad_handle))
        blockname = CStr(out_data_raw(i, col_acad_blockname))
        If InStr(out_data_raw(i, col_acad_pos), "@") Then
            arr_pos = Split(out_data_raw(i, col_acad_pos), "@")
            arr_qty = Split(out_data_raw(i, col_acad_qty), "@")
            arr_diametr = Split(out_data_raw(i, col_acad_diametr), "@")
            arr_Length = Split(out_data_raw(i, col_acad_length), "@")
            If col_acad_gnut > 0 Then
                arr_gnut = Split(out_data_raw(i, col_acad_gnut), "@")
            End If
            If col_acad_mp > 0 Then
                arr_mp = Split(out_data_raw(i, col_acad_mp), "@")
            End If
            If col_acad_klass > 0 Then
                arr_klass = Split(out_data_raw(i, col_acad_klass), "@")
            End If
            n_pos = UBound(arr_pos)
            flag_multipos = True
        Else
            n_pos = 0
            flag_multipos = False
        End If
        For k = 0 To n_pos
            If flag_multipos Then
                pos = arr_pos(k)
                qty = arr_qty(k)
                diametr = arr_diametr(k)
                Length = arr_Length(k)
                If col_acad_gnut > 0 Then
                    gnut = arr_gnut(k)
                Else
                    gnut = "0"
                End If
                If col_acad_mp > 0 Then
                    mp = arr_mp(k)
                Else
                    mp = "0"
                End If
                If col_acad_klass > 0 Then
                    klass = arr_klass(k)
                Else
                    klass = def_klass
                End If
                Handle = CStr(out_data_raw(i, col_acad_handle)) + "_" + CStr(k + 1)
                blockname = CStr(out_data_raw(i, col_acad_blockname)) + "_" + CStr(k + 1)
            Else
                pos = out_data_raw(i, col_acad_pos)
                qty = out_data_raw(i, col_acad_qty)
                diametr = out_data_raw(i, col_acad_diametr)
                Length = out_data_raw(i, col_acad_length)
                If col_acad_gnut > 0 Then
                    gnut = out_data_raw(i, col_acad_gnut)
                Else
                    gnut = "0"
                End If
                If col_acad_mp > 0 Then
                    mp = out_data_raw(i, col_acad_mp)
                Else
                    mp = "0"
                End If
                If col_acad_klass > 0 Then
                    klass = out_data_raw(i, col_acad_klass)
                Else
                    klass = def_klass
                End If
            End If
            pos = Trim$(pos)
            qty = ConvTxt2Num(qty)
            diametr = ConvTxt2Num(diametr)
            Length = ConvTxt2Num(Length)
            If Trim$(gnut) = "1" Then
                gnut = True
            Else
                gnut = False
            End If
            If Trim$(mp) = "1" Then
                mp = True
            Else
                mp = False
            End If
            klass = Trim$(klass)
            If klass = "<>" Then klass = def_klass
            klass = Replace(klass, "А", "A")
            If InStr(klass, "240") Then klass = "A-I(A240)"
            If InStr(klass, "400") Then klass = "A-III(A400)"
            flag_wrtite = True
            If pos = "<>" Or InStr(pos, "!!") Then flag_wrtite = False
            If Left$(pos, 1) = "Г" Then
                tpos = pos + "@" + CStr(diametr) + "@" + klass
                If Not summ_arm.Exists(tpos) Then
                    summ_arm.Item(tpos) = 0
                    data_arm.Item(tpos) = Array(subpos, Replace(pos, "Г", vbNullString), "АВТОКАД_" + Handle + "_" + blockname + "Г", diametr, klass, mp, gnut)
                End If
                summ_arm.Item(tpos) = summ_arm.Item(tpos) + qty * Length
                flag_wrtite = False
            End If
            If IsNumeric(qty) And IsNumeric(diametr) And IsNumeric(Length) And flag_wrtite Then
                n_row_out = n_row_out + 1
                For j = 1 To UBound(out_data, 2)
                    out_data(n_row_out, j) = Empty
                Next j
                out_data(n_row_out, col_man_subpos) = subpos
                out_data(n_row_out, col_man_naen) = "АВТОКАД_" + Handle + "_" + blockname
                out_data(n_row_out, col_man_qty) = qty
                out_data(n_row_out, col_man_pos) = pos
                out_data(n_row_out, col_man_diametr) = diametr
                If Length > lenght_ed_arm Or mp Then
                    out_data(n_row_out, col_man_prim) = "п.м."
                End If
                out_data(n_row_out, col_man_length) = Length
                If gnut Then out_data(n_row_out, col_man_prim) = "*"
                out_data(n_row_out, col_man_klass) = klass
            End If
        Next k
    Next i
    If summ_arm.Count > 0 Then
        For Each tpos In summ_arm.keys
            subpos = data_arm.Item(tpos)(1)
            pos = data_arm.Item(tpos)(2)
            naen = data_arm.Item(tpos)(3)
            diametr = data_arm.Item(tpos)(4)
            klass = data_arm.Item(tpos)(5)
            mp = data_arm.Item(tpos)(6)
            gnut = data_arm.Item(tpos)(7)
            Length = summ_arm.Item(tpos)
            qty = 1
            If IsNumeric(qty) And IsNumeric(diametr) And IsNumeric(Length) Then
                n_row_out = n_row_out + 1
                For j = 1 To UBound(out_data, 2)
                    out_data(n_row_out, j) = Empty
                Next j
                out_data(n_row_out, col_man_subpos) = subpos
                out_data(n_row_out, col_man_naen) = naen
                out_data(n_row_out, col_man_qty) = qty
                out_data(n_row_out, col_man_pos) = pos
                out_data(n_row_out, col_man_diametr) = diametr
                If Length > lenght_ed_arm Or mp Then
                    out_data(n_row_out, col_man_prim) = "п.м."
                End If
                out_data(n_row_out, col_man_length) = Length
                If gnut Then out_data(n_row_out, col_man_prim) = "*"
                out_data(n_row_out, col_man_klass) = klass
            End If
        Next
    End If
    If n_row_out <> n_row Then out_data = ArrayRedim(out_data, n_row_out)
    out_data = ArrayCombine(out_data_up, out_data)
    DataReadAutoArm_2way = out_data
r = functime("DataReadAutoArm_2way", tfunctime)
End Function


Function ManualAddAuto(ByVal nm As String) As Boolean
    If nm = "Сводная_спец" Then
        ManualAddAuto = False
        Exit Function
    End If
    Set arm_data = CreateObject("Scripting.Dictionary")
    If Not SheetExist(nm) Then
        ManualAddAuto = False
        Exit Function
    End If
    Set data_out = wbk.Sheets(nm)
    n_row = SheetGetSize(data_out)(1)
    col = max_col_man
    spec = data_out.Range(data_out.Cells(1, 1), data_out.Cells(n_row, max_col_man))
    subpos_arr = ArrayUniqValColumn(spec, col_man_subpos)
    del_row = vbNullString
    For Each subpos In subpos_arr
        If Len(Trim$(subpos)) > 0 And InStr(subpos, "Марка") = 0 Then
            out_data = DataReadAutoArm_2way(subpos) 'Ищем файлы с извлечением данных и сводим их в массив
            If Not IsEmpty(out_data) Then
                'Если позиции в извелчении не заданы - поищем их в существующем массиве
                block_data = ArraySelectParam_2(spec, "АВТОКАД_?", col_man_naen)
                If Not IsEmpty(block_data) Then
                    For i = 1 To UBound(out_data, 1)
                        If out_data(i, col_man_pos) = Empty Or Len(out_data(i, col_man_pos)) = 0 Then
                            old_data = ArraySelectParam_2(block_data, out_data(i, col_man_naen), col_man_naen)
                            If Not IsEmpty(old_data) Then out_data(i, col_man_pos) = old_data(1, col_man_pos)
                        End If
                    Next i
                End If
                'Проходим по таблице удаляем строки со старым извлечением данных
                For i = 1 To n_row
                    If Not IsError(spec(i, col_man_subpos)) And Not IsError(spec(i, col_man_naen)) Then
                        If spec(i, col_man_subpos) = subpos And InStr(spec(i, col_man_naen), "ВТОКАД_") > 0 Then data_out.Rows(i).ClearContents
                        If InStr(spec(i, col_man_naen), "АВТОКАД_Извлечение данных") > 0 Then data_out.Rows(i).ClearContents
                    End If
                Next i
                arm_data.Item(subpos) = out_data
            End If
        End If
    Next
    If arm_data.Count = 0 Then
        ManualAddAuto = False
        Exit Function
    End If
    Set data_out = wbk.Sheets(nm)
    n_row = SheetGetSize(data_out)(1)
    n_row_end = n_row + 4
    data_out.Cells(n_row_end, col_man_subpos) = "!!!"
    data_out.Cells(n_row_end, col_man_pos) = "!!!"
    data_out.Cells(n_row_end, col_man_obozn) = "НИЖЕ ЭТИХ СТРОК НИЧЕГО ВРУЧНУЮ НЕ ВВОДИТЬ"
    data_out.Cells(n_row_end, col_man_naen) = "АВТОКАД_Извлечение данных"
    n_row_end = n_row_end + 2
    For Each subpos In arm_data.keys()
        out_data = arm_data.Item(subpos)
        r = ManualPasteIzd2Sheet(out_data, n_row_end, subpos, nm)
        n_row_end = n_row_end + UBound(out_data, 1)
    Next
    n_row_end = n_row_end + 1
    data_out.Cells(n_row_end, col_man_subpos) = "!!!"
    data_out.Cells(n_row_end, col_man_pos) = "!!!"
    data_out.Cells(n_row_end, col_man_obozn) = "НИЖЕ ЭТИХ СТРОК НИЧЕГО ВРУЧНУЮ НЕ ВВОДИТЬ"
    data_out.Cells(n_row_end, col_man_naen) = "АВТОКАД_Извлечение данных"
    ManualAddAuto = True
End Function

Function SheetAddTxt() As Boolean
    coll = GetListFile(nm + ".txt")
    If IsEmpty(coll) Then
        SheetAddTxt = False
        Exit Function
    End If
    import_txt_arr = ArraySelectParam_2(coll, "?_сист?", 1)
    If IsEmpty(import_txt_arr) Then
        SheetAddTxt = False
        Exit Function
    End If
    import_sheet_arr = ArraySelectParam_2(GetListOfSheet(wbk), "?из архикада?")
    Set import_sheet = CreateObject("Scripting.Dictionary")
        If Not IsEmpty(import_sheet_arr) Then
        For Each nm In import_sheet_arr
            Set data_out = wbk.Sheets(nm)
            n_col = SheetGetSize(data_out)(2)
            header_sheet = data_out.Range(data_out.Cells(1, 1), data_out.Cells(2, n_col))
            header_sheet = Join(ArrayRow(header_sheet, 1)) + Join(ArrayRow(header_sheet, 2))
            header_sheet = Replace(header_sheet, " ", vbNullString)
            header_sheet = Replace(header_sheet, "_", vbNullString)
            header_sheet = LCase$(header_sheet)
            If InStr(header_sheet, "@@") > 0 Then
                header_sheet = Split(header_sheet, "@@")(2)
            End If
            If Len(header_sheet) > 0 Then import_sheet.Item(header_sheet) = nm
            Set data_out = Nothing
        Next
    End If
    
    For i = 1 To UBound(import_txt_arr, 1)
        short_fname = import_txt_arr(i, 1)
        tdate_txt = FileDateTime(import_txt_arr(i, 2))
        data_txt = ReadFile(import_txt_arr(i, 1) + ".txt", 1, vbTab, vbNewLine)
        If Not IsEmpty(data_txt) Then
            header_txt = Join(ArrayRow(data_txt, 1)) + Join(ArrayRow(data_txt, 2))
            header_txt = Replace(header_txt, " ", vbNullString)
            header_txt = Replace(header_txt, "_", vbNullString)
            header_txt = LCase$(header_txt)
            If import_sheet.Exists(header_txt) Then
                For Each conn In wbk.Connections
                    connName = Replace(conn.Name, " ", vbNullString)
                    connName = Replace(connName, "_", vbNullString)
                    connName = Replace(connName, "excel", vbNullString)
                    connName = LCase$(connName)
                    If InStr(header_txt, connName) > 0 And Len(connName) > 0 Then
                        conn.Delete
                    End If
                Next conn
                sheet_name = import_sheet.Item(header_txt)
            Else
                sheet_name = "из архикада"
                head_txt = Trim$(Join(ArrayRow(data_txt, 1)))
                If InStr(head_txt, " ") > 0 Then
                    n = Split(head_txt, " ")
                    For nn = 1 To UBound(n)
                        sheet_name = sheet_name + " " + n(nn)
                    Next nn
                Else
                    sheet_name = sheet_name + " " + head_txt
                End If
                sheet_name = Replace(sheet_name, "ЖБ", vbNullString)
                sheet_name = Replace(sheet_name, "  ", " ")
                sheet_name = Trim$(Left$(sheet_name, 31))
            End If
            data_txt(1, 1) = CStr(tdate_txt) + " @@ " + CStr(import_txt_arr(i, 2)) + " @@ "
            n_row = UBound(data_txt, 1)
            n_col = UBound(data_txt, 2)
            For k = 1 To n_col - 1
                p1 = (IsEmpty(data_txt(2, k)) And IsEmpty(data_txt(2, k + 1)))
                p2 = (Len(Trim$(data_txt(2, k))) = 0 And Len(Trim$(data_txt(2, k + 1))) = 0)
                If (p1 Or p2) And n_col = UBound(data_txt, 2) Then n_col = k - 1
            Next k
            For j = 3 To n_row
                For k = 1 To n_col
                    If InStr(data_txt(j, k), ".") > 0 Then
                        vval = ConvTxt2Num(data_txt(j, k))
                        If IsNumeric(vval) Then
                            vval = ConvNum2Txt(vval)
                            vval = Replace(vval, ".", ",")
                            data_txt(j, k) = vval
                        End If
                    End If
                Next k
            Next j
            add_flag = False
            If SheetExist(sheet_name) Then
                Set Sh_old = wbk.Sheets(sheet_name)
                size_sheet = SheetGetSize(Sh_old)
                n_row_sh = size_sheet(1)
                n_col_sh = size_sheet(2)
                n_row_add = Application.WorksheetFunction.Max(n_row_sh, n_row)
                If n_col_sh > n_col Then
                    add_col = Sh_old.Range(Sh_old.Cells(1, n_col + 1), Sh_old.Cells(n_row_add, n_col_sh))
                    For j = 1 To UBound(add_col, 1)
                        For k = 1 To UBound(add_col, 1)
                            If Sh_old.Cells(j, n_col + k).HasFormula Then add_col(j, k) = Sh_old.Cells(j, n_col + k).Formula
                            jj = 1
                        Next k
                    Next j
                    add_flag = True
                End If
            End If
            jj = 1
            r = SheetNew(sheet_name)
            Set Sh = wbk.Sheets(sheet_name)
            Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col)) = data_txt
            If checktxt_on_load Then r = ManualCheck_txt(Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col)))
            If add_flag Then
                For j = 1 To n_row_add
                    For k = n_col + 1 To n_col_sh
                        jj = 1
                    Next k
                Next j
                Sh.Range(Sh.Cells(1, n_col + 1), Sh.Cells(n_row_add, n_col_sh)).Formula = add_col
            End If
        End If
    Next i
End Function

Function ManualAdd(ByVal lastfileadd As String) As Boolean
    nm = ActiveSheet.Name
    If SpecGetType(nm) <> 7 Then
        MsgBox ("Перейдите на лист с ручной спецификацией (заканчивается на _спец) и повторите")
        ManualAdd = False
        Exit Function
    End If
    If Right$(lastfileadd, 4) = "_поз" Then
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
    r = LogWrite(nm, "add", str(UBound(add_array, 1)))
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
                                .InputTitle = vbNullString
                                .ErrorTitle = vbNullString
                                .InputMessage = vbNullString
                                .ErrorMessage = vbNullString
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
                                .InputTitle = vbNullString
                                .ErrorTitle = vbNullString
                                .InputMessage = vbNullString
                                .ErrorMessage = vbNullString
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

Function ManualCeilAlert(ByVal ceil As Variant, ByVal txt As String, Optional ByVal type_alert As String = "alert")
    On Error Resume Next
    ceil.AddComment (txt)
    ceil.Comment.Shape.TextFrame.AutoSize = True
    ceil.Comment.Visible = False
    If type_alert = "alert" Then tcolor = 255
    If type_alert = "info" Then tcolor = 65535
    With ceil.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = tcolor
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

Function ManualCheck_txt(ByVal data_out As Range) As Boolean
    spec = data_out
    n_row = UBound(spec, 1)
    n_col = UBound(spec, 2)
    If n_row < 3 Then Exit Function
    n_col_area = 0
    n_col_volume = 0
    n_col_thicknes = 0
    
    koeff_edizm_area = 1
    koeff_edizm_thicknes = 1000
    koeff_edizm_volume = 1
    For i = 1 To 3
        For j = 1 To n_col
            If Len(spec(i, j)) > 3 Then
                tval = Trim$(LCase$(spec(i, j)))
                If InStr(tval, "толщина") > 0 Then
                    n_col_thicknes = j
                    If InStr(tval, "мм") > 0 Then
                        koeff_edizm_thicknes = 1000
                    Else
                        If InStr(tval, "cм") > 0 Then
                            koeff_edizm_thicknes = 100
                        Else
                            If InStr(tval, "м") > 0 Then koeff_edizm_thicknes = 1
                        End If
                    End If
                End If
                If InStr(tval, "площадь") > 0 Then
                    n_col_area = j
                End If
                If InStr(tval, "объем") > 0 Or InStr(tval, "объём") > 0 Then
                    n_col_volume = j
                End If
                If n_col_area > 0 And n_col_volume > 0 And n_col_thicknes > 0 Then
                    i = 3
                    j = n_col
                End If
            End If
        Next j
    Next i
    
    For i = 3 To n_row
        For j = 1 To n_col
            If data_out.Cells(i, j).HasFormula = True Then
                With data_out.Cells(i, j).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 65535
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Next j
    Next i
    If n_col_area = 0 Or n_col_volume = 0 Or n_col_thicknes = 0 Then Exit Function
    For i = 3 To n_row
        flag = 0
        thicknes = ConvTxt2Num(spec(i, n_col_thicknes))
        volume = ConvTxt2Num(spec(i, n_col_volume))
        area = ConvTxt2Num(spec(i, n_col_area))
        
        If IsNumeric(thicknes) Then
            thicknes = thicknes / koeff_edizm_thicknes
            If thicknes <= 0 Then thicknes = vbNullString
        End If
        If IsNumeric(volume) Then
            volume = volume / koeff_edizm_volume
            If volume <= 0 Then volume = vbNullString
        End If
        If IsNumeric(area) Then
            area = area / koeff_edizm_area
            If area <= 0 Then area = vbNullString
        End If
        If IsNumeric(thicknes) Then
            If Not IsNumeric(volume) Or Not IsNumeric(area) Then
                If Not IsNumeric(volume) And IsNumeric(area) Then
                    newval = (area * thicknes) * koeff_edizm_volume
                    flag = n_col_volume
                End If
                If Not IsNumeric(area) And IsNumeric(volume) Then
                    newval = (volume / thicknes) * koeff_edizm_area
                    flag = n_col_area
                End If
            End If
        Else
            If IsNumeric(volume) Or IsNumeric(area) Then
                newval = (volume / area) * koeff_edizm_thicknes
                flag = n_col_thicknes
            End If
        End If
        If flag > 0 Then
            data_out.Cells(i, flag).FormulaLocal = "=" + CStr(newval)
            With data_out.Cells(i, flag).Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        End If
    Next i
    ManualCheck_txt = True
End Function

Function ManualCheck(ByVal nm As String) As Boolean
    'Проверка корректности заполнения ручной спецификации
    If IsEmpty(pr_adress) Then r = ReadPrSortament()
    type_spec = SpecGetType(nm)
    If type_spec <> 7 And type_spec <> 15 And type_spec <> 26 Then
        MsgBox ("Перейдите на лист с ручной спецификацией" & vbLf & "(заканчивается на _спец) и повторите")
        Exit Function
    End If
    If Not SheetCheckName(nm) Then Exit Function
Dim tfunctime As Double
tfunctime = Timer
    r = SetWorkbook()
    If Not SheetExist(nm) Then
        MsgBox ("Лист не найден" & vbLf & nm)
        Exit Function
    End If
    Set data_out = wbk.Sheets(nm)
    n_row = SheetGetSize(data_out)(1)
    col = max_col_man
    spec = data_out.Range(data_out.Cells(1, 1), data_out.Cells(n_row, max_col_man))
    If type_spec = 26 Then
        r = ManualCheck_txt(data_out.Range(data_out.Cells(1, 1), data_out.Cells(n_row, max_col_man)))
        Exit Function
    End If
    r = FormatClear(data_out)
    data_out.Cells.ClearFormats
    data_out.Cells.ClearComments
    r = FormatFont(data_out.Range(data_out.Cells(1, 1), data_out.Cells(n_row, max_col_man)), n_row, max_col_man)
    n_err = 0
    Set name_subpos = DataNameSubpos(Empty)
    Set concrsubpos = CreateObject("Scripting.Dictionary")
    Set dsubpos = CreateObject("Scripting.Dictionary")
    Set ank_subpos = CreateObject("Scripting.Dictionary")
    Dim type_row(): ReDim type_row(n_row)
    Dim row_arr
    ReDim row_arr(max_col_man)
    For i = 3 To n_row
        For n = 1 To max_col_man
            row_arr(n) = spec(i, n)
        Next n
        type_el = ManualType(row_arr)
        type_row(i) = type_el
        If type_el <> t_syserror Then
            subpos = spec(i, col_man_subpos) ' Марка элемента
            pos = spec(i, col_man_pos)  ' Поз.
            obozn = spec(i, col_man_obozn) ' Обозначение
            naen = spec(i, col_man_naen) ' Наименование
            qty = spec(i, col_man_qty) ' Кол-во на один элемент
            Weight = spec(i, col_man_weight) ' Масса, кг
            prim = spec(i, col_man_prim) ' Примечание (на лист)
            If Not IsNumeric(qty) And Not IsEmpty(qty) Then
                qty = ConvTxt2Num(qty)
                If Not IsNumeric(qty) Then
                    r = ManualCeilAlert(data_out.Cells(i, col_man_qty), "Проверьте разделитель")
                    n_err = n_err + 1
                Else
                    r = ManualCeilSetValue(data_out.Cells(i, col_man_qty), qty, "check")
                End If
            End If
                If type_el = t_sys Then 'Отмечаем вспомогательные строки
                    With data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, max_col_man)).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .Color = XlRgbColor.rgbLightGrey
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    If InStr(obozn, "сновной") > 0 And InStr(naen, "етон") > 0 And InStr(subpos, "!!") > 0 Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_subpos), "Впишите марку элемента")
                        n_err = n_err + 1
                    End If
                    If (InStr(obozn, "ейсмика") > 0 Or InStr(naen, "ейсмика") > 0) And InStr(subpos, "!!") > 0 Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_subpos), "Впишите марку элемента")
                        n_err = n_err + 1
                    End If
                    If InStr(obozn, "сновной") > 0 And InStr(naen, "етон") > 0 And InStr(subpos, "!!") = 0 Then
                        ank_subpos.Item(subpos & "_бет") = naen
                        With data_out.Range(data_out.Cells(i, col_man_obozn), data_out.Cells(i, col_man_qty)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent1
                            .TintAndShade = 0.399975585192419
                            .PatternTintAndShade = 0
                        End With
                    End If
                    If (InStr(obozn, "ейсмика") > 0 Or InStr(naen, "ейсмика") > 0) And InStr(subpos, "!!") = 0 Then
                        ank_subpos.Item(subpos & "_kseism") = 1.3
                        If InStr(obozn, "ейсмика") > 0 Then
                            tcol = col_man_obozn
                        Else
                            tcol = col_man_naen
                        End If
                        With data_out.Cells(i, tcol).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent6
                            .TintAndShade = -0.249977111117893
                            .PatternTintAndShade = 0
                        End With
                    End If
                    If InStr(spec(i, col_man_subpos), "!") > 0 And (InStr(obozn, "заголов") > 0 Or InStr(obozn, "назван") > 0 Or InStr(obozn, "шапка") > 0) Then
                        With data_out.Range(data_out.Cells(i, col_man_obozn), data_out.Cells(i, col_man_naen)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 5296274
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End If
                Else
                    If InStr(pos, "на_все@") > 0 Then
                        With data_out.Range(data_out.Cells(i, col_man_pos), data_out.Cells(i, col_man_pos)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent6
                            .TintAndShade = 0.799981688894314
                            .PatternTintAndShade = 0
                        End With
                    End If
                    If InStr(pos, "arch_") > 0 Then
                        With data_out.Range(data_out.Cells(i, col_man_obozn), data_out.Cells(i, max_col_man)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent1
                            .TintAndShade = 0.799981688894314
                            .PatternTintAndShade = 0
                        End With
                    End If
                End If
                If type_el = t_arm Then 'Правила для арамтуры
                    Length = spec(i, col_man_length) ' Арматура
                    diametr = spec(i, col_man_diametr) ' Диаметр
                    klass = spec(i, col_man_klass) ' Класс
                    gost = GetGOSTForKlass(klass)
                    If StrComp(gost, spec(i, col_man_obozn), 1) <> 0 Then
                        data_out.Cells(i, col_man_obozn).Value = gost
                    End If
                    'Массу п.м. посчитаем автоматом
                    Weight = GetWeightForDiametr(diametr, klass)
                    If Weight <= 0 Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_diametr), "Проверить диаметр")
                        r = ManualCeilAlert(data_out.Cells(i, col_man_weight), "Проверить диаметр")
                        n_err = n_err + 1
                    Else
                        If spec(i, col_man_weight) <> Weight Then
                            data_out.Cells(i, col_man_weight).Value = GetWeightForDiametr(diametr, klass)
                        End If
                        data_out.Cells(i, col_man_weight).Interior.Color = XlRgbColor.rgbLightGrey
                    End If

                    If qty = Empty And prim <> "п.м." Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_qty), "Необходимо указать количество")
                        n_err = n_err + 1
                    End If
                    If Length > lenght_ed_arm And prim <> "п.м." Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_prim), "Стержни длиной выше" + ConvNum2Txt(lenght_ed_arm / 1000) + "должны идти в п.м.")
                        n_err = n_err + 1
                    End If
                    If Length < 100 Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_length), "Подозрительно малая длина.")
                        n_err = n_err + 1
                    End If
                    If InStr(naen, "жатая") > 0 Then ank_subpos.Item(subpos & pos & CStr(i) & "тип") = "сжатая"
                    If InStr(naen, "астянутая") > 0 Then ank_subpos.Item(subpos & pos & CStr(i) & "тип") = "растянутая"
                    If InStr(naen, "войная") > 0 Then ank_subpos.Item(subpos & pos & CStr(i) & "тип") = "двойная"
                    If InStr(data_out.Cells(i, col_man_length).Formula, "Арм_ПоПлощади") > 0 Or InStr(data_out.Cells(i, col_man_length).Formula, "Арм_ОдинСлойПоПлощади") > 0 Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_length), "Длина ОДНОГО слоя, всё должно быть в мм", "info")
                        r = ManualCeilAlert(data_out.Cells(i, col_man_qty), "Кол-во слоёв", "info")
                        If prim <> "п.м." Then r = ManualCeilAlert(data_out.Cells(i, col_man_prim), "Должны идти в п.м.")
                    End If
                    If InStr(naen, "ВТОКАД_") > 0 Then
                        With data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, max_col_man)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .ThemeColor = xlThemeColorAccent2
                            .TintAndShade = 0.799981688894314
                            .PatternTintAndShade = 0
                        End With
                        If Length > lenght_ed_arm And data_out.Cells(i, col_man_length).HasFormula = False And nm <> "Сводная_спец" Then
                            addr_nahl = data_out.Cells(i, col_man_nahl).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                            form = "=Арм_Длина_ПМ(" + ConvNum2Txt(Length) + "," + addr_nahl + "," + ConvNum2Txt(lenght_ed_arm) + ")"
                            data_out.Cells(i, col_man_length).ClearContents
                            data_out.Cells(i, col_man_length).Formula = form
                        End If
                    End If
                End If
                If type_el = t_mat Then
                    If Not ArrayHasElement(material_ed_izm, prim) Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_prim), "Проверьте единицы измерения.")
                        n_err = n_err + 1
                    End If
                    If InStr(naen, "Бетон") <> 0 Then
                        concrsubpos.Item(subpos) = True
                        concrsubpos.Item(subpos & "_" & naen) = i
                        With data_out.Range(data_out.Cells(i, 3), data_out.Cells(i, col_man_qty)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = XlRgbColor.rgbLightBlue
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        data_out.Cells(i, col_man_weight).Interior.Color = XlRgbColor.rgbLightGrey
                        If IsEmpty(obozn) Then
                            r = ManualCeilAlert(data_out.Cells(i, col_man_obozn), "Отсутствует ГОСТ на бетон")
                            n_err = n_err + 1
                        End If
                        If prim <> "куб.м." Then
                            r = ManualCeilAlert(data_out.Cells(i, col_man_obozn), "Бетон должен быть в куб.м.")
                            n_err = n_err + 1
                        End If
                        data_out.Cells(i, col_man_weight).Value = "-"
                    End If
                End If
                If type_el = t_prokat Then
                    pr_length = spec(i, col_man_pr_length) ' Прокат
                    pr_gost_pr = spec(i, col_man_pr_gost_pr) ' ГОСТ профиля
                    pr_prof = spec(i, col_man_pr_prof) ' Профиль
                    pr_type = spec(i, col_man_pr_type) ' Тип конструкции
                    pr_st = spec(i, col_man_pr_st) ' Сталь
                    If IsEmpty(pr_st) Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_pr_st), "Не указана марка стали.")
                        n_err = n_err + 1
                    End If
                    data_out.Cells(i, col_man_weight).Interior.Color = XlRgbColor.rgbLightGrey
                    If IsEmpty(qty) Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_qty), "Необходимо указать количество")
                        n_err = n_err + 1
                    End If
                    If IsEmpty(pr_length) Then
                        r = ManualCeilAlert(data_out.Cells(i, col_man_pr_length), "Необходимо указать длину")
                        n_err = n_err + 1
                    End If
                    If InStr(data_out.Cells(i, col_man_pr_gost_pr).Value, "Лист_") Then
                        If InStr(pr_prof, "--") = 0 Then
                            r = ManualCeilAlert(data_out.Cells(i, col_man_pr_prof), "Проверьте толщину, должно начинаться с --")
                            n_err = n_err + 1
                        Else
                            If GetAreaList(data_out.Cells(i, col_man_naen).Value2) <> data_out.Cells(i, col_man_pr_length).Value2 Then
                                data_out.Cells(i, col_man_pr_length).Value = GetAreaList(data_out.Cells(i, col_man_naen).Value)
                            End If
                            data_out.Cells(i, col_man_pr_length).Interior.Color = XlRgbColor.rgbLightGrey
                        End If
                    End If
                    If Not IsEmpty(pr_adress.Item(pr_gost_pr)) Then data_out.Cells(i, col_man_obozn) = pr_adress.Item(pr_gost_pr)(2)
                    If Not IsEmpty(pr_adress.Item(pr_gost_pr & pr_prof)) Then data_out.Cells(i, col_man_weight) = pr_adress.Item(pr_gost_pr & pr_prof)(1)
                    If Not IsEmpty(pr_length) And Not IsEmpty(pr_gost_pr) And Not IsEmpty(pr_prof) And Not IsEmpty(qty) And Not IsEmpty(pr_st) Then
                        With data_out.Range(data_out.Cells(i, col_man_pos), data_out.Cells(i, col_man_qty)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = XlRgbColor.rgbLightGoldenrodYellow
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                    End If
                End If
                If type_el = t_subpos Then 'Правила для маркировки сборок
                    If InStr(prim, "авто") > 0 Then
                        If name_subpos.Exists(subpos) Then
                            If InStr(prim, "поз") > 0 Then
                                tnaen = name_subpos.Item(subpos)(1)
                                tobozn = name_subpos.Item(subpos)(2)
                                data_out.Cells(i, col_man_obozn) = tobozn
                                data_out.Cells(i, col_man_naen) = tnaen
                                naen = tnaen
                                obozn = tobozn
                            End If
                            If InStr(prim, "кол") > 0 Then
                                tqty = name_subpos.Item(subpos)(3)
                                If tqty = 0 Then tqty = Empty
                                data_out.Cells(i, col_man_qty) = tqty
                                qty = tqty
                            End If
                        End If
                    End If
                    If qty = Empty Then
                        With data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, max_col_man)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = XlRgbColor.rgbLightGreen
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        suff = vbNullString
                        If IsEmpty(obozn) Then
                            r = ManualCeilAlert(data_out.Cells(i, col_man_obozn), "Нужна ссылка на лист")
                            n_err = n_err + 1
                        End If
                    Else
                        With data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, max_col_man)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = XlRgbColor.rgbLightCoral
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        suff = "_par"
                    End If
                    If InStr(prim, "авто") > 0 Then
                            With data_out.Cells(i, col_man_prim).Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                .Color = 1072322
                                .TintAndShade = 0
                                .PatternTintAndShade = 0
                            End With
                            If InStr(prim, "поз") > 0 Then
                                With data_out.Range(data_out.Cells(i, col_man_obozn), data_out.Cells(i, col_man_naen)).Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .Color = 1072322
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                            End If
                            If InStr(prim, "кол") > 0 Then
                                With data_out.Cells(i, col_man_qty).Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .Color = 1072322
                                    .TintAndShade = 0
                                    .PatternTintAndShade = 0
                                End With
                            End If
                    End If
                    ky = pos & " " & obozn & " " & naen & suff
                    If dsubpos.Exists(ky) Then
                        dsubpos.Item(ky) = dsubpos.Item(ky) + 1
                        dsubpos.Item(ky + "_adr") = dsubpos.Item(ky + "_adr") + "+" + data_out.Cells(i, 1).Address
                    Else
                        dsubpos.Item(ky) = 1
                        dsubpos.Item(ky + "_adr") = data_out.Cells(i, 1).Address
                    End If
                End If
                If type_el = t_error Then
                    r = ManualCeilAlert(data_out.Cells(i, col_man_length), "Проверьте правильность заполнения.")
                    r = ManualCeilAlert(data_out.Cells(i, col_man_pr_length), "Проверьте правильность заполнения.")
                    n_err = n_err + 1
                End If
                If type_el = -2 Then
                    r = ManualCeilAlert(data_out.Cells(i, col_man_subpos), "Пустая строка")
                    n_err = n_err + 1
                End If
                If type_el = 0 Then
                    With data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, max_col_man)).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.08
                        .PatternTintAndShade = 0
                    End With
                    With data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, max_col_man))
                        .Borders(xlDiagonalDown).LineStyle = xlNone
                        .Borders(xlDiagonalUp).LineStyle = xlNone
                        .Borders(xlEdgeLeft).LineStyle = xlNone
                        .Borders(xlEdgeRight).LineStyle = xlNone
                        .Borders(xlInsideVertical).LineStyle = xlNone
                        .Borders(xlInsideHorizontal).LineStyle = xlNone
                    End With
                End If
        Else
            With data_out.Range(data_out.Cells(i, 1), data_out.Cells(i, max_col_man)).Interior
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
        type_el = type_row(i)
        If type_el = t_arm Then
            subpos = spec(i, col_man_subpos) ' Марка элемента
            pos = spec(i, col_man_pos)  ' Поз.
            diametr = spec(i, col_man_diametr) ' Диаметр
            klass = spec(i, col_man_klass) ' Класс
            r_opr = Арм_МинРадиус(diametr, klass) - 0.5 * diametr
            If spec(i, col_man_dgib) <> r_opr Then data_out.Cells(i, col_man_dgib) = r_opr
            If Not ank_subpos.Exists(subpos & "_бет") Then
                If spec(i, col_man_ank) <> "НЕТ БЕТОНА" Then data_out.Cells(i, col_man_ank) = "НЕТ БЕТОНА"
                If spec(i, col_man_nahl) <> "НЕТ БЕТОНА" Then data_out.Cells(i, col_man_nahl) = "НЕТ БЕТОНА"
            Else
                beton = ank_subpos.Item(subpos & "_бет")
                kseism = 1
                If ank_subpos.Exists(subpos & "_kseism") Then kseism = 1.3
                type_arm = "растянутая"
                If ank_subpos.Exists(subpos & pos & CStr(i) & "тип") Then type_arm = ank_subpos.Item(subpos & pos & CStr(i) & "тип")
                type_out = "L"
                l_ank = Арм_Анкеровка(diametr, klass, beton, kseism, type_arm, type_out)
                l_nahl = Арм_Нахлёст(diametr, klass, beton, kseism, type_arm, type_out)
                If spec(i, col_man_ank) <> l_ank Then data_out.Cells(i, col_man_ank) = l_ank
                If spec(i, col_man_nahl) <> l_nahl Then data_out.Cells(i, col_man_nahl) = l_nahl
            End If
        End If
    Next
    un_man_subpos = ArrayUniqValColumn(spec, col_man_subpos)
    If Not IsEmpty(un_man_subpos) Then
        For Each subpos In un_man_subpos
            If ank_subpos.Exists(subpos & "_бет") And concrsubpos.Exists(subpos) Then
                flag_eq = 0
                i = 0
                bet_ank = ank_subpos.Item(subpos & "_бет")
                bet_ank = GetClassBeton(bet_ank)
                For Each bet In concrsubpos.keys()
                    If InStr(bet, "_") > 0 And InStr(bet, subpos) > 0 Then
                        i = concrsubpos.Item(bet)
                        bet = GetClassBeton(bet)
                        If InStr(bet, bet_ank) > 0 Then flag_eq = 1
                    End If
                Next
                If flag_eq = 0 And i > 0 Then
                    r = ManualCeilAlert(data_out.Cells(i, col_man_naen), "Марка отличается от марки для расчёта анкеровки (" + ank_subpos.Item(subpos & "_бет") + ")")
                    n_err = n_err + 1
                Else
                    concrsubpos.Item(subpos & "@бет") = bet_ank
                End If
            End If
        Next
    End If
    If SheetExist(izd_sheet_name) And nm <> izd_sheet_name Then
        Set spec_izd_sheet = wbk.Sheets(izd_sheet_name)
        spec_izd_size = SheetGetSize(spec_izd_sheet)
        n_izd_row = spec_izd_size(1)
        spec_izd = spec_izd_sheet.Range(spec_izd_sheet.Cells(3, 1), spec_izd_sheet.Cells(n_izd_row, max_col_man))
        For i = 1 To UBound(spec_izd, 1)
        row = ArrayRow(spec_izd, i)
            If ManualType(row) = t_subpos Then
                pos = row(col_man_pos)  ' Поз.
                obozn = row(col_man_obozn) ' Обозначение
                naen = row(col_man_naen) ' Наименование
                ky = pos & " " & obozn & " " & naen & vbNullString
                If dsubpos.Exists(ky) Then
                    dsubpos.Item(ky) = dsubpos.Item(ky) + 1
                    dsubpos.Item(ky + "_adr") = vbNullString
                Else
                    dsubpos.Item(ky) = 1
                    dsubpos.Item(ky + "_adr") = vbNullString
                End If
            End If
        Next
    End If

    For Each ky In dsubpos.keys()
        If InStr(ky, "_adr") = 0 Then
            If dsubpos.Item(ky) > 1 Then
                For Each adr In Split(dsubpos.Item(ky + "_adr"), "+")
                    adr = Replace(adr, "$", vbNullString)
                    If InStr(ky, "_par") = 0 Then
                        r = ManualCeilAlert(data_out.Range(adr), "Повторное определение вложенной сборки (" & dsubpos.Item(ky) & " раза)")
                        n_err = n_err + 1
                    Else
                        r = ManualCeilAlert(data_out.Range(adr), "Эта сборка повторяется " & dsubpos.Item(ky) & " раза. Не ошибка, но подозрительно.")
                    End If
                Next
            End If
            For i = 3 To n_row
                type_el = type_row(i)
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
                        With data_out.Range(data_out.Cells(i, col_man_pos), data_out.Cells(i, col_man_qty)).Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = XlRgbColor.rgbBurlyWood
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                        data_out.Cells(i, col_man_weight).Interior.Color = XlRgbColor.rgbLightGrey
                        If Not IsEmpty(prim) Then
                            r = ManualCeilAlert(data_out.Cells(i, col_man_prim), "Вхождения сборок - только в штуках. Удали " + prim)
                            n_err = n_err + 1
                        End If
                        If Int(qty) - qty <> 0 Then
                            r = ManualCeilAlert(data_out.Cells(i, col_man_qty), "Дробное количество сборок")
                            n_err = n_err + 1
                        End If
                        If Not IsEmpty(Weight) Then
                            r = ManualCeilAlert(data_out.Cells(i, col_man_weight), "Масса для сборки считается автоматически. Удали " + str(Weight))
                            n_err = n_err + 1
                        End If

                    End If
                End If
            Next i
        End If
    Next
    r = FormatManual(nm)
    If (n_err) Then
        MsgBox ("Обнаружено " & str(n_err) & " ошибок, см. примечания к ячейкам")
        ManualCheck = False
    Else
        ManualCheck = True
    End If
    r = LogWrite(nm, "check", str(n_err))
tfunctime = functime("ManualCheck", tfunctime)
End Function

Function ManualDiff(ByVal add_array As Variant, ByVal man_arr As Variant, ByVal type_el As Long) As Variant
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
End Function

Function ManualPaste2Sheet(ByRef array_in As Variant) As Boolean
    Set Sh = wbk.ActiveSheet
    If SpecGetType(Sh.Name) <> 7 And SpecGetType(Sh.Name) <> 15 Then
        MsgBox ("Перейдите на лист с ручной спецификацией (заканчивается на _спец) и повторите")
        ManualPaste2Sheet = False
        Exit Function
    End If
    startpos = SheetGetSize(Sh)(1) + 2
    endpos = startpos + UBound(array_in, 1) - 1
    Sh.Range(Sh.Cells(startpos, 1), Sh.Cells(endpos, max_col_man)) = array_in
    r = ManualCheck(Sh.Name)
End Function

Function ManualPasteIzd2Sheet(ByRef array_in As Variant, Optional ByVal n_first_row_t As Long, Optional ByVal subpos As String, Optional ByVal nm As String) As Boolean
    If IsEmpty(array_in) Then
        ManualPasteIzd2Sheet = False
        Exit Function
    End If
    If nm = vbNullString Then
        Set Sh = wbk.ActiveSheet
    Else
        Set Sh = wbk.Sheets(nm)
    End If
    If SpecGetType(Sh.Name) <> 7 And SpecGetType(Sh.Name) <> 15 Then
        MsgBox ("Перейдите на лист с ручной спецификацией (заканчивается на _спец) и повторите")
        ManualPasteIzd2Sheet = False
        Exit Function
    End If
    n_add = 1
    If n_first_row_t = 0 Then
        n_first_row = ActiveCell.row
        is_row_epmty = True
        For i = 1 To max_col_man
            If Len(Sh.Cells(n_first_row, i).Value) > 0 Then is_row_epmty = False
        Next i
        If is_row_epmty Then n_add = 0
    Else
        n_first_row = n_first_row_t
    End If
    If ArrayIsSecondDim(array_in) Then
        n_col = UBound(array_in, 2)
        n_row = UBound(array_in, 1) - 1
    Else
        n_col = UBound(array_in)
        n_row = 0
    End If
    If subpos = vbNullString Then
        subpos = Sh.Cells(n_first_row - 1, 1).Value
        addr = Sh.Cells(n_first_row - 1, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If
    If subpos = vbNullString Then
        For i = n_first_row To 3 Step -1
            If Len(Trim$(Sh.Cells(i, 1).Value)) > 0 And subpos = vbNullString Then
                subpos = Sh.Cells(i, 1).Value
                addr = Sh.Cells(i, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
            End If
        Next i
    End If
    n_first_row = n_first_row + n_add
    For i = n_first_row To n_first_row + n_row
        Rows(i).Insert Shift:=xlDown
    Next i
    Sh.Range(Sh.Cells(n_first_row, 1), Sh.Cells(n_first_row + n_row, n_col)) = array_in
    If addr <> vbNullString Then
        Sh.Range(Sh.Cells(n_first_row, 1), Sh.Cells(n_first_row + n_row, 1)).Formula = "=" + addr
    Else
        Sh.Range(Sh.Cells(n_first_row, 1), Sh.Cells(n_first_row + n_row, 1)).Value = subpos
    End If
    Application.Calculation = xlCalculationAutomatic
    r = ManualCheck(Sh.Name)
    Application.Calculation = xlCalculationManual
    ManualPasteIzd2Sheet = True
End Function

Function ManualUndoPos(ByVal nm As String) As Boolean
    istart = 2
    Set spec_sheet = wbk.Sheets(nm)
    sheet_size = SheetGetSize(spec_sheet)
    n_row = sheet_size(1)
    If n_row = istart Then n_row = n_row + 1
    Dim pos_out(): ReDim pos_out(max_col)
    spec = spec_sheet.Range(spec_sheet.Cells(1, 1), spec_sheet.Cells(n_row, max_col_man + 1))
    For i = istart To n_row
        If spec(i, max_col_man + 1) <> Empty Then
            spec_sheet.Cells(i, col_man_pos) = spec(i, max_col_man + 1)
            spec_sheet.Cells(i, max_col_man + 1) = Empty
        End If
    Next
    ManualUndoPos = True
End Function

Function posarmsort(ByRef chksum_pos As Variant, ByVal arm As Variant, ByVal cur_pos As Long, ByVal type_pos As Long) As Long
Dim tfunctime As Double
tfunctime = Timer
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
        arm_temp = ArraySort_2(arm_temp, Array(col_diametr, col_length))
        For i = UBound(arm_temp, 1) To LBound(arm_temp, 1) Step -1
            cur_pos = cur_pos + 1
            chksum_pos.Item(arm_temp(i, col_chksum)) = cur_pos
        Next i
    End If
    'Остальное сортируем по длине
    'Берём прямые стержни
    arm_temp = ArraySelectParam_2(arm, 0, col_fon, 0, col_gnut)
    If Not IsEmpty(arm_temp) Then
        arm_temp = ArraySort_2(arm_temp, Array(col_diametr, col_length))
        For i = UBound(arm_temp, 1) To LBound(arm_temp, 1) Step -1
            cur_pos = cur_pos + 1
            chksum_pos.Item(arm_temp(i, col_chksum)) = cur_pos
        Next i
    End If
    'Теперь - гнутые
    arm_temp = ArraySelectParam_2(arm, 1, col_gnut)
    If Not IsEmpty(arm_temp) Then
        arm_temp = ArraySort_2(arm_temp, Array(col_diametr, col_length))
        For i = UBound(arm_temp, 1) To LBound(arm_temp, 1) Step -1
            cur_pos = cur_pos + 1
            chksum_pos.Item(arm_temp(i, col_chksum)) = cur_pos
        Next i
    End If
    posarmsort = cur_pos
tfunctime = functime("posarmsort", tfunctime)
End Function

Function ManualPos(ByVal nm As String, ByVal type_pos As Long) As Boolean
    floor_txt = "all_floor"
    istart = 2
    Set spec_sheet = wbk.Sheets(nm)
    sheet_size = SheetGetSize(spec_sheet)
    n_row = sheet_size(1)
    If n_row = istart Then n_row = n_row + 1
    Dim pos_out(): ReDim pos_out(max_col)
    spec = spec_sheet.Range(spec_sheet.Cells(1, 1), spec_sheet.Cells(n_row, max_col_man + 1))
    If IsEmpty(pr_adress) Then r = ReadPrSortament()
    out_data = DataRead(nm)
    'Словарь, где будем хранить контрольную сумму и позицию
    Set chksum_pos = CreateObject("Scripting.Dictionary")
    'Лишние элементы убираем
    un_parent = ArraySort(ArrayCombine(pos_data.Item(floor_txt).Item("parent").keys(), Array("-"))) 'для всех сборок и элементов вне сборок
    arm = ArraySelectParam_2(out_data, t_arm, col_type_el, un_parent, col_sub_pos)
    un_parent = ArrayUniqValColumn(arm, col_sub_pos)
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
            subpos = Trim$(Replace(row(col_man_subpos), subpos_delim, "@"))  ' Марка элемента
            If spec(i, max_col_man + 1) = Empty Then spec_sheet.Cells(i, max_col_man + 1) = spec_sheet.Cells(i, col_man_pos)
            pos = Trim$(Replace(row(col_man_pos), subpos_delim, "@"))  ' Поз.
            obozn = Trim$(row(col_man_obozn)) ' Обозначение
            naen = Trim$(row(col_man_naen)) ' Наименование
            prim = Trim$(row(col_man_prim)) ' Примечание (на лист)
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
    r = ManualCheck(nm)
    ManualPos = True
End Function

Function ManualCopyIzd(ByVal spec As Variant, ByVal spec_izd As Variant) As Variant
    If IsEmpty(spec_izd) Or IsEmpty(spec) Then
        ManualCopyIzd = Empty
        Exit Function
    End If
    unic_pos_mun = ArrayUniqValColumn(spec, col_man_pos)
    unic_subpos_sheet = ArrayUniqValColumn(spec, col_man_subpos)
    unic_subpos_izd = ArrayUniqValColumn(spec_izd, col_man_subpos)
    If IsEmpty(unic_subpos_izd) Then
        ManualCopyIzd = Empty
        Exit Function
    End If
    For i = 1 To UBound(unic_subpos_izd)
        flag_use = False
        For j = 1 To UBound(unic_pos_mun)
            If unic_subpos_izd(i) = unic_pos_mun(j) Then
                flag_use = True
                Exit For
            End If
        Next j
        'Если сборка(изделие) уже объявлена на листе - копировать её с листа изделий не нужно
        If flag_use = True Then
            For j = 1 To UBound(unic_subpos_sheet)
                If unic_subpos_izd(i) = unic_subpos_sheet(j) Then
                    flag_use = False
                    Exit For
                End If
            Next j
        End If
        If flag_use = False Then unic_subpos_izd(i) = Empty
    Next i
    For Each subpos_izd In unic_subpos_izd
        If Not IsEmpty(subpos_izd) Then
            subpos_spec_izd = ArraySelectParam(spec_izd, subpos_izd, col_man_subpos)
            spec_add = ArrayCombine(spec_add, subpos_spec_izd)
        End If
    Next
    ManualCopyIzd = spec_add
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
    If Not SheetExist(nm) Then Exit Function
    r = ManualAddAuto(nm)
    Set spec_sheet = wbk.Sheets(nm)
    sheet_size = SheetGetSize(spec_sheet)
    n_row = sheet_size(1)
    If n_row = istart Then n_row = n_row + 1
    spec = spec_sheet.Range(spec_sheet.Cells(1, 1), spec_sheet.Cells(n_row, max_col_man))
    If SheetExist(izd_sheet_name) And nm <> izd_sheet_name Then
        Set spec_izd_sheet = wbk.Sheets(izd_sheet_name)
        spec_izd_size = SheetGetSize(spec_izd_sheet)
        n_izd_row = spec_izd_size(1)
        spec_izd = spec_izd_sheet.Range(spec_izd_sheet.Cells(3, 1), spec_izd_sheet.Cells(n_izd_row, max_col_man))
        spec_add = ManualCopyIzd(spec, spec_izd)
        spec = ArrayCombine(spec, spec_add)
    End If
    n_row = UBound(spec, 1)
    Dim pos_out(): ReDim pos_out(n_row - istart, max_col): n_row_out = 0
    Dim param
    Dim add_okr_array
    n_add_okr = 0
    For i = istart To n_row
        If Not IsEmpty(spec(i, col_man_pr_okr)) And spec(i, col_man_pr_okr) <> "-" Then n_add_okr = n_add_okr + 1
    Next i
    ReDim add_okr_array(n_add_okr, max_col)
    n_add_okr = 0
    Set manual_title = CreateObject("Scripting.Dictionary")
    For i = istart To n_row
        row = ArrayRow(spec, i)
        type_el = ManualType(row)
        obozn = Trim$(row(col_man_obozn)) ' Обозначение
        naen = Trim$(row(col_man_naen)) ' Наименование
        If type_el = t_sys And InStr(row(col_man_subpos), "!") > 0 And (InStr(obozn, "заголов") > 0 Or InStr(obozn, "назван") > 0 Or InStr(obozn, "шапка") > 0) Then
            ttspec = SpecGetType(obozn)
            manual_title.Item(ttspec) = naen
        End If
        If type_el > 0 And type_el <> t_sys Then
            flag_qtyall = 0
            If InStr(row(col_man_pos), "arch_") > 0 Then row(col_man_pos) = Replace(row(col_man_pos), "arch_", vbNullString)
            pos = row(col_man_pos)
            If InStr(pos, "на_все@") > 0 Then
                flag_qtyall = 1
                row(col_man_pos) = Replace(pos, "на_все@", vbNullString)
            End If
            subpos = Trim$(Replace(row(col_man_subpos), subpos_delim, "@"))  ' Марка элемента
            pos = Trim$(Replace(row(col_man_pos), subpos_delim, "@"))  ' Поз.
            qty = row(col_man_qty) ' Кол-во на один элемент
            Weight = row(col_man_weight) ' Масса, кг
            prim = row(col_man_prim) ' Примечание (на лист)
            prim = Replace(prim, "авто", vbNullString)
            prim = Replace(prim, "поз", vbNullString)
            prim = Replace(prim, "кол", vbNullString)
            prim = Trim$(prim)
            If qty = Empty Or qty <= 0 Then qty = 1
            If type_el = t_subpos Then nSubPos = qty
            If nSubPos = Empty Or nSubPos <= 0 Then nSubPos = 1
            n_row_out = n_row_out + 1
            pos_out(n_row_out, col_marka) = pos
            pos_out(n_row_out, col_sub_pos) = subpos
            pos_out(n_row_out, col_type_el) = type_el
            pos_out(n_row_out, col_pos) = pos
            If flag_qtyall Then
                pos_out(n_row_out, col_qty) = qty
            Else
                pos_out(n_row_out, col_qty) = qty * nSubPos
            End If
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
                    If Len(prim) > 1 Then pos_out(n_row_out, col_pr_naen) = pos_out(n_row_out, col_pr_naen) + "@@" + prim
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
    Set this_sheet_option.Item("title") = manual_title
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
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_obozn), vbNullString, "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_weight), vbNullString, "add")
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
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_prim), vbNullString, "add")
                Case Else
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_obozn), add_array(i, col_m_obozn), "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_naen), add_array(i, col_m_naen), "add")
                    r = ManualCeilSetValue(spec_sheet.Cells(end_row, col_man_weight), vbNullString, "add")
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
                        If pos_out(j, col_sub_pos) = pos Then pos_out(j, col_type_el) = vbNullString
                    Next j
                End If
            Next i
            'Осталось ещё раз пройти по элементам и добавить элементы из сборок второго уровня,
            'Поменяв при этом обозначение вхождения сборки с t_izd на t_subpos
            Dim subarray(): ReDim subarray(max_col, 1)
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
End Function

Function ManualSpec_batch(type_out)
    r = LogWrite("Автовывод", "Начало", "-")
    If mem_option Then r = LogWrite("Автовывод", "Включена автонастройка листов", "-")
    n_out = 0
    r = OutPrepare()
    For Each objWh In wbk.Worksheets
        nm = objWh.Name
        type_spec = SpecGetType(nm)
        If type_spec = 7 Then
            For Each tspec In type_out
                If Not IsEmpty(tspec) Then
                    'If mem_option Then r = SheetSetOption(nm)
                    sheet_out = Spec_Select(nm, tspec, True)
                    r = ExportSheet(sheet_out)
                    n_out = n_out + 1
                End If
            Next
        End If
    Next objWh
    r = SheetIndex()
    r = LogWrite("Автовывод", "Конец", str(n_out))
    r = OutEnded()
End Function

Function ManualSpec_NewSubpos()
    r = OutPrepare()
    nm_out = izd_sheet_name
    If SheetExist(nm_out) Then
        Worksheets(nm_out).Activate
    Else
        wbk.Worksheets.Add.Name = nm_out
    End If
    Worksheets(nm_out).Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set out_sheet = ThisWorkbook.Sheets(nm_out)
    r = FormatClear(data_out)
    r = FormatManual(nm_out)
    r = FormatManual(nm_out)
    n_last = SheetGetSize(out_sheet)(1) + 2
    flag = Empty
    If UserForm2.fromthiswbCB.Value Then
        For Each objWh In wbk.Worksheets
            nm = objWh.Name
            type_spec = SpecGetType(nm)
            If type_spec = 7 Then
                Set spec_sheet = wbk.Sheets(nm)
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

Function Sheet2Dict(ByRef sheet_data As Variant, ByRef objWh As Worksheet) As Long
    nm$ = objWh.Name
    If mem_option And Left(nm$, 1) <> "|" And Left(nm$, 1) <> "!" Then
        Set sheet_option_param_tmp = OptionSheetGet(nm)
        If sheet_option_param_tmp.Item("defult") = True Then MsgBox ("Для спецификации " & nm & " не найдены сохранённые параметры. Можно задать их  при ручнов выводе или на листе с содержанием.")
        Set this_sheet_option = sheet_option_param_tmp
    End If
    type_spec = SpecGetType(nm$)
    If (type_spec = 7 Or type_spec = 15) And StrComp(nm$, "Сводная_спец") <> 0 Then
        n_row = SheetGetSize(objWh)(1)
        sheet_data.Item(nm$) = objWh.Range(objWh.Cells(3, 1), objWh.Cells(n_row, max_col_man))
        un_subpos = ArrayUniqValColumn(ReadPos(nm$), col_sub_pos)
        out_mat = DataReadAutoMat(nm$, un_subpos)
        n_mat = 0
        If Not IsEmpty(out_mat) Then
            n_mat = UBound(out_mat, 1) + 1
            Dim out_mat_man(): ReDim out_mat_man(n_mat, max_col_man)
            out_mat_man(1, col_man_subpos) = "!!!"
            out_mat_man(1, col_man_pos) = "!!!"
            out_mat_man(1, col_man_naen) = "РАСХОД ДАН НА ВСЁ КОЛ-ВО СБОРОК"
            out_mat_man(1, col_man_obozn) = "АРХИКАД_материалы"
            For n = 2 To n_mat
                out_mat_man(n, col_man_subpos) = out_mat(n - 1, col_sub_pos)
                out_mat_man(n, col_man_pos) = "arch_на_все@" + out_mat(n - 1, col_pos)
                out_mat_man(n, col_man_naen) = out_mat(n - 1, col_m_naen)
                out_mat_man(n, col_man_obozn) = out_mat(n - 1, col_m_obozn)
                out_mat_man(n, col_man_qty) = out_mat(n - 1, col_qty)
                If out_mat(n - 1, col_m_weight) > 0 Then out_mat_man(n, col_man_weight) = out_mat(n - 1, col_m_weight)
                out_mat_man(n, col_man_prim) = out_mat(n - 1, col_m_edizm)
            Next n
            sheet_data.Item(nm$ + "_mat") = out_mat_man
        End If
        If StrComp(izd_sheet_name, nm$) = 0 Then
            Sheet2Dict = 0
        Else
            Sheet2Dict = n_mat + n_row + 1
        End If
    Else
        Sheet2Dict = 0
    End If
End Function

Function ManualSpec_MergeSheet(ByRef spec_book As Workbook, ByRef out_sheet As Worksheet) As Long
    Set sheet_data = CreateObject("Scripting.Dictionary")
    Dim objWh As Worksheet
    n_row = 1
    For Each objWh In spec_book.Worksheets
        n_row = n_row + Sheet2Dict(sheet_data, objWh)
    Next
    Dim spec(): ReDim spec(n_row, max_col_man)
    With sheet_data
        n_row = 1
        spec(n_row, col_man_subpos) = "!!!"
        spec(n_row, col_man_pos) = "!!!"
        spec(n_row, col_man_obozn) = "ИЗ ФАЙЛА"
        spec(n_row, col_man_naen) = spec_book.Name
        flag_izd = 0
        For Each Key In sheet_data.keys
            If StrComp(izd_sheet_name, Key) <> 0 Then
                temp_arr = sheet_data.Item(Key)
                n_row_t = UBound(temp_arr, 1)
                n_row = n_row + 1
                spec(n_row, col_man_subpos) = "!!!"
                spec(n_row, col_man_pos) = "!!!"
                spec(n_row, col_man_obozn) = "C ЛИСТА"
                spec(n_row, col_man_naen) = Key
                For j = 1 To n_row_t
                    flag_del = 0
                    pos = temp_arr(j, col_man_pos)
                    If InStr(pos, "!!") > 0 Then
                        subpos = temp_arr(j, col_man_subpos)
                        obozn = temp_arr(j, col_man_obozn)
                        If InStr(subpos, "!!") = 0 Then
                            flag_del = 1
                            naen = temp_arr(j, col_man_naen)
                            If InStr(obozn, "сновно") Or InStr(naen, "етон") Then flag_del = 0
                            If InStr(obozn, "ейсмик") Or InStr(naen, "ейсмик") Then flag_del = 0
                            If InStr(obozn, "АВТОКАД") Or InStr(naen, "АВТОКАД") Then flag_del = 0
                        Else
                            If InStr(obozn, "C ЛИСТА") = 0 And InStr(naen, "ИЗ ФАЙЛА") = 0 Then flag_del = 1
                            If InStr(obozn, "АРХИКАД") = 0 Or InStr(naen, "АРХИКАД") = 0 Then flag_del = 1
                        End If
                    End If
                    n_row = n_row + 1
                    If flag_del = 0 Then
                        For k = 1 To max_col_man
                            spec(n_row, k) = temp_arr(j, k)
                        Next k
                    End If
                Next j
            Else
                flag_izd = 1
            End If
        Next
    End With
    n_row_start = SheetGetSize(out_sheet)(1)
    n_row_t = n_row
    out_sheet.Range(out_sheet.Cells(n_row_start, 1), out_sheet.Cells(n_row_start + n_row, max_col_man)) = spec
    If flag_izd Then
        spec_add = ManualCopyIzd(spec, sheet_data.Item(izd_sheet_name))
        If Not IsEmpty(spec_add) Then
            n_row_t = n_row + UBound(spec_add, 1) + 1
            n_row = n_row + 1
            out_sheet.Cells(n_row + n_row_start, 1) = "!!!"
            out_sheet.Cells(n_row + n_row_start, 2) = "!!!"
            out_sheet.Cells(n_row + n_row_start, 3) = "C ЛИСТА"
            out_sheet.Cells(n_row + n_row_start, 4) = izd_sheet_name
            n_row = n_row + 1
            out_sheet.Range(out_sheet.Cells(n_row + n_row_start, 1), out_sheet.Cells(n_row_t + n_row_start, max_col_man)) = spec_add
        End If
    End If
    Set sheet_data = Nothing
    ManualSpec_MergeSheet = n_row_t
End Function


Function ManualSpec_Merge()
    r = OutPrepare()
    r = SetWorkbook()
    nm_out = "Сводная_спец"
    If SheetExist(nm_out) Then wbk.Sheets(nm_out).Delete
    wbk.Worksheets.Add.Name = nm_out
    Dim out_sheet As Worksheet
    Dim spec_book As Workbook
    Set out_sheet = wbk.Sheets(nm_out)
    out_sheet.Cells.Clear
    r = FormatClear(out_sheet)
    r = FormatManual(nm_out)
    r = FormatManual(nm_out)
    n_row_out = 0
    If UserForm2.fromthiswbCB.Value Then
       n_row_out = ManualSpec_MergeSheet(wbk, out_sheet)
    End If
    If UserForm2.fromfileCB.Value Then
        Set coll = FilenamesCollection(ThisWorkbook.path, ".xlsm")
        For Each snm In coll
            If InStr(snm, "Спец") And InStr(snm, "~$") = 0 And InStr(snm, ThisWorkbook.Name) = 0 Then
                Set spec_book = GetObject(snm)
                n_row_out = n_row_out + ManualSpec_MergeSheet(spec_book, out_sheet)
                spec_book.Close SaveChanges:=False
                Set spec_book = Nothing
            End If
        Next
    End If
    For i = n_row_out To 3 Step -1
        If Application.CountA(out_sheet.Rows(i)) = 0 Then
            out_sheet.Rows(i).Delete
        End If
    Next
    r = ManualCheck(nm_out)
    r = OutEnded()
End Function


Function ManualType(ByVal row As Variant) As Long
    If IsEmpty(row) Then
        ManualType = t_syserror
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
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
    If isArm And Not (IsNumeric(Length) And IsNumeric(diametr) And Not IsNumeric(klass)) Then
        ManualType = t_error
        Exit Function
    End If
    isProkat = ((Not IsEmpty(pr_length) Or Not IsEmpty(pr_gost_pr) Or Not IsEmpty(pr_prof) Or Not IsEmpty(pr_prof)) And Not isSys)
    If isProkat And Not IsNumeric(pr_length) Then
        ManualType = t_error
        Exit Function
    End If
    
    ismat = ((ArrayHasElement(material_ed_izm, prim) Or InStr(naen, "Бетон") > 0) And Not isSys)
    isEr = ((isSPos And isArm) Or (isSPos And isProkat) Or (isSPos And ismat) Or (isArm And isProkat) Or (isArm And ismat) Or (isProkat And ismat)) 'Проверим - не подходит ли элемент к нескольким типам
    
    If Not isSys And tempt < 3 Then type_el = -2
    If isSys Then type_el = t_sys
    If isSPos Then type_el = t_subpos
    If isArm Then type_el = t_arm
    If isProkat Then type_el = t_prokat
    If ismat Then type_el = t_mat
    If isEr Then type_el = t_error
    
    ManualType = type_el
r = functime("ManualType", tfunctime)
End Function

Function NRowOut(ByRef arr As Variant) As Variant
Dim tfunctime As Double
tfunctime = Timer
    n = 0
    If Not (ArrayIsSecondDim(arr)) Then
        n = 1
    Else
        n_row = UBound(arr, 1)
        n_col = UBound(arr, 2)
        For i = 1 To n_row
            Fl = 0
            If i = 6 Then
            H = 1
            End If
            For j = 1 To n_col
                el = Trim$(arr(i, j))
                If el = vbNullString Or el = " " Or el = 0 Or IsEmpty(el) Then Fl = Fl + 1
                
                If i < n_row Then
                    next_el = Trim$(arr(i + 1, j))
                    If el <> vbNullString And el <> " " And el <> 0 And Not IsEmpty(el) Then Fl = Fl - 1
                End If

            Next j
            If Fl < n_col Then n = n + 1
        Next i
    End If
    NRowOut = n
tfunctime = functime("NRowOut", tfunctime)
End Function

Function OutEnded() As Boolean
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    ActiveSheet.DisplayPageBreaks = True
    Application.DisplayStatusBar = True
    Application.DisplayAlerts = True
    Application.Calculate
    ThisWorkbook.Save
    OutEnded = True
End Function

Function OutPrepare() As Boolean
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    ActiveSheet.DisplayPageBreaks = False
    Application.DisplayStatusBar = False
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    OutPrepare = True
End Function

Function ReadFile(ByVal mask As String, Optional ByVal FirstRow& = 1, Optional ByVal ColumnsSeparator$ = ";", Optional ByVal RowsSeparator$ = vbNewLine, Optional ByVal read_sys_file As Boolean = False) As Variant
Dim tfunctime As Double
tfunctime = Timer
    On Error Resume Next
    Set coll = FilenamesCollection(ThisWorkbook.path & "\import\", mask)
    For Each File In coll
        arr = ArrayCombine(arr, ReadTxt(File, FirstRow&, ColumnsSeparator$, RowsSeparator$, read_sys_file))
    Next
    ReadFile = arr

tfunctime = functime("ReadFile", tfunctime)
End Function

Function ReadMetall() As Boolean
Dim tfunctime As Double
tfunctime = Timer
    SortamentPath = SetPath()
    nf_prof = SortamentPath & "Имена профилей.csv"
    If Len(Dir$(nf_prof)) > 0 Then
        name_gost = ReadTxt(nf_prof, 1, ";", vbNewLine, True)
    Else
        MsgBox ("Нет файла с именами профилей")
        r = LogWrite("Ошибка профилей", vbNullString, "Нет файла с именами профилей")
    End If
tfunctime = functime("ReadMetall", tfunctime)
End Function

Function ReadPos(ByVal lastfileadd As String) As Variant
    r = SetWorkbook()
    Set add_sheet = wbk.Sheets(lastfileadd)
    sheet_size = SheetGetSize(add_sheet)
    istart = 2
    n_row = sheet_size(1)
    n_col = 6
    spec = add_sheet.Range(add_sheet.Cells(1, 1), add_sheet.Cells(n_row, n_col))
    Dim add_array(): ReDim add_array(n_row - istart + 1, max_col): n_row_out = 0
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
    'erase add_array
End Function

Function ReadPrSortament()
Dim tfunctime As Double
tfunctime = Timer
    r = SetWorkbook()
    If Not SheetExist("!System") Then ThisWorkbook.Worksheets.Add.Name = "!System"
    Set Sh = wbk.Sheets("!System") 'На этом скрытом листе будем хранить данные для списков
    Set tpr_adress = CreateObject("Scripting.Dictionary") 'В этом словаре будем хранить адреса
    Set swap_gost = CreateObject("Scripting.Dictionary") 'Для срочной замены ГОСТов
    'Сначала - металл
    SortamentPath = SetPath()
    File = SortamentPath & "Сортаменты.txt"
    If Not CBool(Len(Dir$(File))) Then r = Download_Sortament()
    f_list_sort = ReadTxt(File, 1, vbTab, vbNewLine, True)
    f_list_file = ArrayCol(f_list_sort, 3)
    f_list_gost = ArrayCol(f_list_sort, 2)
    n_sort = UBound(f_list_file)
    tpr_adress.Item("ГОСТпрокат") = "'!System'!" & Sh.Range(Sh.Cells(1, 1), Sh.Cells(1, n_sort)).Address
    Dim tmp_arr(3)
    For n_col = 2 To n_sort
        File = f_list_file(n_col)
        Sh.Cells(1, n_col - 1) = File
        If Dir(SortamentPath & File & ".txt") = vbNullString Then
            MsgBox ("Файл не найден " + File)
            r = Download_Sortament()
            Exit Function
        End If
        f_prof = ReadTxt(SortamentPath & File & ".txt", 1, vbTab, vbNewLine, True)
        f_list_prof = ArrayCol(f_prof, 2)
        f_list_weight = ArrayCol(f_prof, 3)
        If IsEmpty(f_list_prof) Or Not IsArray(f_prof) Then
            MsgBox ("Ошибка чтения файла " + File)
            Exit Function
        End If
        n_prof = UBound(f_list_prof) + 1
        Sh.Range(Sh.Cells(2, n_col - 1), Sh.Cells(n_prof, n_col - 1)) = ArrayTranspose(f_list_prof)
        tmp_arr(1) = "'!System'!" & Sh.Range(Sh.Cells(3, n_col - 1), Sh.Cells(n_prof, n_col - 1)).Address
        tmp_arr(2) = f_list_gost(n_col)
        tpr_adress.Item(File) = tmp_arr
        type_prof = f_list_sort(n_col, 5)
        For j = 2 To n_prof - 1
            If Not IsEmpty(f_list_weight(j)) And IsNumeric(f_prof(j, 4)) And IsNumeric(f_list_weight(j)) Then
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
                tpr_adress.Item(File & prof) = tmp_arr
            End If
        Next j
    Next
    n_start = n_sort + 1
    
    'Теперь арматура
    File = SortamentPath & "Сортамент_арматуры.txt"
    f_list_sort = ReadTxt(File, 1, ";", vbNewLine, True)
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
    File = SortamentPath & "Сталь.txt"
    f_list_stal = ReadTxt(File, 1, vbTab, vbNewLine, True)
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
    For i = 1 To UBound(material_ed_izm)
        Sh.Cells(2 + i, n_start) = material_ed_izm(i)
    Next i
    tpr_adress.Item("Примечания") = "'!System'!" & Sh.Range(Sh.Cells(1, n_start), Sh.Cells(6, n_start)).Address
    
    n_start = n_end + 2
    'Типы окраски
    Sh.Cells(1, n_start) = "-"
    Sh.Cells(2, n_start) = "Тип 1"
    Sh.Cells(3, n_start) = "Тип 2"
    Sh.Cells(4, n_start) = "Тип 3"
    Sh.Cells(5, n_start) = "Тип 4"
    tpr_adress.Item("Окраска") = "'!System'!" & Sh.Range(Sh.Cells(1, n_start), Sh.Cells(5, n_start)).Address
    
    n_start = n_end + 2
    'Замена ГОСТов
    File = SortamentPath & "Замена ГОСТов.txt"
    f_list_gost = ReadTxt(File, 1, vbTab, vbNewLine, True)
    If Not IsEmpty(f_list_gost) Then
        n_gost = UBound(f_list_gost, 1)
        For i = 1 To n_gost
            swap_gost.Item(f_list_gost(i, 1)) = f_list_gost(i, 2)
        Next
    End If
    Set pr_adress = tpr_adress
    ReadPrSortament = True
tfunctime = functime("ReadPrSortament", tfunctime)
End Function

Function ReadReinforce() As Boolean
Dim tfunctime As Double
tfunctime = Timer
    'Чтение сортамента
    SortamentPath = SetPath()
    nf_sort = SortamentPath & "Сортамент_арматуры.txt"
    If Len(Dir$(nf_sort)) > 0 Then
        reinforcement_specifications = ReadTxt(nf_sort, 1, ";", vbNewLine, True)
    Else
        MsgBox ("Нет файла сортамента арматуры")
        r = LogWrite("Ошибка арматуры", vbNullString, "Нет файла сортамента арматуры")
        Exit Function
    End If
    Set gost2fklass = CreateObject("Scripting.Dictionary")
    'Массив соответсвия классов и гостов
    For i = 1 To UBound(reinforcement_specifications, 1)
        klass = reinforcement_specifications(i, col_klass_spec)
        gost = reinforcement_specifications(i, col_gost_spec)
        If InStr(klass, "Класс") = 0 Then gost2fklass.Item(klass) = gost
    Next i
tfunctime = functime("ReadReinforce", tfunctime)
End Function

Function ReadTxt(ByVal filename$, Optional ByVal FirstRow& = 1, Optional ByVal ColumnsSeparator$ = ";", Optional ByVal RowsSeparator$ = vbNewLine, Optional ByVal read_sys_file As Boolean = False) As Variant
Dim tfunctime As Double
tfunctime = Timer
    On Error Resume Next
    Set FSO = CreateObject("scripting.filesystemobject")
    Set ts = FSO.OpenTextFile(filename$, 1, True): txt$ = ts.ReadAll: ts.Close
    If read_sys_file = False Then
        If def_decode Then this_sheet_option.Item("decode") = True
        If this_sheet_option.decode_CB.Value = True Then
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
    End If
    Set ts = Nothing: Set FSO = Nothing
    txt = Trim$(txt): Err.Clear
    If txt Like "*" & RowsSeparator$ Then txt = Left$(txt, Len(txt) - Len(RowsSeparator$))
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
        tmpArr2 = Split(Trim$(tmpArr1(i)), ColumnsSeparator$)
        For j = 1 To UBound(tmpArr2) + 1
            arr(i + 1, j) = ConvTxt2Num(Trim$(tmpArr2(j - 1)))
        Next j
    Next i
    If Len(txt) = 0 Then
        ReadTxt = Empty
    Else
        ReadTxt = arr
    End If
    'erase arr
tfunctime = functime("ReadTxt", tfunctime)
End Function

Function RelFName(ByVal fname As String) As String
    n_slash = InStrRev(fname, "\")
    n_len = Len(fname)
    n_dot = 4
    wt_dot = Left$(fname, n_len - n_dot)
    n_len = Len(wt_dot)
    wt_path = Right$(wt_dot, n_len - n_slash)
    RelFName = wt_path
End Function

Function Round_w(ByVal arg As Variant, ByVal nokrg As Variant, Optional ByVal hard_round As Boolean = False) As Variant
Dim tfunctime As Double
tfunctime = Timer
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
        If arg = vbNullString Or arg = " " Then
            Round_w = 0
        Else
            Round_w = arg
        End If
    End If
tfunctime = functime("Round_w", tfunctime)
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
Dim tfunctime As Double
tfunctime = Timer
    r = SetWorkbook()
    On Error Resume Next
    Dim objWh As Excel.Worksheet
    Dim NameLst As String
    For Each objWh In wbk.Worksheets
        NameLst = objWh.Name
        If NameLst = NameSheet Then
            SheetExist = True
tfunctime = functime("SheetExist", tfunctime)
            Exit Function
        End If
    Next objWh
    SheetExist = False
tfunctime = functime("SheetExist", tfunctime)
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
    For Each objWh In wbk.Worksheets
        nm = objWh.Name
        If InStr(nm, "архикад") = 0 And Left$(nm, 1) <> "|" And Left$(nm, 1) <> "!" Then
            type_spec = SpecGetType(nm)
            If type_out(1) = -1 Then
                Select Case type_spec
                    Case 1, 2, 4, 5, 11, 12, 13, 14, 20
                        wbk.Sheets(nm).Delete
                        n_del = n_del + 1
                        r = LogWrite(nm, vbNullString, "DEL")
                End Select
            Else
                For Each tdel In type_out
                    If Not IsEmpty(tdel) And tdel = type_spec Then
                        On Error Resume Next
                        wbk.Sheets(nm).Delete
                        n_del = n_del + 1
                        r = LogWrite(nm, vbNullString, "DEL")
                    End If
                Next
            End If
        End If
    Next objWh
    SheetClear = n_del
End Function

Function SheetGetSize(ByVal objLst As Variant) As Variant
Dim tfunctime As Double
tfunctime = Timer
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
    'erase out
tfunctime = functime("SheetGetSize", tfunctime)
End Function

Function SheetHideAll()
    Worksheets(inx_name).Activate
    Dim sheet As Worksheet
    With wbk
        For Each sheet In ThisWorkbook.Worksheets
            If Left$(sheet.Name, 1) = "!" Then Sheets(sheet.Name).Visible = False
        Next
    End With
End Function

Function SheetImport(ByVal nm As String) As Boolean
    Dim importbook As Object
    Set importbook = Nothing
    On Error Resume Next
    Workbooks(nm).Close SaveChanges:=False
    On Error Resume Next
    Set importbook = GetObject(nm)
    If Not importbook Is Nothing Then
        fname = GetFileName(nm)
        fname_1 = "[" + fname + "]"
        nm = Replace(nm, fname, fname_1)
        listsheet = GetListOfSheet(importbook)
        For Each sheet_name In listsheet
            If SpecGetType(sheet_name) > 0 Then
                If SheetExist(sheet_name) Then
                    For n = 1 To 100
                        sn = str(n) + " " + sheet_name
                        If Not SheetExist(sn) Then Exit For
                    Next n
                    importbook.Sheets(sheet_name).Name = sn
                    sheet_name = sn
                End If
                importbook.Sheets(sheet_name).Copy Before:=wbk.Sheets(1)
                Set spec_sheet = wbk.Sheets(sheet_name)
                sheet_size = SheetGetSize(spec_sheet)
                Set spec = spec_sheet.Range(spec_sheet.Cells(1, 1), spec_sheet.Cells(sheet_size(1), sheet_size(2)))
                For Each cel In spec.Cells
                    If InStr(cel.Formula, "[") Then
                        cel.Formula = Replace(cel.Formula, nm, vbNullString)
                        cel.Formula = Replace(cel.Formula, fname_1, vbNullString)
                    End If
                Next
            End If
        Next
        importbook.Close SaveChanges:=False
        Kill nm
        SheetImport = True
    Else
        SheetImport = False
    End If
End Function

Function SheetActivate(ByVal sheetn As String)
    If ModeType() = True Then Exit Function
    If Not isINIset Then r = INISet()
    If sheetn = inx_name And check_on_active And UserForm2.Visible Then
        r = OutPrepare()
        r = SheetIndex()
        r = OutEnded()
    Else
        type_spec = SpecGetType(sheetn)
        If (type_spec = 7 Or type_spec = 15) And check_on_active And UserForm2.Visible Then
            n_row = SheetGetSize(wbk.Sheets(sheetn))(1)
            If n_row < 1000 Then
                r = OutPrepare()
                r = ManualCheck(sheetn)
                r = OutEnded()
            End If
        End If
        If type_spec > 0 And UserForm2.Visible Then
            r = UserForm2.set_sheet(sheetn)
            If mem_option Then r = OptionSheetSet(sheetn)
        End If
    End If
End Function

Function SheetIndex_NCol(ByVal sheetn As String)
    n_col = 0
    start_symb = Left$(sheetn, 1)
    If start_symb <> "|" And start_symb <> "!" Then
        tspec = SpecGetType(sheetn)
        Select Case tspec
        Case 7, 10
            n_col = 1
        Case 1, 2, 3, 4, 5, 13, 0, 20, 21
            n_col = 2
        Case Else
            n_col = 3
        End Select
    End If
    SheetIndex_NCol = n_col
End Function

Function SheetIndex()
    If Not isINIset Then r = INISet()
    tcheck_on_active = check_on_active
    check_on_active = False
    tmem_option = mem_option
    mem_option = False
    If SheetExist(inx_name) Then
        wbk.Worksheets(inx_name).Activate
    Else
        wbk.Worksheets.Add.Name = inx_name
        wbk.Worksheets(inx_name).Activate
    End If
    Set sheets_params = OptionSheetGet_All()
    wbk.Worksheets(inx_name).Move Before:=ThisWorkbook.Sheets(1)
    Dim sheet As Worksheet
    Dim cell As Range
    Worksheets(inx_name).Cells.Clear
    r = FormatClear(Worksheets(inx_name))
    Worksheets(inx_name).Cells(3) = "Ведомости"
    Worksheets(inx_name).Cells(2) = "Автоматические"
    Worksheets(inx_name).Cells(1) = "Ручные"
    Dim sheetnames(): j = 0
    With wbk
        For Each sheet In wbk.Worksheets
            j = j + 1
            ReDim Preserve sheetnames(j)
            sheetnames(j) = sheet.Name
            If InStr(sheetnames(j), "из архикада") > 0 Then sheetnames(j) = "яяя" + sheetnames(j)
            If InStr(sheetnames(j), "_спец") > 0 Then sheetnames(j) = fin_str + sheetnames(j)
        Next
    End With
    sheetnames = ArraySort(sheetnames)
    For j = 1 To UBound(sheetnames)
        sheetn = Replace(Replace(sheetnames(j), "яяя", vbNullString), fin_str, vbNullString)
        tspec = SpecGetType(sheetn)
        Select Case tspec
            Case 1, 2, 3, 4, 5, 13, 20, 21
                With wbk.Sheets(sheetn).Tab
                    .Color = 0
                    .TintAndShade = 0
                End With
            Case 6, 7, 9, 10, 11, 12, 14, 8
                wbk.Worksheets(sheetn).Move After:=wbk.Sheets(2)
                With wbk.Sheets(sheetn).Tab
                    .ThemeColor = xlThemeColorAccent4
                    .TintAndShade = 0.4
                End With
            Case 15
                wbk.Worksheets(sheetn).Move After:=wbk.Sheets(2)
                With wbk.Sheets(sheetn).Tab
                    .ThemeColor = xlThemeColorAccent5
                    .TintAndShade = 0.5
                End With
            Case 26
                wbk.Worksheets(sheetn).Move After:=wbk.Sheets(UBound(sheetnames))
                With wbk.Sheets(sheetn).Tab
                    .ThemeColor = xlThemeColorAccent5
                    .TintAndShade = 0.5
                End With
            Case Else
                With wbk.Sheets(sheetn).Tab
                    .Color = 0
                    .TintAndShade = 1
                End With
        End Select
        If sheetn = inx_name Then
            With wbk.Sheets(sheetn).Tab
                .Color = 5287936
                .TintAndShade = 0
            End With
        End If
    Next
    k = 1
    For j = 2 To UBound(sheetnames) + 1
        sheetn = Replace(Replace(sheetnames(j - 1), "яяя", vbNullString), fin_str, vbNullString)
        n_col = SheetIndex_NCol(sheetn)
        If n_col > 0 Then
            k = k + 1
            Set cell = Worksheets(inx_name).Cells(k, n_col)
            ThisWorkbook.Worksheets(inx_name).Hyperlinks.Add anchor:=cell, Address:=vbNullString, SubAddress:="'" & sheetn & "'" & "!A1"
            cell.Formula = sheetn
            'Set sheet_param_tmp = sheets_params.Item(sheetn)
            r = OptionSheetWrite(sheetn, sheets_params.Item(sheetn), k)
            Sheets(sheetn).Visible = True
        Else
            If sheetn <> inx_name Then Sheets(sheetn).Visible = False
        End If
        Sheets(inx_name).Visible = True
    Next
    ThisWorkbook.Worksheets(inx_name).Activate
    ThisWorkbook.Worksheets(inx_name).Rows("1:1").Font.Bold = True
    With ThisWorkbook.Worksheets(inx_name).Rows("1:1").Font
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
    ThisWorkbook.Worksheets(inx_name).Rows("1:1").RowHeight = 80
    ThisWorkbook.Worksheets(inx_name).Rows("2:100").RowHeight = 15
    ThisWorkbook.Worksheets(inx_name).Range("A2:AC100").Rows.AutoFit
    ThisWorkbook.Worksheets(inx_name).Columns("A:C").ColumnWidth = 24
    ThisWorkbook.Worksheets(inx_name).Columns("D:E").ColumnWidth = 8
    ThisWorkbook.Worksheets(inx_name).Columns("F").ColumnWidth = 15
    ThisWorkbook.Worksheets(inx_name).Columns("I:AA").ColumnWidth = 5
    ThisWorkbook.Worksheets(inx_name).Columns("AB:AE").ColumnWidth = 10
    ThisWorkbook.Worksheets(inx_name).Range("A1:AE100").Select
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

    Set sheet_option_inx = OptionDefultDict()
    For Each varKey In sheet_option_inx.keys
        arr = sheet_option_inx.Item(varKey)
        Worksheets(inx_name).Cells(1, arr(3)) = arr(1)
        With Worksheets(inx_name).Cells(1, arr(3))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 90
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With Worksheets(inx_name).Cells(1, arr(3)).Font
            .Name = "Calibri"
            .FontStyle = "полужирный"
            .Size = 9
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
    Next
    With ThisWorkbook.Worksheets(inx_name).Range("G2:AA100").Font
        .Name = "Calibri"
        .Size = 6
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
    ThisWorkbook.Worksheets(inx_name).Range("G2:AA100").Select
    Selection.FormatConditions.Add Type:=xlTextString, String:="ЛОЖЬ", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlTextString, String:="ИСТИНА", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    ThisWorkbook.Worksheets(inx_name).Range("A1").Select
    check_on_active = tcheck_on_active
    mem_option = tmem_option
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

Function regexp(ByVal txt As String, ByVal expr As String) As Boolean
Dim tfunctime As Double
tfunctime = Timer
    flag = False
    expr = Replace(expr, "?", "*")
    If expr = "*" Then
        flag = True
    Else
        If InStr(expr, "*") > 0 Then
            'Подстановочные знаки
            rparam = (StrComp(Right$(expr, 1), "*", vbBinaryCompare) = 0) 'стоит ли сзади знак *
            lParam = (StrComp(Left$(expr, 1), "*", vbBinaryCompare) = 0)  'стоит ли спереди знак *
            texpr = Trim$(Replace(expr, "*", vbNullString))
            If rparam And lParam Then
                If InStr(txt, texpr) > 0 Then flag = 1
            Else
                If lParam Then tparam = Right$(txt, Len(texpr))
                If rparam Then tparam = Left$(txt, Len(texpr))
                If StrComp(tparam, texpr, vbBinaryCompare) = 0 Then flag = True
            End If
        Else
            'Простое сравнение
            If StrComp(txt, expr, vbBinaryCompare) = 0 Then flag = True
        End If
    End If
    regexp = flag
tfunctime = functime("regexp", tfunctime)
End Function

Function DataFilter(ByVal txt As String, ByVal arr_add As Variant, ByVal arr_del As Variant) As Boolean
Dim tfunctime As Double
tfunctime = Timer
    txt = LCase$(txt)
    txt = Trim$(txt)
    flag_ignore = False
    flag_add = True
    For i = LBound(arr_del) To UBound(arr_del)
        del_type = arr_del(i)
        If Len(del_type) > 0 Then
            flag_ignore = regexp(txt, del_type)
            If flag_ignore Then Exit For
        End If
    Next
    If flag_ignore = False Then
        For i = LBound(arr_add) To UBound(arr_add)
            add_type = arr_add(i)
            If Len(add_type) > 0 Then
                flag_add = regexp(txt, add_type)
                If flag_add Then Exit For
            End If
        Next
    End If
    flag_filter = (flag_add = True And flag_ignore = False)
    DataFilter = flag_filter
tfunctime = functime("DataFilter", tfunctime)
End Function

Function OptionSplit(ByVal txt As String) As Variant
Dim tfunctime As Double
tfunctime = Timer
    If IsEmpty(txt) Then txt = vbNullString
    If Len(txt) > 0 Then
        txt = LCase$(txt)
        txt = Trim$(txt)
    End If
    If IsEmpty(txt) Then txt = vbNullString
    If InStr(txt, ";") > 0 Then
        out = Split(txt, ";")
        For i = LBound(out) To UBound(out)
            out(i) = Trim$(out(i))
        Next i
    Else
        out = Array(txt)
    End If
    OptionSplit = out
tfunctime = functime("OptionSplit", tfunctime)
End Function

Function OptionJoin(ByVal arr As Variant) As String
Dim tfunctime As Double
tfunctime = Timer
    If IsEmpty(arr) Then
        txt = ""
    End If
    If IsArray(arr) Then
        txt = Join(arr, ";")
    Else
        txt = CStr(arr)
    End If
    OptionJoin = txt
tfunctime = functime("OptionJoin", tfunctime)
End Function

Function OptionDefultDict()
Dim tfunctime As Double
tfunctime = Timer
    Set sheet_option_inx = CreateObject("Scripting.Dictionary")
    sheet_option_inx.comparemode = 1
    '[1]Название столбца [2]Значение по умолчанию [3]Номер столбца в листе с содержанием
    sheet_option_inx.Item("k_zap") = Array("Коэффицент запаса", "1.0", 5)
    sheet_option_inx.Item("date") = Array("Дата последнего обновления", "---", 6)
    sheet_option_inx.Item("arm_pm") = Array("Всю арматуру в п.м.", False, 7)
    sheet_option_inx.Item("pr_pm") = Array("Весь прокат в п.м.", False, 8)
    sheet_option_inx.Item("keep_pos") = Array("Сохранять разбивку по поз.", False, 9)
    sheet_option_inx.Item("qty_one_subpos") = Array("Расход на одну сборку", False, 10)
    sheet_option_inx.Item("show_subpos") = Array("Сборки в общестрое", True, 11)
    sheet_option_inx.Item("only_subpos") = Array("Только элементы в сборках", False, 12)
    sheet_option_inx.Item("qty_one_floor") = Array("Разбивать по этажам", False, 13)
    sheet_option_inx.Item("ignore_subpos") = Array("Игнорировать разделение на сборки", False, 14)
    sheet_option_inx.Item("merge_material") = Array("[Отделка] Объединять ячейки с одинаковыми материалами", True, 15)
    sheet_option_inx.Item("separate_material") = Array("[Отделка] Разбивать материалы", True, 16)
    sheet_option_inx.Item("otd_by_type") = Array("[Отделка] Разбивать по типам", True, 17)
    sheet_option_inx.Item("ed_izm_km") = Array("Техничка КМ в кг", False, 18)
    sheet_option_inx.Item("show_type") = Array("Показ типов элементов", False, 19)
    sheet_option_inx.Item("add_row") = Array("Добавить пустую строку после типов", False, 20)
    sheet_option_inx.Item("show_qty_spec") = Array("Показ ', на все', ', на 1 шт.'", False, 21)
    sheet_option_inx.Item("decode") = Array("Декодировать txt", True, 22)
    sheet_option_inx.Item("isarm") = Array("Арматура", True, 23)
    sheet_option_inx.Item("isizd") = Array("Изделия", True, 24)
    sheet_option_inx.Item("isprok") = Array("Прокат", True, 25)
    sheet_option_inx.Item("ismat") = Array("Материал", True, 26)
    sheet_option_inx.Item("issubpos") = Array("Сборки", True, 27)
    sheet_option_inx.Item("arr_subpos_add") = Array("Имена добавляемых сборок через ;", "*", 28)
    sheet_option_inx.Item("arr_subpos_del") = Array("Имена исключаемых сборок через ;", " ", 29)
    sheet_option_inx.Item("arr_typeKM_add") = Array("Типы добавляемых конструкций КМ через ;", "*", 30)
    sheet_option_inx.Item("arr_typeKM_del") = Array("Типы исключаемых конструкций КМ через ;", " ", 31)
    Set OptionDefultDict = sheet_option_inx
tfunctime = functime("OptionDefultDict", tfunctime)
End Function

Function OptionSheetGet(ByVal sheetn As String) As Variant
    'Выдаёт словарь с параметрами для заданного листа
    'Если лист ранее не сохранялся - выдаёт значения по умолчанию
Dim tfunctime As Double
tfunctime = Timer
    Set sheet_option_param = CreateObject("Scripting.Dictionary")
    sheet_option_param.comparemode = 1
    inx_row = OptionSheetInx(sheetn)
    If inx_row < 0 Then
        If IsEmpty(spec_type_suffix) Then r = INISet()
        For Each k In spec_type_suffix.keys
            inx_row = OptionSheetInx(sheetn + "_" + k)
            If inx_row > 0 Then
                Exit For
            End If
        Next
    End If
    Set sheet_option_inx = OptionDefultDict()
    If inx_row > 0 Then
        end_col = sheet_option_inx.Item("arr_typeKM_del")(3)
        existssheet = ThisWorkbook.Worksheets(inx_name).Range(ThisWorkbook.Worksheets(inx_name).Cells(inx_row, 1), ThisWorkbook.Worksheets(inx_name).Cells(inx_row, end_col))
        sheet_option_param.Item("defult") = False
    Else
        sheet_option_param.Item("defult") = True
        existssheet = Empty
    End If
    For Each parname In sheet_option_inx.keys
        inx_col = sheet_option_inx.Item(parname)(3)
        defult_value = sheet_option_inx.Item(parname)(2)
        If Not IsEmpty(existssheet) Then
            parval = existssheet(1, inx_col)
            If IsEmpty(parval) Then parval = defult_value
        Else
            parval = defult_value
        End If
        If ArrayHasElement(Array("arr_subpos_add", "arr_subpos_del", "arr_typeKM_add", "arr_typeKM_del"), parname) Then
            If (parname = "arr_subpos_add" Or parname = "arr_typeKM_add") And Len(Trim(parval)) = 0 Then parval = "*"
            If (parname = "arr_subpos_del" Or parname = "arr_typeKM_del") And Trim(parval) = "*" Then parval = ""
            sheet_option_param.Item(parname) = OptionSplit(parval)
        Else
            sheet_option_param.Item(parname) = parval
        End If
    Next
    sheet_option_param.Item("sheet_name") = sheetn
    Set OptionSheetGet = sheet_option_param
tfunctime = functime("OptionSheetGet", tfunctime)
End Function

Function OptionSheetGet_All() As Variant
    r = SetWorkbook()
    Set sheets_params = CreateObject("Scripting.Dictionary")
    With wbk
        For Each sheet In wbk.Worksheets
            sheetn = sheet.Name
            n_col = SheetIndex_NCol(sheetn)
            If n_col > 0 Then
                Set sheets_params_tmp = OptionSheetGet(sheetn)
                Set sheets_params.Item(sheetn) = sheets_params_tmp
            End If
        Next
    End With
    Set OptionSheetGet_All = sheets_params
End Function

Function OptionSheetInx(ByVal sheetn As String) As Long
    'Если имя листа есть в содержании - вернёт положительное значение с номером строки
    'Если лист не найден - отрицатиельное значение последней пустой строки
Dim tfunctime As Double
tfunctime = Timer
    If Not isINIset Then r = INISet()
    n_row = SheetGetSize(ThisWorkbook.Worksheets(inx_name))(1)
    Set existssheet = ThisWorkbook.Worksheets(inx_name).Range(ThisWorkbook.Worksheets(inx_name).Cells(1, 1), ThisWorkbook.Worksheets(inx_name).Cells(n_row, 3))
    Set myCell = existssheet.Find(sheetn, LookAt:=1, MatchCase:=0)
    If myCell Is Nothing Then Set myCell = existssheet.Find(sheetn + "_экспл", LookAt:=1, MatchCase:=0)
    If myCell Is Nothing Then Set myCell = existssheet.Find(sheetn + "_вед", LookAt:=1, MatchCase:=0)
    If myCell Is Nothing Then Set myCell = existssheet.Find(sheetn + "_спец", LookAt:=1, MatchCase:=0)
    If Not myCell Is Nothing Then
        OptionSheetInx = myCell.row
    Else
        OptionSheetInx = -(n_row + 1)
    End If
tfunctime = functime("OptionSheetInx", tfunctime)
End Function

Function OptionSheetSet(ByVal sheetn As String)
Dim tfunctime As Double
tfunctime = Timer
    If sheetn = inx_name Or sheetn = izd_sheet_name Then Exit Function
    Set sheet_option_tmp = OptionSheetGet(sheetn)
    Set set_sheet_option = sheet_option_tmp
    UserForm2.Kzap.Text = sheet_option_tmp.Item("k_zap")
    UserForm2.arm_pm_CB.Value = sheet_option_tmp.Item("arm_pm")
    UserForm2.pr_pm_CB.Value = sheet_option_tmp.Item("pr_pm")
    UserForm2.keep_pos_CB.Value = sheet_option_tmp.Item("keep_pos")
    UserForm2.qtyOneSubpos_CB.Value = sheet_option_tmp.Item("qty_one_subpos")
    UserForm2.show_subpos_CB.Value = sheet_option_tmp.Item("show_subpos")
    UserForm2.ignore_subpos_CB.Value = sheet_option_tmp.Item("ignore_subpos")
    UserForm2.merge_material_CB.Value = sheet_option_tmp.Item("merge_material")
    UserForm2.otd_by_type_CB.Value = sheet_option_tmp.Item("otd_by_type")
    UserForm2.add_row_CB.Value = sheet_option_tmp.Item("add_row")
    UserForm2.ed_izm_km_CB.Value = sheet_option_tmp.Item("ed_izm_km")
    UserForm2.separate_material_CB.Value = sheet_option_tmp.Item("separate_material")
    UserForm2.show_type_CB.Value = sheet_option_tmp.Item("show_type")
    UserForm2.show_qty_spec.Value = sheet_option_tmp.Item("show_qty_spec")
    UserForm2.decode_CB.Value = sheet_option_tmp.Item("decode")
    UserForm2.only_subpos_CB.Value = sheet_option_tmp.Item("only_subpos")
    UserForm2.qtyOneFloor_CB.Value = sheet_option_tmp.Item("qty_one_floor")
    If def_decode Then UserForm2.decode_CB.Value = True
    UserForm1.CheckBox_isarm.Value = sheet_option_tmp.Item("isarm")
    UserForm1.CheckBox_isizd.Value = sheet_option_tmp.Item("isizd")
    UserForm1.CheckBox_isprok.Value = sheet_option_tmp.Item("isprok")
    UserForm1.CheckBox_ismat.Value = sheet_option_tmp.Item("ismat")
    UserForm1.CheckBox_issubpos.Value = sheet_option_tmp.Item("issubpos")
    UserForm1.TextBox_subpos_add.Value = OptionJoin(sheet_option_tmp.Item("arr_subpos_add"))
    UserForm1.TextBox_subpos_del.Value = OptionJoin(sheet_option_tmp.Item("arr_subpos_del"))
    UserForm1.TextBox_typeKM_add.Value = OptionJoin(sheet_option_tmp.Item("arr_typeKM_add"))
    UserForm1.TextBox_typeKM_del.Value = OptionJoin(sheet_option_tmp.Item("arr_typeKM_del"))
    OptionSheetSet = True
tfunctime = functime("OptionSheetSet", tfunctime)
End Function

Function OptionGetForm(ByVal nm As String) As Variant
Dim tfunctime As Double
tfunctime = Timer
    tdate = Right$(str(DatePart("yyyy", Now)), 2) & str(DatePart("m", Now)) & str(DatePart("d", Now))
    stamp = tdate + "/" + str(DatePart("h", Now)) + str(DatePart("n", Now)) + str(DatePart("s", Now))
    Set sheet_option_tmp = CreateObject("Scripting.Dictionary")
    sheet_option_tmp.comparemode = 1
    '[1]Название столбца [2]Значение по умолчанию [3]Номер столбца
    sheet_option_tmp.Item("k_zap") = UserForm2.Kzap.Text
    sheet_option_tmp.Item("date") = stamp
    sheet_option_tmp.Item("arm_pm") = UserForm2.arm_pm_CB.Value
    sheet_option_tmp.Item("pr_pm") = UserForm2.pr_pm_CB.Value
    sheet_option_tmp.Item("keep_pos") = UserForm2.keep_pos_CB.Value
    sheet_option_tmp.Item("qty_one_subpos") = UserForm2.qtyOneSubpos_CB.Value
    sheet_option_tmp.Item("show_subpos") = UserForm2.show_subpos_CB.Value
    sheet_option_tmp.Item("ignore_subpos") = UserForm2.ignore_subpos_CB.Value
    sheet_option_tmp.Item("only_subpos") = UserForm2.only_subpos_CB.Value '!!!!!!!!!
    sheet_option_tmp.Item("qty_one_floor") = UserForm2.qtyOneFloor_CB.Value '!!!!!!!!!
    sheet_option_tmp.Item("merge_material") = UserForm2.merge_material_CB.Value
    sheet_option_tmp.Item("separate_material") = UserForm2.separate_material_CB.Value
    sheet_option_tmp.Item("otd_by_type") = UserForm2.otd_by_type_CB.Value
    sheet_option_tmp.Item("ed_izm_km") = UserForm2.ed_izm_km_CB.Value
    sheet_option_tmp.Item("show_type") = UserForm2.show_type_CB.Value
    sheet_option_tmp.Item("add_row") = UserForm2.add_row_CB.Value
    sheet_option_tmp.Item("show_qty_spec") = UserForm2.show_qty_spec.Value
    sheet_option_tmp.Item("decode") = UserForm2.decode_CB.Value
    sheet_option_tmp.Item("isarm") = UserForm1.CheckBox_isarm.Value
    sheet_option_tmp.Item("isizd") = UserForm1.CheckBox_isizd.Value
    sheet_option_tmp.Item("isprok") = UserForm1.CheckBox_isprok.Value
    sheet_option_tmp.Item("ismat") = UserForm1.CheckBox_ismat.Value
    sheet_option_tmp.Item("issubpos") = UserForm1.CheckBox_issubpos.Value
    If Len(Trim(UserForm1.TextBox_subpos_add.Value)) = 0 Then UserForm1.TextBox_subpos_add.Value = "*"
    If Len(Trim(UserForm1.TextBox_typeKM_add.Value)) = 0 Then UserForm1.TextBox_typeKM_add.Value = "*"
    sheet_option_tmp.Item("arr_subpos_add") = OptionSplit(UserForm1.TextBox_subpos_add.Value)
    sheet_option_tmp.Item("arr_subpos_del") = OptionSplit(UserForm1.TextBox_subpos_del.Value)
    sheet_option_tmp.Item("arr_typeKM_add") = OptionSplit(UserForm1.TextBox_typeKM_add.Value)
    sheet_option_tmp.Item("arr_typeKM_del") = OptionSplit(UserForm1.TextBox_typeKM_del.Value)
    sheet_option_tmp.Item("sheet_name") = nm
    Set OptionGetForm = sheet_option_tmp
tfunctime = functime("OptionGetForm", tfunctime)
End Function

Function OptionSheetWrite(ByVal sheetn As String, ByVal sheet_option_tmp As Variant, Optional ByVal inx_row As Long = 0)
Dim tfunctime As Double
tfunctime = Timer
    If Not isINIset Then r = INISet()
    If sheetn = inx_name Or sheetn = izd_sheet_name Or sheetn = "|Лог|" Then Exit Function
    Set sheet_option_inx = OptionDefultDict()
    If IsEmpty(sheet_option_tmp) Then Set sheet_option_tmp = OptionGetForm(sheetn)
    first_col = sheet_option_inx.Item("k_zap")(3)
    end_col = sheet_option_inx.Item("arr_typeKM_del")(3)
    Dim write_parm
    ReDim write_parm(end_col - first_col + 1)
    For Each parname In sheet_option_inx.keys
        inx_col = sheet_option_inx.Item(parname)(3) - first_col + 1
        If IsArray(sheet_option_tmp.Item(parname)) Then
            write_parm(inx_col) = OptionJoin(sheet_option_tmp.Item(parname))
        Else
            write_parm(inx_col) = sheet_option_tmp.Item(parname)
        End If
    Next
    If inx_row = 0 Then inx_row = OptionSheetInx(sheetn)
    If inx_row < 0 Then
        inx_row = Abs(inx_row)
        inx_col = SheetIndex_NCol(sheetn)
        ThisWorkbook.Worksheets(inx_name).Cells(inx_row, inx_col).Formula = sheetn
    End If
    ThisWorkbook.Worksheets(inx_name).Range(ThisWorkbook.Worksheets(inx_name).Cells(inx_row, first_col), ThisWorkbook.Worksheets(inx_name).Cells(inx_row, end_col)) = write_parm
tfunctime = functime("OptionSheetWrite", tfunctime)
    OptionSheetWrite = True
End Function


Function SpecGetType(ByVal nm As String) As Long
Dim tfunctime As Double
tfunctime = Timer
    On Error Resume Next
    form = ThisWorkbook.VBProject.VBComponents("UserForm2").Name
    If IsEmpty(form) Then
        SpecGetType = 7
        Exit Function
    End If
    If Left$(nm, 1) <> "!" And Left$(nm, 1) <> "|" Then
        If InStr(nm, "_") > 0 Then
            type_spec = Split(nm, "_")
            suffix = Trim$(type_spec(UBound(type_spec)))
            If IsEmpty(spec_type_suffix) Then r = INISet()
            If spec_type_suffix.Exists(suffix) Then
                spec = spec_type_suffix.Item(suffix)
            Else
                spec = 2
            End If
            If Left(nm, Len("из архикада")) = "из архикада" And spec = 2 Then spec = 26
        Else
            spec = 3
            If InStr(nm, "Фас") > 0 Then spec = 8
            If Left(nm, Len("из архикада")) = "из архикада" Then spec = 26
        End If
    Else
        spec = 0
    End If
tfunctime = functime("SpecGetType", tfunctime)
    SpecGetType = spec
End Function

Function SpecArm(ByVal arm As Variant, ByVal n_arm As Long, ByVal type_spec As Long, ByVal nSubPos As Long) As Variant
    n_txt = get_nsubpos_txt(nSubPos)
    If this_sheet_option.Item("qty_one_subpos") = False Then nSubPos = 1
    Dim pos_out
    un_chsum_arm = ArrayUniqValColumn(arm, col_chksum)
    pos_chsum_arm = UBound(un_chsum_arm, 1)
    If type_spec = 1 Or this_sheet_option.Item("arm_pm") Or (type_spec = 3 And this_sheet_option.Item("ignore_subpos")) Then
        'Нам нужны уникальные суммы только для диаметра и класса
        'Поэтому сформируем новый массив, где от архикадовской суммы отрежем лишнее
        For i = 1 To pos_chsum_arm
            If this_sheet_option.Item("arm_pm") Then
                If this_sheet_option.Item("keep_pos") Then
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
    If type_spec = 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 1
    If type_spec <> 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 2
    ReDim pos_out(pos_chsum_arm, n_col_spec)
    For i = 1 To pos_chsum_arm
        For j = 1 To n_arm
            If type_spec = 1 Or this_sheet_option.Item("arm_pm") Or (type_spec = 3 And this_sheet_option.Item("ignore_subpos")) Then
                If type_spec = 1 Or (type_spec = 3 And this_sheet_option.Item("ignore_subpos")) Then current_chksum = Split(arm(j, col_chksum), "_")(0) & Split(arm(j, col_chksum), "_")(2)
                If this_sheet_option.Item("arm_pm") And Not this_sheet_option.Item("keep_pos") Then current_chksum = Split(arm(j, col_chksum), "_")(0)
                If this_sheet_option.Item("arm_pm") And this_sheet_option.Item("keep_pos") Then current_chksum = Split(arm(j, col_chksum), "_")(0) & Split(arm(j, col_chksum), "_")(2)
            Else
                current_chksum = arm(j, col_chksum)
            End If
            chksum = un_chsum_arm(i)
            If current_chksum = chksum Then
                klass = arm(j, col_klass)
                diametr = arm(j, col_diametr)
                weight_pm = GetWeightForDiametr(diametr, klass)
                fon = arm(j, col_fon)
                mp = arm(j, col_mp)
                gnut = arm(j, col_gnut)
                prim = " ": If arm(j, col_gnut) And Not this_sheet_option.Item("arm_pm") Then prim = "*"
                qty = arm(j, col_qty)
                n_el = qty / nSubPos
                length_pos = arm(j, col_length) / 1000
                Select Case type_spec
                Case 1
                    pos_out(i, 1) = arm(j, col_sub_pos) & n_txt
                    If (this_sheet_option.Item("keep_pos") And this_sheet_option.Item("arm_pm")) Or Not this_sheet_option.Item("arm_pm") Then
                        pos_out(i, 2) = arm(j, col_pos)
                    Else
                        pos_out(i, 2) = " "
                    End If
                    If fon Or this_sheet_option.Item("arm_pm") Then
                        l_pos = Round_w(length_pos * k_zap_total, n_round_l) * n_el
                        If prim = "п.м." Then prim = " "
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L= п.м." & prim
                        pos_out(i, 4) = pos_out(i, 4) + l_pos
                        pos_out(i, 5) = weight_pm
                    Else
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L=" & length_pos * 1000 & "мм." & prim
                        pos_out(i, 4) = pos_out(i, 4) + n_el
                        If zap_only_mp Then
                            pos_out(i, 5) = Round_w(weight_pm * length_pos, n_round_w)
                        Else
                            pos_out(i, 5) = Round_w(weight_pm * length_pos * k_zap_total, n_round_w)
                        End If
                    End If
                Case Else
                    If (this_sheet_option.Item("keep_pos") And this_sheet_option.Item("arm_pm")) Or Not this_sheet_option.Item("arm_pm") Then
                        pos_out(i, 1) = arm(j, col_pos)
                    Else
                        pos_out(i, 1) = " "
                    End If
                    pos_out(i, 2) = GetGOSTForKlass(klass)
                    If fon Or this_sheet_option.Item("arm_pm") Then
                        l_pos = Round_w(length_pos * k_zap_total, n_round_l) * n_el
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L= п.м."
                        pos_out(i, 4) = pos_out(i, 4) + l_pos
                        pos_out(i, 5) = weight_pm
                        If show_sum_prim Then
                            pos_out(i, 6) = pos_out(i, 6) + (l_pos * weight_pm)
                            pos_out(i, 3) = pos_out(i, 3) & prim
                        Else
                            pos_out(i, 6) = prim
                        End If
                    Else
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L=" & length_pos * 1000 & "мм."
                        pos_out(i, 4) = pos_out(i, 4) + n_el
                        'Запас для одиночных элементов применяется только если это включено в setting
                        If zap_only_mp Then
                            pos_out(i, 5) = Round_w(weight_pm * length_pos, n_round_w)
                        Else
                            pos_out(i, 5) = Round_w(weight_pm * length_pos * k_zap_total, n_round_w)
                        End If
                        If show_sum_prim Then
                            pos_out(i, 3) = pos_out(i, 3) & prim
                            pos_out(i, 6) = pos_out(i, 6) + n_el * pos_out(i, 5)
                        Else
                            pos_out(i, 6) = prim
                        End If
                    End If
                End Select
            End If
        Next j
    Next i
    
    For i = 1 To UBound(pos_out, 1)
        If type_spec = 13 Or this_sheet_option.Item("qty_one_floor") Then pos_out(i, 7) = t_arm
        If InStr(pos_out(i, 3), " L= п.м.") > 0 Then
            pos_out(i, 4) = ConvNum2Txt(Round_w(pos_out(i, 4), 0))
        Else
            pos_out(i, 5) = ConvNum2Txt(pos_out(i, 5), n_round_w, True)
        End If
    Next
    If type_spec = 1 Then
        n_col_pos = 2
    Else
        n_col_pos = 1
    End If
    pos_out = ArraySort(pos_out, n_col_pos)
    SpecArm = pos_out
    'erase arm, pos_out
End Function


Function SpecIzd(ByVal izd As Variant, ByVal n_izd As Long, ByVal type_spec As Long, ByVal nSubPos As Long) As Variant
    n_txt = get_nsubpos_txt(nSubPos)
    If this_sheet_option.Item("qty_one_subpos") = False Then nSubPos = 1
    un_chsum_izd = ArrayUniqValColumn(izd, col_chksum)
    pos_chsum_izd = UBound(un_chsum_izd, 1)
    If type_spec = 1 Or ((type_spec = 3 Or type_spec = 2) And this_sheet_option.Item("ignore_subpos")) Then
        For i = 1 To pos_chsum_izd
            un_chsum_izd(i) = Split(un_chsum_izd(i), "_")(0) & Split(un_chsum_izd(i), "_")(2)
        Next i
        un_chsum_izd = ArrayUniqValColumn(un_chsum_izd, 1)
        pos_chsum_izd = UBound(un_chsum_izd, 1)
    End If
    
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    If type_spec = 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 1
    If type_spec <> 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 2
    Dim pos_out(): ReDim pos_out(pos_chsum_izd, n_col_spec)
    For i = 1 To pos_chsum_izd
        For j = 1 To n_izd
            If type_spec = 1 Or ((type_spec = 3 Or type_spec = 2) And this_sheet_option.Item("ignore_subpos")) Then
                current_chksum = Split(izd(j, col_chksum), "_")(0) & Split(izd(j, col_chksum), "_")(2)
            Else
                current_chksum = izd(j, col_chksum)
            End If
            If current_chksum = un_chsum_izd(i) Then
                qty = izd(j, col_qty)
                n_el = qty / nSubPos
                'Запас для одиночных элементов применяется только если это включено в setting
                If k_zap_total > 1 Then
                    If zap_only_mp Then
                        If izd(j, col_m_edizm) = "п.м." Then n_el = (qty / nSubPos) + Round((k_zap_total - 1) * (qty / nSubPos), 0)
                    Else
                        n_el = (qty / nSubPos) + Round((k_zap_total - 1) * (qty / nSubPos), 0)
                    End If
                End If
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
                        Weight = ConvTxt2Num(Replace(izd(j, col_m_edizm), "п.м.#дерев", vbNullString)) 'теперь тут площадь сечения
                        pos_out(i, 1) = pos_out(i, 1) + "#дерев"
                        pos_out(i, 3) = naen + ", L=п.м."
                        pos_out(i, 4) = pos_out(i, 4) + n_el
                        pos_out(i, 5) = "-"
                        If IsNumeric(Weight) Then pos_out(i, 6) = Weight / 1000
                    Else
                        pos_out(i, 3) = naen
                        pos_out(i, 4) = pos_out(i, 4) + n_el
                        pos_out(i, 5) = Weight
                        If IsNumeric(Weight) And Len(izd(j, col_m_edizm)) = 0 Then
                            If show_sum_prim Then
                                If IsNumeric(pos_out(i, 6)) Then
                                    pos_out(i, 6) = pos_out(i, 6) + n_el * Weight
                                Else
                                    pos_out(i, 6) = n_el * Weight
                                End If
                            End If
                        Else
                            pos_out(i, 6) = izd(j, col_m_edizm)
                        End If
                    End If
                End Select
            End If
        Next j
    Next i
    For i = 1 To UBound(pos_out, 1)
        If type_spec = 13 Or this_sheet_option.Item("qty_one_floor") Then pos_out(i, 7) = t_izd
        If type_spec <> 1 Then
            If InStr(pos_out(i, 1), "#дерев") > 0 Then
                pos_out(i, 1) = Replace(pos_out(i, 1), "#дерев", vbNullString)
                pos_out(i, 4) = Round_w(pos_out(i, 4), 0)
                pos_out(i, 6) = ConvNum2Txt(Round_w((pos_out(i, 6) * pos_out(i, 4)), 3)) + " куб.м." 'площадь сечения на длину
            End If
        End If
        If Int(pos_out(i, 5)) - pos_out(i, 5) > 0 Then pos_out(i, 5) = ConvNum2Txt(pos_out(i, 5), n_round_l)
    Next
    
    If type_spec = 1 Then
        n_col_pos = 2
        n_col_obozn = 3
    Else
        n_col_pos = 1
        n_col_obozn = 2
    End If
    pos_out = ArraySort(pos_out, n_col_pos)
    SpecIzd = pos_out
    'erase izd, pos_out
End Function

Function get_nsubpos_txt(ByVal nSubPos As Variant) As String
    n_txt = ",**"
    If this_sheet_option.Item("qty_one_subpos") Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
    Else
        If nSubPos > 1 Then n_txt = "," & vbLf & "на все " & nSubPos & " шт."
        nSubPos = 1
    End If
    If this_sheet_option.Item("show_qty_spec") Then n_txt = ",**"
    get_nsubpos_txt = n_txt
End Function

Function SpecMaterial(ByVal mat As Variant, ByVal n_mat As Long, ByVal type_spec As Long, ByVal nSubPos As Long) As Variant
    n_txt = get_nsubpos_txt(nSubPos)
    If this_sheet_option.Item("qty_one_subpos") = False Then nSubPos = 1
    un_pos_mat = ArrayUniqValColumn(mat, col_chksum)
    pos_mat = UBound(un_pos_mat, 1)
    
    If type_spec = 1 Or (type_spec = 3 And this_sheet_option.Item("ignore_subpos")) Then
        For i = 1 To pos_mat
            un_pos_mat(i) = Split(un_pos_mat(i), "_")(0) & Split(un_pos_mat(i), "_")(2)
        Next i
        un_pos_mat = ArrayUniqValColumn(un_pos_mat, 1)
        pos_mat = UBound(un_pos_mat, 1)
    End If
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    If type_spec = 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 1
    If type_spec <> 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 2
    Dim pos_out(): ReDim pos_out(pos_mat, n_col_spec)
    For i = 1 To pos_mat
        For j = 1 To n_mat
            If type_spec = 1 Or (type_spec = 3 And this_sheet_option.Item("ignore_subpos")) Then
                current_chksum = Split(mat(j, col_chksum), "_")(0) & Split(mat(j, col_chksum), "_")(2)
            Else
                current_chksum = mat(j, col_chksum)
            End If
            If current_chksum = un_pos_mat(i) Then
                gost = mat(j, col_m_obozn)
                If Len(swap_gost.Item(gost)) > 0 Then gost = swap_gost.Item(gost)
                Select Case type_spec
                Case 1
                    pos_out(i, 1) = mat(j, col_sub_pos) & n_txt
                    pos_out(i, 2) = " "
                    pos_out(i, 3) = mat(j, col_m_naen) & " по " & gost & ", " & mat(j, col_m_edizm)
                    pos_out(i, 4) = pos_out(i, 4) + (Round_w(mat(j, col_qty) * k_zap_total, n_round_mat) / nSubPos)
                    pos_out(i, 5) = "-"
                Case Else
                    pos_out(i, 1) = " "
                    pos_out(i, 2) = gost
                    pos_out(i, 3) = mat(j, col_m_naen)
                    pos_out(i, 4) = pos_out(i, 4) + (Round_w(mat(j, col_qty) * k_zap_total, n_round_mat) / nSubPos)
                    pos_out(i, 5) = "-"
                    pos_out(i, 6) = mat(j, col_m_edizm)
                End Select
            End If
        Next j
    Next i
    For i = 1 To UBound(pos_out, 1)
        If type_spec = 13 Or this_sheet_option.Item("qty_one_floor") Then pos_out(i, 7) = t_mat
        pos_out(i, 4) = ConvNum2Txt(Round_w(pos_out(i, 4), n_round_mat), n_round_mat)
    Next
    If type_spec = 1 Then
        n_col_pos = 2
    Else
        n_col_pos = 1
    End If
    pos_out = ArraySort(pos_out, n_col_pos)
    SpecMaterial = pos_out
    'erase mat, un_pos_mat, pos_out
End Function

Function SpecOneSubpos(ByVal all_data As Variant, ByVal subpos As String, ByVal type_spec As Long, ByVal floor_txt As String) As Variant
    If IsEmpty(all_data) Then
        SpecOneSubpos = Empty
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
'r = functime("ArrayCombine", tfunctime)
    nSubPos = GetNSubpos(subpos, type_spec, floor_txt)
    If Not this_sheet_option.Item("qty_one_subpos") Then nSubPos = 1
    If (type_spec = 3 And this_sheet_option.Item("ignore_subpos")) Then nSubPos = 1
    'Добавляем загаловок для сборки
    Dim pos_naen
    If this_sheet_option.Item("add_row") Then
        n_n = 2
    Else
        n_n = 1
    End If
    sb_naen = "@"
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    If type_spec = 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 1
    If type_spec <> 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 2
    If type_spec = 2 Then
        ReDim pos_naen(n_n, n_col_spec)
        If subpos <> "-" Then
            naen = subpos
            If pos_data.Item(floor_txt).Item("name").Count > 0 Then
                If pos_data.Item(floor_txt).Item("name").Exists(subpos) Then naen = pos_data.Item(floor_txt).Item("name").Item(subpos)(1)
                If this_sheet_option.Item("qty_one_subpos") Then
                    pos_naen(n_n, 1) = fin_str & naen
                    If this_sheet_option.Item("qty_one_floor") Then
                        pos_naen(n_n, 6) = nSubPos
                    Else
                        If nSubPos > 1 Then
                            pos_naen(n_n, 1) = pos_naen(n_n, 1) & ", на 1 шт. (всего " & nSubPos & " шт.)"
                        Else
                            pos_naen(n_n, 1) = pos_naen(n_n, 1) & ",**"
                        End If
                    End If
                Else
                    pos_naen(n_n, 1) = fin_str & naen
                    If this_sheet_option.Item("qty_one_floor") Then
                        pos_naen(n_n, 6) = -1
                    Else
                        If nSubPos > 1 Then
                            pos_naen(n_n, 1) = pos_naen(n_n, 1) & ", на все " & nSubPos & " шт."
                        Else
                            pos_naen(n_n, 1) = pos_naen(n_n, 1) & ",**"
                        End If
                    End If
                End If
                If this_sheet_option.Item("show_qty_spec") Then
                    pos_naen(n_n, 1) = fin_str & naen
                    If this_sheet_option.Item("qty_one_floor") Then
                        pos_naen(n_n, 6) = -1
                    Else
                        pos_naen(n_n, 1) = pos_naen(n_n, 1) & ",**"
                    End If
                End If
            End If
        Else
            pos_naen(n_n, 1) = fin_str & " Прочие элементы,**"
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
                    u2 = (pos_data.Item(floor_txt).Item("-").Exists(сurrent_subpos) And (сurrent_parent = "-") And (сurrent_type_el = t_subpos))   'Элементы вложенных сборок
                    usl = u1 Or u2
                Else
                    u1 = ((сurrent_parent = "-") And (сurrent_subpos = subpos) And (сurrent_type_el <> t_subpos)) 'Элементы главной сборки
                    u2 = (сurrent_parent = subpos) And (сurrent_type_el = t_subpos) 'Маркировка вложенных сборок
                    usl = (u1 Or u2)
                End If
            'Общестроительная
            'Только наименование сборки и все элементы без сборок
            Case 3
                If this_sheet_option.Item("ignore_subpos") Then
                    usl = (сurrent_type_el <> t_subpos)
                Else
                    u1 = (сurrent_subpos = "-")
                    u2 = ((сurrent_parent = "-") And (сurrent_type_el = t_subpos) And this_sheet_option.Item("show_subpos"))
                    usl = (u1 Or u2)
                End If
        End Select
        If usl Then
            Select Case сurrent_type_el
                Case t_arm And this_sheet_option.Item("isarm")
                    n_arm = n_arm + 1
                    For j = 1 To max_col
                        arm(n_arm, j) = all_data(i, j)
                    Next j
                Case t_prokat And this_sheet_option.Item("isprok")
                    n_prokat = n_prokat + 1
                    For j = 1 To max_col
                        prokat(n_prokat, j) = all_data(i, j)
                    Next j
                Case t_mat And this_sheet_option.Item("ismat")
                    n_mat = n_mat + 1
                    For j = 1 To max_col
                        mat(n_mat, j) = all_data(i, j)
                    Next j
                Case t_izd And this_sheet_option.Item("isizd")
                    n_izd = n_izd + 1
                    For j = 1 To max_col
                        izd(n_izd, j) = all_data(i, j)
                    Next j
                    If izd(n_izd, col_m_weight) = "-" Then izd(n_izd, col_m_weight) = 0
                Case t_subpos And this_sheet_option.Item("issubpos")
                        n_subpos = n_subpos + 1
                        For j = 1 To max_col
                            subp(n_subpos, j) = all_data(i, j)
                        Next j
                        If pos_data.Item(floor_txt).Item("weight").Item(сurrent_subpos) > 0 Then
                            subp(n_subpos, col_m_weight) = pos_data.Item(floor_txt).Item("weight").Item(сurrent_subpos)
                        End If
                End Select
        End If
    Next
    
    ReDim pos_naen(n_n, n_col_spec)
    If n_subpos > 0 Then
        'subp = ArrayRedim(subp, n_subpos)
        pos_naen(n_n, 3) = type_el_name.Item(t_subpos)
        If this_sheet_option.Item("qty_one_floor") And type_spec <> 13 Then pos_naen(n_n, 1) = fin_str
        If type_spec = 13 Or this_sheet_option.Item("qty_one_floor") Then
            pos_naen(n_n, 7) = t_subpos
            pos_naen(1, 7) = pos_naen(n_n, 7)
        End If
        If type_spec <> 1 And this_sheet_option.Item("show_type") Then pos_out = ArrayCombine(pos_out, pos_naen)
        pos_out = ArrayCombine(pos_out, SpecSubpos(subp, n_subpos, type_spec, nSubPos, floor_txt))
    End If
    
    If n_arm > 0 Then
        'arm = ArrayRedim(arm, n_arm)
        pos_naen(n_n, 3) = type_el_name.Item(t_arm)
        If this_sheet_option.Item("qty_one_floor") And type_spec <> 13 Then pos_naen(n_n, 1) = fin_str
        If type_spec = 13 Or this_sheet_option.Item("qty_one_floor") Then
            pos_naen(n_n, 7) = t_arm
            pos_naen(1, 7) = pos_naen(n_n, 7)
        End If
        If type_spec <> 1 And this_sheet_option.Item("show_type") Then pos_out = ArrayCombine(pos_out, pos_naen)
        pos_out = ArrayCombine(pos_out, SpecArm(arm, n_arm, type_spec, nSubPos))
    End If

    If n_prokat > 0 Then
        'prokat = ArrayRedim(prokat, n_prokat)
        pos_naen(n_n, 3) = type_el_name.Item(t_prokat)
        If this_sheet_option.Item("qty_one_floor") And type_spec <> 13 Then pos_naen(n_n, 1) = fin_str
        If type_spec = 13 Or this_sheet_option.Item("qty_one_floor") Then
            pos_naen(n_n, 7) = t_prokat
            pos_naen(1, 7) = pos_naen(n_n, 7)
        End If
        If type_spec <> 1 And this_sheet_option.Item("show_type") Then pos_out = ArrayCombine(pos_out, pos_naen)
        pos_out = ArrayCombine(pos_out, SpecProkat(prokat, n_prokat, type_spec, nSubPos))
    End If
    
    If n_izd > 0 Then
        'izd = ArrayRedim(izd, n_izd)
        pos_naen(n_n, 3) = type_el_name.Item(t_izd)
        If this_sheet_option.Item("qty_one_floor") And type_spec <> 13 Then pos_naen(n_n, 1) = fin_str
        If type_spec = 13 Or this_sheet_option.Item("qty_one_floor") Then
            pos_naen(n_n, 7) = t_izd
            pos_naen(1, 7) = pos_naen(n_n, 7)
        End If
        If type_spec <> 1 And this_sheet_option.Item("show_type") Then pos_out = ArrayCombine(pos_out, pos_naen)
        pos_out = ArrayCombine(pos_out, SpecIzd(izd, n_izd, type_spec, nSubPos))
    End If

    If n_mat > 0 Then
        'mat = ArrayRedim(mat, n_mat)
        pos_naen(n_n, 3) = type_el_name.Item(t_mat)
        If this_sheet_option.Item("qty_one_floor") And type_spec <> 13 Then pos_naen(n_n, 1) = fin_str
        If type_spec = 13 Or this_sheet_option.Item("qty_one_floor") Then
            pos_naen(n_n, 7) = t_mat
            pos_naen(1, 7) = pos_naen(n_n, 7)
        End If
        If type_spec <> 1 And this_sheet_option.Item("show_type") Then pos_out = ArrayCombine(pos_out, pos_naen)
        pos_out = ArrayCombine(pos_out, SpecMaterial(mat, n_mat, type_spec, nSubPos))
    End If
    
    If IsEmpty(pos_out) Or n_subpos + n_izd + n_prokat + n_arm + n_mat = 0 Then
        SpecOneSubpos = Empty
    Else
        Select Case type_spec
            Case 1
                If Not this_sheet_option.Item("show_type") Then
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
                If this_sheet_option.Item("qty_one_subpos") Then
                    subpos_we_group = pos_out(1, 6) / nSubPos
                    subpos_we_spec = pos_data.Item(floor_txt).Item("weight").Item(subpos) / nSubPos
                Else
                    subpos_we_group = pos_out(1, 6) / nSubPos
                    subpos_we_spec = pos_data.Item(floor_txt).Item("weight").Item(subpos) * GetNSubpos(subpos, type_spec, floor_txt)
                End If
                If Abs(subpos_we_group - subpos_we_spec) > 0.01 Then
                    r = LogWrite(lastfilespec, subpos, "Небивка массы на " & str(subpos_we_group - subpos_we_spec) & " груп=" & str(subpos_we_group) & ", общая=" & str(subpos_we_spec))
                End If
                If subpos_we_group <= 0.01 Then
                    r = LogWrite(lastfilespec, subpos, "Проверьте вес " & str(subpos_we_group))
                End If
                If subpos_we_spec <= 0.01 Then
                    r = LogWrite(lastfilespec, subpos, "Проверьте вес " & str(subpos_we_spec))
                End If
                For i = 1 To UBound(pos_out, 1)
                    pos_out(1, 6) = Round_w(pos_out(1, 6), n_round_w)
                Next i
            Case Else
                If Not this_sheet_option.Item("show_type") Then
                    'Если мы игнорируем сборки - то сортировка будет в общей спецификации, не будем тратить время зря
                    If this_sheet_option.Item("ignore_subpos") = False Then
                        pos_out_sort = ArraySort_2(pos_out, Array(1, 2, 3), 2)
                    Else
                        pos_out_sort = pos_out
                    End If
                    If sb_naen = "@" Then
                        n_row = 0
                    Else
                        n_row = n_n
                    End If
                    For i = 1 To UBound(pos_out, 1)
                        If pos_out_sort(i, 1) <> "Поз." And pos_out_sort(i, 1) <> sb_naen And pos_out_sort(i, 3) <> vbNullString And Not IsEmpty(pos_out_sort(i, 3)) Then
                            n_row = n_row + 1
                            For j = 1 To UBound(pos_out, 2)
                                pos_out(n_row, j) = pos_out_sort(i, j)
                            Next j
                        Else
                            hh = 1
                        End If
                    Next i
                    If n_row <> UBound(pos_out, 1) Then pos_out = ArrayRedim(pos_out, n_row)
                End If
        End Select
        If type_spec <> 13 And this_sheet_option.Item("qty_one_floor") And Not IsEmpty(pos_out) Then
            For i = 1 To UBound(pos_out, 1)
                pos_out(i, 8) = nSubPos
            Next i
        End If
        SpecOneSubpos = pos_out
        'If Not IsEmpty(pos_out) Then erase pos_out
    End If
r = functime("SpecOneSubpos", tfunctime)
End Function

Function SpecProkat(ByVal prokat As Variant, ByVal n_prokat As Long, ByVal type_spec As Long, Optional ByVal nSubPos As Long = 1) As Variant
    n_txt = get_nsubpos_txt(nSubPos)
    If this_sheet_option.Item("qty_one_subpos") = False Then nSubPos = 1
    un_chsum_prokat = ArrayUniqValColumn(prokat, col_chksum)
    pos_chsum_prokat = UBound(un_chsum_prokat, 1)
    If type_spec = 1 Or this_sheet_option.Item("pr_pm") Or (type_spec = 3 And this_sheet_option.Item("ignore_subpos")) Then
        For i = 1 To pos_chsum_prokat
            If this_sheet_option.Item("pr_pm") Then
                If this_sheet_option.Item("keep_pos") Then
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
    If type_spec = 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 1
    If type_spec <> 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 2
    ReDim pos_out(pos_chsum_prokat, n_col_spec)
    For i = 1 To pos_chsum_prokat
        For j = 1 To n_prokat
            If type_spec = 1 Or this_sheet_option.Item("pr_pm") Or (type_spec = 3 And this_sheet_option.Item("ignore_subpos")) Then
                If type_spec = 1 Or (type_spec = 3 And this_sheet_option.Item("ignore_subpos")) Then current_chksum = Split(prokat(j, col_chksum), "_")(0) & Split(prokat(j, col_chksum), "_")(2)
                If this_sheet_option.Item("pr_pm") And Not this_sheet_option.Item("keep_pos") Then current_chksum = Split(prokat(j, col_chksum), "_")(0)
                If this_sheet_option.Item("pr_pm") And this_sheet_option.Item("keep_pos") Then current_chksum = Split(prokat(j, col_chksum), "_")(0) & Split(prokat(j, col_chksum), "_")(2)
            Else
                current_chksum = prokat(j, col_chksum)
            End If
            If current_chksum = un_chsum_prokat(i) Then
                name_pr = GetShortNameForGOST(prokat(j, col_pr_gost_prof))
                If InStr(prokat(j, col_pr_naen), "@@") > 0 Then prokat(j, col_pr_naen) = Split(prokat(j, col_pr_naen), "@@")(0)
                pm = False: If InStr(prokat(j, col_chksum), "lpm") > 0 Then pm = True
                n_el = prokat(j, col_qty) / nSubPos
                If (n_el = 0) Or IsEmpty(prokat(j, col_qty)) Then n_el = 1
                If InStr(1, name_pr, "Лист") Then
                    naen_plate = SpecMetallPlate(prokat(j, col_pr_prof), prokat(j, col_pr_naen), prokat(j, col_pr_length), prokat(j, col_pr_weight), prokat(j, col_chksum))
                    we = naen_plate(4)
                    L = naen_plate(2)
                Else
                    L = Round_w(prokat(j, col_pr_length) / 1000, n_round_l)
                    we = prokat(j, col_pr_weight) * L
                End If
                If this_sheet_option.Item("pr_pm") Or pm Then
                    we = Round(Round_w(we * k_zap_total, n_round_w) / L, n_round_w)
                Else
                    'Запас для одиночных элементов применяется только если это включено в setting
                    If zap_only_mp Then
                        we = Round_w(we, n_round_w)
                    Else
                        we = Round_w(we * k_zap_total, n_round_w)
                    End If
                End If
                gost = prokat(j, col_pr_gost_prof)
                If Len(swap_gost.Item(gost)) > 0 Then gost = swap_gost.Item(gost)
                Select Case type_spec
                    Case 1
                        If this_sheet_option.Item("pr_pm") Or pm Then
                            pos_out(i, 1) = prokat(j, col_sub_pos) & n_txt
                            If this_sheet_option.Item("keep_pos") Or pm Then
                                pos_out(i, 2) = prokat(j, col_pos)
                            Else
                                pos_out(i, 2) = " "
                            End If
                            If InStr(1, name_pr, "Лист") Then
                                pos_out(i, 3) = name_pr & gost & " " & naen_plate(1)
                            Else
                                pos_out(i, 3) = name_pr & gost & " " & prokat(j, col_pr_prof) & " L = п.м."
                            End If
                            pos_out(i, 4) = pos_out(i, 4) + (L * n_el)
                            pos_out(i, 5) = we
                        Else
                            pos_out(i, 1) = prokat(j, col_sub_pos) & n_txt
                            pos_out(i, 2) = prokat(j, col_pos)
                            If InStr(1, name_pr, "Лист") Then
                                pos_out(i, 3) = name_pr & gost & " " & naen_plate(1)
                            Else
                                pos_out(i, 3) = name_pr & gost & " " & prokat(j, col_pr_prof) & " L=" & L * 1000 & "мм."
                            End If
                            pos_out(i, 4) = pos_out(i, 4) + n_el
                            pos_out(i, 5) = we
                        End If
                    Case Else
                        If this_sheet_option.Item("pr_pm") Or pm Then
                            If this_sheet_option.Item("keep_pos") Or pm Then
                                pos_out(i, 1) = prokat(j, col_pos)
                            Else
                                pos_out(i, 1) = " "
                            End If
                            pos_out(i, 2) = gost
                            If InStr(1, name_pr, "Лист") Then
                                pos_out(i, 3) = name_pr & " " & naen_plate(1)
                            Else
                                pos_out(i, 3) = name_pr & " " & prokat(j, col_pr_prof) & " L = п.м."
                            End If
'TODO Лист при прокате в п.м. выдаёт и площадь и наименование
                            pos_out(i, 4) = pos_out(i, 4) + (L * n_el)
                            pos_out(i, 5) = we
                            If show_sum_prim Then pos_out(i, 6) = pos_out(i, 6) + Round_w(L * we, n_round_w) * n_el
                        Else
                            pos_out(i, 1) = prokat(j, col_pos)
                            pos_out(i, 2) = gost
                            If InStr(1, name_pr, "Лист") Then
                                pos_out(i, 3) = name_pr & " " & naen_plate(1)
                            Else
                                pos_out(i, 3) = name_pr & prokat(j, col_pr_prof) & " L=" & L * 1000 & "мм."
                            End If
                            pos_out(i, 4) = pos_out(i, 4) + n_el
                            pos_out(i, 5) = we
                            If show_sum_prim Then pos_out(i, 6) = pos_out(i, 6) + n_el * pos_out(i, 5)
                        End If
                End Select
            End If
        Next j
    Next i
    
    For i = 1 To UBound(pos_out, 1)
        If type_spec = 13 Or this_sheet_option.Item("qty_one_floor") Then pos_out(i, 7) = t_prokat
        If Int(pos_out(i, 4)) - pos_out(i, 4) > 0 Then
            pos_out(i, 4) = ConvNum2Txt(pos_out(i, 4), n_round_l)
        Else
            pos_out(i, 5) = ConvNum2Txt(pos_out(i, 5), n_round_w)
        End If
    Next
    If type_spec = 1 Then
        n_col_pos = 2
    Else
        n_col_pos = 1
    End If
    pos_out = ArraySort(pos_out, n_col_pos)
    SpecProkat = pos_out
    'erase prokat, pos_out
End Function

Function SpecMetallPlate(ByVal prokat_prof As String, ByVal prokat_naen As String, ByVal area_plate As Variant, ByVal weight_plate As Variant, ByVal checksum As String) As Variant
'TODO Добавить обработку п.м. и кв.м.
    pm = False: If InStr(checksum, "lpm") > 0 Then pm = True
    If this_sheet_option.Item("pr_pm") Then pm = True
    If IsNumeric(ConvTxt2Num(area_plate)) Then
        area_plate = area_plate / 1000
    Else
        area_plate = 0
    End If
    If Not IsNumeric(ConvTxt2Num(weight_plate)) Then
        weight_plate = 0
    End If
    Dim array_out(): ReDim array_out(7)
    prokat_naen_t = prokat_naen
    prokat_prof = Replace(prokat_prof, " ", vbNullString)
    prokat_prof = Replace(prokat_prof, "-", vbNullString)
    prokat_prof = Trim$(prokat_prof)
    prokat_naen = Replace(prokat_naen, "Лист", vbNullString)
    prokat_naen = Replace(prokat_naen, " ", vbNullString)
    prokat_naen = Replace(prokat_naen, "-", vbNullString)
    prokat_naen = Replace(prokat_naen, "X", "*")
    prokat_naen = Replace(prokat_naen, "x", "*")
    prokat_naen = Replace(prokat_naen, "Х", "*")
    prokat_naen = Replace(prokat_naen, "х", "*")
    prokat_naen = Trim$(prokat_naen)
    t_list = ConvTxt2Num(prokat_prof)
    If IsNumeric(t_list) Then t_list = t_list / 1000
    flag_read = False
    a = 0: b = 0: t = 0: S = 0
    
    'Если в наименовании указаны размеры листа - попробуем вытащить их оттуда
    abc = Split(prokat_naen, "*")
    If UBound(abc) = 2 Then
        flag_read = True
        a = 0: b = 0: t = 100000: S = 0
        For nn = 0 To UBound(abc)
            k = ConvTxt2Num(abc(nn))
            If IsNumeric(k) Then
                k = k / 1000
                If k > a Then a = k
                If k < t Then t = k
                S = S + k
            End If
        Next nn
        b = S - a - t
        b = Round(b, 3)
        a = Round(a, 3)
        t = Round(t, 3)
        prokat_prof = "--" + ConvNum2Txt(t * 1000)
        prokat_naen = "--" + ConvNum2Txt(t * 1000) + "x" + ConvNum2Txt(b * 1000) + "x" + ConvNum2Txt(a * 1000)
        we_plate = t * 7850
        area_plate = b * a
        we_plate_one = we_plate * area_plate
        If pm Then
            naen_plate = prokat_prof & " S=кв.м."
        Else
            naen_plate = prokat_naen
        End If
        If IsNumeric(t_list) Then
            If t <> t_list Then
                MsgBox ("Разница толщины листа в наименовании и типе профиля " + prokat_naen + "<>" + prokat_prof)
                flag_read = False
            End If
        End If
    End If
    
    If IsNumeric(t_list) And t_list > 1 / 10000 And flag_read = False And area_plate > 0 Then
        t = t_list
        we_plate_one = t * weight_plate
        tarea_plate = weight_plate / t
        prokat_prof = "--" + ConvNum2Txt(t * 1000)
        prokat_naen = "--" + ConvNum2Txt(t * 1000)
        If pm Then
            naen_plate = prokat_prof & " S=кв.м."
        Else
            naen_plate = prokat_naen & " S=" & area_plate & "кв.м."
        End If
        a = Sqr(area_plate)
        b = a
        flag_read = True
    End If
    
    
    If flag_read = False Then
        MsgBox ("Ошибка в имени типа профиля листа " + prokat_prof)
        array_out(1) = "ОШИБКА ТИПА ПРОФИЛЯ ЛИСТА"
        array_out(2) = 0.001
        array_out(3) = 0.001
        array_out(4) = 0.001
        array_out(5) = 0.001
        array_out(6) = 0.001
        array_out(7) = 0.001
        SpecMetallPlate = array_out
        Exit Function
    End If
    array_out(1) = naen_plate 'Имя листа
    array_out(2) = area_plate 'Площадь
    array_out(3) = we_plate 'Вес кв.м.
    array_out(4) = we_plate_one 'Вес единицы
    array_out(5) = a 'бОльшая сторона
    array_out(6) = b 'меньшая сторона
    array_out(7) = t 'толщина
    SpecMetallPlate = array_out
End Function

Function SpecSubpos(ByVal subp As Variant, ByVal n_subp As Long, ByVal type_spec As Long, ByVal nSubPos As Long, ByVal floor_txt As String) As Variant
    n_txt = get_nsubpos_txt(nSubPos)
    If this_sheet_option.Item("qty_one_subpos") = False Then nSubPos = 1
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
    If type_spec = 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 1
    If type_spec <> 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 2
    Dim pos_out(): ReDim pos_out(pos_chsum_subp, n_col_spec)
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
                    If pos = naen Then pos_out(i, 1) = " "
                Case Else
                    obozn = subp(j, col_m_obozn)
                    naen = subp(j, col_m_naen)
                    If InStr(naen, "!!!") <> 0 Or InStr(obozn, "!!!") <> 0 Then
                        If pos_data.Item(floor_txt).Item("name").Exists(pos) Then
                            naen = pos_data.Item(floor_txt).Item("name").Item(pos)(1)
                            obozn = pos_data.Item(floor_txt).Item("name").Item(pos)(2)
                        End If
                    End If
                    pos_out(i, 1) = pos
                    pos_out(i, 2) = obozn
                    pos_out(i, 3) = naen
                    pos_out(i, 4) = pos_out(i, 4) + n_el
                    pos_out(i, 5) = Weight
                    If show_sum_prim And IsNumeric(Weight) Then pos_out(i, 6) = pos_out(i, 6) + n_el * Weight
                    If pos = naen Then pos_out(i, 1) = " "
                End Select
            End If
        Next j
    Next i
    If type_spec = 13 Or this_sheet_option.Item("qty_one_floor") Then
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
    'erase subp, pos_out
End Function

Function Spec_AS(ByRef all_data As Variant, ByVal type_spec As Long) As Variant
    n_col_spec = 6
    n_col_end = 4
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    If type_spec = 13 And this_sheet_option.Item("qty_one_floor") Then n_col_spec = n_col_spec + 1
    If type_spec <> 13 And this_sheet_option.Item("qty_one_floor") Then
        n_col_spec = n_col_spec + 2
        Set pos_out_dic = CreateObject("Scripting.Dictionary")
        n_row_out = 0
    End If
    Dim pos_out(): ReDim pos_out(1, n_col_spec)
    If IsEmpty(all_data) Then Spec_AS = Empty: Exit Function
    all_data = ArraySelectParam_2(all_data, Array(t_arm, t_prokat, t_mat, t_izd, t_subpos), col_type_el)
    nfloor = 1
    floor_txt = "all_floor"
    If spec_version > 1 And this_sheet_option.Item("qty_one_floor") Then 'Учтём кол-во этажей
        nfloor = UBound(floor_txt_arr, 1)
        all_data_allfloor = all_data
    End If
    qty_parent = UBound(pos_data.Item(floor_txt).Item("parent").keys()) + 1
    qty_child = UBound(pos_data.Item(floor_txt).Item("child").keys()) + 1
    qty_empty = pos_data.Item(floor_txt).Exists("-")
    For inxfloor = 1 To nfloor
        If spec_version > 1 And this_sheet_option.Item("qty_one_floor") Then
            t_floor = floor_txt_arr(inxfloor, 2)
            floor_txt = floor_txt_arr(inxfloor, 3)
            all_data = ArraySelectParam_2(all_data_allfloor, t_floor, col_floor)
            qty_parent_floor = UBound(pos_data.Item(floor_txt).Item("parent").keys()) + 1
            qty_child_floor = UBound(pos_data.Item(floor_txt).Item("child").keys()) + 1
            qty_empty_floor = pos_data.Item(floor_txt).Exists("-")
        Else
            qty_parent_floor = qty_parent
            qty_child_floor = qty_child
            qty_empty_floor = qty_empty
        End If
        If qty_parent < 0 And qty_child < 0 And (type_spec = 2 Or type_spec = 13) Then
            r = LogWrite("Ошибка спецификации", vbNullString, "Сборки отсутвуют. Создана общестроительная спецификаця")
            MsgBox ("Сборки отсутвуют. Создана общестроительная спецификаця")
            type_spec = 3
        End If
        If type_spec = 13 And ((qty_parent <= 1) Or (qty_parent < 1 And qty_empty)) Then
            MsgBox ("Сборок меньше двух. Создана общестроительная спецификаця")
            r = LogWrite("Ошибка спецификации", vbNullString, "Сборок меньше двух. Создана общестроительная спецификаця")
            Set pos_out_dic = CreateObject("Scripting.Dictionary")
            type_spec = 3
        End If
        If inxfloor = 1 Then
            Select Case type_spec
                Case 1
                    end_col = 5
                    pos_out(1, 1) = "Марка" & vbLf & "изделия."
                    pos_out(1, 2) = "Поз." & vbLf & "дет."
                    pos_out(1, 3) = "Наименование"
                    pos_out(1, 4) = "Кол-во*"
                    If this_sheet_option.Item("qty_one_subpos") Then
                        pos_out(1, 6) = "Масса изделия, кг."
                        pos_out(1, 5) = "Масса 1 дет., кг."
                    Else
                        pos_out(1, 6) = "Масса изделий, кг."
                        pos_out(1, 5) = "Масса, кг."
                    End If
                Case 13
                    end_col = 6 + qty_parent
                    If pos_data.Item("all_floor").Exists("-") Then end_col = end_col + 1
                    ReDim pos_out(2, end_col)
                    pos_out(1, 1) = "Поз."
                    pos_out(1, 2) = "Обозначение"
                    pos_out(1, 3) = "Наименование"
                    If this_sheet_option.Item("qty_one_subpos") Then
                        pos_out(1, 4) = "Кол-во на 1 шт."
                    Else
                        pos_out(1, 4) = "Кол-во на все"
                    End If
                    pos_out(1, end_col - 2) = "Всего"
                    pos_out(1, end_col - 1) = "Масса ед., кг."
                    pos_out(1, end_col) = "Примечание"
                    If this_sheet_option.Item("qty_one_floor") Then pos_out(1, end_col - 2) = "Всего на отм."
                    Dim pos_out_floor
                    Dim pos_out_subpos
                    Dim pos_out_arm
                    Dim pos_out_prokat
                    Dim pos_out_izd
                    Dim pos_out_mat
                    Dim n_row_subpos As Long
                    Dim n_row_arm As Long
                    Dim n_row_prokat As Long
                    Dim n_row_izd As Long
                    Dim n_row_mat As Long
                Case Else
                    end_col = 6
                    pos_out(1, 1) = "Поз."
                    pos_out(1, 2) = "Обозначение"
                    pos_out(1, 3) = "Наименование"
                    pos_out(1, 4) = "Кол-во"
                    pos_out(1, end_col - 1) = "Масса ед., кг."
                    pos_out(1, end_col) = "Примечание"
            End Select
            pos_out_up = pos_out
        End If
        If this_sheet_option.Item("qty_one_floor") And type_spec = 13 And Not IsEmpty(all_data) Then
            ReDim pos_out_floor(1, end_col)
            pos_out_floor(1, 1) = "Элементы на отм. " + Replace(ConvNum2Otm(t_floor), "'", vbNullString)
        End If
        Dim ch_key As String
        ch_key = "child"
        If qty_child_floor <= 0 And ((type_spec = 1) Or (type_spec = 2)) Then
            If qty_parent_floor >= 0 Then
                ch_key = "parent"
            Else
                ll = 1
            End If
        End If
        If type_spec = 1 And Not IsEmpty(all_data) Then
            Dim pos_end: ReDim pos_end(1, 6)
            If this_sheet_option.Item("qty_one_subpos") Then
                pos_end(1, 1) = Space(60) & "* расход на одно изделие"
            Else
                pos_end(1, 1) = Space(60) & "* расход на все изделия"
            End If
            subpos_arr = pos_data.Item(floor_txt).Item(ch_key).keys()
            If UBound(subpos_arr) - LBound(subpos_arr) + 1 > 0 Then
                For Each subpos In ArraySort(pos_data.Item(floor_txt).Item(ch_key).keys(), 1)
                    pos_out_onesubpos = SpecOneSubpos(all_data, subpos, type_spec, floor_txt)
                    If delim_group_ved Then
                        If UBound(pos_out, 1) > 1 Then pos_out_onesubpos = ArrayCombine(pos_out_up, pos_out_onesubpos)
                        pos_out_onesubpos = ArrayCombine(pos_out_onesubpos, pos_end)
                        pos_out = ArrayCombine(pos_out, pos_out_onesubpos)
                    Else
                        pos_out = ArrayCombine(pos_out, SpecOneSubpos(all_data, subpos, type_spec, floor_txt))
                    End If
                Next
                If Not delim_group_ved Then pos_out = ArrayCombine(pos_out, pos_end)
            Else
                MsgBox ("Сборки отсутвуют.")
            End If
        End If
        
        If type_spec = 2 And Not IsEmpty(all_data) Then
            If qty_parent_floor > 0 Then
                For Each subpos In ArraySort(pos_data.Item(floor_txt).Item("parent").keys(), 1)
                    pos_out_onesubpos = SpecOneSubpos(all_data, subpos, type_spec, floor_txt)
                    If Not IsEmpty(pos_out_onesubpos) Then
                        If this_sheet_option.Item("qty_one_floor") Then
                            If Not pos_out_dic.Exists(subpos) Then Set pos_out_dic.Item(subpos) = CreateObject("Scripting.Dictionary")
                            For i = 1 To UBound(pos_out_onesubpos, 1)
                                type_el = CStr(pos_out_onesubpos(i, 7))
                                If Not pos_out_dic.Item(subpos).Exists(type_el) Then Set pos_out_dic.Item(subpos).Item(type_el) = CreateObject("Scripting.Dictionary")
                                pos = CStr(pos_out_onesubpos(i, 1))
                                obozn = CStr(pos_out_onesubpos(i, 2))
                                naen = CStr(pos_out_onesubpos(i, 3))
                                ves = CStr(pos_out_onesubpos(i, 5))
                                pos_out_onesubpos(i, 7) = pos_out_onesubpos(i, 8)
                                pos_out_onesubpos(i, 8) = floor_txt
                                row_key = pos + "%" + obozn + "%" + naen + "%" + ves
                                row_type = ArrayRow(pos_out_onesubpos, i, True)
                                If Not pos_out_dic.Item(subpos).Item(type_el).Exists(row_key) Then
                                    pos_out_dic.Item(subpos).Item(type_el).Item(row_key) = row_type
                                Else
                                    pos_out_dic.Item(subpos).Item(type_el).Item(row_key) = ArrayCombine(pos_out_dic.Item(subpos).Item(type_el).Item(row_key), row_type)
                                End If
                                n_row_out = n_row_out + 1
                            Next i
                        End If
                    pos_out = ArrayCombine(pos_out, pos_out_onesubpos)
                    End If
                Next
            End If
            subpos = "-"
            If pos_data.Item(floor_txt).Exists(subpos) Then
                pos_out_onesubpos = SpecOneSubpos(all_data, subpos, type_spec, floor_txt)
                If Not IsEmpty(pos_out_onesubpos) Then
                    If this_sheet_option.Item("qty_one_floor") Then
                        If Not pos_out_dic.Exists(subpos) Then Set pos_out_dic.Item(subpos) = CreateObject("Scripting.Dictionary")
                        For i = 1 To UBound(pos_out_onesubpos, 1)
                            type_el = CStr(pos_out_onesubpos(i, 7))
                            If Not pos_out_dic.Item(subpos).Exists(type_el) Then Set pos_out_dic.Item(subpos).Item(type_el) = CreateObject("Scripting.Dictionary")
                            pos = CStr(pos_out_onesubpos(i, 1))
                            obozn = CStr(pos_out_onesubpos(i, 2))
                            naen = CStr(pos_out_onesubpos(i, 3))
                            ves = CStr(pos_out_onesubpos(i, 5))
                            pos_out_onesubpos(i, 7) = pos_out_onesubpos(i, 8)
                            pos_out_onesubpos(i, 8) = floor_txt
                            row_key = pos + "%" + obozn + "%" + naen + "%" + ves
                            row_type = ArrayRow(pos_out_onesubpos, i, True)
                            If Not pos_out_dic.Item(subpos).Item(type_el).Exists(row_key) Then
                                pos_out_dic.Item(subpos).Item(type_el).Item(row_key) = row_type
                            Else
                                pos_out_dic.Item(subpos).Item(type_el).Item(row_key) = ArrayCombine(pos_out_dic.Item(subpos).Item(type_el).Item(row_key), row_type)
                            End If
                            n_row_out = n_row_out + 1
                        Next i
                    End If
                    pos_out = ArrayCombine(pos_out, pos_out_onesubpos)
                End If
            End If
        End If
        
        If type_spec = 3 And Not IsEmpty(all_data) Then
            If (pos_data.Item(floor_txt).Exists("-") Or (this_sheet_option.Item("show_subpos") And (UBound(pos_data.Item(floor_txt).Item("parent").keys()) >= 0))) Then
                pos_out_onesubpos = SpecOneSubpos(all_data, "-", type_spec, floor_txt)
                If this_sheet_option.Item("qty_one_floor") Then
                    subpos = "-"
                    If Not pos_out_dic.Exists(subpos) Then Set pos_out_dic.Item(subpos) = CreateObject("Scripting.Dictionary")
                    For i = 1 To UBound(pos_out_onesubpos, 1)
                        type_el = CStr(pos_out_onesubpos(i, 7))
                        If Not pos_out_dic.Item(subpos).Exists(type_el) Then Set pos_out_dic.Item(subpos).Item(type_el) = CreateObject("Scripting.Dictionary")
                        pos = CStr(pos_out_onesubpos(i, 1))
                        obozn = CStr(pos_out_onesubpos(i, 2))
                        naen = CStr(pos_out_onesubpos(i, 3))
                        ves = CStr(pos_out_onesubpos(i, 5))
                        pos_out_onesubpos(i, 7) = pos_out_onesubpos(i, 8)
                        pos_out_onesubpos(i, 8) = floor_txt
                        row_key = pos + "%" + obozn + "%" + naen + "%" + ves
                        row_type = ArrayRow(pos_out_onesubpos, i, True)
                        If Not pos_out_dic.Item(subpos).Item(type_el).Exists(row_key) Then
                            pos_out_dic.Item(subpos).Item(type_el).Item(row_key) = row_type
                        Else
                            pos_out_dic.Item(subpos).Item(type_el).Item(row_key) = ArrayCombine(pos_out_dic.Item(subpos).Item(type_el).Item(row_key), row_type)
                        End If
                        n_row_out = n_row_out + 1
                    Next i
                End If
                pos_out = ArrayCombine(pos_out, pos_out_onesubpos)
            End If
        End If
        
        If type_spec = 13 And Not IsEmpty(all_data) Then
            ReDim pos_out_subpos(UBound(all_data, 1), end_col)
            ReDim pos_out_arm(UBound(all_data, 1), end_col)
            ReDim pos_out_prokat(UBound(all_data, 1), end_col)
            ReDim pos_out_izd(UBound(all_data, 1), end_col)
            ReDim pos_out_mat(UBound(all_data, 1), end_col)
            n_row_subpos = 0
            n_row_arm = 0
            n_row_prokat = 0
            n_row_izd = 0
            n_row_mat = 0
            subpos_arr = pos_data.Item(floor_txt).Item("parent").keys()
            n_subpos = UBound(subpos_arr, 1) - LBound(subpos_arr, 1)
            If n_subpos >= 0 Then
                For Each subpos In ArraySort(subpos_arr, 1)
                    'Получаем спецификацию на одну сборку и начинаем её делить по типам
                    pos_out_tmp = SpecOneSubpos(all_data, subpos, type_spec, floor_txt)
                    If Not IsEmpty(pos_out_tmp) Then
                        If this_sheet_option.Item("qty_one_subpos") Then
                            'Количество на один этаж
                            nSubPos = GetNSubpos(subpos, type_spec, floor_txt)
                            'В названии - количество на все этажи
                            naen_col_subpos = subpos & vbLf & "(" & GetNSubpos(subpos, type_spec, "all_floor") & "шт)"
                        Else
                            nSubPos = 1
                            naen_col_subpos = subpos
                        End If
                        'Ищем номер столбца с именем сборки
                        flag = 0
                        For i = 4 To UBound(pos_out, 2)
                            If pos_out(2, i) = naen_col_subpos Then flag = i
                        Next i
                        'Если не нашли - создаём новый
                        If flag = 0 Then
                            n_col_sb = n_col_end
                            n_col_end = n_col_end + 1
                            pos_out(2, n_col_sb) = naen_col_subpos
                        Else
                            n_col_sb = flag
                        End If
                        'Если не нужно указывать количество
                        If this_sheet_option.Item("show_qty_spec") Then pos_out(2, n_col_sb) = subpos & ",**"
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
                    End If
                Next
            End If
            If pos_data.Item(floor_txt).Exists("-") Then
                pos_out_tmp = SpecOneSubpos(all_data, "-", type_spec, floor_txt)
                If Not IsEmpty(pos_out_tmp) Then
                    n_col_sb = end_col - 3
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
            End If
            pos_out_subpos = ArraySort(ArrayRedim(pos_out_subpos, n_row_subpos), 1)
            pos_out_arm = ArraySort(ArrayRedim(pos_out_arm, n_row_arm), 1)
            pos_out_prokat = ArraySort(ArrayRedim(pos_out_prokat, n_row_prokat), 1)
            pos_out_izd = ArraySort(ArrayRedim(pos_out_izd, n_row_izd), 1)
            pos_out_mat = ArraySort(ArrayRedim(pos_out_mat, n_row_mat), 1)
            If n_row_subpos > 0 Then pos_out_floor = ArrayCombine(pos_out_floor, pos_out_subpos)
            If n_row_arm > 0 Then pos_out_floor = ArrayCombine(pos_out_floor, pos_out_arm)
            If n_row_prokat > 0 Then pos_out_floor = ArrayCombine(pos_out_floor, pos_out_prokat)
            If n_row_izd > 0 Then pos_out_floor = ArrayCombine(pos_out_floor, pos_out_izd)
            If n_row_mat > 0 Then pos_out_floor = ArrayCombine(pos_out_floor, pos_out_mat)
            For i = 1 To UBound(pos_out_floor, 1)
                If Not IsEmpty(pos_out_floor(i, end_col - 1)) Then
                    For j = 4 To end_col - 1
                        If IsEmpty(pos_out_floor(i, j)) Then pos_out_floor(i, j) = "-"
                    Next j
                End If
                If IsNumeric(pos_out_floor(i, end_col)) Then
                    If Round_w(pos_out_floor(i, end_col), 0) > 0 Then
                        pos_out_floor(i, end_col) = Trim$(ConvNum2Txt(Round_w(pos_out_floor(i, end_col), n_round_w)) & " кг.")
                        If Left$(pos_out_floor(i, end_col), 1) = "." Then pos_out_floor(i, end_col) = "0" + pos_out_floor(i, end_col)
                    Else
                        pos_out_floor(i, end_col) = "-"
                    End If
                End If
            Next i
            If Not this_sheet_option.Item("show_type") Then
                pos_out_sort = ArraySort(pos_out_floor, 1)
                n_row = 1
                For i = 1 To UBound(pos_out_sort, 1)
                    If pos_out_sort(i, 1) <> "Поз." And pos_out_sort(i, 3) <> vbNullString And InStr(pos_out_sort(i, 1), "на отм.") = 0 Then
                        n_row = n_row + 1
                        If n_row <= UBound(pos_out_sort, 1) Then
                            For j = 1 To UBound(pos_out_floor, 2)
                                pos_out_floor(n_row, j) = pos_out_sort(i, j)
                            Next j
                        End If
                    End If
                Next i
            End If
            pos_out = ArrayCombine(pos_out, pos_out_floor)
            If this_sheet_option.Item("qty_one_floor") Then
                ReDim pos_out_subpos(1, 1)
                ReDim pos_out_arm(1, 1)
                ReDim pos_out_prokat(1, 1)
                ReDim pos_out_izd(1, 1)
                ReDim pos_out_mat(1, 1)
                pos_out_floor = Empty
            End If
        End If
    Next inxfloor
    
    If this_sheet_option.Item("qty_one_floor") And (type_spec = 2 Or type_spec = 3) And spec_version > 1 Then
        end_col = 6 + nfloor
        ReDim pos_out_floor(n_row_out + 2, end_col)
        n_row = 2
        pos_out_floor(1, 4) = "Кол-во на отм."
        pos_out_floor(1, 1) = "Поз."
        pos_out_floor(1, 2) = "Обозначение"
        pos_out_floor(1, 3) = "Наименование"
        pos_out_floor(1, 4) = "Кол-во на отм."
        pos_out_floor(1, end_col - 2) = "Всего"
        pos_out_floor(1, end_col - 1) = "Масса ед., кг."
        pos_out_floor(1, end_col) = "Примечание"
        subpos_array = ArraySort(pos_out_dic.keys)
        For Each subpos In subpos_array
            If subpos <> "-" Then
                For Each type_el In pos_out_dic.Item(subpos).keys
                    For Each row In pos_out_dic.Item(subpos).Item(type_el).keys
                        arr = pos_out_dic.Item(subpos).Item(type_el).Item(row)
                        arr = ArraySort(arr, 1)
                        n_row = n_row + 1
                        For inxfloor = 1 To nfloor
                            n_col_floor = 3 + inxfloor
                            t_floor = floor_txt_arr(inxfloor, 2)
                            floor_txt = floor_txt_arr(inxfloor, 3)
                            pos_out_floor(2, n_col_floor) = ConvNum2Otm(t_floor)
                            el_floor = ArraySelectParam(arr, floor_txt, 8)
                            If Not IsEmpty(el_floor) Then
                                For i = 1 To UBound(el_floor, 1)
                                    If Len(type_el) = 0 Then
                                        qty = el_floor(i, 6)
                                        If qty < 0 Then qty = 0
                                    Else
                                        qty = el_floor(i, 4)
                                        If Not IsNumeric(qty) Then qty = ConvTxt2Num(qty)
                                        If Not IsNumeric(qty) Then qty = 0
                                    End If
                                    nSubPos = el_floor(i, 7)
                                    pos_out_floor(n_row, 1) = el_floor(i, 1)
                                    pos_out_floor(n_row, 2) = el_floor(i, 2)
                                    pos_out_floor(n_row, 3) = el_floor(i, 3)
                                    pos_out_floor(n_row, n_col_floor) = pos_out_floor(n_row, n_col_floor) + qty * nSubPos
                                    pos_out_floor(n_row, end_col - 2) = pos_out_floor(n_row, end_col - 2) + qty * nSubPos
                                    pos_out_floor(n_row, end_col - 1) = el_floor(i, 5)
                                    If IsNumeric(el_floor(i, 5)) Then
                                        pos_out_floor(n_row, end_col) = pos_out_floor(n_row, end_col) + el_floor(i, 5) * qty * nSubPos
                                    Else
                                        pos_out_floor(n_row, end_col) = el_floor(i, 6)
                                    End If
                                Next i
                            End If
                        Next inxfloor
                    Next
                Next
            End If
        Next
        subpos = "-"
        If pos_out_dic.Exists(subpos) Then
            For Each type_el In pos_out_dic.Item(subpos).keys
                For Each row In pos_out_dic.Item(subpos).Item(type_el).keys
                    arr = pos_out_dic.Item(subpos).Item(type_el).Item(row)
                    arr = ArraySort(arr, 1)
                    n_row = n_row + 1
                    For inxfloor = 1 To nfloor
                        n_col_floor = 3 + inxfloor
                        t_floor = floor_txt_arr(inxfloor, 2)
                        floor_txt = floor_txt_arr(inxfloor, 3)
                        pos_out_floor(2, n_col_floor) = ConvNum2Otm(t_floor)
                        el_floor = ArraySelectParam(arr, floor_txt, 8)
                        If Not IsEmpty(el_floor) Then
                            For i = 1 To UBound(el_floor, 1)
                                If Len(type_el) = 0 Then
                                    qty = el_floor(i, 6)
                                    If qty < 0 Then qty = 0
                                Else
                                    qty = el_floor(i, 4)
                                End If
                                nSubPos = el_floor(i, 7)
                                pos_out_floor(n_row, 1) = el_floor(i, 1)
                                pos_out_floor(n_row, 2) = el_floor(i, 2)
                                pos_out_floor(n_row, 3) = el_floor(i, 3)
                                If IsNumeric(qty) Then
                                    pos_out_floor(n_row, n_col_floor) = pos_out_floor(n_row, n_col_floor) + qty * nSubPos
                                    pos_out_floor(n_row, end_col - 2) = pos_out_floor(n_row, end_col - 2) + qty * nSubPos
                                End If
                                pos_out_floor(n_row, end_col - 1) = el_floor(i, 5)
                                If IsNumeric(el_floor(i, 5)) And IsNumeric(qty) Then
                                    pos_out_floor(n_row, end_col) = pos_out_floor(n_row, end_col) + el_floor(i, 5) * qty * nSubPos
                                Else
                                    pos_out_floor(n_row, end_col) = el_floor(i, 6)
                                End If
                            Next i
                        End If
                    Next inxfloor
                Next
            Next
        End If
        pos_out = ArrayRedim(pos_out_floor, n_row)
    End If
    
    If Not IsEmpty(pos_out) Then
        For i = 1 To UBound(pos_out, 1)
            If pos_out(i, 3) <> vbNullString Then
                If IsNumeric(ConvTxt2Num(pos_out(i, end_col))) Then
                    If Round_w(pos_out(i, end_col), 0) > 0 Then
                        pos_out(i, end_col) = Trim$(ConvNum2Txt(Round_w(pos_out(i, end_col), n_round_w)) & " кг.")
                        If Left$(pos_out(i, end_col), 1) = "." Then pos_out(i, end_col) = "0" + pos_out(i, end_col)
                    Else
                        If Not (IsNumeric(Application.Match(pos_out(i, 3), type_el_name.items, 0))) Then pos_out(i, end_col) = "'"
                    End If
                End If
                For kk = 4 To end_col
                    If (Len(pos_out(i, kk)) = 0 Or pos_out(i, kk) = " " Or pos_out(i, kk) = 0) And Not (IsNumeric(Application.Match(pos_out(i, 3), type_el_name.items, 0))) Then pos_out(i, kk) = "-"
                Next kk
            End If
            If InStr(pos_out(i, 1), fin_str) > 0 Then
                pos_out(i, 1) = Replace(pos_out(i, 1), fin_str, vbNullString)
                pos_out(i, end_col) = "'"
            End If
        Next i
        n_col_naen = 3: n_col_pos = 1
        If type_spec = 1 Then n_col_pos = 2
        If show_sum_prim Then
            For i = 2 To UBound(pos_out, 1)
                If (Right$(Trim$(UCase$(pos_out(i, n_col_naen))), 1) = "*") Then
                    pos_out(i, n_col_naen) = Left$(pos_out(i, n_col_naen), Len(pos_out(i, n_col_naen)) - 1)
                    pos_out(i, 1) = pos_out(i, 1) & "*"
                End If
            Next i
        End If
    End If
    If this_sheet_option.Item("ignore_subpos") = True And this_sheet_option.Item("show_type") = False Then
        istart = 1
        For i = 2 To Application.WorksheetFunction.Min(5, UBound(pos_out, 1))
            If Len(pos_out(i, 1)) > 0 And InStr(pos_out(i, 1), "Поз.") = 0 And istart = 1 Then istart = i
        Next i
        
        Dim pos_out_sort_end
        ReDim pos_out_sort_end(UBound(pos_out, 1) - istart + 1, UBound(pos_out, 2))
        Dim pos_out_head
        ReDim pos_out_head(Application.WorksheetFunction.Max(1, istart - 1), UBound(pos_out, 2))
        For i = 1 To istart - 1
            For j = 1 To UBound(pos_out, 2)
                pos_out_head(i, j) = pos_out(i, j)
            Next j
        Next i
        ii = 0
        For i = istart To UBound(pos_out, 1)
            ii = ii + 1
            For j = 1 To UBound(pos_out, 2)
                pos_out_sort_end(ii, j) = pos_out(i, j)
            Next j
        Next i
        hh = 1
        n_col_naen = 3: n_col_pos = 1: n_col_obozn = 1
        If type_spec = 1 Then
            n_col_naen = 3
            n_col_pos = 2
            n_col_obozn = 4
        End If
        pos_out_sort_end = ArraySort_2(pos_out_sort_end, Array(n_col_pos, n_col_obozn, n_col_naen))
        pos_out = ArrayCombine(pos_out_head, pos_out_sort_end)
    End If
    Spec_AS = pos_out
End Function
Function SpecPeremMarka(ByRef all_data_perem As Variant) As Boolean
    rez = False
    all_data_marka = ArraySelectParam(all_data_perem, t_perem_m, col_type_el)
    un_marka = ArrayUniqValColumn(all_data_marka, col_m_naen)
    If Not IsEmpty(un_marka) Then
        n_mark = UBound(un_marka)
        n_znak = Len(CStr(n_mark))
        Dim mark_otm
        ReDim mark_otm(n_mark, 3)
        For i = 1 To n_mark
            marka = un_marka(i)
            mark_otm(i, 1) = ArraySelectParam(all_data_marka, marka, col_m_naen)(1, col_m_naen)
            mark_otm(i, 2) = vbNullString
            t_marka = ArraySelectParam(all_data_marka, marka, col_m_naen)
            un_otm = ArrayUniqValColumn(t_marka, col_param)
            For j = 1 To UBound(un_otm)
                un_otm(j) = GetZoneParam(un_otm(j), "Z")
            Next j
            un_otm = ArraySort(un_otm)
            For j = 1 To UBound(un_otm)
                otm_txt = ConvNum2Otm(un_otm(j)) & Space(4)
                mark_otm(i, 2) = mark_otm(i, 2) & otm_txt
            Next j
            zero_txt = vbNullString
            If n_znak > 1 Then
                n_zero = n_znak - Len(CStr(i))
                For n = 1 To n_zero
                    zero_txt = zero_txt + "0"
                Next n
            End If
            mark_otm(i, 3) = "ПР" + zero_txt + CStr(i)
        Next i
        If Not IsEmpty(mark_otm) Then
            CSVfilename$ = ThisWorkbook.path & "\list\Марки_" & ThisWorkbook.ActiveSheet.Name & ".txt"
            n = ExportArray2CSV(mark_otm, CSVfilename$)
            rez = True
        End If
    End If
    If rez Then MsgBox ("Данные о марках перемычек записаны в файл" & vbLf & "\list\Марки_" & ThisWorkbook.ActiveSheet.Name & ".txt")
    SpecPeremMarka = rez
End Function

Function Spec_WIN(ByRef all_data As Variant) As Variant
    all_data_perem = ArraySelectParam_2(all_data, Array(t_perem, t_perem_m), col_type_el)
    all_data = ArraySelectParam(all_data, t_wind, col_type_el)
    If IsEmpty(all_data) And IsEmpty(all_data_perem) Then Spec_WIN = Empty: Exit Function
    Dim out_data(): ReDim out_data(2)
    Dim myRegExp As Object
    Dim myObj As Object
    If Not IsEmpty(all_data) Then
        un_chsum = ArrayUniqValColumn(all_data, col_chksum)
        pos_chsum = UBound(un_chsum, 1)
        un_floor = ArrayUniqValColumn(all_data, col_floor)
        For i = 1 To pos_chsum
            un_chsum(i) = Split(un_chsum(i), "_")(0)
        Next i
        un_chsum = ArrayUniqValColumn(un_chsum, 1)
        pos_chsum = UBound(un_chsum, 1)
        Dim pos_zag
        If this_sheet_option.Item("qty_one_floor") Then
            ReDim pos_zag(2, UBound(un_floor) + 6)
            pos_zag(1, 1) = "Позиция"
            pos_zag(1, 2) = "Обозначение"
            pos_zag(1, 3) = "Наименование"
            pos_zag(1, 4) = "Количество"
            i = 4
            floor_start = i
            For Each tfloor In un_floor
                For j = 1 To UBound(all_data, 1)
                    If tfloor = all_data(j, col_floor) Then
                        pos_zag(2, i) = all_data(j, col_floor)
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
        pos_wind1 = "о"
        pos_door1 = "д"
        n_row_out = 0
        For Each t In Array(pos_wind, pos_wind1, pos_door, pos_door1)
            el_data = ArraySelectParam(all_data, t, col_pos)
            un_sub_pos_el = ArrayUniqValColumn(el_data, col_sub_pos)
            If Not IsEmpty(un_sub_pos_el) Then
                For Each sub_pos In un_sub_pos_el
                    pos_dat = ArraySelectParam(all_data, sub_pos, col_sub_pos)
                    'Поставим заполнение впереди
                    un_pos = ArrayDelElement(ArrayUniqValColumn(pos_dat, col_pos), t)
                    un_pos = ArrayCombine(Array(t), un_pos)
                    For Each pos_el In un_pos
                        If Not IsEmpty(pos_el) Then
                            n_row_out = n_row_out + 1
                            pos_out(n_row_out, n_col_qty + 2) = Array(0, 0, vbNullString)
                        End If
                        For i = 1 To UBound(pos_dat)
                            tpos = pos_dat(i, col_pos)
                            If tpos = pos_el Then
                                sub_pos = pos_dat(i, col_sub_pos)
                                pos = Replace(pos_dat(i, col_pos), t, vbNullString)
                                If pos = "-" Then pos = vbNullString
                                pos = Trim$(pos)
                                If Len(pos) > 0 Then pos = pos + ":"
                                obozn = pos_dat(i, col_w_obozn)
                                naen = pos & pos_dat(i, col_w_naen)
                                qty = pos_dat(i, col_qty)
                                Weight = pos_dat(i, col_w_weight)
                                prim = pos_dat(i, col_w_prim)
                                If prim = "-" Then prim = vbNullString
                                If prim = "п.м." Then
                                    naen = naen + " L=п.м."
                                    prim = vbNullString
                                End If
                                area = GetZoneParam(pos_dat(i, col_param), "S")
                                'Проверка площади
                                Set myRegExp = CreateObject("VBScript.RegExp")
                                myStr2 = ""
                                myRegExp.Global = True
                                myRegExp.Pattern = "[0-9]{2,6}\(h\)x[0-9]{2,6}"
                                Set myObj = myRegExp.Execute(naen)
                                For Each myStr1 In myObj
                                    myStr2 = myStr2 & myStr1 & vbNewLine
                                Next
                                If Len(myStr2) > 0 Then
                                    razm = Split(myStr2, "(h)x")
                                    If UBound(razm) > 0 Then
                                        a = 0
                                        b = 0
                                        If IsNumeric(Trim(razm(0))) Then a = CDbl(Trim(razm(0))) / 1000
                                        If IsNumeric(Trim(razm(1))) Then b = CDbl(Trim(razm(1))) / 1000
                                        If a > 0 And b > 0 Then
                                            area_temp = a * b
                                            If Abs(area_temp - area) > 0.01 Then area = 0
                                        End If
                                    End If
                                End If
                                
                                pos_out(n_row_out, 1) = sub_pos
                                pos_out(n_row_out, 2) = obozn
                                pos_out(n_row_out, 3) = naen
                                If this_sheet_option.Item("qty_one_floor") Then
                                    t_floor = pos_dat(i, col_floor)
                                    For k = floor_start To floor_end
                                        If pos_zag(2, k) = t_floor Then pos_out(n_row_out, k) = pos_out(n_row_out, k) + qty
                                    Next k
                                End If
                                pos_out(n_row_out, n_col_qty) = pos_out(n_row_out, n_col_qty) + qty
                                pos_out(n_row_out, n_col_qty + 1) = Weight
                                pos_out(n_row_out, n_col_qty + 2)(1) = pos_out(n_row_out, n_col_qty + 2)(1) + qty * Weight
                                pos_out(n_row_out, n_col_qty + 2)(2) = pos_out(n_row_out, n_col_qty + 2)(2) + area * qty
                                If Len(pos_out(n_row_out, n_col_qty + 2)(3)) = 0 Then pos_out(n_row_out, n_col_qty + 2)(3) = prim
                            End If
                        Next i
                    Next
                Next
            End If
        Next
        For i = 1 To n_row_out
            prim = vbNullString
            If Len(pos_out(i, n_col_qty + 2)(3)) = 0 Then
                If pos_out(i, n_col_qty + 2)(1) > 0 Then prim = prim + ConvNum2Txt(pos_out(i, n_col_qty + 2)(1)) + "кг. " & vbLf
                If pos_out(i, n_col_qty + 2)(2) > 0 Then prim = prim + ConvNum2Txt(pos_out(i, n_col_qty + 2)(2)) + "кв.м."
            Else
                prim = pos_out(i, n_col_qty + 2)(3)
            End If
            pos_out(i, n_col_qty + 2) = prim
        Next i
        If this_sheet_option.Item("qty_one_floor") Then
            For k = floor_start To floor_end
                For Each deltxt In Array("План", "НА", "этаж", "отм.")
                    pos_zag(2, k) = Replace(pos_zag(2, k), deltxt, vbNullString)
                Next
                pos_zag(2, k) = ConvNum2Otm(pos_zag(2, k))
                For i = 1 To n_row_out
                    If IsEmpty(pos_out(i, k)) Then pos_out(i, k) = "-"
                Next i
            Next k
        End If
        out_data(1) = ArrayCombine(pos_zag, pos_out)
    Else
        out_data(1) = Empty
    End If
    'Чтоб дважды не вставать - бахнем спец-ю для перемычек
    type_spec = 3
    this_sheet_option.Item("ignore_subpos") = True
    If Not IsEmpty(all_data_perem) Then
        r = SpecPeremMarka(all_data_perem)
        For i = 1 To UBound(all_data_perem, 1)
            If all_data_perem(i, col_type_el) = t_perem Then all_data_perem(i, col_type_el) = t_izd
            If all_data_perem(i, col_type_el) = t_perem_m Then all_data_perem(i, col_type_el) = t_subpos
        Next i
    End If
    If Not IsEmpty(all_data_perem) Then
        all_data_perem = DataPrepare(all_data_perem)
        out_data(2) = Spec_AS(all_data_perem, type_spec)
    End If
    Spec_WIN = out_data
End Function

Function Spec_KM(ByRef all_data As Variant) As Variant
    prokat = ArraySelectParam(all_data, t_prokat, col_type_el)
    If IsEmpty(prokat) Then
        n_prokat = 0
        MsgBox ("Прокат в файле/листе не найден")
        r = LogWrite("Ошибка спецификации", vbNullString, "Прокат в файле/листе не найден")
        Spec_KM = Empty
        Exit Function
    Else
        n_prokat = UBound(prokat, 1)
    End If

    If this_sheet_option.Item("ed_izm_km") Then
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
                n_t_konstr = 0
                If Not IsEmpty(unique_konstr) Then n_t_konstr = UBound(unique_konstr)
                For k = 1 To n_t_konstr
                    'Все элементы с заданной конструкцией
                    type_konstr = unique_konstr(k) 'Текущий тип конструкции
                    elem_m = ArraySelectParam(el, type_konstr, col_pr_type_konstr) 'Выбираем все элементы с этим типом конструкций
                    n_el_m = UBound(elem_m, 1)
                    Weight# = 0 'Начинаем считать вес для каждого типа
                    For kk = 1 To n_el_m
                        'Вес одной строки, с учётом того, что масса дана на п.м. в кг.
                        wp = elem_m(kk, col_pr_weight) * elem_m(kk, col_qty) * elem_m(kk, col_pr_length) / 1000
                        Weight# = Weight# + wp * k_zap_total
                    Next kk
                    'Итоговый вес для отдельного типа конструкции, в тоннах
                    'Из-за особенностей ГОСТа минимальное значение - 100 кг.
                    'Не плохой такой источник экономии
                    Weight# = Round_w(Weight# / koef, n_okr)
                    w_min = 1 / (10 ^ n_okr)
                    If Weight# < w_min Then
                        If hard_round_km Then
                            Weight# = w_min
                        Else
                            If Weight# < 0.00001 Then
                                Weight# = 0.00001
                            End If
                            wt = ConvNum2Txt(Weight#)
                            If Len(wt) > Len(w_format) Then
                                w_format_t = "0."
                                For nnul = 1 To Len(wt) - Len(w_format_t)
                                    w_format_t = w_format_t + "0"
                                Next nnul
                                w_format = w_format_t
                            End If
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
    If Not IsEmpty(mat) And this_sheet_option.Item("ismat") Then
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
                    obozn = mat(j, col_m_obozn): If obozn <> vbNullString Then obozn = " по " & obozn
                    ed = mat(j, col_m_edizm)
                    qty = mat(j, col_qty)
                    usl = 0
                    For Each n In naen_mat
                        usl = usl + InStr(naen, n)
                    Next
                    If usl > 0 Then
                        pos_out(row, 1) = naen & obozn & ", " & ed
                        pos_out(row, n_type_konstr + 5) = pos_out(row, n_type_konstr + 5) + qty * k_zap_total
                    End If
                End If
            Next j
            If pos_out(row, n_type_konstr + 5) <> 0 Then row = row + 1
        Next i
    End If
    'erase prokat, unique_gost_prof, unique_stal, prof_stal, unique_prof_stal, unique_type_konstr, prof, unique_prof, el, unique_konstr, elem_m, weight_stal, weight_gost_prof, weight_total, weight_stal_total
    pos_out = ArrayRedim(pos_out, row - 1)
    Spec_KM = pos_out
End Function

Function Spec_KZH(ByRef all_data As Variant) As Variant
    'Из всех родительсвких сборок получим только сборки с армированием
Dim tfunctime As Double
tfunctime = Timer
    If this_sheet_option.Item("isarm") And this_sheet_option.Item("isprok") Then
        find_elem = Array(t_arm, t_prokat)
    Else
        If this_sheet_option.Item("isarm") Then find_elem = Array(t_arm)
        If this_sheet_option.Item("isprok") Then find_elem = Array(t_prokat)
    End If
    all_data_arm = ArraySelectParam_2(all_data, find_elem, col_type_el)
    If IsEmpty(all_data_arm) Then
        MsgBox ("Элементы армирования не найдены")
        Spec_KZH = Empty
        Exit Function
    End If
    un_arm_subpos_1 = ArrayUniqValColumn(all_data_arm, col_sub_pos)
    un_arm_subpos_2 = ArrayUniqValColumn(all_data_arm, col_parent)
    un_arm_subpos = ArrayUniqValColumn(ArrayCombine(un_arm_subpos_1, un_arm_subpos_2))
    has_nosubpos_arm = False
    For j = 1 To UBound(un_arm_subpos_1)
        If un_arm_subpos_1(j) = "-" Then
            has_nosubpos_arm = True
            j = UBound(un_arm_subpos_1)
        End If
    Next j
    floor_txt = "all_floor"
    Set name_subpos = pos_data.Item(floor_txt).Item("name") 'Словарь с именами сборок
    un_child = ArraySort(pos_data.Item(floor_txt).Item("child").keys())
    If IsEmpty(un_child) Then un_child = Array()
    un_parent = ArraySort(pos_data.Item(floor_txt).Item("parent").keys())
    If IsEmpty(un_parent) Then un_parent = Array()
    un_parent = ArraySelectParam_2(un_parent, un_arm_subpos, col_type_el)
    If IsEmpty(un_parent) Then un_parent = Array()
    'Получаем словарь с количеством и типом бетона для сборок
    is_bet = False
    If show_bet_wkzh And this_sheet_option.Item("ismat") Then is_bet = Spec_CONC(all_data, un_parent)
tfunctime = functime("Spec_KZH_выбор элементов", tfunctime)
    'Выясняем - какие диаметры и какие классы арматуры есть для всех сборок
    'заодно отсортируем арматуру в закладных деталях и прокат
    n_row = UBound(all_data_arm, 1)
    Dim arm_arr(): ReDim arm_arr(8)
    Dim temp_arr(): ReDim temp_arr(n_row, max_col)
    For i = 1 To 4
        arm_arr(i) = temp_arr
    Next i
    n_arm_a = 0: n_arm_z = 0: n_prokat_a = 0: n_prokat_z = 0
    For i = 1 To n_row
        сurrent_type_el = all_data_arm(i, col_type_el)
        If сurrent_type_el = t_arm Or сurrent_type_el = t_prokat Then
            сurrent_subpos = all_data_arm(i, col_sub_pos)
            naen = " "
            If name_subpos.Exists(сurrent_subpos) Then naen = name_subpos.Item(сurrent_subpos)(1)
            flag = False
            Select Case сurrent_type_el
                Case t_arm
                    If InStr(naen, "Заклад") = 0 Then
                        n_arm_a = n_arm_a + 1
                        For j = 1 To max_col
                            arm_arr(1)(n_arm_a, j) = all_data_arm(i, j)
                        Next j
                        arm_arr(4 + 1) = n_arm_a
                    Else
                        n_arm_z = n_arm_z + 1
                        For j = 1 To max_col
                            arm_arr(3)(n_arm_z, j) = all_data_arm(i, j)
                        Next j
                        arm_arr(4 + 3) = n_arm_z
                    End If
                Case t_prokat
                    If InStr(naen, "Заклад") = 0 Then
                        n_prokat_a = n_prokat_a + 1
                        For j = 1 To max_col
                            arm_arr(2)(n_prokat_a, j) = all_data_arm(i, j)
                        Next j
                        arm_arr(4 + 2) = n_prokat_a
                    Else
                        n_prokat_z = n_prokat_z + 1
                        For j = 1 To max_col
                            arm_arr(4)(n_prokat_z, j) = all_data_arm(i, j)
                        Next j
                        arm_arr(4 + 4) = n_prokat_z
                    End If
            End Select
        End If
    Next
tfunctime = functime("Spec_KZH_массив элементов", tfunctime)
    'Теперь у нас есть массив с отсортированной арматурой для всех сборок
    '1 - Арматура общая
    '2 - Прокат общий
    '3 - Арматура в закладных
    '4 - Прокат в закладных
    'Сформируем общую таблицу диаметров и классов арматуры
    n_row = 5
    If UBound(un_parent) >= 0 Then n_row = n_row + UBound(un_parent)
    If has_nosubpos_arm Then n_row = n_row + 1
    If this_sheet_option.Item("ignore_subpos") = True Then sum_row_wkzh = True
    sum_row = 0: If n_row > 6 And sum_row_wkzh = True Then sum_row = 1
    Dim pos_out(): ReDim pos_out(n_row + sum_row, 1)
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
    Dim pos_out_nsubpos()
    ReDim pos_out_nsubpos(n_row + sum_row, 2)
    If is_bet = True Then
        n_conc = 0
        For Each sub_bet In concrsubpos.keys()
            If InStr(sub_bet, "tot_") > 0 Then n_conc = n_conc + 1
        Next
        If n_conc > 1 Then n_conc = n_conc + 1
        ReDim pos_out_bet(n_row + sum_row, n_conc)
        pos_out_bet(1, 1) = "Объём бетона, куб.м."
        If n_conc > 2 Then pos_out_bet(1, n_conc) = "Всего"
        n_conc = 0
        For Each sub_bet In concrsubpos.keys()
            If InStr(sub_bet, "tot_") > 0 Then
                bet = Split(sub_bet, "_")(1)
                n_conc = n_conc + 1
                pos_out_bet(2, n_conc) = bet
            End If
        Next
    End If
    For kk = 1 To 5
        pos_out(kk, current_size + 1) = "Всего"
    Next kk
tfunctime = functime("Spec_KZH_выбор диаметров", tfunctime)
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
    'Заполняем словарь со строками для сборок
    k = 1
    If Not IsEmpty(un_parent) Then
        For k = 1 To UBound(un_parent)
            subpos = un_parent(k)
            nSubPos = pos_data.Item(floor_txt).Item("qty").Item("-_" & subpos)
            If nSubPos < 1 Then
                r = LogWrite("Ошибка спецификации", subpos, "Не определено кол-во сборок")
                MsgBox ("Не определено кол-во сборок " & subpos & ", принято 1 шт.")
                nSubPos = 1
            End If
            n_txt = subpos & get_nsubpos_txt(nSubPos)
            If this_sheet_option.Item("qty_one_subpos") = False Then nSubPos = 1
            k_row = k + 5
            pos_out(k_row, 1) = n_txt
            weight_index.Item("row" & subpos) = k_row
            pos_out_nsubpos(k_row, 1) = subpos
            pos_out_nsubpos(k_row, 2) = nSubPos
        Next k
    End If
    If has_nosubpos_arm Then
        subpos = "-"
        nSubPos = 1
        k_row = k + 6
        pos_out(k_row, 1) = "Прочие,**"
        weight_index.Item("row" & subpos) = k_row
        pos_out_nsubpos(k_row, 1) = subpos
        pos_out_nsubpos(k_row, 2) = nSubPos
    End If

    If is_bet = True Then
        n_conc_end_col = 0
        For k = 5 To k_row
            subpos = pos_out_nsubpos(k, 1)
            nSubPos = pos_out_nsubpos(k, 2)
            For Each sub_bet In concrsubpos.keys()
                v_bet = 0: naen_bet = vbNullString: flag = 1
                If InStr(sub_bet, "_") > 0 And Right$(sub_bet, 4) = "_qty" And InStr(sub_bet, "bet") = 0 Then
                    subb = Split(sub_bet, "_")
                    bet_subpos = subb(0)
                    naen_bet = subb(1)
                    v_bet = concrsubpos.Item(sub_bet)
                    If bet_subpos = subpos Then 'Выбираем бетон для текущей сборки
                        For jc = 1 To UBound(pos_out_bet, 2)
                            If pos_out_bet(2, jc) = naen_bet Then pos_out_bet(k, jc) = pos_out_bet(k, jc) + v_bet
                        Next jc
                        pos_out_bet(k, UBound(pos_out_bet, 2)) = concrsubpos.Item(subpos & "_bet_qty")
                    End If
                End If
            Next
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
            u1 = (pos_data.Item(floor_txt).Item("parent").Exists(subpos) Or pos_data.Item(floor_txt).Item("parent").Exists(tparent))
            If has_nosubpos_arm Then u2 = ((subpos = "-") Or (pos_data.Item(floor_txt).Item("-").Exists(subpos) And tparent = "-"))
            If u1 Or u2 Then
                If u2 Then
                    nSubPos = 1
                    k = weight_index.Item("row" & tparent)
                End If
                If u1 Then
                    If pos_data.Item(floor_txt).Item("parent").Exists(subpos) Then
                        nSubPos = pos_data.Item(floor_txt).Item("qty").Item("-_" & subpos)
                        k = weight_index.Item("row" & subpos)
                    End If
                    If pos_data.Item(floor_txt).Item("parent").Exists(tparent) Then
                        nSubPos = pos_data.Item(floor_txt).Item("qty").Item("-_" & tparent)
                        k = weight_index.Item("row" & tparent)
                    End If
                End If
                If IsEmpty(k) Then
                    hh = 1
                End If
                If Not this_sheet_option.Item("qty_one_subpos") Then nSubPos = 1
                If arm_arr(i)(j, col_type_el) = t_arm Then
                    diametr = arm_arr(i)(j, col_diametr)
                    klass = arm_arr(i)(j, col_klass)
                    gost = GetGOSTForKlass(klass)
                    length_pos = arm_arr(i)(j, col_length) / 1000
                    weight_pm = GetWeightForDiametr(diametr, klass)
                    qty = arm_arr(i)(j, col_qty)
                    If (qty = 0) Or IsEmpty(qty) Then qty = 1
                    fon = arm_arr(i)(j, col_fon)
                    If fon Or this_sheet_option.Item("arm_pm") Then
                        length_pos = Round_w(length_pos * k_zap_total, n_round_l)
                        w_pos = length_pos * weight_pm * qty / nSubPos
                    Else
                        'Запас для одиночных элементов применяется только если это включено в setting
                        If zap_only_mp Then
                            w_pos = Round_w(weight_pm * length_pos, n_round_w) * qty / nSubPos
                        Else
                            w_pos = Round_w(weight_pm * length_pos * k_zap_total, n_round_w) * qty / nSubPos
                        End If
                    End If
                    tkeyd = "Арматура" & n_type & klass & gost & symb_diam & diametr
                    tkesum_1 = "Арматура" & n_type & klass & gost & "Всего"
                Else
                    prof = arm_arr(i)(j, col_pr_prof)
                    gost_prof = arm_arr(i)(j, col_pr_gost_prof)
                    stal = arm_arr(i)(j, col_pr_st)
                    gost_stal = arm_arr(i)(j, col_pr_gost_st)
                    qty = arm_arr(i)(j, col_qty)
                    name_pr = GetShortNameForGOST(arm_arr(i)(j, col_pr_gost_prof))
                    If InStr(1, name_pr, "Лист") Then
                        naen_plate = SpecMetallPlate(arm_arr(i)(j, col_pr_prof), arm_arr(i)(j, col_pr_naen), length_pos, weight_pm, arm_arr(i)(j, col_chksum))
                        length_pos = naen_plate(2)
                        weight_ed = naen_plate(4)
                    Else
                        length_pos = Round_w(arm_arr(i)(j, col_pr_length) / 1000, 3)
                        weight_ed = arm_arr(i)(j, col_pr_weight) * length_pos
                    End If
                    '---------------------
                    pm = False: If InStr(arm_arr(i)(j, col_chksum), "lpm") > 0 Then pm = True
                    If this_sheet_option.Item("pr_pm") Or pm Then
                        w_pos = Round_w(weight_ed * k_zap_total, n_round_w) * qty / nSubPos
                    Else
                        'Запас для одиночных элементов применяется только если это включено в setting
                        If zap_only_mp Then
                            w_pos = Round_w(weight_ed, n_round_w) * qty / nSubPos
                        Else
                            w_pos = Round_w(weight_ed * k_zap_total, n_round_w) * qty / nSubPos
                        End If
                    End If
                    '-----------------
                    'w_pos = Round_w(weight_ed * k_zap_total, n_round_w) * qty / nSubPos
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
tfunctime = functime("Spec_KZH_заполнение", tfunctime)
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
    If sum_row Then
        pos_out(n_row + sum_row, 1) = "Итого"
        If this_sheet_option.Item("qty_one_subpos") Then pos_out(n_row + sum_row, 1) = pos_out(n_row + sum_row, 1) + ", на все"
        For i = 2 To UBound(pos_out, 2)
            pos_out(n_row + sum_row, i) = 0
            For j = 6 To n_row
                nSubPos = pos_out_nsubpos(j, 2)
                If pos_out(j, i) <> "-" Then pos_out(n_row + sum_row, i) = pos_out(n_row + sum_row, i) + pos_out(j, i) * nSubPos
            Next
        Next
    End If
tfunctime = functime("Spec_KZH_сумма", tfunctime)
    Spec_KZH = pos_out
End Function

Function Spec_POL(ByRef all_data As Variant) As Variant
    out_data = all_data(1)
    rules = all_data(2)
    rules_mod = all_data(3)
    'erase all_data
    If IsEmpty(out_data) Then
        Spec_POL = Empty
        Exit Function
    End If
    isrim = 0
    Set zone = CreateObject("Scripting.Dictionary")
    zone.comparemode = 1
    un_n_zone = ArrayUniqValColumn(out_data, col_s_numb_zone, 2)
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
    un_pol = ArrayUniqValColumn(pol, col_s_type_pol, 2)
    n_type_pol = UBound(un_pol, 1)
    Dim pos_out(): ReDim pos_out(n_type_pol, 4)
    n_row_tot = 0
    For i = 1 To n_type_pol
        un_pol(i) = ConvTxt2Num(un_pol(i))
    Next i
    un_pol = ArraySort(un_pol, 1, 2)
    For i = 1 To n_type_pol
        un_pol(i) = ConvNum2Txt(un_pol(i))
    Next i
    For j = 1 To n_type_pol
        type_pol = un_pol(j)
        t_pol = ArraySelectParam(pol, type_pol, col_s_type_pol)
        t_un_zone = ArrayUniqValColumn(t_pol, col_s_numb_zone, 2)
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
        add_perim = 0
        If del_dor_perim Then add_perim = add_perim - door_len
        If del_freelen_perim Then add_perim = add_perim - free_len
        If add_holes_perim Then add_perim = add_perim + perim_hole
        pol_perim = pol_perim + add_perim
        'А теперь - магия. Что делать если периметр меньше 0?
        If pol_perim < 0.01 Then pol_perim = pol_perim_el + add_perim
        If pol_perim < 0.01 Then pol_perim = wall_len + add_perim
        If pol_perim < 0.01 Then pol_perim = perim_total + add_perim
        If pol_perim < 0.01 Then pol_perim = perim_total 'Если и это не помогает - чистите карму.
        t_zone = vbNullString
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
    tt = ConvTxt2Num(this_sheet_option.Item("k_zap"))
    If IsNumeric(tt) Then
        If tt > 1 And tt < 2 Then
            k_zap_total = tt
        Else
            k_zap_total = 1
        End If
    End If
    this_sheet_option.Item("k_zap") = CStr(k_zap_total)
End Function

Function Spec_Select(ByVal lastfilespec As String, ByVal suffix As String, Optional quiet As Boolean = False) As String
    If Not isINIset Then r = INISet()
    If SpecGetType(lastfilespec) = 7 Then
        nm = Split(lastfilespec, "_")(0) & suffix
    Else
        nm = lastfilespec & suffix
    End If
    type_spec = SpecGetType(nm)
    If quiet Then
        Set sheet_option_param_tmp = OptionSheetGet(nm)
        If sheet_option_param_tmp.Item("defult") = True Then
            Set sheet_option_param_tmp = OptionSheetGet(lastfilespec)
            If sheet_option_param_tmp.Item("defult") = True Then
                MsgBox ("Для спецификации " & nm & " не найдены сохранённые параметры. Можно задать их  при ручнов выводе или на листе с содержанием.")
            End If
        End If
        Set this_sheet_option = sheet_option_param_tmp
    Else
        If type_spec <> 11 And type_spec <> 12 Then
            r = UserForm1.show_form(nm)
        Else
            Set this_sheet_option = OptionGetForm(nm)
        End If
    End If
    If type_spec <> 11 And type_spec <> 12 Then
        this_sheet_option.Item("title_on") = True
    Else
        this_sheet_option.Item("title_on") = False
    End If
    r = SetKzap()
    Set set_sheet_option = this_sheet_option
    'Записываем данные для созданного листа
    tdate = Right$(str(DatePart("yyyy", Now)), 2) & str(DatePart("m", Now)) & str(DatePart("d", Now))
    stamp = tdate + "/" + str(DatePart("h", Now)) + str(DatePart("n", Now)) + str(DatePart("s", Now))
    this_sheet_option.Item("date") = stamp
    r = OptionSheetWrite(nm, this_sheet_option)
    'Записываем данные для файла(строки), выбранной при создании
    If lastfilespec <> nm Then
        r = OptionSheetWrite(lastfilespec, this_sheet_option)
    End If
    'Если есть файл _спец, то и для него запишем
    If InStr(lastfilespec, "_спец") = 0 Then
        nm_manual = Split(lastfilespec, "_")(0) & "_спец"
        If SheetExist(nm_manual) Then r = OptionSheetWrite(nm_manual, this_sheet_option)
    End If
    
    If Not SheetCheckName(nm) Then
        r = LogWrite(lastfilespec, suffix, "Ошибка имени листа или файла")
        If Not (quiet) Then MsgBox ("Данные отсутвуют")
        Exit Function
    End If
    type_spec = SpecGetType(nm)
    If type_spec = 1 Then this_sheet_option.Item("qty_one_floor") = False
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
        File = ArraySelectParam(listFile, Split(lastfilespec, "_")(0) & "_разб", 1)
        If Not IsEmpty(File) Then
            split_data = VedSplitFile(Split(lastfilespec, "_")(0))
            flag_split = True
        End If
    End If
    pos_out_all = Empty
    msg_zap_mat = vbNullString
    If ignore_zap_material And type_spec <> 4 Then msg_zap_mat = msg_zap_mat & vbLf & "Запас на раскрой материала не учитывается"
    If zap_only_mp And type_spec <> 4 Then msg_zap_mat = msg_zap_mat & vbLf & "!!! Запас применяется только к элементам, выводимым в п.м. (арматуре, прокату, изделиям) и материалам !!!"
    defult_add = this_sheet_option.Item("arr_subpos_add")(1) = "*" And UBound(this_sheet_option.Item("arr_subpos_add")) = 1
    defult_del = Trim(this_sheet_option.Item("arr_subpos_del")(1)) = "" And UBound(this_sheet_option.Item("arr_subpos_del")) = 1
    
    defult_type_add = this_sheet_option.Item("arr_typeKM_add")(1) = "*" And UBound(this_sheet_option.Item("arr_typeKM_add")) = 1
    defult_type_del = Trim(this_sheet_option.Item("arr_typeKM_del")(1)) = "" And UBound(this_sheet_option.Item("arr_typeKM_del")) = 1
    If Not defult_add Or Not defult_del Then msg_zap_mat = msg_zap_mat & vbLf & "--- Включена фильтрация по сборкам ---"
    If Not defult_type_add Or Not defult_type_del Then msg_zap_mat = msg_zap_mat & vbLf & "--- Включена фильтрация по типам конструкций КМ ---"
    Dim pos_zag()
    Select Case type_spec
        Case 1, 2, 3, 13
            If Not (quiet) Then MsgBox ("Коэффицент запаса для объёма, площади и длин " & ConvNum2Txt(k_zap_total) & msg_zap_mat)
            pos_out = Spec_AS(all_data, type_spec)
        Case 4
            If Not (quiet) Then MsgBox ("Коэффицент запаса для веса и площади " & ConvNum2Txt(k_zap_total) & msg_zap_mat)
            pos_out = Spec_KM(all_data)
        Case 5
            If Not (quiet) Then MsgBox ("Коэффицент запаса для веса " & ConvNum2Txt(k_zap_total) & msg_zap_mat)
            pos_out = Spec_KZH(all_data)
        Case 11
            If Not (quiet) Then MsgBox ("Коэффицент запаса площади отделки -" & ConvNum2Txt(k_zap_total))
            'Проверка возможности разделения на типы (если они заданы)
            If this_sheet_option.Item("otd_by_type") Then
                zone_el = ArraySelectParam(all_data(1), "ЗОНА", col_s_type)
                flag = Empty
                If Not IsEmpty(zone_el) Then
                    For jj = LBound(zone_el, 1) To UBound(zone_el, 1)
                        is_type_otd = zone_el(1, col_s_type_otd)
                        If is_type_otd = 0 Or Len(is_type_otd) = 0 Then
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
            out_data = Spec_WIN(all_data)
            If Not IsEmpty(out_data) Then
                pos_out = out_data(1)
                If Not IsEmpty(out_data(2)) Then
                    suffix_perem = vbNullString
                    shname = Replace(nm, "_", vbNullString) + suffix_perem
                    r = Spec_OUT(out_data(2), shname, suffix_perem, quiet)
                End If
            End If
        Case 25
            pos_out = Spec_RSK(all_data)
    End Select
    If Not IsEmpty(pos_out_all) Then pos_out = pos_out_all
    If flag_split = False Or (delim_by_sheet = False And flag_split = True) Then Spec_Select = Spec_OUT(pos_out, nm, suffix, quiet)
    r = print_functime()
End Function

Function VedFinByDraft(ByRef mat_fin As String, ByRef mat_draft As String)
 'Если в черновой отделке есть % - добавит в чистовую разделение по типам
 'Если в чистовой % - разделения не будет
    'Проверяем наличие модификатора для чистовой отделки
    fin_modify = ""
    If InStr(mat_draft, "%") And InStr(mat_fin, "%") = 0 Then
        n = Split(mat_draft, ";")
        For i = 0 To UBound(n)
            If InStr(n(i), "%") Then
                fin_modify = Trim(Replace(n(i), "%", ""))
                mat_draft = Replace(mat_draft, n(i), "")
                If (Right(Trim(mat_draft), 1) = UCase(";")) Then
                    mat_draft = Left(mat_draft, Len(mat_draft) - 1)
                End If
                Exit For
            End If
        Next i
    End If
    If mat_fin <> "---" And Len(mat_fin) > 1 And Len(fin_modify) > 0 Then mat_fin = mat_fin + "%" + fin_modify
End Function


Function VedAddAreaGR(ByVal area As Double, ByVal mat_fin As String, ByVal type_constr As String, ByVal type_name As String, ByVal mat_draft As String, ByRef rules_mod As Variant, ByRef materials_by_type As Variant, Optional ByVal num As String) As Long
    If area < 0.001 Then
        VedAddAreaGR = 0
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
    flag_fin = 1
    flag_draft = 1
    'Если есть черновая отделка - запишем её
    If Len(mat_draft) > 1 Then
        'Если в названии черновой отделки стоит = - чистовая отделка не нужна
        If InStr(mat_draft, "=") > 0 Then
            mat_draft = Trim$(Left$(mat_draft, Len(mat_draft) - 1))
            flag_fin = 0
        End If
    Else
        flag_draft = 0
    End If
    key_total_area = "tot;" + type_constr
    If Not materials_by_type.Exists(key_total_area) Then
        materials_by_type.Item(key_total_area) = 0
    End If
    materials_by_type.Item(key_total_area) = materials_by_type.Item(key_total_area) + area
    
    num = Replace(num, ",", ".")
    mat_fin = Replace(mat_fin, "<>", vbNullString)
    mat_draft = Replace(mat_draft, "<>", vbNullString)
    'Если в названии чистовой отделки стоит --- или УНИВЕРСАЛЬНЫЙ - чистовая отделка не нужна
    If InStr(mat_fin, "--") > 0 Or InStr(mat_fin, "УНИВЕРСАЛЬНОЕ") > 0 Or mat_fin = "0" Or InStr(mat_fin, "е задан") > 0 Then flag_fin = 0
    If mat_draft = "0" Or InStr(mat_draft, "е задан") > 0 Then flag_draft = 0
    If flag_draft Then
        'Черновая отделка с учётом исключений
        r = VedFinByDraft(mat_fin, mat_draft)
        all_name_mat = Split(Replace(VedModMat(Replace(mat_fin, fin_str, vbNullString), mat_draft, rules_mod), "@", ";"), ";")
        mat_fin = Replace(mat_fin, "%", " ")
        For Each mat In all_name_mat
            If InStr(mat, "%") = 0 Then
                mat = Trim$(mat)
                If Len(mat) > 0 Then
                    materials_by_type.Item(type_name + type_constr + mat) = materials_by_type.Item(type_name + type_constr + mat) + area
                    If InStr(type_constr, ";pot;") > 0 And zonenum_pot Then
                        If materials_by_type.Exists(type_name + ";pot_num" + mat) Then
                            materials_by_type.Item(type_name + ";pot_num" + mat) = materials_by_type.Item(type_name + ";pot_num" + mat) + ";" + Trim$(num)
                        Else
                            materials_by_type.Item(type_name + ";pot_num" + mat) = Trim$(num)
                        End If
                    End If
                End If
            End If
        Next
    End If
    If flag_fin Then
        'Чистовая отделка
        all_name_mat = Split(Replace(mat_fin, "@", ";"), ";")
        For ni = 0 To UBound(all_name_mat)
            mat = Trim$(all_name_mat(ni))
            If InStr(mat, "%") = 0 Then
                materials_by_type.Item(type_name + type_constr + mat) = materials_by_type.Item(type_name + type_constr + mat) + area
                If InStr(type_constr, ";pot;") > 0 And zonenum_pot Then
                    If materials_by_type.Exists(type_name + ";pot_num" + mat) Then
                        materials_by_type.Item(type_name + ";pot_num" + mat) = materials_by_type.Item(type_name + ";pot_num" + mat) + ";" + Trim$(num)
                    Else
                        materials_by_type.Item(type_name + ";pot_num" + mat) = Trim$(num)
                    End If
                End If
            End If
        Next ni
    End If
tfunctime = functime("VedAddAreaGR", tfunctime)
    VedAddAreaGR = flag_draft + flag_fin
End Function
Function Spec_OUT(ByRef pos_out As Variant, ByVal nm As String, ByVal suffix As String, ByVal quiet As Boolean) As String
    If IsEmpty(pos_out) Then
        r = LogWrite(nm, suffix, "Данные отсутвуют")
        If Not (quiet) Then MsgBox ("Данные отсутвуют")
        Exit Function
    End If
    type_spec = SpecGetType(nm)
    Dim pos_out_title
    ReDim pos_out_title(1, UBound(pos_out, 2))
    If this_sheet_option.Exists("title") Then
        If this_sheet_option.Item("title").Exists(ttspec) Then pos_out_title(1, 1) = this_sheet_option.Item("title").Item(ttspec)
    End If
    If SheetExist(nm) Then
        If IsEmpty(pos_out_title(1, 1)) Then pos_out_title(1, 1) = wbk.Worksheets(nm).Cells(1, 1).Value
        wbk.Worksheets(nm).Activate
        wbk.Worksheets(nm).Cells.Clear
    Else
        wbk.Worksheets.Add.Name = nm
    End If
    If this_sheet_option.Item("title_on") = True Then pos_out = ArrayCombine(pos_out_title, pos_out)
    wbk.Worksheets(nm).Move After:=wbk.Sheets(wbk.Sheets.Count)
    str_kzap = "k=" + ConvNum2Txt(k_zap_total)
    r = FormatTable(nm, pos_out, str_kzap)
    r = FormatTable(nm)
    r = LogWrite(nm, suffix, "ОК")
    If inx_on_new And Not (quiet) Then
        r = SheetIndex()
        wbk.Worksheets(nm).Activate
    End If
    Spec_OUT = nm
End Function
Function VedAddArea(ByRef zone As Variant, ByRef materials As Variant, ByVal mat_draft As String, ByVal mat_fin As String, ByVal num As String, ByVal area_mat As Double, ByVal rules_mod As Variant, Optional ByVal perim As Double = 0, Optional ByVal h_pan As Double = 0) As Long
Dim tfunctime As Double
tfunctime = Timer
    type_ = "Низ лестничных маршей: "
    If this_sheet_option.Item("separate_material") Then
        razd = ";"
        mat_fin = Trim$(mat_fin)
    Else
        razd = "&"
        mat_fin = " " + mat_fin
    End If
    r = VedFinByDraft(mat_fin, mat_draft)
    mat_fin = Replace(mat_fin, "@", ";a@")
    If Trim$(mat_fin) = "0" Then mat_fin = "---"
    mat_draft = VedModMat(Replace(mat_fin, fin_str, vbNullString), mat_draft, rules_mod)
    mat_fin = Replace(mat_fin, "%", " ")
    mat_draft = Trim$(mat_draft)
    mat_draft = "b@" & Replace(mat_draft, razd, ";b@")
    mat_draft = Replace(mat_draft, "@ ", "@")
    mat_draft = Replace(mat_draft, "<>", vbNullString)
    mat_fin = Replace(mat_fin, "<>", vbNullString)
    If InStr(mat_draft, "=") > 0 Then
        name_mat = Array(Trim$(Left$(mat_draft, Len(mat_draft) - 1)))
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
        key_total_area = "tot;" + Split(num, ";")(1)
        If Not zone.Exists(key_total_area) Then
            zone.Item(key_total_area) = 0
        End If
        zone.Item(key_total_area) = zone.Item(key_total_area) + area_mat
        For Each mat In name_mat
            mat = Trim$(mat)
            naen_mat = Trim$(Replace(Replace(mat, "b@", vbNullString), "a@", vbNullString))
            If Left$(naen_mat, 1) <> vbNullString Then naen_mat = StrConv(Left$(naen_mat, 1), vbUpperCase) + Right$(naen_mat, Len(naen_mat) - 1)
            If Left$(naen_mat, 1) = ";" Then naen_mat = Trim$(Right$(naen_mat, Len(naen_mat) - 1))
            If InStr(naen_mat, "%") = 0 And naen_mat <> vbNullString And Not IsEmpty(naen_mat) And InStr(naen_mat, "--") = 0 And InStr(naen_mat, "УНИВЕРСАЛЬНОЕ") = 0 And InStr(naen_mat, "е задан") = 0 Then
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
                naen_mat = VedCleanName(naen_mat)
                mat = VedCleanName(mat)
                If InStr(num, ";pot") > 0 Then
                    mat = mat + ";pot"
                    fname = vbNullString: If InStr(mat_clear, type_) = 0 Then fname = "Потолок: "
                    naen_mat = fname + naen_mat
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
tfunctime = functime("VedAddArea", tfunctime)
    VedAddArea = flag
End Function

Function VedCleanName(ByVal mat As String) As String
    If InStr(mat, "%%") > 0 Then
        end_p = InStr(mat, "%%") + 1
        start_p = InStr(mat, fin_str + "До ") - 1
        If start_p = -1 Then start_p = InStr(mat, fin_str_sec + "Выше ") - 1
        txt = Left$(mat, end_p)
        end_txt = Right$(txt, Len(txt) - start_p)
        mat = Replace(mat, end_txt, fin_str)
        ll = 1
    End If
    VedCleanName = mat
End Function

Function Spec_CONC(ByRef all_data As Variant, Optional ByVal un_parent As Variant = Empty) As Boolean
    floor_txt = "all_floor"
    all_bet = ArraySelectParam_2(all_data, t_mat, col_type_el, "?етон?", col_m_naen)
    If IsEmpty(concrsubpos) Then Set concrsubpos = CreateObject("Scripting.Dictionary")
    flag = False
    concrsubpos.Item("bet_qty") = 0
    If IsEmpty(un_parent) Then un_parent = pos_data.Item(floor_txt).Item("parent").keys()
    For Each subpos In un_parent
        all_bet_subpos = ArraySelectParam(all_bet, subpos, col_sub_pos)
        concrsubpos.Item(subpos & "_bet_qty") = 0
        If Not IsEmpty(all_bet_subpos) Then
            nSubPos = GetNSubpos(subpos, 1, floor_txt)
            n_mat = UBound(all_bet_subpos, 1)
            spec_subpos = SpecMaterial(all_bet_subpos, n_mat, 2, nSubPos)
            For j = 1 To UBound(spec_subpos, 1)
                bet = spec_subpos(j, 3)
                If InStr(bet, "(") > 0 And InStr(bet, ")") > 0 And clear_bet_name Then
                    str_to_del = Mid$(bet, InStr(bet, "("), InStr(bet, ")"))
                    bet = Trim$(Replace(bet, str_to_del, vbNullString))
                End If
                qty = Round_w(spec_subpos(j, 4), n_round_wkzh)
                concrsubpos.Item(subpos & "_" & bet & "_qty") = qty
                concrsubpos.Item(subpos & "_bet_qty") = concrsubpos.Item(subpos & "_bet_qty") + qty
                concrsubpos.Item("bet_qty") = concrsubpos.Item("bet_qty") + qty
                concrsubpos.Item("tot_" & bet) = concrsubpos.Item("tot_" & bet) + qty
            Next j
            concrsubpos.Item(subpos & "_bet") = ArrayUniqValColumn(spec_subpos, 3)
            flag = True
        End If
    Next
    Spec_CONC = flag
End Function

Function Spec_RSK(ByRef all_data As Variant) As Variant
    all_arm = ArraySelectParam_2(all_data, t_arm, col_type_el)
    Set rsk_arm = CreateObject("Scripting.Dictionary")
    Set usage_arm = CreateObject("Scripting.Dictionary")
    un_class = ArrayUniqValColumn(all_arm, col_klass)
    un_subpos = ArraySort(pos_data.Item("all_floor").Item("parent").keys())
    txt_2_arr = vbNullString
    n_row_out = 0
    If IsEmpty(un_class) Then
        Spec_RSK = Empty
        Exit Function
    End If
    For i = 1 To UBound(un_class)
        class = un_class(i)
        Set class_dict = CreateObject("Scripting.Dictionary")
        diam_arr = ArraySelectParam_2(all_arm, class, col_klass)
        un_diam = ArrayUniqValColumn(diam_arr, col_diametr)
        For j = 1 To UBound(un_diam)
            diametr = un_diam(j)
            Set diametr_dict = CreateObject("Scripting.Dictionary")
            leght_arr = ArraySelectParam_2(diam_arr, diametr, col_diametr)
            For k = 1 To UBound(leght_arr, 1)
                qty = leght_arr(k, col_qty)
                leght = Round(leght_arr(k, col_length), 0)
                tkey = class + "_" + CStr(diametr)
                tparent = leght_arr(k, col_parent)
                If tparent = "-" Then tparent = leght_arr(k, col_sub_pos)
                
                txt_2_arr = tparent
                If leght_arr(k, col_parent) <> "-" Then txt_2_arr = txt_2_arr + "-" + leght_arr(k, col_sub_pos)
                txt_2_arr = txt_2_arr + "-" + leght_arr(k, col_pos)
                
                If leght > lenght_ed_arm Then
                    If Not diametr_dict.Exists(lenght_ed_arm) Then
                        diametr_dict.Item(lenght_ed_arm) = 0
                        n_row_out = n_row_out + 1
                    End If
                    If Not usage_arm.Exists(tkey + "_" + CStr(lenght_ed_arm)) Then
                        usage_arm.Item(tkey + "_" + CStr(lenght_ed_arm)) = txt_2_arr
                    Else
                        usage_arm.Item(tkey + "_" + CStr(lenght_ed_arm)) = usage_arm.Item(tkey + "_" + CStr(lenght_ed_arm)) + "; " + txt_2_arr
                    End If
                    qty = Int(leght / lenght_ed_arm)
                    diametr_dict.Item(lenght_ed_arm) = diametr_dict.Item(lenght_ed_arm) + qty
                    If Not usage_arm.Exists(tkey + "_" + CStr(lenght_ed_arm) + "_" + leght_arr(k, col_sub_pos)) Then usage_arm.Item(tkey + "_" + CStr(lenght_ed_arm) + "_" + leght_arr(k, col_sub_pos)) = 0
                    usage_arm.Item(tkey + "_" + CStr(lenght_ed_arm) + "_" + tparent) = usage_arm.Item(tkey + "_" + CStr(lenght_ed_arm) + "_" + tparent) + qty
                    leght = leght - qty * lenght_ed_arm
                    qty = 1
                End If
                tkey = tkey + "_" + CStr(leght)
                If Not usage_arm.Exists(tkey) Then
                    usage_arm.Item(tkey) = txt_2_arr
                Else
                    usage_arm.Item(tkey) = usage_arm.Item(tkey) + "; " + txt_2_arr
                End If
                If Not usage_arm.Exists(tkey + "_" + tparent) Then usage_arm.Item(tkey + "_" + tparent) = 0
                usage_arm.Item(tkey + "_" + tparent) = usage_arm.Item(tkey + "_" + tparent) + qty
                If Not diametr_dict.Exists(leght) Then diametr_dict.Item(leght) = 0
                diametr_dict.Item(leght) = diametr_dict.Item(leght) + qty
                n_row_out = n_row_out + 1
            Next k
            Set class_dict.Item(diametr) = diametr_dict
        Next j
        Set rsk_arm.Item(class) = class_dict
    Next i
    Dim pos_out(): ReDim pos_out(n_row_out + 2, 5 + UBound(un_subpos))
    n_out = 1
    pos_out(n_out, 1) = "Выборка арматуры на " + Join(un_subpos, ", ")
    n_out = 2
    pos_out(n_out, 1) = "Класс"
    pos_out(n_out, 2) = "Диаметр"
    pos_out(n_out, 3) = "Длина, мм."
    pos_out(n_out, 4) = "Кол-во всего"
    For i = 1 To UBound(un_subpos)
        pos_out(n_out, 4 + i) = un_subpos(i)
    Next i
    pos_out(n_out, 5 + UBound(un_subpos)) = "Использование (Марка-Позиция)"
    For Each class In rsk_arm.keys
        For Each diametr In rsk_arm(class).keys
            leght_arr = ArraySort(rsk_arm(class)(diametr).keys)
            For Each leght In leght_arr
                qty = rsk_arm.Item(class)(diametr)(leght)
                tkey = class + "_" + CStr(diametr) + "_" + CStr(leght)
                n_out = n_out + 1
                pos_out(n_out, 1) = class
                pos_out(n_out, 2) = diametr
                pos_out(n_out, 3) = leght
                pos_out(n_out, 4) = qty
                pos_out(n_out, 5 + UBound(un_subpos)) = usage_arm.Item(tkey)
                For i = 1 To UBound(un_subpos)
                    pos_out(n_out, 4 + i) = usage_arm.Item(tkey + "_" + un_subpos(i))
                Next i
            Next
        Next
    Next
    pos_out = ArrayRedim(pos_out, n_out)
    Spec_RSK = pos_out
End Function


Function Spec_NRM(ByRef all_data As Variant) As Variant
    floor_txt = "all_floor"
    this_sheet_option.Item("qty_one_subpos") = False
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
    subpos = pos_data.Item(floor_txt).Item("parent").keys()
    Dim pos_out_norm(): ReDim pos_out_norm(UBound(subpos, 1) + 3, 5)
    n_out = 1
    pos_out_norm(n_out, 1) = "Поз."
    pos_out_norm(n_out, 2) = "Марка бетона"
    pos_out_norm(n_out, 3) = "Объём" & vbLf & "бетона, куб.м."
    pos_out_norm(n_out, 4) = "Вес арматуры, кг."
    pos_out_norm(n_out, 5) = "Расход, кг/куб.м"
    sum_bet = 0: sum_arm = 0
    For Each subpos In pos_data.Item(floor_txt).Item("parent").keys()
        v_bet = 0
        v_arm = 0
        naen_bet = vbNullString
        If concrsubpos.Exists(subpos & "_bet_qty") Then
            If concrsubpos.Exists(subpos & "@бет") Then
                bet_ank = concrsubpos.Item(subpos & "@бет")
                For Each sub_bet In concrsubpos.keys()
                    If InStr(sub_bet, "_") > 0 And Right$(sub_bet, 4) = "_qty" And InStr(sub_bet, "bet") = 0 Then
                        subb = Split(sub_bet, "_")
                        If subb(0) = subpos And InStr(subb(1), bet_ank) > 0 Then
                            v_bet = v_bet + concrsubpos.Item(sub_bet)
                            naen_bet = naen_bet & vbLf & subb(1)
                        End If
                    End If
                Next
            Else
                If Not IsEmpty(concrsubpos.Item(subpos & "_bet")) Then
                    v_bet = concrsubpos.Item(subpos & "_bet_qty")
                    For Each nbet In concrsubpos.Item(subpos & "_bet")
                        naen_bet = naen_bet & vbLf & nbet
                    Next
                End If
            End If
        End If
        v_bet = Round(v_bet, 0)
        If v_bet > 0 Then
            For k = 1 To UBound(pos_out_arm, 1)
                If Left$(pos_out_arm(k, 1), Len(subpos)) = subpos Then v_arm = pos_out_arm(k, n_col_arm)
            Next k
        End If
        If IsNumeric(v_arm) Then
            v_arm = Round(v_arm, 0)
        Else
            v_arm = 0
        End If
        n_out = n_out + 1
        pos_out_norm(n_out, 1) = subpos
        pos_out_norm(n_out, 2) = naen_bet
        pos_out_norm(n_out, 3) = v_bet
        pos_out_norm(n_out, 4) = v_arm
        If v_bet > 0 And v_arm > 0 Then
            pos_out_norm(n_out, 5) = Round(v_arm / v_bet, 0)
        Else
            pos_out_alarm = vbNullString
            If v_arm <= 0 Then pos_out_alarm = pos_out_alarm + "!Арматуры нет!"
            If v_bet <= 0 Then pos_out_alarm = pos_out_alarm + " !Бетона нет!"
            pos_out_norm(n_out, 5) = pos_out_alarm
        End If
        
        sum_bet = sum_bet + v_bet
        sum_arm = sum_arm + v_arm
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
    'erase all_data
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
    tot_area_delta_h_zone = 0
    h_pot = 0
    type_pot_zone = vbNullString
    type_pol_zone = vbNullString
    mwall_up_zone = vbNullString
    param_zone = vbNullString
    Set materials_by_type = CreateObject("Scripting.Dictionary")
    Set materials = CreateObject("Scripting.Dictionary")
    materials_by_type.comparemode = 1
    materials.comparemode = 1
    spec_type = 1 'И лестницы, и полы
    If UBound(out_data, 2) = max_col_type_2 Then spec_type = 2 'Без лестниц
    If UBound(out_data, 2) = max_col_type_1 Then spec_type = 3 'Только зоны
    '------- Предварительная выборка элементов --------------------
    zones_el_all = ArraySelectParam_2(out_data, "ЗОНА", col_s_type)
    walls_el_all = ArraySelectParam_2(out_data, "СТЕНА", col_s_type)
    pots_el_all = ArraySelectParam_2(out_data, "Потолок", col_s_type_el)
    pols_el_all = ArraySelectParam_2(out_data, "Пол", col_s_type_el)
    '--------------------------------------------------------------
    un_type_otd = ArrayUniqValColumn(zones_el_all, col_s_type_otd)
    materials_by_type.Item("list") = un_type_otd
    For Each type_name In un_type_otd
        If InStr(type_name, "Без отделки") = 0 Then
            If IsNumeric(type_name) Then type_name = CStr(type_name)
            zone_bytype_el = ArraySelectParam_2(zones_el_all, type_name, col_s_type_otd) 'Список зон с этим типом отделки
            un_n_zone_type = ArrayUniqValColumn(zone_bytype_el, col_s_numb_zone) 'Список номеров зон
            materials_by_type.Item(type_name + ";zone") = un_n_zone_type
            '------- Предварительная выборка элементов --------------------
            zones_el = ArraySelectParam_2(zones_el_all, un_n_zone_type, col_s_numb_zone)
            walls_el = ArraySelectParam_2(walls_el_all, un_n_zone_type, col_s_numb_zone)
            pots_el = ArraySelectParam_2(pots_el_all, un_n_zone_type, col_s_numb_zone)
            pols_el = ArraySelectParam_2(pols_el_all, un_n_zone_type, col_s_numb_zone)
            If otd_version = 2 Then
                un_zone_h = ArrayUniqValColumn(zones_el, col_s_h_pot_zone)
                If IsEmpty(un_zone_h) Then
                    n_zone_h = 0
                Else
                    n_zone_h = UBound(un_zone_h, 1)
                End If
                If n_zone_h > 1 Or un_zone_h(1) > 0 Then
                    is_delta_h = True
                Else
                    is_delta_h = False
                End If
            Else
                is_delta_h = False
            End If
            '--------------------------------------------------------------
            For Each num In un_n_zone_type
                ' Теперь для каждой зоны с этим типом отделки считаем всё что можем
                is_error = 0
                If IsNumeric(num) Then num = CStr(num)
                zone_el = ArraySelectParam_2(zone_bytype_el, num, col_s_numb_zone)
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
                        fin_column = Replace(fin_column, "<>", vbNullString)
                    Else
                        column_is_wall = False
                    End If
                    manual_column = False
                    ' ------------------------------------------
                    area_total = zone_el(1, col_s_area_zone)
                    perim_total = zone_el(1, col_s_perim_zone) / 1000
                    perim_hole = zone_el(1, col_s_perimhole_zone) / 1000
                    h_zone = zone_el(1, col_s_h_zone) / 1000
                    free_len = zone_el(1, col_s_freelen_zone) / 1000
                    h_pan = zone_el(1, col_s_hpan_zone) / 1000
                    materials_by_type.Item(type_name + ";zoneh;") = Application.WorksheetFunction.Max(materials_by_type.Item(type_name + ";zoneh;"), h_zone)
                    free_len_wall = 0
                    tfin_pot = vbNullString
                    If otd_version = 2 Then
                        h_pot = zone_el(1, col_s_h_pot_zone) / 1000
                        type_pot_zone = zone_el(1, col_s_type_pot_zone)
                        type_pol_zone = zone_el(1, col_s_type_pol_zone)
                        If zone_el(1, col_s_mwall_up_zone) = fin_wall Or zone_el(1, col_s_mwall_up_zone) = fin_column Or h_pot <= 0.0001 Then
                            h_pot = h_zone
                            tfin_up_zone = zone_el(1, col_s_mwall_up_zone)
                        Else
                            If delim_zone_fin Then
                                tfin_up_zone = "Выше " + CStr(h_pot) + "м.:%%" + CStr(zone_el(1, col_s_mwall_up_zone))
                            Else
                                tfin_up_zone = "Выше подвесного потолка:%%" + CStr(zone_el(1, col_s_mwall_up_zone))
                            End If
                        End If
                        fin_up_zone = fin_str_sec + Replace(tfin_up_zone, "@", "; ")
                        param_zone = zone_el(1, col_s_param_zone)
                        column_perim_total = GetZoneParam(zone_el(1, col_s_param_zone), "Cp")
                        If Not IsEmpty(column_perim_total) Then
                            manual_column = True
                            column_perim_total = value_param / 1000
                            If Abs(column_perim_total) < 0.002 Then column_perim_total = -1
                        End If
                        wall_razd = ArraySelectParam_2(walls_el, "Разделитель", col_s_mat_wall)
                        If Not IsEmpty(wall_razd) Then
                            For wi = 1 To UBound(wall_razd, 1)
                                free_len_wall = free_len_wall + wall_razd(wi, col_s_freelen_zone) / 1000
                            Next wi
                        End If
                    End If
                    If h_pot < 1 And h_pot > 0 Then
                        r = LogWrite("Проверьте высоту помещения, должна быть в мм. - " + CStr(h_pot), "Ошибка", num)
                        is_error = is_error + 1
                        h_pot = h_pot * 1000
                    End If
                    If h_pot > h_zone Then
                        r = LogWrite("Проверьте высота потолка выше помещения. - " + CStr(h_pot), "Ошибка", num)
                        is_error = is_error + 1
                        h_pot = h_zone
                    End If
                    If h_pot < 0.01 Then h_pot = h_zone
                    delta_h_zone = h_zone - h_pot 'Полоска над потолком
                    If is_delta_h Then
                        If delim_zone_fin Then
                            fin_wall = fin_str + "До " + CStr(h_pot) + "м.:%%" + CStr(zone_el(1, col_s_mwall_zone))
                            fin_column = fin_str + "До " + CStr(h_pot) + "м.:%%" + CStr(zone_el(1, col_s_mcolumn_zone))
                        Else
                            fin_wall = fin_str + CStr(zone_el(1, col_s_mwall_zone))
                            fin_column = fin_str + CStr(zone_el(1, col_s_mcolumn_zone))
                        End If
                    End If
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
                    If free_len_wall > 0 Then free_len = free_len_wall
                    tot_area_wall_zone = tot_area_wall_zone + (perim_total - free_len) * h_zone
                    tot_area_zone = tot_area_zone + (perim_total - free_len) * h_zone
                    tot_perim_zone = tot_perim_zone + perim_total
                    ' --- Длины стен и дверей ---
                    wall = ArraySelectParam_2(walls_el, num, col_s_numb_zone)
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
                    If manual_column = False Then column_perim_total = perim_total - wall_len - free_len + perim_hole * (hole_in_zone = True)
                    If column_perim_total < 0 Then column_perim_total = 0
                    column_pan_area = column_perim_total * h_pan
                    column_pan_area_delta_h = column_perim_total * delta_h_zone
                    column_area = column_perim_total * (h_zone - h_pan - delta_h_zone)
                    tot_area_column = tot_area_column + column_area
                    tot_area_pan = tot_area_pan + column_pan_area
                    tot_area_delta_h_zone = tot_area_delta_h_zone + column_pan_area_delta_h
                    colm = VedNameMat("Колонны", "ЖБ", rules)
                    If column_is_wall Then
                        'Стены
                        r = VedAddAreaGR(column_area + column_pan_area + column_pan_area_delta_h, fin_column, ";wall;", type_name, vbNullString, rules_mod, materials_by_type, num)
                    Else
                        If column_area > 0.01 Then is_column = True
                        If column_pan_area > 0.01 Then is_pan = True
                        colm = VedNameMat("Колонны", "ЖБ", rules)
                        'Колонны
                        r = VedAddAreaGR(column_area, fin_column, ";column;", type_name, colm, rules_mod, materials_by_type, num)
                        'Панели на колоннах
                        r = VedAddAreaGR(column_pan_area, fin_pan, ";pan;", type_name, colm, rules_mod, materials_by_type, num)
                        'Колонны выше потолка
                        r = VedAddAreaGR(column_pan_area_delta_h, fin_up_zone, ";column;", type_name, colm, rules_mod, materials_by_type, num)
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
                            wall_up_area = 0
                            wall_c_len = 0
                            wall_c_area = 0
                            wall_c_up_area = 0
                            pan_area = 0
                            pan_c_area = 0
                            twall_area = 0
                            twall_h = 0
                            tpan_area = 0
                            tup_area = 0
                            wall_by_key = ArraySelectParam_2(wall, w, col_s_mat_wall)
                            For i = 1 To UBound(wall_by_key, 1)
                                twall_area = wall_by_key(i, col_s_area_wall)
                                If twall_area > 0 Then
                                    tdoor_len = wall_by_key(i, col_s_doorlen_zone) / 1000
                                    twall_len = wall_by_key(i, col_s_walllen_zone) / 1000
                                    If otd_version = 2 Then
                                        th_wall = wall_by_key(i, col_s_h_wall) / 1000
                                        If th_wall > h_zone Then th_wall = h_zone
                                    Else
                                        If twall_len > tdoor_len Then th_wall = twall_area / (twall_len - tdoor_len)
                                    End If
                                    If th_wall > h_pan Then
                                        If th_wall > h_pot Then
                                            tup_area = twall_len * delta_h_zone
                                        Else
                                            tup_area = 0
                                        End If
                                        tpan_area = (twall_len - tdoor_len) * h_pan
                                        twall_area = twall_area - tpan_area - tup_area
                                    Else
                                        If twall_len > tdoor_len Then
                                            tpan_area = twall_area
                                            twall_area = 0
                                            r = LogWrite("Панели на всю высоту стен? " + CStr(h_pan), vbNullString, num)
                                        Else
                                            tpan_area = 0
                                            twall_area = twall_area
                                            r = LogWrite("Стена полностью скрыта дверью? " + CStr(h_pan), vbNullString, num)
                                        End If
                                    End If
                                    If InStr(wall_by_key(i, col_s_layer), "Колонн") > 0 Then
                                        wall_c_area = wall_c_area + twall_area
                                        pan_c_area = pan_c_area + tpan_area
                                        wall_c_up_area = wall_c_up_area + tup_area
                                    Else
                                        wall_area = wall_area + twall_area
                                        pan_area = pan_area + tpan_area
                                        wall_up_area = wall_up_area + tup_area
                                    End If
                                End If
                            Next i
                            wall_area_zone = wall_area_zone + wall_area + wall_c_area + pan_c_area + pan_area + wall_up_area + wall_c_up_area
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
                            'Стены выше потолка
                            r = VedAddAreaGR(wall_up_area, fin_up_zone, ";wall;", type_name, w, rules_mod, materials_by_type, num)
                            'Колонны выше потолка
                            r = VedAddAreaGR(wall_c_up_area, fin_up_zone, ";column;", type_name, w, rules_mod, materials_by_type, num)
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
                    noPot = False
                    If spec_type < 3 Then
                        pot = ArraySelectParam_2(pots_el, num, col_s_numb_zone, "Потолок", col_s_type_el)
                        un_pot = ArrayUniqValColumn(pot, col_s_type_pol)
                        If Not IsEmpty(un_pot) Then
                            For Each p In un_pot
                                pot_by_type = ArraySelectParam_2(pot, p, col_s_type_pol)
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
'                                r = VedAddAreaGR(diff_area_pot, fin_pot, ";pot;", type_name, vbNullString, rules_mod, materials_by_type, num)
'                                r = LogWrite("Добавлена окраска" & num, "Ошибка", diff_area_pot)
'                                is_error = is_error + 1
'                            End If
                            If Abs(diff_area_pot) > 1 Then
                                r = LogWrite("Разница площади помещения и потолка в помещении " & num, "Ошибка", diff_area_pot)
                                is_error = is_error + 1
                            End If
                        Else
                            noPot = True
                        End If
                    Else
                        noPot = True
                    End If
                    If noPot Then
                        pot_perim = perim_total
                        If del_freelen_perim Then pot_perim = pot_perim - free_len
                        If add_holes_perim Then pot_perim = pot_perim + perim_hole
                        r = VedAddAreaGR(area_total, fin_pot, ";pot;", type_name, vbNullString, rules_mod, materials_by_type, num)
                        materials_by_type.Item(type_name + ";pot_perim;") = materials_by_type.Item(type_name + ";pot_perim;") + pot_perim
                    End If
                    '----------------------
                    '        ПОЛЫ
                    '----------------------
                    area_total_pol = 0
                    diff_area_pol = 0
                    If spec_type < 3 Then
                        pol = ArraySelectParam_2(pols_el, num, col_s_numb_zone)
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
                            zone_error.Item(num + "_pot_diff") = vbNullString
                        End If
                        If Abs(diff_area_pol) > 1 Then
                            zone_error.Item(num + "_pol_diff") = diff_area_pol
                        Else
                            zone_error.Item(num + "_pol_diff") = vbNullString
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
                    If (Left$(mat, len_find_pot) = type_name + ";pot;") Then
                        npot = npot + 1
                        ReDim Preserve all_mat_pot(npot)
                        all_mat_pot(npot) = Right$(mat, len_mat - len_find_pot)
                    End If
    
                    If (Left$(mat, len_find_wall) = type_name + ";wall;") Then
                        nwall = nwall + 1
                        ReDim Preserve all_mat_wall(nwall)
                        all_mat_wall(nwall) = Right$(mat, len_mat - len_find_wall)
                    End If
    
                    If (Left$(mat, len_find_column) = type_name + ";column;") Then
                        ncolumn = ncolumn + 1
                        ReDim Preserve all_mat_column(ncolumn)
                        all_mat_column(ncolumn) = Right$(mat, len_mat - len_find_column)
                    End If
    
                    If (Left$(mat, len_find_pan) = type_name + ";pan;") Then
                        npan = npan + 1
                        ReDim Preserve all_mat_pan(npan)
                        all_mat_pan(npan) = Right$(mat, len_mat - len_find_pan)
                    End If
                End If
            Next
            If npot > 0 Then
                materials_by_type.Item(type_name + ";pot") = ArrayUniqValColumn(all_mat_pot, 1)
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
    type_ = "Низ лестничных маршей: "
    n_col_out = 7
    If is_pan Then n_col_out = n_col_out + 3
    If is_column Then n_col_out = n_col_out + 2
    If show_surf_area And delim_by_sheet = True Then n_col_out = Application.WorksheetFunction.Max(n_col_out, 8)
    Dim pos_out(): ReDim pos_out(3400, n_col_out)
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
                    mat_clear = VedCleanName(mat)
                    area = Round_w(materials_by_type.Item(type_name + ";pot;" + mat) * k_zap_total, n_round_area)
                    If area > 0.001 Then
                        If InStr(mat, fin_str) > 0 Or InStr(mat, fin_str_sec) > 0 Then
                            sum_pot = sum_pot + area
                            fname = vbNullString: If InStr(mat_clear, type_) = 0 Then fname = "Потолок: "
                            mat_clear = Replace(Replace(Replace(mat_clear, "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString)
                            materials.Item(fname + mat_clear) = materials.Item(fname + mat_clear) + area
                        Else
                            fname = vbNullString: If InStr(mat_clear, type_) = 0 Then fname = "Потолок: "
                            mat_clear = Replace(Replace(Replace(mat_clear, "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString)
                            materials.Item(fname + mat_clear) = materials.Item(fname + mat_clear) + area
                        End If
                        num_zone = vbNullString
                        If Not IsEmpty(materials_by_type.Item(type_name + ";pot_num" + mat)) And zonenum_pot Then
                            num_zone = materials_by_type.Item(type_name + ";pot_num" + mat)
                            If InStr(num_zone, ";") > 0 Then
                                un_num_pot = ArrayUniqValColumn(Split(num_zone, ";"), 1)
                                num_zone = vbNullString
                                For Each nnum In un_num_pot
                                    If Len(num_zone) = 0 Then
                                        num_zone = nnum
                                    Else
                                        num_zone = num_zone + "; " + nnum
                                    End If
                                Next
                            End If
                            pos_out(n_row_p, 2) = num_zone
                        End If
                        pos_out(n_row_p, 3) = Replace(Replace(Replace(mat, "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString)
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
                num_zone = vbNullString
                For Each num In ArrayUniqValColumn(materials_by_type.Item(type_name + ";zone"), 1)
                    If IsNumeric(num) Then num = CStr(num)
                    num = Replace(num, ",", ".")
                    If Len(num_zone) = 0 Then
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
                    mat_clear = VedCleanName(mat)
                    area = materials_by_type.Item(type_name + ";wall;" + mat)
                    If area > 0.001 Then
                        If InStr(mat, fin_str) > 0 Or InStr(mat, fin_str_sec) > 0 Then sum_wall = sum_wall + Round_w(area * k_zap_total, n_round_area)
                        materials.Item(mat_clear) = materials.Item(mat_clear) + Round_w(area * k_zap_total, n_round_area)
                        pos_out(n_row_w, 5) = Replace(Replace(mat, "%%", vbNullString), fin_str, vbNullString)
                        pos_out(n_row_w, 5) = Replace(pos_out(n_row_w, 5), fin_str_sec, vbNullString)
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
                        mat_clear = VedCleanName(mat)
                        area = materials_by_type.Item(type_name + ";column;" + mat)
                        If area > 0.001 Then
                            If InStr(mat, fin_str) > 0 Or InStr(mat, fin_str_sec) > 0 Then sum_column = sum_column + Round_w(area * k_zap_total, n_round_area)
                            materials.Item(mat_clear) = materials.Item(mat_clear) + Round_w(area * k_zap_total, n_round_area)
                            pos_out(n_row_c, 7) = Replace(Replace(Replace(mat, "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString)
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
                        mat_clear = VedCleanName(mat)
                        area = materials_by_type.Item(type_name + ";pan;" + mat)
                        If area > 0.001 Then
                            If InStr(mat, fin_str) > 0 Or InStr(mat, fin_str_sec) > 0 Then sum_pan = sum_pan + Round_w(area * k_zap_total, n_round_area)
                            materials.Item(mat_clear) = materials.Item(mat_clear) + Round_w(area * k_zap_total, n_round_area)
                            pos_out(n_row_pan, col_end + 1) = Replace(Replace(Replace(mat, "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString)
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
            If show_perim Then
                prim = "Lперим=" + CStr(Round_w(materials_by_type.Item(type_name + ";pot_perim;") * k_zap_total, n_round_area)) + "п.м."
                prim = prim & vbLf & "Hmax=" + CStr(materials_by_type.Item(type_name + ";zoneh;")) + "м."
                pos_out(n_row, n_col_out) = prim
            End If
            n_row = Application.WorksheetFunction.Max(n_row_p, n_row_w, n_row_c, n_row_pan)
        End If
    Next
    If (show_surf_area Or show_mat_area) And delim_by_sheet = True Then n_row = n_row + 1
    If show_surf_area And delim_by_sheet = True Then
        nn_col = 5
        If Not show_mat_area Then nn_col = 1
        n_row_surf_area = n_row
        key_total_area = "tot;;pot;"
        If materials_by_type.Exists(key_total_area) Then
            n_row_surf_area = n_row_surf_area + 1
            pos_out(n_row_surf_area, nn_col) = "Потолки"
            pos_out(n_row_surf_area, nn_col + 3) = materials_by_type.Item(key_total_area)
        End If
        key_total_area = "tot;;wall;"
        If materials_by_type.Exists(key_total_area) Then
            n_row_surf_area = n_row_surf_area + 1
            pos_out(n_row_surf_area, nn_col) = "Стены(за вычетом панелей)"
            pos_out(n_row_surf_area, nn_col + 3) = materials_by_type.Item(key_total_area)
        End If
        key_total_area = "tot;;column;"
        If materials_by_type.Exists(key_total_area) Then
            n_row_surf_area = n_row_surf_area + 1
            pos_out(n_row_surf_area, nn_col) = "Колонны"
            pos_out(n_row_surf_area, nn_col + 3) = materials_by_type.Item(key_total_area)
        End If
        key_total_area = "tot;;pan;"
        If materials_by_type.Exists(key_total_area) Then
            n_row_surf_area = n_row_surf_area + 1
            pos_out(n_row_surf_area, nn_col) = "Панели"
            pos_out(n_row_surf_area, nn_col + 3) = materials_by_type.Item(key_total_area)
        End If
        If n_row_surf_area <> n_row Then pos_out(n_row, nn_col) = "Общяя площадь поверхности, кв.м."
    End If
    If show_mat_area And delim_by_sheet = True Then
        pos_out(n_row, 1) = "Общяя площадь отделки, кв.м."
        n_row = n_row + 1
        pos_out(n_row, 1) = "Отделка потолка"
        pos_out(n_row, 4) = vbNullString
        For Each type_mat In Array("@Потолок: ", "Потолок: ", fin_str, "@@@")
            Select Case type_mat
                Case fin_str
                    n_row = n_row + 1
                    pos_out(n_row, 1) = "Финишная отделка"
                    pos_out(n_row, 4) = vbNullString
                Case "@@@"
                    n_row = n_row + 1
                    pos_out(n_row, 1) = "Подготовка поверхности стен, колонн"
                    pos_out(n_row, 4) = vbNullString
            End Select
            
            len_type_mat = Len(type_mat)
            For Each mat In materials.keys()
                If Len(mat) > 2 And (Left$(mat, len_type_mat) = type_mat Or (type_mat = "@@@" And InStr(mat, "@") = 0 And InStr(mat, "Потолок: ") = 0 And InStr(mat, fin_str) = 0)) Then
                    n_row = n_row + 1
                    pos_out(n_row, 1) = Replace(Replace(Replace(Replace(Replace(mat, "@", vbNullString), "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString), "Потолок:", vbNullString)
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
    'erase all_data
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
            zone.Item(num + ";zoneh;") = h_zone
            fin_pot = fin_str + CStr(zone_el(1, col_s_mpot_zone))
            fin_pan = fin_str + CStr(zone_el(1, col_s_mpan_zone))
            fin_wall = fin_str + CStr(zone_el(1, col_s_mwall_zone))
            fin_column = fin_str + CStr(zone_el(1, col_s_mcolumn_zone))
            If InStr(fin_column, "<>") > 0 Then
                column_is_wall = True
                fin_column = Replace(fin_column, "<>", vbNullString)
            Else
                column_is_wall = False
            End If
            wall = ArraySelectParam(out_data, num, col_s_numb_zone, "СТЕНА", col_s_type)
            free_len_wall = 0
            manual_column = False
            If otd_version = 2 Then
                h_pot = zone_el(1, col_s_h_pot_zone) / 1000
                If zone_el(1, col_s_mwall_up_zone) = fin_wall Or zone_el(1, col_s_mwall_up_zone) = fin_column Or h_pot <= 0.0001 Then
                    h_pot = h_zone
                    tfin_up_zone = zone_el(1, col_s_mwall_up_zone)
                Else
                    If delim_zone_fin Then
                        tfin_up_zone = "Выше " + CStr(h_pot) + "м.:%%" + CStr(zone_el(1, col_s_mwall_up_zone))
                    Else
                        tfin_up_zone = "Выше подвесного потолка:%%" + CStr(zone_el(1, col_s_mwall_up_zone))
                    End If
                End If
                fin_up_zone = fin_str_sec + Replace(tfin_up_zone, "@", "; ")
                column_perim_total = GetZoneParam(zone_el(1, col_s_param_zone), "Cp")
                If Not IsEmpty(column_perim_total) Then
                    manual_column = True
                    column_perim_total = value_param / 1000
                    If Abs(column_perim_total) < 0.002 Then column_perim_total = -1
                End If
                wall_razd = ArraySelectParam(wall, "Разделитель", col_s_mat_wall)
                If Not IsEmpty(wall_razd) Then
                    For wi = 1 To UBound(wall_razd, 1)
                        free_len_wall = free_len_wall + wall_razd(wi, col_s_freelen_zone) / 1000
                    Next wi
                End If
            End If
            If h_pot < 1 And h_pot > 0 Then
                r = LogWrite("Проверьте высоту помещения, должна быть в мм. - " + CStr(h_pot), "Ошибка", num)
                is_error = is_error + 1
                h_pot = h_pot * 1000
            End If
            If h_pot > h_zone Then
                r = LogWrite("Проверьте высота потолка выше помещения. - " + CStr(h_pot), "Ошибка", num)
                is_error = is_error + 1
                h_pot = h_zone
            End If
            If h_pot < 0.01 Then h_pot = h_zone
            delta_h_zone = h_zone - h_pot 'Полоска над потолком
            If delta_h_zone < 0 Then delta_h_zone = 0
            If is_delta_h Then
                If delim_zone_fin Then
                    fin_wall = fin_str + "До " + CStr(h_pot) + "м.:%%" + CStr(zone_el(1, col_s_mwall_zone))
                    fin_column = fin_str + "До " + CStr(h_pot) + "м.:%%" + CStr(zone_el(1, col_s_mcolumn_zone))
                Else
                    fin_wall = fin_str + CStr(zone_el(1, col_s_mwall_zone))
                    fin_column = fin_str + CStr(zone_el(1, col_s_mcolumn_zone))
                End If
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
            If free_len_wall > 0 Then free_len = free_len_wall
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
            If manual_column = False Then column_perim_total = perim_total - wall_len - free_len + perim_hole * (hole_in_zone = True)
            column_pan_area = column_perim_total * h_pan
            column_pan_area_delta_h = column_perim_total * delta_h_zone
            column_area = column_perim_total * (h_zone - h_pan - delta_h_zone)
            name_mat_column = "ЖБ"
            colmn = VedNameMat("Колонны", name_mat_column, rules)
            If column_is_wall Then
                If column_pan_area_delta_h + column_area + column_pan_area > 0.1 Then n_row_pn = n_row_pn + VedAddArea(zone, materials, colmn, fin_column, num + ";c", column_pan_area_delta_h + column_area + column_pan_area, rules_mod)
            Else
                If column_area > 0.01 Then is_column = True
                If column_pan_area > 0.01 Then is_pan = True
                If column_area > 0.1 Then n_row_c = n_row_c + VedAddArea(zone, materials, colmn, fin_column, num + ";c", column_area, rules_mod)
                If column_pan_area > 0.1 Then n_row_pn = n_row_pn + VedAddArea(zone, materials, colmn, fin_pan, num + ";pn", column_pan_area, rules_mod, 0, h_pan)
                If column_pan_area_delta_h > 0.1 Then n_row_pn = n_row_pn + VedAddArea(zone, materials, colmn, fin_up_zone, num + ";c", column_pan_area_delta_h, rules_mod)
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
                    wall_up_area = 0
                    wall_c_len = 0
                    wall_c_area = 0
                    wall_c_up_area = 0
                    pan_area = 0
                    pan_c_area = 0
                    twall_area = 0
                    twall_h = 0
                    tpan_area = 0
                    tup_area = 0
                    wall_by_key = ArraySelectParam(wall, w, col_s_mat_wall)
                    For i = 1 To UBound(wall_by_key, 1)
                        twall_area = wall_by_key(i, col_s_area_wall)
                        If twall_area > 0 Then
                            tdoor_len = wall_by_key(i, col_s_doorlen_zone) / 1000
                            twall_len = wall_by_key(i, col_s_walllen_zone) / 1000
                            If otd_version = 2 Then
                                th_wall = wall_by_key(i, col_s_h_wall) / 1000
                                If th_wall > h_zone Then th_wall = h_zone
                            Else
                                If twall_len > tdoor_len Then th_wall = twall_area / (twall_len - tdoor_len)
                            End If
                            If th_wall > h_pan Then
                                If th_wall > h_pot Then
                                    tup_area = twall_len * delta_h_zone
                                Else
                                    tup_area = 0
                                End If
                                tpan_area = (twall_len - tdoor_len) * h_pan
                                twall_area = twall_area - tpan_area - tup_area
                            Else
                                If twall_len > tdoor_len Then
                                    tpan_area = twall_area
                                    twall_area = 0
                                    r = LogWrite("Панели на всю высоту стен? " + CStr(h_pan), vbNullString, num)
                                Else
                                    tpan_area = 0
                                    twall_area = twall_area
                                    r = LogWrite("Стена полностью скрыта дверью? " + CStr(h_pan), vbNullString, num)
                                End If
                            End If
                            If InStr(wall_by_key(i, col_s_layer), "Колонн") > 0 Then
                                wall_c_area = wall_c_area + twall_area
                                pan_c_area = pan_c_area + tpan_area
                                wall_c_up_area = wall_c_up_area + tup_area
                            Else
                                wall_area = wall_area + twall_area
                                pan_area = pan_area + tpan_area
                                wall_up_area = wall_up_area + tup_area
                            End If
                        End If
                    Next i
                    result = VedAddArea(zone, materials, w, fin_column, num + ";c", wall_c_area, rules_mod)
                    n_row_c = n_row_c + result
                    result = VedAddArea(zone, materials, w, fin_wall, num + ";w", wall_area, rules_mod)
                    n_row_w = n_row_w + result
                    result = VedAddArea(zone, materials, w, fin_up_zone, num + ";w", wall_up_area, rules_mod)
                    n_row_w = n_row_w + result
                    result = VedAddArea(zone, materials, w, fin_pan, num + ";pn", pan_area + pan_c_area, rules_mod, 0, h_pan)
                    n_row_pn = n_row_pn + result
                    wall_area_zone = wall_area_zone + wall_area + wall_c_area + pan_c_area + pan_area + wall_up_area + wall_c_up_area
                Next
            End If
            If wall_c_area > 0.1 Then is_column = True
            If h_pan > 0 Then is_pan = True
            If wall_area_zone < 0.1 Then
                r = LogWrite("Почти нет стен" & num, "Ошибка", wall_area_zone)
                is_error = is_error + 1
            End If
            '----------------------
            '        ПОТОЛКИ
            '----------------------
            noPot = False
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
                    noPot = True
                End If
            Else
                noPot = True
            End If
            If noPot Then
                n_row_pot = n_row_pot + VedAddArea(zone, materials, vbNullString, fin_pot, num + ";pot", area_total, rules_mod, perim_total)
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
    'erase out_data
    n_col_out = 7
    If is_pan Then n_col_out = n_col_out + 3
    If is_column Then n_col_out = n_col_out + 2
    n_un_mat = (materials.Count / 2)
    If (n_un_mat - Int(n_un_mat)) <> 0 Then MsgBox ("Ошибка записи в словарь")
    Dim pos_out(): ReDim pos_out(3 + n_row_tot + n_un_mat + 60, n_col_out)
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
    n_row = 2
    For Each num In un_n_zone
        n_row_p = n_row
        n_row_w = n_row
        n_row_c = n_row
        n_row_pan = n_row
        pos_out(n_row + 1, 1) = "'" + Replace(num, ",", ".")
        pos_out(n_row + 1, 2) = zone.Item(num + ";name")
        If show_perim Then
            prim = "Lперим=" + CStr(Round_w(zone.Item(num + ";potperim;") * k_zap_total, n_round_area)) + "п.м."
            prim = prim & vbLf & "Hпом=" + CStr(zone.Item(num + ";zoneh;")) + "м."
            pos_out(n_row + 1, n_col_out) = prim
        End If
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
                pos_out(n_row_p, 3) = Replace(Replace(Replace(mat, "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString)
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
                pos_out(n_row_w, 5) = Replace(Replace(Replace(mat, "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString)
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
                    pos_out(n_row_c, 7) = Replace(Replace(Replace(mat, "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString)
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
                    pos_out(n_row_pan, col_end + 1) = Replace(Replace(Replace(mat, "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString)
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
        nn_col = 5
        If Not show_mat_area Then nn_col = 1
        n_row_surf_area = n_row
        key_total_area = "tot;pot"
        If zone.Exists(key_total_area) Then
            n_row_surf_area = n_row_surf_area + 1
            pos_out(n_row_surf_area, nn_col) = "Потолки"
            pos_out(n_row_surf_area, nn_col + 3) = zone.Item(key_total_area)
        End If
        key_total_area = "tot;w"
        If zone.Exists(key_total_area) Then
            n_row_surf_area = n_row_surf_area + 1
            pos_out(n_row_surf_area, nn_col) = "Стены(за вычетом панелей)"
            pos_out(n_row_surf_area, nn_col + 3) = zone.Item(key_total_area)
        End If
        key_total_area = "tot;c"
        If zone.Exists(key_total_area) Then
            n_row_surf_area = n_row_surf_area + 1
            pos_out(n_row_surf_area, nn_col) = "Колонны"
            pos_out(n_row_surf_area, nn_col + 3) = zone.Item(key_total_area)
        End If
        key_total_area = "tot;pn"
        If zone.Exists(key_total_area) Then
            n_row_surf_area = n_row_surf_area + 1
            pos_out(n_row_surf_area, nn_col) = "Панели"
            pos_out(n_row_surf_area, nn_col + 3) = zone.Item(key_total_area)
        End If
        If n_row_surf_area <> n_row Then pos_out(n_row, nn_col) = "Общяя площадь поверхности, кв.м."
    End If
    If show_mat_area Then
        pos_out(n_row, 1) = "Общяя площадь отделки, кв.м."
        material_all = ArraySort(materials.keys())
        For Each mat In material_all
            If (Right$(mat, 2) <> ";a") Then
                n_row = n_row + 1
                pos_out(n_row, 1) = Replace(Replace(Replace(materials.Item(mat), "%%", vbNullString), fin_str, vbNullString), fin_str_sec, vbNullString)
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
    nm_rule = vbNullString
    nm = Trim$(Split(nm, "_")(0))
    If UBound(Split(add_rule(0), ";"), 1) < 1 Then Exit Function
    listsheet = GetListOfSheet(ThisWorkbook)
    For Each nlist In listsheet
        spec_type = SpecGetType(nlist)
        name_list = Split(nlist, "_")
        If spec_type = 10 Then
            If name_list(0) = nm Then nm_rule = nlist
        End If
    Next
    If nm_rule <> vbNullString Then
        Set rule_sheet = wbk.Sheets(nm_rule)
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
        Set data_out = rule_sheet.Range(rule_sheet.Cells(1, 1), rule_sheet.Cells(n_row_sheet + n_row, n_col))
        r = FormatClear(data_out)
        r = FormatSpec_Rule(data_out)
        VedAddRules = True
    Else
        VedAddRules = False
        r = VedNewListRules(nm)
        MsgBox ("Не найден лист с правилами отделки (оканчивается на '_правила')")
    End If
End Function

Function VedGetRules(ByVal nm As String) As Variant
    nm_rule = vbNullString
    nm = Trim$(Split(nm, "_")(0))
    listsheet = GetListOfSheet(ThisWorkbook)
    For Each nlist In listsheet
        spec_type = SpecGetType(nlist)
        name_list = Split(nlist, "_")
        If spec_type = 10 Then
            If name_list(0) = nm Then nm_rule = nlist
        End If
    Next
    If nm_rule <> vbNullString Then
        Set rule_sheet = wbk.Sheets(nm_rule)
        lsize = SheetGetSize(rule_sheet)
        n_row = lsize(1)
        n_col = lsize(2)
        If n_row = 1 Then n_row = 2
        Set data_out = rule_sheet.Range(rule_sheet.Cells(1, 1), rule_sheet.Cells(n_row, n_col))
        Worksheets(nm_rule).Activate
        r = FormatClear(Worksheets(nm_rule))
        r = FormatSpec_Rule(data_out)
        Dim rules(): ReDim rules(n_row - 1, 3)
        Dim rules_mod(): ReDim rules_mod(n_row - 1, 3)
        n_rules = 0
        n_rules_mod = 0
        For i = 2 To n_row
            If Not IsEmpty(data_out(i, 1)) And Len(data_out(i, 1)) > 0 And Left$(data_out(i, 1), 2) <> "!!" Then
                If InStr(data_out(i, 1), "Исключить") = 0 And InStr(data_out(i, 1), "Добавить") = 0 Then
                    If InStr(data_out(i, 1), "Стены-разделители зон") = 0 Then
                        n_rules = n_rules + 1
                        rules(n_rules, 1) = ConvNum2Txt(data_out(i, 1))
                        rules(n_rules, 2) = ConvNum2Txt(data_out(i, 2))
                        rules(n_rules, 3) = ConvNum2Txt(data_out(i, 3))
                    Else
                        n_rules_mod = n_rules_mod + 1
                        rules_mod(n_rules_mod, 1) = "Разделитель"
                        rules_mod(n_rules_mod, 2) = ConvNum2Txt(Trim$(data_out(i, 2)))
                        rules_mod(n_rules_mod, 3) = ConvNum2Txt(Trim$(data_out(i, 3)))
                    End If
                Else
                    n_rules_mod = n_rules_mod + 1
                    rules_mod(n_rules_mod, 1) = Trim$(ConvNum2Txt(data_out(i, 1)))
                    rules_mod(n_rules_mod, 1) = Replace(rules_mod(n_rules_mod, 1), "Исключить", "-")
                    rules_mod(n_rules_mod, 1) = Replace(rules_mod(n_rules_mod, 1), "Добавить", "+")
                    rules_mod(n_rules_mod, 2) = Trim$(ConvNum2Txt(data_out(i, 2)))
                    rules_mod(n_rules_mod, 3) = Trim$(ConvNum2Txt(data_out(i, 3)))
                End If
            End If
        Next i
        rules = ArrayRedim(rules, n_rules)
        rules_mod = ArrayRedim(rules_mod, n_rules_mod)
        VedGetRules = Array(rules, rules_mod)
        'erase rules
    Else
        VedGetRules = Array(Empty, Empty)
        r = VedNewListRules(nm)
        MsgBox ("Создан лист с правилами.")
    End If
End Function

Function VedModMat(ByVal fin_material As String, ByVal all_material As String, ByRef rules_mod As Variant) As String
Dim tfunctime As Double
tfunctime = Timer
    fin_material_t = fin_material
    If InStr(fin_material, "%") > 0 Then
        n = Split(fin_material, "%")
        fin_material_t = n(0)
    End If
    If Not IsEmpty(rules_mod) Then
        For i = 1 To UBound(rules_mod, 1)
            If Trim$(fin_material_t) = Trim$(rules_mod(i, 2)) Or (InStr(fin_material_t, "Выше") > 0 And InStr(fin_material_t, rules_mod(i, 2)) > 0) Then
                If rules_mod(i, 1) = "-" Then
                    arr_mat = Split(all_material, ";")
                    arr_mod = Split(rules_mod(i, 3), ";")
                    all_material_out = vbNullString
                    For Each mat In arr_mat
                        mat = Trim$(mat)
                        flag_in = True
                        For Each modd In arr_mod
                            modd = Trim$(modd)
                            If mat = modd Then flag_in = False
                        Next modd
                        If flag_in = True Then
                            If Len(all_material_out) = 0 Then
                                all_material_out = mat
                            Else
                                all_material_out = all_material_out & ";" & mat
                            End If
                        End If
                    Next mat
                    all_material = Trim$(all_material_out)
                End If
                'all_material = Replace(all_material, rules_mod(i, 3), vbNullString)
                If rules_mod(i, 1) = "+" Then all_material = all_material + ";" + rules_mod(i, 3)
            End If
        Next i
        all_material = Replace(all_material, "; ;", ";")
        all_material = Replace(all_material, ";;", ";")
        all_material = Trim$(all_material)
        If all_material = ";" Then all_material = vbNullString
    End If
tfunctime = functime("VedModMat", tfunctime)
    VedModMat = all_material
End Function

Function VedNameMat(ByVal layer As String, ByVal material As String, ByRef rules As Variant) As String
Dim tfunctime As Double
tfunctime = Timer
    name_m = vbNullString
    flag = 0
    For i = 1 To UBound(rules, 1) 'Ищем точное соответсвие
        m = rules(i, 1)
        L = rules(i, 2)
        If (layer = L Or Len(layer) = 0) And m = material Then
            name_m = rules(i, 3)
            flag = flag + 1
        End If
    Next i
    If flag < 1 Then 'Если ничего не нашли - попробуем поискать похожий материал при заданном слое
        For i = 1 To UBound(rules, 1)
            m = rules(i, 1)
            L = rules(i, 2)
            If (layer = L Or Len(layer) = 0) And InStr(material, m) > 0 Then
                name_m = rules(i, 3)
                flag = flag + 1
            End If
        Next i
    End If
    If flag < 1 Then 'Отчаявшись ищем похожие материалы на похожих слоях
        For i = 1 To UBound(rules, 1)
            m = rules(i, 1)
            L = rules(i, 2)
            If InStr(layer, L) > 0 And InStr(material, m) > 0 Then
                name_m = rules(i, 3)
                flag = flag + 1
            End If
        Next i
    End If
    If flag = 1 Then
        If InStr(name_m, "ез отделк") > 0 Then
            name_m = Replace(name_m, "; БЕЗ ОТДЕЛКИ", vbNullString)
            name_m = Replace(name_m, "без отделки", vbNullString)
            name_m = Trim$(name_m)
            If Right$(name_m, 1) = ";" Then name_m = Trim$(Left$(name_m, Len(name_m) - 1))
            name_m = name_m + "="
        End If
        VedNameMat = name_m
    Else
        VedNameMat = material + ";" + layer + ";ОШИБКА"
        If flag > 1 Then
            MsgBox ("Несколько правил для одного материала - " + material + " слой" + layer)
        End If
    End If
tfunctime = functime("VedNameMat", tfunctime)
End Function

Function VedNewListRules(ByVal nm As String) As Boolean
    wbk.Worksheets.Add.Name = nm & "_правила"
    wbk.Worksheets(nm & "_правила").Move After:=wbk.Sheets(wbk.Sheets.Count)
    wbk.Worksheets(nm & "_правила").Activate
    Cells(1, 1).Value = "Имя многослойной конструкции (целиком или часть имени, строки с !! не учитываются)"
    Cells(1, 2).Value = "Слой"
    Cells(1, 3).Value = "Черновая отделка (разделитель ';')"
    
    Cells(2, 1).Value = "!!Исключить"
    Cells(2, 2).Value = "Облицовка керам. плиткой"
    Cells(2, 3).Value = "Шпатлёвка"
    
    Cells(3, 1).Value = "!!Добавить"
    Cells(3, 2).Value = "Облицовка керам. плиткой"
    Cells(3, 3).Value = "Обработка бетоноконтактом"
    
    Cells(4, 1).Value = "Стены-разделители зон"
    Cells(4, 2).Value = "Зоны.АР"
    Cells(4, 3).Value = "!"
    
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
    fin_str = Trim$(fin_str)
    lastfilespec = Left$(lastfilespec, Len(lastfilespec) - Len("_вед"))
    out_data = ReadFile(lastfilespec & ".txt")
    If Not DataIsOtd(out_data) Then
        MsgBox ("Неверный формат файла")
        VedRead = Empty
        Exit Function
    End If
Dim tfunctime As Double
tfunctime = Timer
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
    layer_razd = vbNullString
    material_razd = vbNullString
    For i = 1 To UBound(rules_mod, 1)
        If rules_mod(i, 1) = "Разделитель" Then
            layer_razd = rules_mod(i, 2)
            material_razd = rules_mod(i, 3)
        End If
    Next i
    n_row_a = UBound(out_data, 1) - 2
    n_col_a = UBound(out_data, 2)
    If n_col_a < max_s_col Then
        Dim out_data_t(): ReDim out_data_t(UBound(out_data, 1), max_s_col)
        For i = 1 To UBound(out_data, 1)
            For j = 1 To UBound(out_data, 2)
                out_data_t(i, j) = out_data(i, j)
            Next j
            For j = UBound(out_data, 2) + 1 To max_s_col
                out_data_t(i, j) = 0
            Next j
        Next i
        out_data = out_data_t
    End If
    Dim add_pol(): ReDim add_pol(max_s_col, n_row_a)
    n_add = 0
    n_zone = 999999
    For i = 1 To n_row_a
        out_data(i, col_s_type_otd) = ConvNum2Txt(out_data(i, col_s_type_otd))
        If out_data(i, col_s_numb_zone) = 0 Or Len(out_data(i, col_s_numb_zone)) = 0 Then
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
        If n_col_a > max_col_type_2 Then ' Если есть лестницы
            out_data(i, col_s_tipverh_l) = ConvNum2Txt(out_data(i, col_s_tipverh_l))
            out_data(i, col_s_tipniz_l) = ConvNum2Txt(out_data(i, col_s_tipniz_l))
            out_data(i, col_s_tippl_l) = ConvNum2Txt(out_data(i, col_s_tippl_l))
            out_data(i, col_s_tipl_l) = ConvNum2Txt(out_data(i, col_s_tipl_l))
        End If
        If out_data(i, col_s_type) = "СТЕНА" Then
            layer = out_data(i, col_s_layer)
            material = out_data(i, col_s_mat_wall)
            name_mat = VedNameMat(layer, material, rules)
            If Len(layer_razd) > 0 And InStr(layer, layer_razd) > 0 And InStr(material, material_razd) > 0 And otd_version = 2 Then
                out_data(i, col_s_area_wall) = 0
                out_data(i, col_s_h_wall) = 0
                out_data(i, col_s_mat_wall) = "Разделитель"
                out_data(i, col_s_freelen_zone) = out_data(i, col_s_walllen_zone)
                out_data(i, col_s_walllen_zone) = 0
                out_data(i, col_s_doorlen_zone) = 0
            End If
            If out_data(i, col_s_area_wall) > 0 Then
                If InStr(name_mat, "ОШИБКА") > 0 Then
                    If Not add_rule.Exists(name_mat) Then add_rule.Item(name_mat) = name_mat
                    out_data(i, col_s_mat_wall) = "ОШИБКА"
                Else
                    out_data(i, col_s_mat_wall) = name_mat
                End If
            End If
        End If
        If n_col_a > max_col_type_1 Then 'Если есть пол или потолок
            out_data(i, col_s_type_pol) = ConvNum2Txt(out_data(i, col_s_type_pol))
            If otd_version > 1 Then
                out_data(i, col_s_type_pot_zone) = ConvNum2Txt(out_data(i, col_s_type_pot_zone))
                out_data(i, col_s_type_pol_zone) = ConvNum2Txt(out_data(i, col_s_type_pol_zone))
            End If
            If out_data(i, col_s_type) = "ОБЪЕКТ" Then
                If out_data(i, col_s_type_el) = "Потолок" Then
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
                out_data(i, col_s_n_mun_zone) = ConvNum2Txt(out_data(i, col_s_n_mun_zone))
                If out_data(i, col_s_n_mun_zone) <> vbNullString And out_data(i, col_s_n_mun_zone) <> out_data(i, col_s_numb_zone) Then
                    If out_data(i, col_s_mun_zone) = 1 Then
                        out_data(i, col_s_numb_zone) = out_data(i, col_s_n_mun_zone)
                    Else
                        r = LogWrite("Проверьте пол/потолок номер помещения " & out_data(i, col_s_numb_zone) & " или " & out_data(i, col_s_n_mun_zone), "Ошибка", num)
                    End If
                End If
                If out_data(i, col_s_type_el) = "Потолок" Then zone_error.Item(out_data(i, col_s_numb_zone) + "_pot_qty") = zone_error.Item(out_data(i, col_s_numb_zone) + "_pot_qty") + 1
                If out_data(i, col_s_type_el) = "Пол" Then zone_error.Item(out_data(i, col_s_numb_zone) + "_pol_qty") = zone_error.Item(out_data(i, col_s_numb_zone) + "_pol_qty") + 1
                If out_data(i, col_s_area_pol) < 0.0001 Then
                    For j = 1 To UBound(out_data, 2)
                        out_data(i, j) = 0
                    Next j
                End If
            End If
        End If
        'Добавляем поверхность низа ж/б лестниц
        If n_col_a > max_col_type_2 And out_data(i, col_s_type) = "ОБЪЕКТ" And otd_version = 2 Then 'Если есть лестницы
            out_data(i, col_s_tipverh_l) = ConvNum2Txt(out_data(i, col_s_tipverh_l))
            out_data(i, col_s_tipniz_l) = ConvNum2Txt(out_data(i, col_s_tipniz_l))
            out_data(i, col_s_tippl_l) = ConvNum2Txt(out_data(i, col_s_tippl_l))
            out_data(i, col_s_tipl_l) = ConvNum2Txt(out_data(i, col_s_tipl_l))
            If out_data(i, col_s_tipverh_l) <> vbNullString Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_tipniz_l) <> vbNullString Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_tippl_l) <> vbNullString Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_tipl_l) <> vbNullString Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_type_el) = "Лестница" Then
                s_pot = GetZoneParam(out_data(i, col_s_param_zone), "Spl")
                If IsEmpty(s_pot) Then
                    add_s_pol = 0
                Else
                    s_pot = s_pot / 1000
                End If
                If s_pot > 0.1 Then
                    material = out_data(i, col_s_type_pot_zone)
                    If InStr(material, "е задан") > 0 Then material = "Низ ж/б лестниц"
                    layer = "Потолок"
                    name_mat = VedNameMat(layer, material, rules)
                    If InStr(name_mat, "ОШИБКА") > 0 Then
                       If Not add_rule.Exists(name_mat) Then add_rule.Item(name_mat) = name_mat
                    Else
                        type_ = "Низ лестничных маршей: "
                        name_mat = type_ + out_data(i, col_s_mpot_zone) + ";" + type_ + name_mat
                        n_add = n_add + 1
                        add_pol(col_s_numb_zone, n_add) = out_data(i, col_s_numb_zone)
                        add_pol(col_s_type, n_add) = "ОБЪЕКТ"
                        add_pol(col_s_type_el, n_add) = "Потолок"
                        add_pol(col_s_type_pol, n_add) = name_mat
                        add_pol(col_s_area_pol, n_add) = s_pot
                        add_pol(col_s_perim_pol, n_add) = 0
                        add_pol(col_s_param_zone, n_add) = "@MatPot=" + out_data(i, col_s_mpot_zone)
                    End If
                End If
            End If
        End If
        For j = 1 To n_col_a
            If Len(out_data(i, j)) = 0 Then out_data(i, j) = 0
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
    'Обработка пола и потолка, заданного в зонах
    If otd_version = 2 Then
        zone_el = ArraySelectParam(out_data, "ЗОНА", col_s_type)
        For i = 1 To UBound(zone_el, 1)
            add_s_pol = GetZoneParam(zone_el(1, col_s_param_zone), "AL")
            If IsEmpty(add_s_pol) Then
                add_s_pol = 0
            Else
                add_s_pol = add_s_pol / 1000
            End If
            add_s_pot = GetZoneParam(zone_el(1, col_s_param_zone), "AT")
            If IsEmpty(add_s_pot) Then
                add_s_pot = 0
            Else
                add_s_pot = add_s_pot / 1000
            End If
            If Len(zone_el(i, col_s_type_pol_zone)) > 0 And zone_el(i, col_s_type_pol_zone) <> "0" And InStr(zone_el(i, col_s_type_pol_zone), "е задан") = 0 Then
                'Вычисляем площадь пола, заданной аксессуарами и добавляем недостающий пол
                s_pol = 0
                s_pol_accs = 0
                pol = ArraySelectParam(out_data, "Пол", col_s_type_el, zone_el(i, col_s_numb_zone), col_s_numb_zone)
                If Not IsEmpty(pol) Then
                    For j = 1 To UBound(pol, 1)
                        s_pol_accs = s_pol_accs + pol(j, col_s_area_pol)
                    Next j
                End If
                s_pol = zone_el(i, col_s_area_zone) + add_s_pol - s_pol_accs
                If s_pol > 0.1 Then
                    n_add = n_add + 1
                    add_pol(col_s_numb_zone, n_add) = zone_el(i, col_s_numb_zone)
                    add_pol(col_s_type, n_add) = "ОБЪЕКТ"
                    add_pol(col_s_type_el, n_add) = "Пол"
                    add_pol(col_s_type_pol, n_add) = zone_el(i, col_s_type_pol_zone)
                    add_pol(col_s_area_pol, n_add) = s_pol
                    add_pol(col_s_perim_pol, n_add) = zone_el(i, col_s_perim_zone)
                End If
            End If
            If Len(zone_el(i, col_s_type_pot_zone)) > 0 And zone_el(i, col_s_type_pot_zone) <> "0" And InStr(zone_el(i, col_s_type_pot_zone), "е задан") = 0 Then
                'Вычисляем площадь потолка, заданной аксессуарами и добавляем недостающий потолок
                s_pot = 0
                s_pot_accs = 0
                pot = ArraySelectParam(out_data, "Потолок", col_s_type_el, zone_el(i, col_s_numb_zone), col_s_numb_zone)
                If Not IsEmpty(pot) Then
                    For j = 1 To UBound(pol, 1)
                        s_pot_accs = s_pot_accs + pot(j, col_s_area_pol)
                    Next j
                End If
                s_pot = zone_el(i, col_s_area_zone) + add_s_pot - s_pot_accs
                If s_pot > 0.1 Then
                    layer = "Потолок"
                    material = zone_el(i, col_s_type_pot_zone)
                    name_mat = VedNameMat(layer, material, rules)
                    If InStr(name_mat, "ОШИБКА") > 0 Then
                        If Not add_rule.Exists(name_mat) Then add_rule.Item(name_mat) = name_mat
                    Else
                        n_add = n_add + 1
                        add_pol(col_s_numb_zone, n_add) = zone_el(i, col_s_numb_zone)
                        add_pol(col_s_type, n_add) = "ОБЪЕКТ"
                        add_pol(col_s_type_el, n_add) = "Потолок"
                        add_pol(col_s_type_pol, n_add) = name_mat
                        add_pol(col_s_area_pol, n_add) = s_pot
                        add_pol(col_s_perim_pol, n_add) = zone_el(i, col_s_perim_zone)
                    End If
                End If
            End If
        Next i
    End If
    If n_add > 0 Then
        ReDim Preserve add_pol(max_s_col, n_add)
        add_pol = ArrayTranspose(add_pol)
        out_data = ArrayCombine(out_data, add_pol)
        'erase add_pol
    End If
    Dim pos_out(): ReDim pos_out(3)
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
tfunctime = functime("VedRead", tfunctime)
    VedRead = pos_out
End Function

Function VedReadPol(ByVal lastfilespec As String) As Variant
    fin_str = Trim$(fin_str)
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
Dim tfunctime As Double
tfunctime = Timer
    n_row_a = UBound(out_data, 1)
    n_col_a = UBound(out_data, 2)
    If n_col_a <= col_s_layer Then
        VedReadPol = Empty
        Exit Function
    End If
    If n_col_a < max_s_col Then
        Dim out_data_t: ReDim out_data_t(n_row_a, max_s_col)
        For i = 1 To n_row_a
            For j = 1 To n_col_a
                out_data_t(i, j) = out_data(i, j)
            Next j
        Next i
        out_data = out_data_t
    End If
    Dim add_pol: ReDim add_pol(max_s_col, n_row_a)
    n_add = 0
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
            If out_data(i, col_s_n_mun_zone) <> vbNullString And out_data(i, col_s_n_mun_zone) <> out_data(i, col_s_numb_zone) Then
                If out_data(i, col_s_mun_zone) = 1 Then
                    out_data(i, col_s_numb_zone) = out_data(i, col_s_n_mun_zone)
                Else
                    r = LogWrite("Проверьте пол/потолок номер помещения " & out_data(i, col_s_numb_zone) & " или " & out_data(i, col_s_n_mun_zone), "Ошибка", num)
                End If
            End If
        End If
        If n_col_a >= col_s_type_pol Then
            out_data(i, col_s_type_pol) = ConvNum2Txt(out_data(i, col_s_type_pol))
        End If
        If n_col_a >= col_s_tipverh_l And out_data(i, col_s_type) = "ОБЪЕКТ" Then
            out_data(i, col_s_tipverh_l) = ConvNum2Txt(out_data(i, col_s_tipverh_l))
            out_data(i, col_s_tipniz_l) = ConvNum2Txt(out_data(i, col_s_tipniz_l))
            out_data(i, col_s_tippl_l) = ConvNum2Txt(out_data(i, col_s_tippl_l))
            out_data(i, col_s_tipl_l) = ConvNum2Txt(out_data(i, col_s_tipl_l))
            If out_data(i, col_s_tipverh_l) <> vbNullString Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_tipniz_l) <> vbNullString Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_tippl_l) <> vbNullString Then out_data(i, col_s_type_el) = "Лестница"
            If out_data(i, col_s_tipl_l) <> vbNullString Then out_data(i, col_s_type_el) = "Лестница"
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
                If Len(out_data(i, j)) = 0 Then out_data(i, j) = 0
            End If
        Next j
    Next i
    'Обработка пола, заданного в зона
    If otd_version = 2 Then
        zone_el = ArraySelectParam(out_data, "ЗОНА", col_s_type)
        For i = 1 To UBound(zone_el, 1)
            If Len(zone_el(i, col_s_type_pol)) > 0 And zone_el(i, col_s_type_pol) <> "0" And InStr(zone_el(i, col_s_type_pol), "е задан") = 0 Then
                'Вычисляем площадь пола, заданной аксессуарами и добавляем недостающий пол
                s_pol = 0
                s_pol_accs = 0
                pol = ArraySelectParam(out_data, "Пол", col_s_type_el, zone_el(i, col_s_numb_zone), col_s_numb_zone)
                If Not IsEmpty(pol) Then
                    For j = 1 To UBound(pol, 1)
                        s_pol_accs = s_pol_accs + pol(j, col_s_area_pol)
                    Next j
                End If
                add_s_pol = GetZoneParam(zone_el(1, col_s_param_zone), "AL")
                If IsEmpty(add_s_pol) Then
                    add_s_pol = 0
                Else
                    add_s_pol = add_s_pol / 1000
                End If
                s_pol = zone_el(i, col_s_area_zone) + add_s_pol - s_pol_accs
                If s_pol > 0.1 Then
                    n_add = n_add + 1
                    add_pol(col_s_numb_zone, n_add) = zone_el(i, col_s_numb_zone)
                    add_pol(col_s_type, n_add) = "ОБЪЕКТ"
                    add_pol(col_s_type_el, n_add) = "Пол"
                    add_pol(col_s_type_pol, n_add) = zone_el(i, col_s_type_pol)
                    add_pol(col_s_area_pol, n_add) = s_pol
                    add_pol(col_s_perim_pol, n_add) = zone_el(i, col_s_perim_zone)
                End If
            End If
        Next i
    End If
    If n_add > 0 Then
        ReDim Preserve add_pol(max_s_col, n_add)
        add_pol = ArrayTranspose(add_pol)
        out_data = ArrayCombine(out_data, add_pol)
    End If
    VedReadPol = Array(out_data, Empty, Empty)
tfunctime = functime("VedReadPol", tfunctime)
End Function

Function VedWriteLog(ByVal nm As String)
    ilg = 1
    If Debug_mode = False Or ilg = 1 Then Exit Function
    nm_log = Right$(nm, 24) & "_log"
    If SheetExist(nm_log) Then
        wbk.Worksheets(nm_log).Activate
        wbk.Worksheets(nm_log).Cells.Clear
    Else
        wbk.Worksheets.Add.Name = nm_log
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
    Dim pos_out(): ReDim pos_out(n_row, n_col)
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
            If InStr(nkey, "_qty") > 0 And zone_error.Item(num + "_" + nkey) = 1 Then pos_out(i, j) = vbNullString
        Next j
    Next i
    Set Sh = wbk.Sheets(nm_log)
    Sh.Range(Sh.Cells(2, 1), Sh.Cells(n_row + 1, n_col)) = pos_out
    Set data_out = Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col))
End Function

Function VedSplitData(ByVal all_data As Variant, ByVal split_data As Variant, ByVal lastfilespec As Variant, ByVal suffix As String) As Variant
    n_split = UBound(split_data, 1)
    Dim out_data: ReDim out_data(n_split, 2)
    raw_data = all_data(1)
    rules = all_data(2)
    rules_mod = all_data(3)
    zones_el_all = Empty
    For i = 1 To n_split
        nm = Right$(lastfilespec & "-" & split_data(i, 1) & suffix, 31)
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
    'erase all_data, split_data
End Function

Function VedSplitSheet(ByVal lastfilespec As String)
    Set split_sheet = wbk.Sheets(Split(lastfilespec, "_")(0) & "_разб")
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
                    num_zone(j) = Trim$(Trim$(num_zone(j)))
                Next
                n_row = n_row + 1
                split_data(n_row, 1) = Left$(nm, 10)
                split_data(n_row, 2) = num_zone
                split_data(n_row, 3) = n_col_param
            End If
        End If
    Next i
    split_data = ArrayRedim(split_data, n_row)
    VedSplitSheet = split_data
End Function

Function VedSplitFile(ByVal lastfilespec As String)
    list_razb = GetListFile("_разб" & ".txt")
    raw_data = ReadTxt(list_razb(1, 2), 1, vbTab, vbNewLine)
    sheet_name = ArrayUniqValColumn(raw_data, 1)
    n_split = UBound(sheet_name, 1)
    Dim split_data: ReDim split_data(n_split, 3)
    For i = 1 To n_split
        If Not IsEmpty(sheet_name(i)) Then
            nm = sheet_name(i)
            For Each del_txt In Array("План", "Кровля", "на", "отм.", "отметке", "отметка", "этаж", "  ")
                nm = Replace(nm, del_txt, vbNullString)
            Next
            nm = Trim$(Trim$(nm)) 'Безусловное удаление пробелов
            num_zone = ArrayUniqValColumn(ArraySelectParam(raw_data, sheet_name(i), 1), 2)
            If Not IsEmpty(num_zone) Then
                For j = 1 To UBound(num_zone)
                    If IsNumeric(num_zone(j)) Then num = CStr(num_zone(j))
                    num_zone(j) = Trim$(Trim$(num_zone(j)))
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




