Attribute VB_Name = "calc"
Option Compare Text
Option Base 1

Public Const macro_version As String = "2.5"
'Тип округления
' 1 - округление в большую сторону
' 2 - округление стандартным round
' 3 - округление отключено
Public Const type_okrugl As Integer = 1
'Кол-во знаков после запятой для длины и веса
Public Const n_round_l As Integer = 2
Public Const n_round_w As Integer = 2
'-------------------------------------------------------
'Типы элементов (столбец col_type_el)
Public Const t_arm As Integer = 10
Public Const t_prokat As Integer = 20
Public Const t_mat As Integer = 30
Public Const t_mat_spc As Integer = 35
Public Const t_izd As Integer = 40
Public Const t_subpos As Integer = 45
Public Const t_else As Integer = 50
Public Const t_error As Integer = -1 'Ошибка распознавания типов
Public Const t_sys As Integer = -10 'Вспомогательный тип
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
'Общее количество столбцов во входном массиве
Public Const max_col As Integer = 15
'-------------------------------------------------------
'Описание таблицы с именами сборок (суффикс "_поз")
Public Const col_add_pos As Integer = 1
Public Const col_add_obozn As Integer = 2
Public Const col_add_naen As Integer = 3
Public Const col_add_qty As Integer = 4
Public Const col_add_prim As Integer = 5

Public symb_diam As String 'Символ диаметра в спецификацию
Public w_format As String 'Формат выводав техничку
Public pos_data As Variant

Function ControlSumAddVar(ByVal var As Variant) As String
    If IsNumeric(var) Then var = Trim(Str(var))
    If var = "_" Then
        ControlSumAddVar = "_"
    Else
        var = Trim(Replace(var, " ", ""))
        var = Trim(Replace(var, "--", ""))
        var = Trim(Replace(var, "x", ""))
        var = Trim(Replace(var, "х", ""))
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
                param(9) = 0
                param(10) = 1
                param(11) = 0
                param(12) = (gnut = 1) * 3
            Else
                param(9) = Int(Length / 10)
                param(10) = 0
                param(11) = 0
                param(12) = (gnut = 1) * 3
            End If
        Case t_prokat
            isel = 1
            type_konstr = array_in(col_pr_type_konstr)
            gost_st = array_in(col_pr_gost_st)
            st = array_in(col_pr_st)
            gost_prof = array_in(col_pr_gost_prof)
            prof = array_in(col_pr_prof)
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
            param(11) = Int(Length)
            
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
            param(10) = edizm
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
        If InStr(current_subpos, "-") = 0 Then
            'Проеряем - есть ли маркировка для главных сборок
            seach_subpos = ArraySelectParam(array_in, current_subpos, col_sub_pos, t_subpos, col_type_el)
            If IsEmpty(seach_subpos) Then
                If IsEmpty(add_txt) Then
                    add_txt = current_subpos
                Else
                    add_txt = current_subpos & ", " & add_txt
                End If
                If name_subpos.exists(current_subpos) Then
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
                add_subpos(1, col_m_naen) = naen
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
    Dim out_data: ReDim out_data(UBound(array_in, 1), n_col): n_row = 0
    For i = 1 To UBound(array_in, 1)
        type_el = array_in(i, col_type_el)
        'Вложенные сборки определяем по точке в первом столбце
        'Если это строка относится к вложенной сборке, то формат будет ИмяГлавнойСборки.ПозЭлемента
        array_in(i, col_parent) = Empty
        If InStr(array_in(i, col_marka), ".") Then
            parent_subpos = Split(array_in(i, col_marka), ".")(0)
            pos = Split(array_in(i, col_marka), ".")(1)
            array_in(i, col_parent) = parent_subpos
            array_in(i, col_marka) = parent_subpos
        End If
        'Также проверяем поле сборки, сравнивая его с маркой
        'Если это вложенная сброрка, то в графе сборки будет МАРКА.СБОРКА
        If InStr(array_in(i, col_sub_pos), ".") Then
            parent_subpos = Split(array_in(i, col_sub_pos), ".")(0)
            subpos = Split(array_in(i, col_sub_pos), ".")(1)
            array_in(i, col_sub_pos) = subpos
            array_in(i, col_parent) = parent_subpos
            array_in(i, col_marka) = parent_subpos
            If type_el = t_subpos Then array_in(i, col_pos) = subpos
        End If

        If type_el <> "" Then
            'Ищем материалы, привязанные к стандартным элементам архикада (t_mat_spc)
            If type_el = t_mat_spc Then
                If InStr(array_in(i, col_marka), ".") Then
                    array_in(i, col_sub_pos) = Split(array_in(i, col_marka), ".")(1)
                Else
                    array_in(i, col_sub_pos) = array_in(i, col_marka)
                End If
                array_in(i, col_pos) = Empty
                array_in(i, col_type_el) = t_mat
                array_in(i, col_m_weight) = "-"
            End If
            If type_el <> Empty Then
                If array_in(i, col_sub_pos) = "" Then array_in(i, col_sub_pos) = "-"
                If array_in(i, col_sub_pos) = " " Then array_in(i, col_sub_pos) = "-"
                If array_in(i, col_sub_pos) = 0 Then array_in(i, col_sub_pos) = "-"
                If array_in(i, col_sub_pos) = "-" Then array_in(i, col_parent) = "-"
                If IsEmpty(array_in(i, col_parent)) Then array_in(i, col_parent) = "-"
            End If
            'Вычисление и проверка контрольных сумм
            array_in(i, col_chksum) = ControlSumEl(ArrayRow(array_in, i))
            n_row = n_row + 1
            For j = 1 To n_col
                out_data(n_row, j) = array_in(i, j)
            Next j
        End If
    Next i
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
    Next i
    If n = n_row Then DataIsSpec = True Else DataIsSpec = False
End Function


Function DataIsShort(ByVal array_in As Variant) As Boolean
'Если номер столбца с типами элементов отличается от col_type_el - то первый столбец, скорее всего - количество элементов
    colum = 0
    n_row = Int(UBound(array_in, 1) / 2) + 1
    For j = 1 To col_type_el + 1
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
    nm = ActiveWorkbook.ActiveSheet.Name
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
                If name_subpos.exists(subpos) Then name_subpos.Remove subpos
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
                    out_data = ManualSpec(nsht)
                End If
            Else
                'Читаем из файла
                out_data = ReadFile(file(1, 1) & ".txt")
                if instr(out_data(1,1), ""
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
        If Not DataIsSpec(out_data) And SpecGetType(nm) <> "7" Then
            MsgBox ("Неверный формат файла")
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
            If pos_data.exists("-") Then
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
            Select Case out(n_row, col_type_el)
                Case t_arm
                    out(n_row, col_qty) = out(n_row, col_qty) * qty
                Case t_prokat
                    out(n_row, col_qty) = out(n_row, col_qty) * qty
                Case t_else
                    out(n_row, col_qty) = out(n_row, col_qty) * qty
            End Select
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
    For Each t_el In Array(t_arm, t_prokat, t_mat, t_izd, t_subpos)
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
                            'Суммируем
                            sum_by_type(i, col_qty) = sum_by_type(i, col_qty) + arr_by_type(j, col_qty)
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
                        If Not dchild.exists(subpos) Then dchild.Item(subpos) = 1
                    End If
                Next i
                If flag And (Not dparent.exists(subpos)) Then dparent.Item(subpos) = 1
            End If
        Next
        For i = 1 To UBound(sub_pos_arr, 1)
            spos = sub_pos_arr(i, col_sub_pos)
            tparent = sub_pos_arr(i, col_parent)
            qty = sub_pos_arr(i, col_qty)
            If dqty.exists(tparent & "_" & spos) Then
                dqty.Item(tparent & "_" & spos) = dqty.Item(tparent & "_" & spos) + qty
            Else
                dqty.Item(tparent & "_" & spos) = qty
            End If
            If Not (tparent = "-" And dparent.exists(spos)) Then
                If dqty.exists("all" & spos) Then
                    dqty.Item("all" & spos) = dqty.Item("all" & spos) + qty
                Else
                    dqty.Item("all" & spos) = qty
                End If
            End If
            If tparent = "-" And dchild.exists(spos) Then
                If Not dfirst.exists(spos) Then dfirst.Item(spos) = 1
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
                    length_pos = array_in(i, col_length) / 1000
                    tweight = weight_pm * length_pos * qty
                Case t_prokat
                    qty = array_in(i, col_qty)
                    If (qty = 0) Or IsEmpty(qty) Then qty = 1
                    weight_pm = array_in(i, col_pr_weight)
                    length_pos = array_in(i, col_pr_length) / 1000
                    tweight = weight_pm * length_pos * qty
                Case t_izd
                    qty = array_in(i, col_qty)
                    If (qty = 0) Or IsEmpty(qty) Then qty = 1
                    If array_in(i, col_m_weight) = "-" Then
                        tweight = 0
                    Else
                        tweight = qty * array_in(i, col_m_weight)
                    End If
            End Select
            If tweight Then dweight.Item(subpos) = dweight.Item(subpos) + tweight
        End If
    Next
    'Делим на количество вхождений, чтоб получить массу одной шт.
    For Each subpos In dweight.keys()
        If pos_data.Item("child").exists(subpos) Then
            nSubPos = pos_data.Item("qty").Item("all" & subpos)
        Else
            nSubPos = pos_data.Item("qty").Item("-_" & subpos)
        End If
        If nSubPos < 1 Then
            MsgBox ("Не определено кол-во сборок " & subpos & ", принято 1 шт.")
            nSubPos = 1
        End If
        w = (dweight.Item(subpos) / nSubPos)
        dweight.Item(subpos) = w
    Next
    'Для сборок первого уровня учтём вхождения сборок второго уровня
    For Each subpos In pos_data.Item("parent").keys()
        For Each tchild In pos_data.Item("child").keys()
            If pos_data.Item("qty").exists(subpos & "_" & tchild) Then
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
            lsize = GetSizeSheet(Sh)
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

Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal mask As String = "", Optional ByVal SearchDeep As Long = 2) As Collection
    Set FilenamesCollection = New Collection
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetAllFileNamesUsingFSO FolderPath, mask, FSO, FilenamesCollection, SearchDeep
    Set FSO = Nothing: Application.StatusBar = False
End Function

Function FormatSpec_AS(ByVal Data_out As Range, ByVal n_row As Integer, ByVal n_col As Integer) As Boolean
        For i = 2 To n_row
            If InStr(Data_out(i, 1), ", на ") > 0 Then Range(Cells(i, 1), Cells(i, 6)).Merge
            If InStr(Data_out(i, 1), " Прочие") > 0 Then Range(Cells(i, 1), Cells(i, 6)).Merge
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
    Range(Data_out.Cells(1, 3), Data_out.Cells(1, 3)).ColumnWidth = 0.45 * koeff
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

    dblPoints = Application.CentimetersToPoints(1)
    Range(Data_out.Cells(1, 1), Data_out.Cells(5, n_col)).RowHeight = dblPoints * 1
    Range(Data_out.Cells(6, 1), Data_out.Cells(n_row, n_col)).Rows.AutoFit
    koeff = 6
    Range(Data_out.Cells(1, 1), Data_out.Cells(1, 1)).ColumnWidth = 4 * koeff
    For i = 2 To n_col
        Range(Data_out.Cells(1, i), Data_out.Cells(n_row, i)).ColumnWidth = 1.5 * koeff
    Next i
    r = FormatFont(Data_out, n_row, n_col)
    Range(Data_out.Cells(n_row, 1), Data_out.Cells(n_row, n_col)).Select
    Selection.Font.Bold = True
    Data_out.Cells(n_row, n_col).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
End Function

Function FormatTable(Optional ByVal pos_out As Variant) As Boolean
    r = OutPrepare()
    Set Sh = Application.ThisWorkbook.ActiveSheet
    If IsError(pos_out) Then
        lsize = GetSizeSheet(Sh)
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
            r = FormatSpec_Pol(Data_out, n_row, n_col)
        Case 13
            r = FormatSpec_ASGR(Data_out, n_row, n_col)
    End Select
    r = OutEnded()
    FormatTable = True
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
            nSubPos = 1
        End If
    Else
        nSubPos = 1
    End If
    GetNSubpos = nSubPos
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
    r = ManualSpec(nm, array_out)
    ManualAdd = True
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
            h = InStr(chck_m, chck_a)
            If h Then
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

Function ManualSpec(ByVal nm As String, Optional ByVal add_array As Variant) As Variant
    istart = 2 'Пропускаем шапку
    If IsArray(add_array) Then
        flag_add = 1
        mod_array = ArraySelectParam(add_array, "mod", col_marka)
    Else
        flag_add = 0
        mod_array = Empty
        'If Not ManualCheck() Then
            'ManualSpec = Empty
            'Exit Function
        'End If
    End If
    Set spec_sheet = Application.ThisWorkbook.Sheets(nm)
    sheet_size = GetSizeSheet(spec_sheet)
    n_row = sheet_size(1)
    If n_row = istart Then n_row = n_row + 1
    spec = spec_sheet.Range(spec_sheet.Cells(1, 1), spec_sheet.Cells(n_row, max_col_man))
    Dim pos_out: ReDim pos_out(n_row - istart, max_col): n_row_out = 0
    Dim param
    For i = istart To n_row
        row = ArrayRow(spec, i)

        subpos = row(col_man_subpos)  ' Марка элемента
        pos = row(col_man_pos)  ' Поз.
        obozn = row(col_man_obozn) ' Обозначение
        naen = row(col_man_naen) ' Наименование
        qty = row(col_man_qty) ' Кол-во на один элемент
        Weight = row(col_man_weight) ' Масса, кг
        prim = row(col_man_prim) ' Примечание (на лист)
        
        If qty = Empty Or qty <= 0 Then qty = 1
        type_el = ManualType(row)
        If type_el > 0 Then
            n_row_out = n_row_out + 1

            pos_out(n_row_out, col_marka) = pos
            pos_out(n_row_out, col_sub_pos) = subpos
            pos_out(n_row_out, col_type_el) = type_el
            pos_out(n_row_out, col_pos) = pos
            pos_out(n_row_out, col_qty) = qty
            
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
                
                pos_out(n_row_out, col_pr_type_konstr) = pr_type
                pos_out(n_row_out, col_pr_gost_st) = pr_adress.Item(pr_st)
                pos_out(n_row_out, col_pr_st) = pr_st
                pos_out(n_row_out, col_pr_gost_prof) = pr_adress.Item(pr_gost_pr)(2)
                pos_out(n_row_out, col_pr_prof) = pr_prof
                pos_out(n_row_out, col_pr_length) = pr_length
                pos_out(n_row_out, col_pr_weight) = Weight
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
                pos_out(n_row_out, col_m_weight) = "-"
                pos_out(n_row_out, col_m_edizm) = prim
            End Select
            If flag_add And Not IsEmpty(mod_array) Then
                'Изменяем строки с одинаковыми контрольными суммами
                ReDim param(max_col)
                param = ArrayRow(pos_out, n_row_out)
                current_sum = ControlSumEl(param)
                current_sum = Split(current_sum, "_")(0) & Split(current_sum, "_")(2)
                For kk = 1 To UBound(mod_array, 1)
                    mod_sum = Split(mod_array(kk, col_chksum), "_")(0) & Split(mod_array(kk, col_chksum), "_")(2)
                    If mod_sum = current_sum Then
                        r = CeilSetValue(spec_sheet.Cells(i, col_man_qty), mod_array(kk, col_qty), "mod")
                    End If
                Next kk
            End If
        End If
    Next i

    If flag_add Then
        'Добавим из add_array все новые элементы (в первом столбце нет значения "mod")
        end_row = n_row_out + istart + 1
        For i = 1 To UBound(add_array, 1)
            type_el = add_array(i, col_type_el)
            If add_array(i, col_marka) <> "mod" And type_el <> t_prokat Then
                end_row = end_row + 1
                r = CeilSetValue(spec_sheet.Cells(end_row, col_man_subpos), add_array(i, col_sub_pos), "add")
                r = CeilSetValue(spec_sheet.Cells(end_row, col_man_pos), add_array(i, col_pos), "add")
                r = CeilSetValue(spec_sheet.Cells(end_row, col_man_qty), add_array(i, col_qty), "add")
                Select Case type_el
                Case t_arm
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_naen), "Арматура", "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_obozn), "", "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_weight), "", "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_length), add_array(i, col_length), "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_diametr), add_array(i, col_diametr), "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_klass), add_array(i, col_klass), "add")
                    If add_array(i, col_fon) Then r = CeilSetValue(spec_sheet.Cells(end_row, col_man_prim), "*", "add")
                    If add_array(i, col_gnut) Then r = CeilSetValue(spec_sheet.Cells(end_row, col_man_prim), "п.м.", "add")
                Case t_prokat
                    'r = CeilSetValue(spec_sheet.Cells( end_row, col_man_naen, "Прокат", "add")
                    'r = CeilSetValue(spec_sheet.Cells( end_row, col_man_obozn), add_array(i, col_pr_gost_prof), "add")
                    'r = CeilSetValue(spec_sheet.Cells( end_row, col_man_weight), add_array(i, col_pr_weight), "add")
                    'r = CeilSetValue(spec_sheet.Cells( end_row, col_man_length), add_array(i, col_pr_length), "add")
                    'r = CeilSetValue(spec_sheet.Cells( end_row, col_man_diametr), add_array(i, col_pr_prof), "add")
                    'r = CeilSetValue(spec_sheet.Cells( end_row, col_man_klass), add_array(i, col_pr_st), "add")
                    'r = CeilSetValue(spec_sheet.Cells( end_row, col_man_komment, GetShortNameForGOST(add_array(i, col_pr_gost_prof)), "add")
                Case t_izd
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_obozn), add_array(i, col_m_obozn), "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_naen), add_array(i, col_m_naen), "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_weight), add_array(i, col_m_weight), "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_prim), "", "add")
                Case Else
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_obozn), add_array(i, col_m_obozn), "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_naen), add_array(i, col_m_naen), "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_weight), "", "add")
                    r = CeilSetValue(spec_sheet.Cells(end_row, col_man_prim), add_array(i, col_m_edizm), "add")
                End Select
            End If
        Next i
    Else
        sub_pos_arr = ArraySelectParam(pos_out, t_subpos, col_type_el)
        If Not IsEmpty(sub_pos_arr) Then
            'Из за того, что в дальнейшем количество элементов в сборке делится на количество сборок - нужно домножить количества
            'Для этого сначала получим количество сборок
            'Далее будем вытаскивать элементы для каждой сборки и домножать
            For j = 1 To UBound(sub_pos_arr, 1)
                subpos = sub_pos_arr(j, col_sub_pos)
                qty = sub_pos_arr(j, col_qty)
                For i = 1 To UBound(pos_out, 1)
                    If pos_out(i, col_type_el) <> t_subpos And pos_out(i, col_sub_pos) = subpos Then
                        pos_out(i, col_qty) = pos_out(i, col_qty) * qty
                    End If
                Next i
            Next j
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
            'Поменяв при это обозначение вхождения сборки с t_izd на t_subpos
            Dim subarray: ReDim subarray(max_col, 1)
            For j = 1 To UBound(pos_out, 1)
                If pos_out(j, col_type_el) = t_izd Then
                    pos = pos_out(j, col_pos)
                    naen = pos_out(j, col_m_naen)
                    If subpos_el.exists(pos & naen) Then
                        subpos = pos_out(j, col_sub_pos)
                        pos_out(j, col_marka) = subpos & "." & pos_out(j, col_pos)
                        pos_out(j, col_sub_pos) = pos_out(j, col_pos)
                        pos_out(j, col_type_el) = t_subpos
                        qty = pos_out(j, col_qty)
                        el = subpos_el.Item(pos & naen)
                        qty_from_list = ArraySelectParam(el, t_subpos, col_type_el)(1, col_qty)
                        For k = 1 To UBound(el, 1)
                            If el(k, col_type_el) <> t_subpos Then
                                c_size = UBound(subarray, 2)
                                For i = 1 To max_col
                                    subarray(i, c_size) = el(k, i)
                                Next i
                                subarray(col_marka, c_size) = subpos & "." & el(k, col_pos)
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
            h = 1
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

Function Paste2Sheet(ByRef array_in As Variant) As Boolean
    r = OutPrepare()
    Set Sh = Application.ThisWorkbook.ActiveSheet
    If SpecGetType(Sh.Name) <> 7 Then
        MsgBox ("Перейдите на лист с ручной спецификацией (заканчивается на _спец) и повторите")
        Paste2Sheet = False
        Exit Function
    End If
    startpos = GetSizeSheet(Sh)(1) + 2
    endpos = startpos + UBound(array_in, 1) - 1
    Sh.Range(Sh.Cells(startpos, 1), Sh.Cells(endpos, max_col_man)) = array_in
    r = OutEnded()
    r = ManualCheck()
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

Function ReadPos(ByVal lastfileadd As String) As Variant
    Set add_sheet = Application.ThisWorkbook.Sheets(lastfileadd)
    sheet_size = GetSizeSheet(add_sheet)
    istart = 2
    n_row = sheet_size(1)
    n_col = 6
    spec = add_sheet.Range(add_sheet.Cells(1, 1), add_sheet.Cells(n_row, n_col))
    Dim add_array: ReDim add_array(n_row - istart, max_col): n_row_out = 0
    For i = istart + 1 To n_row
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

Function RelFName(ByVal fname As String) As String
    n_slash = InStrRev(fname, "\")
    n_len = Len(fname)
    n_dot = 4
    wt_dot = Left(fname, n_len - n_dot)
    n_len = Len(wt_dot)
    wt_path = Right(wt_dot, n_len - n_slash)
    RelFName = wt_path
End Function

Function Round_w(ByVal arg As Variant, ByVal nokrg As Integer) As Variant
    If IsNumeric(arg) Then
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

Function Sheet2Pdf(ByVal Data_out As Range, ByVal filename As String, Optional ByVal type_print As Integer = 0) As Boolean
    Data_out.Select
    Application.PrintCommunication = True
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
                .PrintQuality = 600
                .CenterHorizontally = False
                .CenterVertically = False
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA3
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = 100
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
                .PrintTitleRows = "$1:$2"
                .PrintTitleColumns = ""
            End With
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
                .PrintQuality = 600
                .CenterHorizontally = False
                .CenterVertically = False
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA3
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = True
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
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
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, filename:=filename$, Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Sheet2Pdf = True
End Function

Function SheetExport()
    r = OutPrepare()
    nm = ActiveWorkbook.ActiveSheet.Name
    If SpecGetType(nm) <> 7 Then
        Set Sh = ActiveWorkbook.ActiveSheet
        lsize = GetSizeSheet(Sh)
        n_row = lsize(1)
        n_col = lsize(2)
        Set Data_out = Sh.Range(Sh.Cells(1, 1), Sh.Cells(n_row, n_col))
        filename$ = ThisWorkbook.path & "\list\Спец_" & nm & ".pdf"
        type_print = 0
        h = GetHeightSheet()
        If SpecGetType(nm) = "11" Then type_print = 1
        If GetHeightSheet() > 420 Then r = SetPageBreaks(420, 2)
        r = Sheet2Pdf(Data_out, filename, type_print)
    End If
    r = OutEnded()
End Function

Function SheetHideAll()
    Worksheets("|Содержание|").Activate
    Dim sheet As Worksheet
    With ActiveWorkbook
        For Each sheet In ActiveWorkbook.Worksheets
            If Left(sheet.Name, 1) = "!" Then Sheets(sheet.Name).Visible = False
        Next
    End With
End Function

Function SheetIndex()
    Worksheets("|Содержание|").Activate
    Dim sheet As Worksheet
    Dim cell As Range
    Range("A3:D500").ClearContents
    With ActiveWorkbook
        For Each sheet In ActiveWorkbook.Worksheets
            If Sheets(sheet.Name).Visible = 0 And Left(sheet.Name, 1) <> "|" And Left(sheet.Name, 1) <> "!" Then sheet.Name = "!" & sheet.Name
            If Left(sheet.Name, 1) = "|" Then
                Set cell = Worksheets(1).Cells(sheet.Index + 2, 1)
                .Worksheets(1).Hyperlinks.Add anchor:=cell, Address:="", SubAddress:="'" & sheet.Name & "'" & "!A1"
                cell.Formula = sheet.Name
                Sheets(sheet.Name).Visible = True
            Else
                 If Left(sheet.Name, 1) = "!" Then
                    Set cell = Worksheets(1).Cells(sheet.Index + 2, 2)
                    .Worksheets(1).Hyperlinks.Add anchor:=cell, Address:="", SubAddress:="'" & sheet.Name & "'" & "!B2"
                    cell.Formula = sheet.Name
                    Sheets(sheet.Name).Visible = False
                Else
                    If SpecGetType(sheet.Name) = 7 Or Right(sheet.Name, 4) = "_поз" Then
                        Set cell = Worksheets(1).Cells(sheet.Index + 2, 3)
                        .Worksheets(1).Hyperlinks.Add anchor:=cell, Address:="", SubAddress:="'" & sheet.Name & "'" & "!C3"
                        cell.Formula = sheet.Name
                        Sheets(sheet.Name).Visible = True
                    Else
                        Set cell = Worksheets(1).Cells(sheet.Index + 2, 4)
                        .Worksheets(1).Hyperlinks.Add anchor:=cell, Address:="", SubAddress:="'" & sheet.Name & "'" & "!D4"
                        cell.Formula = sheet.Name
                        Sheets(sheet.Name).Visible = True

                    End If
                End If
            End If
            Sheets("|Содержание|").Visible = True
        Next
    End With
    Range("A3:C500").Rows.AutoFit
End Function

Function SheetNew(ByVal NameSheet As String)
    On Error Resume Next
    If SheetExist(NameSheet) Then
       Worksheets(NameSheet).Cells.Clear
    Else
        ThisWorkbook.Worksheets.Add.Name = NameSheet
    End If
End Function

Function SheetShowAll()
    Worksheets("|Содержание|").Activate
    Dim sheet As Worksheet
    With ActiveWorkbook
        For Each sheet In ActiveWorkbook.Worksheets
            Sheets(sheet.Name).Visible = True
        Next
    End With
End Function

Function SpecArm(ByVal arm As Variant, ByVal n_arm As Integer, _
                 ByVal type_spec As Integer, ByVal nSubPos As Integer) As Variant
    Dim pos_out
    If UserForm2.qtyOneSubpos_CB.Value Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
        
    Else
        n_txt = "," & vbLf & "на все"
    End If
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
                prim = "": If arm(j, col_gnut) Then prim = "*"
                qty = arm(j, col_qty)
                n_el = qty / nSubPos
                length_pos = arm(j, col_length) / 1000
                l_spec = (length_pos * n_el)
                Select Case type_spec
                Case 1
                    pos_out(i, 1) = arm(j, col_sub_pos) & n_txt
                    If (UserForm2.keep_pos_CB.Value And UserForm2.arm_pm_CB.Value) Or Not (UserForm2.arm_pm_CB.Value) Then
                        pos_out(i, 2) = arm(j, col_pos) & prim
                    Else
                        pos_out(i, 2) = " "
                    End If
                    If fon Then
                        If prim = "п.м." Then prim = ""
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L= п.м." & prim
                        pos_out(i, 4) = pos_out(i, 4) + l_spec
                        pos_out(i, 5) = weight_pm
                    Else
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L=" & length_pos * 1000 & "мм." & prim
                        pos_out(i, 4) = pos_out(i, 4) + n_el
                        pos_out(i, 5) = Round_w(weight_pm * length_pos, n_round_w)
                    End If
                Case Else
                    If (UserForm2.keep_pos_CB.Value And UserForm2.arm_pm_CB.Value) Or Not (UserForm2.arm_pm_CB.Value) Then
                        pos_out(i, 1) = arm(j, col_pos) & prim
                    Else
                        pos_out(i, 1) = " "
                    End If
                    pos_out(i, 2) = GetGOSTForKlass(klass)
                    If fon Then
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L= п.м."
                        pos_out(i, 4) = pos_out(i, 4) + l_spec
                        pos_out(i, 5) = weight_pm
                        pos_out(i, 6) = pos_out(i, 6) + l_spec * weight_pm
                    Else
                        pos_out(i, 3) = symb_diam & diametr & " " & klass & " L=" & length_pos * 1000 & "мм."
                        pos_out(i, 4) = pos_out(i, 4) + n_el
                        pos_out(i, 5) = Round_w(weight_pm * length_pos, n_round_w)
                        pos_out(i, 6) = pos_out(i, 6) + n_el * pos_out(i, 5)
                    End If
                End Select
            End If
        Next j
    Next i
    
    For i = 1 To UBound(pos_out, 1)
        pos_out(i, 4) = Round_w(pos_out(i, 4), n_round_l)
        If type_spec = 13 Then pos_out(i, 7) = t_arm
    Next
    
    If type_spec = 1 Then
        pos_out = ArraySort(pos_out, 2)
    Else
        pos_out = ArraySort(pos_out, 1)
    End If
    
    SpecArm = pos_out
    Erase arm, pos_out
End Function

Function SpecIzd(ByVal izd As Variant, ByVal n_izd As Integer, _
                 ByVal type_spec As Integer, ByVal nSubPos As Integer) As Variant
    
    If UserForm2.qtyOneSubpos_CB.Value Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
        
    Else
        n_txt = "," & vbLf & "на все"
        nSubPos = 1
    End If

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
                n_el = izd(j, col_qty) / nSubPos
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
                    pos_out(i, 3) = naen
                    pos_out(i, 4) = pos_out(i, 4) + n_el
                    pos_out(i, 5) = Weight
                    If IsNumeric(Weight) Then
                        pos_out(i, 6) = pos_out(i, 6) + Round_w(n_el * Weight, n_round_w)
                    Else
                        pos_out(i, 6) = izd(j, col_m_edizm)
                    End If
                End Select
            End If
        Next j
    Next i
    
    If type_spec = 13 Then
        For i = 1 To UBound(pos_out, 1)
            pos_out(i, 7) = t_izd
        Next
    End If
    
    If type_spec = 1 Then
        pos_out = ArraySort(pos_out, 2)
    Else
        pos_out = ArraySort(pos_out, 1)
    End If
    
    SpecIzd = pos_out
    Erase izd, pos_out
End Function

Function SpecMaterial(ByVal mat As Variant, ByVal n_mat As Integer, _
                      ByVal type_spec As Integer, ByVal nSubPos As Integer) As Variant
   
    If UserForm2.qtyOneSubpos_CB.Value Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
    Else
        n_txt = "," & vbLf & "на все"
        nSubPos = 1
    End If

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
                    pos_out(i, 4) = pos_out(i, 4) + Round_w(mat(j, col_qty) / nSubPos, n_round_w)
                    pos_out(i, 5) = "-"
                Case Else
                    pos_out(i, 1) = " "
                    pos_out(i, 2) = mat(j, col_m_obozn)
                    pos_out(i, 3) = mat(j, col_m_naen)
                    pos_out(i, 4) = pos_out(i, 4) + Round_w(mat(j, col_qty) / nSubPos, n_round_w)
                    pos_out(i, 5) = "-"
                    pos_out(i, 6) = mat(j, col_m_edizm)
                End Select
            End If
        Next j
    Next i
    
    If type_spec = 13 Then
        For i = 1 To UBound(pos_out, 1)
            pos_out(i, 7) = t_mat
        Next
    End If
    
    If type_spec = 1 Then
        pos_out = ArraySort(pos_out, 2)
    Else
        pos_out = ArraySort(pos_out, 1)
    End If
    
    SpecMaterial = pos_out
    Erase mat, un_pos_mat, pos_out
End Function

Function SpecOneSubpos(ByVal all_data As Variant, ByVal subpos As String, _
                       ByVal type_spec As Integer) As Variant
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
    
    n_col_spec = 6
    If type_spec = 13 Then n_col_spec = n_col_spec + 1
    If type_spec = 2 Then
        ReDim pos_naen(n_n, n_col_spec)
        If subpos <> "-" Then
            naen = subpos
            If pos_data.Item("name").Count Then
                If pos_data.Item("name").exists(subpos) Then naen = pos_data.Item("name").Item(subpos)(1)
                If UserForm2.qtyOneSubpos_CB.Value Then
                    pos_naen(n_n, 1) = naen & ", на 1 шт. (всего " & nSubPos & " шт.)"
                Else
                    pos_naen(n_n, 1) = naen & ", на все"
                End If
            End If
        Else
            pos_naen(n_n, 1) = " Прочие элементы "
        End If
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
                    u2 = (pos_data.Item("-").exists(сurrent_subpos) And (сurrent_parent = "-") And (сurrent_type_el = t_subpos))   'Элементы вложенных сборок
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
                        subp(n_subpos, col_m_weight) = pos_data.Item("weight").Item(сurrent_subpos)
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
        g = SpecSubpos(subp, n_subpos, type_spec, nSubPos, name_subpos)
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
        If type_spec = 1 Then
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
            pos_out(1, 6) = Round_w(pos_out(1, 6), 0)
        End If
        SpecOneSubpos = pos_out
        Erase pos_out
    End If
End Function

Function SpecProkat(ByVal prokat As Variant, ByVal n_prokat As Integer, _
                    ByVal type_spec As Integer, Optional ByVal nSubPos As Integer = 1) As Variant
    
    If UserForm2.qtyOneSubpos_CB.Value Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
    Else
        n_txt = "," & vbLf & "на все"
    End If
    
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
                l = Round_w(prokat(j, col_pr_length) / 1000, n_round_l)
                If UserForm2.pr_pm_CB.Value Then
                    we = Round_w(prokat(j, col_pr_weight), n_round_w)
                Else
                    we = Round_w(prokat(j, col_pr_weight) * l, n_round_w)
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
                                pos_out(i, 3) = name_pr & prokat(j, col_pr_gost_prof) & " " & prokat(j, col_pr_naen) & " S = кв.м."
                            Else
                                pos_out(i, 3) = name_pr & prokat(j, col_pr_gost_prof) & " " & prokat(j, col_pr_prof) & " L = п.м."
                            End If
                            pos_out(i, 4) = pos_out(i, 4) + (l * n_el)
                            pos_out(i, 5) = we
                        Else
                            pos_out(i, 1) = prokat(j, col_sub_pos) & n_txt
                            pos_out(i, 2) = prokat(j, col_pos)
                            pos_out(i, 3) = name_pr & prokat(j, col_pr_prof) & " L=" & l * 1000 & "мм."
                            pos_out(i, 4) = pos_out(i, 4) + n_el
                            pos_out(i, 5) = we
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
                                pos_out(i, 3) = name_pr & " " & prokat(j, col_pr_naen) & " S = кв.м."
                            Else
                                pos_out(i, 3) = name_pr & " " & prokat(j, col_pr_prof) & " L = п.м."
                            End If
                            pos_out(i, 4) = pos_out(i, 4) + (l * n_el)
                            If pos_out(i, 5) <> we And pos_out(i, 5) > 0 Then
                                hh = 1
                            End If
                            pos_out(i, 5) = we
                            pos_out(i, 6) = pos_out(i, 6) + Round_w((l * n_el * we), n_round_w)
                        Else
                            pos_out(i, 1) = prokat(j, col_pos)
                            pos_out(i, 2) = prokat(j, col_pr_gost_prof)
                            If InStr(1, name_pr, "Лист") Then
                                pos_out(i, 3) = name_pr & " " & prokat(j, col_pr_naen) & " S=" & l & "кв.м."
                            Else
                                pos_out(i, 3) = name_pr & prokat(j, col_pr_prof) & " L=" & l * 1000 & "мм."
                            End If
                            pos_out(i, 4) = pos_out(i, 4) + n_el
                            pos_out(i, 5) = we
                            pos_out(i, 6) = pos_out(i, 6) + Round_w((n_el * we), n_round_w)
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
        pos_out = ArraySort(pos_out, 2)
    Else
        pos_out = ArraySort(pos_out, 1)
    End If
    
    SpecProkat = pos_out
    Erase prokat, pos_out
End Function

Function SpecSubpos(ByVal subp As Variant, ByVal n_subp As Integer, _
                    ByVal type_spec As Integer, ByVal nSubPos As Integer, _
                    ByVal name_subpos As Variant) As Variant
    
    If UserForm2.qtyOneSubpos_CB.Value Then
        n_txt = vbLf & "(" & nSubPos & " шт.)"
    Else
        n_txt = "," & vbLf & "на все"
    End If

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
                Weight = Round_w(subp(j, col_m_weight), n_round_w)
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
                        If pos_data.Item("name").exists(pos) Then
                            naen = pos_data.Item("name").Item(pos)(1)
                            obozn = pos_data.Item("name").Item(pos)(2)
                        End If
                    End If
                    pos_out(i, 1) = pos
                    pos_out(i, 2) = obozn
                    pos_out(i, 3) = naen
                    pos_out(i, 4) = pos_out(i, 4) + n_el
                    pos_out(i, 5) = Weight
                    pos_out(i, 6) = subp(j, col_m_edizm)
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
        pos_out = ArraySort(pos_out, 2)
    Else
        pos_out = ArraySort(pos_out, 1)
    End If
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
        MsgBox ("Сборки отсутвуют. Создана общестроительная спецификаця")
        type_spec = 3
    End If
    If (qty_parent <= 1) Or (qty_parent < 1 And pos_data.exists("-")) And type_spec = 13 Then
        MsgBox ("Сборок меньше двух. Создана общестроительная спецификаця")
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
            If pos_data.exists("-") Then end_col = end_col + 1
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
        
    Dim ch_key As String
    ch_key = "child"
    If qty_child <= 0 And type_spec = 1 Then
        If qty_parent >= 0 Then
            ch_key = "parent"
        Else
            Spec_AS = Empty
            Exit Function
        End If
    End If
    
    If type_spec = 1 Then
        For Each subpos In ArraySort(pos_data.Item(ch_key).keys(), 1)
            pos_out = ArrayCombine(pos_out, SpecOneSubpos(all_data, subpos, type_spec))
        Next
        Dim pos_end: ReDim pos_end(1, 6)
        If UserForm2.qtyOneSubpos_CB.Value Then
            pos_end(1, 1) = Space(60) & "* расход на одно изделие"
        Else
            pos_end(1, 1) = Space(60) & "* расход на все изделия"
        End If
        pos_out = ArrayCombine(pos_out, pos_end)
    End If
    
    If type_spec = 2 Then
        For Each subpos In ArraySort(pos_data.Item("parent").keys(), 1)
            pos_out = ArrayCombine(pos_out, SpecOneSubpos(all_data, subpos, type_spec))
        Next
        If pos_data.exists("-") Then pos_out = ArrayCombine(pos_out, SpecOneSubpos(all_data, "-", type_spec))
    End If
    
    If type_spec = 3 Then
        If (pos_data.exists("-") Or (UserForm2.show_subpos_CB.Value And (UBound(pos_data.Item("parent").keys()) >= 0))) Then
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
        If pos_data.exists("-") Then
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
        If n_row_subpos > 0 Then pos_out = ArrayCombine(pos_out, ArrayRedim(pos_out_subpos, n_row_subpos))
        If n_row_arm > 0 Then pos_out = ArrayCombine(pos_out, ArrayRedim(pos_out_arm, n_row_arm))
        If n_row_prokat > 0 Then pos_out = ArrayCombine(pos_out, ArrayRedim(pos_out_prokat, n_row_prokat))
        If n_row_izd > 0 Then pos_out = ArrayCombine(pos_out, ArrayRedim(pos_out_izd, n_row_izd))
        If n_row_mat > 0 Then pos_out = ArrayCombine(pos_out, ArrayRedim(pos_out_mat, n_row_mat))
        For i = 3 To UBound(pos_out, 1)
            If Not IsEmpty(pos_out(i, end_col - 1)) Then
                For j = 4 To end_col - 1
                    If IsEmpty(pos_out(i, j)) Then pos_out(i, j) = "-"
                Next j
            End If
            If IsNumeric(pos_out(i, end_col)) Then
                pos_out(i, end_col) = Str(Round_w(pos_out(i, end_col), 0)) & " кг."
            End If
        Next i
    Else
        For i = 2 To UBound(pos_out, 1)
            If IsNumeric(pos_out(i, 6)) And pos_out(i, 6) > 0 Then
                pos_out(i, 6) = Str(Round_w(pos_out(i, 6), 0)) & " кг."
            End If
        Next i
    End If
    Spec_AS = pos_out
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

Function Spec_AS2arr(ByVal filename As String) As Variant
    all_data = DataRead(filename & ".txt")
    data_t = ReadFile(filename & ".txt")
    type_spec = 3
    spec = SpecOneSubpos("-", type_spec)
    Spec_AS2arr = ArrayEmp2Space(spec)
End Function

Function Spec_KM(ByRef all_data As Variant) As Variant
    prokat = ArraySelectParam(all_data, t_prokat, col_type_el)
    If IsEmpty(prokat) Then
        n_prokat = 0
        MsgBox ("Прокат в файле/листе не найден")
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
        n_okr = n_round_w
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
            For l = 1 To n_prof
                'Все элементы с заданным сечением
                konstr = unique_prof(l) 'Текущий типоразмер профиля
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
            Next l
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
    Erase prokat, unique_gost_prof, _
          unique_stal, prof_stal, unique_prof_stal, _
          unique_type_konstr, prof, unique_prof, el, _
          unique_konstr, elem_m, weight_stal, weight_gost_prof, weight_total, weight_stal_total
    pos_out = ArrayRedim(pos_out, row - 1)
    Spec_KM = pos_out
End Function


Function Spec_KZH(ByRef all_data As Variant) As Variant
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
            If name_subpos.exists(сurrent_subpos) Then naen = name_subpos.Item(сurrent_subpos)(1)
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
    If pos_data.exists("-") Then n_row = n_row + 1
    sum_row = 0: If n_row - 5 > 1 Then sum_row = 1
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
                    For l = 1 To n_gost
                        current_gost = un_gost(l)
                        pr_current_gost = ArraySelectParam(pr_current_slal, current_gost, col_pr_gost_prof)
                        un_prof = ArrayUniqValColumn(pr_current_gost, col_pr_prof)
                        n_prof = UBound(un_prof, 1)
                        current_size = UBound(pos_out, 2)
                        ReDim Preserve pos_out(n_row, current_size + n_prof + 1)
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
                    Next l
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
        If (pos_data.exists("-") And k = n_row) Or (UBound(un_parent) <= 0) Then
            subpos = "-"
            nSubPos = 1
        Else
            subpos = un_parent(k - 5)
            nSubPos = pos_data.Item("qty").Item("-_" & subpos)
            If nSubPos < 1 Then
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
        pos_out(k, 1) = n_txt
        If subpos = "-" Then pos_out(k, 1) = "Прочие"
        weight_index.Item("row" & subpos) = k
    Next k
    
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
            u1 = (pos_data.Item("parent").exists(subpos) Or pos_data.Item("parent").exists(tparent))
            If pos_data.exists("-") Then u2 = ((subpos = "-") Or (pos_data.Item("-").exists(subpos) And tparent = "-"))
            If u1 Or u2 Then
                If u2 Then
                    nSubPos = 1
                    k = weight_index.Item("row" & tparent)
                End If
                If u1 Then
                    If pos_data.Item("parent").exists(subpos) Then
                        nSubPos = pos_data.Item("qty").Item("-_" & subpos)
                        k = weight_index.Item("row" & subpos)
                    End If
                    If pos_data.Item("parent").exists(tparent) Then
                        nSubPos = pos_data.Item("qty").Item("-_" & tparent)
                        k = weight_index.Item("row" & tparent)
                    End If
                End If
                If arm_arr(i)(j, col_type_el) = t_arm Then
                    diametr = arm_arr(i)(j, col_diametr)
                    klass = arm_arr(i)(j, col_klass)
                    qty = arm_arr(i)(j, col_qty)
                    gost = GetGOSTForKlass(klass)
                    length_pos = arm_arr(i)(j, col_length) / 1000
                    weight_pm = GetWeightForDiametr(diametr, klass)
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
                    tkeyd = "Прокат" & n_type & stal & vbLf & gost_stal & gost_prof & prof
                    tkesum_1 = "Прокат" & n_type & stal & vbLf & gost_stal & gost_prof & "Всего"
                End If
                If Not UserForm2.qtyOneSubpos_CB.Value Then nSubPos = 1
                l_spec = (length_pos * qty) / nSubPos
                w_pos = Round_w(weight_pm * l_spec, n_round_w)
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
    End If
    For i = 2 To UBound(pos_out, 2)
        For j = 6 To n_row
            If IsEmpty(pos_out(j, i)) Then pos_out(j, i) = "-"
        Next
    Next
    Spec_KZH = pos_out
End Function

Function Spec_Select(ByVal lastfilespec As String, ByVal suffix As String)
    If SpecGetType(lastfilespec) = 7 Then
        nm = Split(lastfilespec, "_")(0) & suffix
    Else
        nm = lastfilespec & suffix
    End If
    type_spec = SpecGetType(nm)
    Select Case type_spec
        Case 11
            all_data = ReadVed(nm)
        Case 12
            all_data = ReadPol(nm)
        Case Else
            If IsEmpty(pr_adress) Then r = ReadPrSortament()
            all_data = DataRead(lastfilespec)
    End Select
    If IsEmpty(all_data) Then
        MsgBox ("Данные отсутвуют")
        Exit Function
    End If
    If SheetExist(nm) Then
        Worksheets(nm).Activate
    Else
        ThisWorkbook.Worksheets.Add.Name = nm
    End If
    Select Case type_spec
        Case 1, 2, 3, 13
            pos_out = Spec_AS(all_data, type_spec)
        Case 4
            pos_out = Spec_KM(all_data)
        Case 5
            pos_out = Spec_KZH(all_data)
        Case 11
            pos_out = Spec_VED(all_data)
        Case 12
            pos_out = Spec_POL(all_data)
    End Select
    If Not IsEmpty(pos_out) Then
        Worksheets(nm).Cells.Clear
        r = FormatTable(pos_out)
        r = FormatTable()
    Else
        MsgBox ("Данные отсутвуют")
    End If
End Function

