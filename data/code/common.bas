Attribute VB_Name = "common"
Option Compare Text
Option Base 1

Public Const common_version As String = "3.95"
Public Const Pi As Double = 3.141592653589
Public ank_data As Variant
Public Function GetLeghtByID(id As String, table As Range, n_col_id As Integer, n_col_l As Integer) As Variant
    GetLeghtByID = Сумма_ПоИД(id, table, n_col_id, n_col_l)
End Function
Public Function Сумма_ПоИД(id As String, table As Range, n_col_id As Integer, n_col_l As Integer) As Variant
Attribute Сумма_ПоИД.VB_Description = "Суммирует элементов с одинаковым ID из таблицы. Ввод в формате ID, Таблица, Номер столбца ID, Номер столбца суммируемой величины"
    'Сумма значений заданного столбца для заданного ID
    Sum_l = 0
    id = Trim(id)
    For i = 1 To table.Rows.Count
        If Trim(table(i, n_col_id)) = id Then
            ll = table(i, n_col_l)
            If IsNumeric(ll) Then Sum_l = Sum_l + ll
        End If
    Next i
    Сумма_ПоИД = Sum_l
End Function

Public Function Макс_ПоИД(id As String, table As Range, n_col_id As Integer, n_col_l As Integer) As Variant
    'Выбор максимального значения из таблицы по заданному ID (макс. высота стен, макс. площадь и т.д.)
    max_l = 0
    id = Trim(id)
    For i = 1 To table.Rows.Count
        If Trim(table(i, n_col_id)) = id Then
            ll = table(i, n_col_l)
            If IsNumeric(ll) Then
                If ll > max_l Then max_l = ll
            End If
        End If
    Next i
    Макс_ПоИД = max_l
End Function

Public Function Мин_ПоИД(id As String, table As Range, n_col_id As Integer, n_col_l As Integer) As Variant
    'Выбор минимального значения из таблицы по заданному ID (мин. высота стен, мин. площадь и т.д.)
    min_l = Макс_ПоИД(id, table, n_col_id, n_col_l)
    id = Trim(id)
    For i = 1 To table.Rows.Count
        If Trim(table(i, n_col_id)) = id Then
            ll = table(i, n_col_l)
            If IsNumeric(ll) Then
                If ll < min_l Then min_l = ll
            End If
        End If
    Next i
    Мин_ПоИД = min_l
End Function

Public Function SetPlast_T(diam As Integer) As String
    Select Case diam
        Case 16
            SetPlast_T = "-- 8"
        Case 20, 22
            SetPlast_T = "-- 10"
        Case 25, 28
            SetPlast_T = "-- 14"
    End Select
End Function

Private Function set_ank_data()
    Set ank_data = CreateObject("Scripting.Dictionary")
    'Rbt
    ank_data.Item("7.5") = 0.48
    ank_data.Item("10") = 0.57
    ank_data.Item("12.5") = 0.66
    ank_data.Item("15") = 0.75
    ank_data.Item("20") = 0.9
    ank_data.Item("25") = 1.05
    ank_data.Item("30") = 1.15
    ank_data.Item("35") = 1.3
    ank_data.Item("40") = 1.4
    ank_data.Item("45") = 1.45
    ank_data.Item("50") = 1.55
    ank_data.Item("55") = 1.6
    ank_data.Item("60") = 1.65
    'Rs
    ank_data.Item("A-I(A240)") = 210
    ank_data.Item("A-III(A400)") = 350
    ank_data.Item("A500C") = 435
    'n
    ank_data.Item("анк_сжатая") = 0.75
    ank_data.Item("анк_растянутая") = 1
    ank_data.Item("анк_двойная") = 2
    
    ank_data.Item("нахл_сжатая") = 0.9
    ank_data.Item("нахл_растянутая") = 1.2
    ank_data.Item("нахл_двойная") = 2
End Function

Private Function get_lo_arm(ByVal diam As Integer, ByVal class As String, beton As String) As Double
    set_ank_data
    beton = GetClassBeton(beton)
    Rs = ank_data.Item(class)
    Rbt = ank_data.Item(beton)
    Rbond = 2.5 * 1 * Rbt
    Areas = Pi * (diam * diam) / 4
    Perims = Pi * diam
    lo = (Rs * Areas) / (Rbond * Perims)
    get_lo_arm = lo
End Function

Public Function Арм_Анкеровка(ByVal diam As Integer, ByVal class As String, ByVal beton As String, _
                            Optional ByVal kseism As Double = 1, _
                            Optional ByVal type_arm As String = "растянутая", _
                            Optional ByVal type_out As String = "L") As Variant
    type_out = StrConv(type_out, vbUpperCase)
    If type_out <> "L" And type_out <> "D" Then type_out = "L"
    type_arm = "анк_" + StrConv(type_arm, vbLowerCase)
    If type_arm <> "анк_растянутая" And type_arm <> "анк_сжатая" And type_arm <> "анк_двойная" Then type_arm = "анк_растянутая"
    If kseism <= 0.9 Then kseism = 1
    If kseism >= 2 Then kseism = 1.3
    
    lo = get_lo_arm(diam, class, beton)
    al = ank_data.Item(type_arm)
    L = al * lo * kseism
    l1 = 15 * diam
    l2 = 200
    l3 = 0.3 * lo
    lout = Application.WorksheetFunction.Max(L, l1, l2, l3)
    krat = "5мм"
    lout = Арм_Округление(lout, krat)
    Select Case StrConv(type_out, vbUpperCase)
        Case "L"
            Арм_Анкеровка = lout
        Case "D"
            Арм_Анкеровка = Round((lout / diam), 2) & "d"
    End Select
    If Not ank_data.exists(class) Then Арм_Анкеровка = "ОШИБКА КЛАССА"
    If Not ank_data.exists(beton) Then Арм_Анкеровка = "ОШИБКА БЕТОНА"
End Function

Public Function Арм_Округление(ByVal L As Long, Optional ByVal krat As String = "10мм") As Long
    lout = L
    krat_n = 5
    If krat = "5мм" Then krat_n = 5
    If krat = "10мм" Then krat_n = 10
    l_round = Round(L / krat_n) * krat_n
    If l_round < L Then
        lout = l_round + krat_n
    Else
        lout = l_round
    End If
    Арм_Округление = lout
End Function

Public Function Арм_Нахлёст(ByVal diam As Integer, ByVal class As String, ByVal beton As String, _
                            Optional ByVal kseism As Double = 1, _
                            Optional ByVal type_arm As String = "растянутая", _
                            Optional ByVal type_out As String = "L") As Variant
    If diam > 40 Then
        Нахлёст = "ДИАМЕТР БОЛЬШЕ 40"
        Exit Function
    End If
    type_out = StrConv(type_out, vbUpperCase)
    If type_out <> "L" And type_out <> "D" Then type_out = "L"
    type_arm = "нахл_" + StrConv(type_arm, vbLowerCase)
    If type_arm <> "нахл_растянутая" And type_arm <> "нахл_сжатая" And type_arm <> "нахл_двойная" Then type_arm = "нахл_растянутая"
    If kseism <= 0.9 Then kseism = 1
    If kseism >= 2 Then kseism = 1.3
    lo = get_lo_arm(diam, class, beton)
    al = ank_data.Item(type_arm)
    L = al * lo * kseism
    l1 = 20 * diam
    l2 = 250
    l3 = 0.4 * al * lo
    lout = Application.WorksheetFunction.Max(L, l1, l2, l3)
    krat = "5мм"
    lout = Арм_Округление(lout, krat)
    Select Case StrConv(type_out, vbUpperCase)
        Case "L"
            Арм_Нахлёст = lout
        Case "D"
            Арм_Нахлёст = Round((lout / diam), 2) & "d"
    End Select
End Function

Public Function Арм_МинРадиус(ByVal diam As Integer, ByVal class As String) As Double
    d_opr = 0
    If class = "A-I(A240)" Or class = "Вр-I" Then
        If diam < 20 Then
            d_opr = 2.5 * diam
        Else
            d_opr = 4 * diam
        End If
    Else
        If diam < 20 Then
            d_opr = 5 * diam
        Else
            d_opr = 8 * diam
        End If
    End If
    r_opr = d_opr / 2
    r_arm = diam / 2
    Арм_МинРадиус = r_opr + r_arm
End Function

Public Function Арм_Элемент_П(ByVal L As Variant, ByVal H As Variant, ByVal diam As Integer, ByVal class As String, Optional ByVal Lniz As Integer = 0, Optional ByVal diam_osn As Integer = 0) As Double
    If Lniz = 0 Then Lniz = L
    agib = 90
    r_min = Арм_МинРадиус(diam, class)
    r_osn = diam_osn / 2 + diam / 2
    'Диаметр гиба - не менее минимального
    If r_osn > r_min Then
        r = r_osn
    Else
        r = r_min
    End If
    lr = (Pi * r * agib) / 180
    lout = L + Lniz + H - 4 * r + 2 * lr
    krat = "10мм"
    lout = Арм_Округление(lout, krat)
    Арм_Элемент_П = lout
End Function

Public Function Арм_Элемент_Г(ByVal L As Variant, ByVal H As Variant, ByVal diam As Integer, ByVal class As String, Optional ByVal diam_osn As Integer = 0) As Double
    agib = 90
    r_min = Арм_МинРадиус(diam, class)
    r_osn = diam_osn / 2 + diam / 2
    'Диаметр гиба - не менее минимального
    If r_osn > r_min Then
        r = r_osn
    Else
        r = r_min
    End If
    lr = (Pi * r * agib) / 180
    lout = L + H - 2 * r + lr
    krat = "10мм"
    lout = Арм_Округление(lout, krat)
    Арм_Элемент_Г = lout
End Function

Public Function Арм_Элемент_Хомут(ByVal L As Integer, ByVal H As Integer, ByVal diam_osn As Integer, ByVal diam As Integer, ByVal class As String) As Double
    r_min = Арм_МинРадиус(diam, class)
    r_osn = diam_osn / 2 + diam / 2
    'Диаметр гиба - не менее минимального
    If r_osn > r_min Then
        r = r_osn
    Else
        r = r_min
    End If
    'Длина загиба на 90
    agib = 90
    lr_90 = (Pi * r * agib) / 180
    'Длина загиба на 135 (сейсмика)
    agib = 135
    lr_135 = (Pi * r * agib) / 180
    'Заведение стержня в тело - 75 или 6d
    If 6 * diam > 75 Then
        ank = 6 * diam
    Else
        ank = 75
    End If
    lout = (L + H + ank) * 2 + lr_90 * 3 + lr_135 * 2
    krat = "10мм"
    lout = Арм_Округление(lout, krat)
    Арм_Элемент_Хомут = lout
End Function

Public Function Арм_Длина_ПМ(ByVal L As Variant, ByVal lnahl As Variant, Optional ByVal led As Variant = 11700) As Long
    n_nahl = Round(L / led)
    If n_nahl * led < L Then n_nahl = n_nahl + 1
    lout = L + lnahl * n_nahl
    lout = Арм_Округление(lout)
    Арм_Длина_ПМ = lout
End Function

Public Function Арм_ПоПлощади(S As Variant, shag As Variant, ByVal lnahl As Integer, Optional ByVal led As Integer = 11700) As Long
    L = Sqr(S) + (S / shag)
    lout = Арм_Длина_ПМ(L, lnahl, led)
    Арм_ПоПлощади = lout
End Function

Public Function Арм_ОдинСлойПоПлощади(S As Variant, shag As Variant, ByVal lnahl As Integer, Optional ByVal led As Integer = 11700) As Long
    Арм_ОдинСлойПоПлощади = Арм_ПоПлощади(S, shag, lnahl, led)
End Function

Public Function SetPlast_Razm(diam As Integer) As String
    Select Case diam
        Case 16
            SetPlast_Razm = "100*100"
        Case 20, 22
            SetPlast_Razm = "120*120"
        Case 25, 28
            SetPlast_Razm = "150*150"
    End Select
End Function
