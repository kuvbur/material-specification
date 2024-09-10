VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   ClientHeight    =   10755
   ClientLeft      =   50
   ClientTop       =   390
   ClientWidth     =   5010
   OleObjectBlob   =   "UserForm2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Text
Option Base 1

Const form_version As String = "4.03"
Const form1_version As String = "4.01"
Public CodePath, MaterialPath, SortamentPath As String
Public lastsheet, lastconstrtype, lastconstr, lastfile, lastfilespec, lastfileadd, materialbook_index, name_izd As Variant

#If Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As LongPtr) As Long
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Private Sub AllPosButton_Click()
    ans = MsgBox("�������� ����� �� ������. ������� ��, ���� ������� ��������� �����", vbYesNo)
    If ans = 6 Then
        r = OutPrepare()
        nm = ThisWorkbook.ActiveSheet.Name
        type_spec = SpecGetType(nm)
        If type_spec = 7 Then
            r = ManualPos(nm, 1)
            r = LogWrite(nm, "��", "�������������� �������� ����������")
        Else
            If Not (quiet) Then MsgBox ("��������� �� ���� _���� � ���������")
        End If
        r = OutEnded()
    End If
    r = print_functime()
End Sub

Private Sub CommandButtonExportData_Click()
    r = OutPrepare()
    nm = ThisWorkbook.ActiveSheet.Name
    type_spec = SpecGetType(nm)
    If type_spec = 7 Then
        r = ExportAttribut(nm)
    Else
        If Not (quiet) Then MsgBox ("��������� �� ���� _���� � ���������")
    End If
    r = OutEnded()
    r = print_functime()
End Sub

Private Sub HelpButton_Click()
    strUrl = "https://docs.google.com/document/d/1bedvuS3quC37ivwVWzWDyfZSt_zxFPhGAQAQ38g3Ubo/edit?usp=sharing"
    r = ShellExecute(0, "open", strUrl, 0, 0, 1)
End Sub

Private Sub PosSubposButton_Click()
    ans = MsgBox("�������� ����� �� ������. ������� ��, ���� ������� ��������� �����", vbYesNo)
    If ans = 6 Then
        r = OutPrepare()
        nm = ThisWorkbook.ActiveSheet.Name
        type_spec = SpecGetType(nm)
        If type_spec = 7 Then
            r = ManualPos(nm, 2)
            r = LogWrite(nm, "��", "�������������� �� �������")
        Else
            If Not (quiet) Then MsgBox ("��������� �� ���� _���� � ���������")
        End If
        r = OutEnded()
    End If
    r = print_functime()
End Sub

Private Sub Raskroy_Button_Click()
    r = OutPrepare()
    nm = ThisWorkbook.ActiveSheet.Name
    type_spec = SpecGetType(nm)
    If type_spec <> 3 Or type_spec <> 7 Then
        If InStr(nm, "_") > 0 Then nm = Split(nm, "_")(0)
    End If
    suffix = "_�����"
    r = Spec_Select(nm, suffix)
    r = OutEnded()
End Sub

Private Sub ReloadTXTButton_Click()
    rv = False
    r = OutPrepare()
    rv = SheetAddTxt()
    r = OutEnded()
    r = print_functime()
End Sub

Private Sub UndoPosButton_Click()
    r = OutPrepare()
    nm = ThisWorkbook.ActiveSheet.Name
    type_spec = SpecGetType(nm)
    If type_spec = 7 Then
        r = ManualUndoPos(nm)
        r = LogWrite(nm, "��", "������ �������������")
    Else
        If Not (quiet) Then MsgBox ("��������� �� ���� _���� � ���������")
    End If
    r = OutEnded()
End Sub

Private Sub ClearSheetButton_Click()
    r = OutPrepare()
    Dim type_out: ReDim type_out(7)
    If UserForm2.ob_CB.Value Then type_out(1) = 2
    If UserForm2.obsh_CB.Value Then type_out(2) = 3
    If UserForm2.groub_CB.Value Then type_out(3) = 1
    If UserForm2.vedkzh_CB.Value Then type_out(4) = 5
    If UserForm2.bysybpos_CB.Value Then type_out(5) = 13
    If UserForm2.byved_CB.Value Then type_out(6) = 11
    If UserForm2.bypol_CB.Value Then type_out(7) = 13
    tdel = vbNullString
    If UserForm2.ob_CB.Value Then tdel = tdel + "_��, "
    If UserForm2.obsh_CB.Value Then tdel = tdel + "���������, "
    If UserForm2.groub_CB.Value Then tdel = tdel + "_��, "
    If UserForm2.vedkzh_CB.Value Then tdel = tdel + "_��, "
    If UserForm2.bysybpos_CB.Value Then tdel = tdel + "_���, "
    If UserForm2.byved_CB.Value Then tdel = tdel + "_���, "
    If UserForm2.bypol_CB.Value Then tdel = tdel + "_�����, "
    r = LogWrite("������� ������", vbNullString, 1)
    r = SheetClear(type_out)
    r = LogWrite("������� " & r & " ������", vbNullString, 1)
    r = SheetIndex()
    r = OutEnded()
End Sub

Private Sub ClearAllButton_Click()
    r = OutPrepare()
    Dim type_out: ReDim type_out(1)
    type_out(1) = -1
    r = LogWrite("������� ���� ������", vbNullString, 1)
    r = SheetClear(type_out)
    r = LogWrite("������� " & r & " ������", vbNullString, 1)
    r = SheetIndex()
    r = OutEnded()
End Sub

Private Sub CommandButtonET_Click()
    suffix = "_���"
    r = OutPrepare()
    r = Spec_Select(lastfilespec, suffix)
    r = OutEnded()
End Sub

Private Sub CommandButtonIns_Click()
    r = OutPrepare()
    arr2paste = materialbook_index.Item(lastconstrtype & "_" & lastconstr)
    r = ManualPaste2Sheet(arr2paste)
    r = OutEnded()
    r = print_functime()
End Sub

Private Sub CommandButtonOTD_Click()
    r = OutPrepare()
    suffix = "_���"
    r = Spec_Select(lastfilespec, suffix)
    r = OutEnded()
End Sub

Private Sub CommandButtonPOL_Click()
    r = OutPrepare()
    suffix = "_�����"
    r = Spec_Select(lastfilespec, suffix)
    r = OutEnded()
End Sub

Private Sub CommandButtonSB_Click()
    r = OutPrepare()
    suffix = "_���"
    r = Spec_Select(lastfilespec, suffix)
    r = OutEnded()
End Sub

Private Sub CommandButtonShowS_Click()
    r = SheetShowAddictions()
End Sub

Private Sub CopyGromButton_Click()
    If Left$(ThisWorkbook.path, 2) <> "\\" Then ChDrive Left$(ThisWorkbook.path, 1)
    ChDir ThisWorkbook.path
    TmpPath = ThisWorkbook.path + "\tmpimport\"
    fileToOpen = Application.GetOpenFilename("XLSM Files (*.xlsm),*.xlsm,XLS Files (*.xls),*.xls,CSV Files (*.csv),*.csv,TXT Files (*.txt),*.txt", Title:="����� ������ ��� �������", MultiSelect:=True)
    If IsArray(fileToOpen) Then
        r = OutPrepare()
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(TmpPath) Then MkDir (TmpPath)
        For n = LBound(fileToOpen) To UBound(fileToOpen)
            If ThisWorkbook.FullName <> fileToOpen(n) Then
                FnameInLoop = Right$(fileToOpen(n), Len(fileToOpen(n)) - InStrRev(fileToOpen(n), Application.PathSeparator, , 1))
                fname = TmpPath + Trim$(str(Round(Rnd(100) * 1000, 0)) + FnameInLoop)
                FileCopy fileToOpen(n), TmpPath + FnameInLoop
                Name TmpPath + FnameInLoop As fname
                If SheetImport(fname) Then r = LogWrite(FnameInLoop, vbNullString, "������ ����� ��������")
            End If
        Next n
        If CreateObject("Scripting.FileSystemObject").FolderExists(TmpPath) Then Shell "cmd /c rd /S/Q """ & TmpPath & """"
        r = SheetIndex()
        r = OutEnded()
    End If
    r = print_functime()
End Sub

Private Sub dischargeButton_Click()
    r = OutPrepare()
    nm = ThisWorkbook.ActiveSheet.Name
    type_spec = SpecGetType(nm)
    If type_spec <> 3 Or type_spec <> 7 Then
        If InStr(nm, "_") > 0 Then nm = Split(nm, "_")(0)
    End If
    suffix = "_����"
    r = Spec_Select(nm, suffix)
    r = OutEnded()
End Sub

Private Sub getizdButton_Click()
    r = ManualSpec_NewSubpos()
    r = print_functime()
End Sub

Private Sub ManualBatchButton_Click()
    Dim type_out: ReDim type_out(5)
    If UserForm2.ob_CB.Value Then type_out(1) = "_��"
    If UserForm2.obsh_CB.Value Then type_out(2) = " "
    If UserForm2.groub_CB.Value Then type_out(3) = "_��"
    If UserForm2.vedkzh_CB.Value Then type_out(4) = "_��"
    If UserForm2.bysybpos_CB.Value Then type_out(5) = "_���"
    r = ManualSpec_batch(type_out)
    r = print_functime()
End Sub

Private Sub MergeManualButton_Click()
    r = ManualSpec_Merge()
    r = print_functime()
End Sub

Private Sub ReSortamentButton_Click()
    r = OutPrepare()
    nm = "!System"
    r = LogWrite(nm, SheetExist(nm), "���������� ��������")
    If SheetExist(nm) Then ThisWorkbook.Sheets(nm).Delete
    r = ReadPrSortament()
    Sheets(nm).Visible = False
    r = OutEnded()
    MsgBox ("���������� ��������")
    r = print_functime()
End Sub

Private Sub use_tmp_CB_Click()
    If IsEmpty(materialbook_index) Then
        Set materialbook_index = ReadConstr()
    End If
End Sub

Private Sub UserForm_Initialize()
    MaterialPath = CheckPath(MaterialPatht.Text)
    SortamentPath = CheckPath(SortamentPatht.Text)
    CodePath = CheckPath(CodePatht.Text)
    f = Split(ThisWorkbook.FullName, "\")
    form_caption = Split(f(UBound(f)), ".")(0)
    If UBound(f) >= 2 Then form_caption = f(UBound(f) - 2) + "/" + form_caption
    UserForm2.Caption = form_caption
    If ModeType() Then
        MsgBox ("����� �� �������. ������ �����.")
        is_create = CreateFolders()
        If is_create And ModeType() Then
            MsgBox ("����� �������")
        Else
            MsgBox ("����")
        End If
        Exit Sub
    Else
        If use_tmp_CB.Value Then Set materialbook_index = ReadConstr()
        r = INISet()
        FormRebild
        If check_version Then r = CheckVersion()
        r = set_sheet(vbNullString)
        If mem_option Then r = OptionSheetSet(lastfilespec)
    End If
End Sub

Function CheckPath(ByVal path) As String
   If InStr(1, path, "\") = 1 Then
      CheckPath = ThisWorkbook.path + path
   Else
      CheckPath = path
   End If
End Function

Private Sub CodePatht_Change()
    CodePath = CheckPath(CodePatht.Text)
End Sub

Private Sub CommandButtonAdd2Man_Click()
    r = OutPrepare()
    r = ManualPasteIzd2Sheet(name_izd.Item(lastfileadd))
    r = OutEnded()
    r = print_functime()
End Sub

Private Sub CommandButtonAS_Click()
    r = OutPrepare()
    r = Spec_Select(lastfilespec, vbNullString)
    r = OutEnded()
End Sub

Private Sub CommandButtonExport_Click()
    r = OutPrepare()
    nm = ThisWorkbook.ActiveSheet.Name
    r = ExportSheet(nm)
    r = OutEnded()
    r = print_functime()
End Sub

Private Sub CommandButtonGR_Click()
    r = OutPrepare()
    suffix = "_��"
    r = Spec_Select(lastfilespec, suffix)
    r = OutEnded()
End Sub

Private Sub CommandButtonKM_Click()
    r = OutPrepare()
    suffix = "_��"
    r = Spec_Select(lastfilespec, suffix)
    r = OutEnded()
End Sub

Private Sub CommandButtonKZH_Click()
    r = OutPrepare()
    suffix = "_��"
    r = Spec_Select(lastfilespec, suffix)
    r = OutEnded()
End Sub

Private Sub CommandButtonManPrep_Click()
    r = OutPrepare()
    nm = Application.ThisWorkbook.ActiveSheet.Name
    r = ManualCheck(nm)
    r = OutEnded()
    r = print_functime()
End Sub

Private Sub CommandButtonOBSH_Click()
    r = OutPrepare()
    suffix = "_��"
    r = Spec_Select(lastfilespec, suffix)
    r = OutEnded()
End Sub

Private Sub CommandButtonUpdate_Click()
    r = OutPrepare()
    sNameSheet = ThisWorkbook.ActiveSheet.Name
    r = Spec_Select(sNameSheet, vbNullString)
    r = OutEnded()
End Sub

Private Sub FormatButton_Click()
    r = OutPrepare()
    sNameSheet = ThisWorkbook.ActiveSheet.Name
    r = FormatTable(sNameSheet)
    r = OutEnded()
    r = print_functime()
End Sub

Sub FormRebild()
    r = INISet()
    calc_ver.Caption = macro_version
    com_ver.Caption = common_version
    form_ver.Caption = form_version
    form1_ver.Caption = form1_version
    symb_diam = ChrW$(8960)
    remat
End Sub

Function set_sheet(ByVal sheetn As String) As Boolean
    If Len(sheetn) = 0 Then sheetn = Trim$(ThisWorkbook.ActiveSheet.Name)
    flag_set = False
    For i = 0 To UserForm2.ListBoxFileSpec.ListCount - 1
        If UserForm2.ListBoxFileSpec.Column(0, i) = sheetn Then
            i = UserForm2.ListBoxFileSpec.ListCount - 1
            flag_set = True
        End If
    Next i
    sheetn_mun = sheetn
    If InStr(sheetn, "_") > 0 Then sheetn_mun = Split(sheetn, "_")(0)
    If flag_set = False Then
        sheetn_mun_ = sheetn_mun + "_����"
        For i = 0 To UserForm2.ListBoxFileSpec.ListCount - 1
            If UserForm2.ListBoxFileSpec.Column(0, i) = sheetn_mun_ Then
                i = UserForm2.ListBoxFileSpec.ListCount - 1
                sheetn = sheetn_mun_
                flag_set = True
            End If
        Next i
    End If
    If flag_set = False Then
        For i = 0 To UserForm2.ListBoxFileSpec.ListCount - 1
            If InStr(UserForm2.ListBoxFileSpec.Column(0, i), sheetn_mun) > 0 Then
                i = UserForm2.ListBoxFileSpec.ListCount - 1
                sheetn = sheetn_mun
                flag_set = True
            End If
        Next i
    End If
    set_sheet = False
    If flag_set = True Then
        not_same_list = Not (lastfilespec = ListBoxFileSpec.Value)
        not_same_val = Not (sheetn = ListBoxFileSpec.Value)
        not_same_sheetn = Not (sheetn = lastfilespec)
        is_emty_last = (IsEmpty(lastfilespec) Or Len(lastfilespec) = 0)
        If not_same_list Or not_same_sheetn Or not_same_val Or is_emty_last Then
            'ListBoxFileSpec.Value = sheetn
            lastfilespec = sheetn
        End If
    Else
        'ListBoxFileSpec.Value = lastfilespec
    End If
End Function

Private Sub HideButton_Click()
    r = SheetHideAll()
End Sub

Private Sub ListBoxFileSpec_Click()
    lastfilespec_new = ListBoxFileSpec.Value
    If lastfilespec <> lastfilespec_new And SpecGetType(lastfilespec_new) > 0 And Len(lastfilespec_new) > 0 And mem_option Then
        r = OptionSheetSet(lastfilespec_new)
    End If
    If Len(lastfilespec_new) > 0 Then lastfilespec = lastfilespec_new
End Sub

Private Sub ListBoxName_Click()
    lastfileadd = ListBoxName.Value
End Sub

Private Sub ListBoxTypeIns_Click()
    lastconstrtype = ListBoxTypeIns.Value
    r = ReList(ListBoxNaenIns, materialbook_index.Item(lastconstrtype & "constr"))
End Sub

Private Sub ListBoxNaenIns_Click()
    lastconstr = ListBoxNaenIns.Value
End Sub

Private Sub ListButton_Click()
    r = OutPrepare()
    r = SheetIndex()
    r = OutEnded()
End Sub

Private Sub MaterialPatht_Change()
    MaterialPath = CheckPath(MaterialPatht.Text)
End Sub

Private Sub MultiPage1_Change()
    FormRebild
End Sub

Function ReList(ByRef objListBox As Variant, ByRef arr As Variant) As _
    Boolean
        objListBox.Clear
    If Not IsEmpty(arr) Then
        For i = 1 To UBound(arr)
            objListBox.AddItem (arr(i))
        Next i
        ReList = False
    Else
        ReList = True
    End If
End Function

Function update_list_spec() As Variant
    Dim listspec: ReDim listspec(1): n_man = 0
    listsheet = GetListOfSheet(ThisWorkbook)
    For Each sheet In listsheet
        type_spec = SpecGetType(sheet)
        If type_spec = 7 Then
            n_man = n_man + 1
            ReDim Preserve listspec(n_man)
            listspec(n_man) = sheet
        End If
    Next
    listFile = GetListFile("*.txt")
    If Not IsEmpty(listFile) Then
        Dim add_spec(): ReDim add_spec(UBound(listFile, 1)): n_add = 0
        For i = 1 To UBound(listFile, 1)
            flag_add = 1
            tf_name = listFile(i, 1)
            If tf_name = "����" Then flag_add = 0
            If tf_name = "�������_���������" Then flag_add = 0
            If tf_name = "����_�����" Then flag_add = 0
            If InStr(tf_name, "_����") > 0 Then flag_add = 0
            'If InStr(tf_name, "_���") > 0 Then flag_add = 0
            If InStr(tf_name, "_���") > 0 Then flag_add = 0
            If SheetExist(tf_name + "_����") Then
                flag_add = 0
            End If
            If flag_add = 1 Then
                type_spec = SpecGetType(tf_name)
                If type_spec <> 7 And type_spec <> 2 And type_spec <> 3 Then flag_add = 0
                If type_spec = 22 Then
                    n_add = n_add + 1
                    add_spec(n_add) = Split(tf_name, "_")(0)
                End If
            End If
            If flag_add = 1 Then
                n_man = n_man + 1
                ReDim Preserve listspec(n_man)
                listspec(n_man) = listFile(i, 1)
            End If
        Next i
    End If
    If n_add > 0 Then
        ReDim Preserve add_spec(n_add)
        listspec = ArrayCombine(listspec, add_spec)
    End If
    listspec = ArrayUniqValColumn(listspec, 1)
    If Not IsEmpty(listspec) Then r = ReList(ListBoxFileSpec, listspec)
    update_list_spec = listspec
End Function

Function update_list_add() As Variant
    Dim listadd: ReDim listadd(1): n_add = 0
    Set name_izd = CreateObject("Scripting.Dictionary")
    Dim adress_array: ReDim adress_array(4)
    For Each objWh In ThisWorkbook.Worksheets
        If SpecGetType(objWh.Name) = 15 Then
            Set spec_izd_sheet = Application.ThisWorkbook.Sheets(objWh.Name)
            spec_izd_size = SheetGetSize(spec_izd_sheet)
            n_izd_row = spec_izd_size(1)
            spec_izd = spec_izd_sheet.Range(spec_izd_sheet.Cells(1, 1), spec_izd_sheet.Cells(n_izd_row, max_col_man))
            For i = 3 To n_izd_row
                subpos = spec_izd(i, col_man_subpos)
                pos = spec_izd(i, col_man_pos)
                If name_izd.Exists(subpos) = False And subpos = pos And Len(subpos) > 1 Then
                    For jj = 1 To 4
                        adress_array(jj) = "=" + objWh.Name + "!" + spec_izd_sheet.Cells(i, jj).Address()
                    Next jj
                    name_izd.Item(subpos) = adress_array
                End If
            Next i
        End If
    Next objWh
    listadd = ArraySort(name_izd.keys())
    If Not IsEmpty(listadd) Then r = ReList(ListBoxName, listadd)
    update_list_add = listadd
End Function

Sub remat()
    If use_tmp_CB.Value Then
        lastconstrtype = materialbook_index.Item("sheet_list")(1)
        r = ReList(ListBoxTypeIns, materialbook_index.Item("sheet_list"))
        lastconstr = materialbook_index.Item(lastconstrtype & "constr")(1)
        r = ReList(ListBoxNaenIns, materialbook_index.Item(lastconstrtype & "constr"))
    End If
    listadd = update_list_add()
    listspec = update_list_spec()
    If Not IsEmpty(listadd) Then
        If lastfileadd = vbNullString Then lastfileadd = listadd(1)
    Else
        lastfileadd = vbNullString
    End If
    
    If Len(lastfilespec) > 0 And calc.ArrayHasElement(listspec, lastfilespec) Then
        ListBoxFileSpec.Value = lastfilespec
    Else
        r = set_sheet(vbNullString)
    End If
End Sub


Private Sub SaveCodeButton_Click()
    r = ExportAllMod()
End Sub

Private Sub ShowButton_Click()
    r = SheetShowAll()
End Sub

Private Sub SortamentPatht_Change()
    SortamentPath = CheckPath(SortamentPatht.Text)
End Sub

'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'If CloseMode = 0 Then
        'FormRebild
        'Cancel = 1
    'End If
'End Sub

Function ReadConstr()
    If IsModEx("calc") Then
        r = OutPrepare()
        Set materialbook = GetObject(MaterialPath & "constr.xlsm")
        Set constr_index = CreateObject("Scripting.Dictionary")
        constr_index.comparemode = 1
'        listsheet = GetListOfSheet(materialbook)
        constr_index.Item("sheet_list") = listsheet
        Dim constr_list: ReDim constr_list(1)
        Dim tarr: ReDim tarr(1, max_col_man)
        For Each sheet_name In listsheet
            ReDim constr_list(1)
            Set sheet = materialbook.Sheets(sheet_name)
            n_row = SheetGetSize(sheet)(1)
            flag = 0: istart = 0
            For i = 1 To n_row
                If InStr(sheet.Cells(i, col_man_pos), "#") > 0 Or InStr(sheet.Cells(i, col_man_subpos), "#") > 0 Or i = n_row Then
                    If istart = 0 Then
                        istart = i + 1
                    Else
                        If i = n_row Then
                            iend = i
                        Else
                            iend = i - 1
                        End If
                        flag = 1
                    End If
                End If
                If flag Then
                    ReDim tarr(iend - istart + 1, max_col_man)
                    For irow = istart To iend
                        For icol = 1 To max_col_man
                            If sheet.Cells(irow, icol).HasFormula Then
                                valc = sheet.Cells(irow, icol).FormulaR1C1
                            Else
                                valc = sheet.Cells(irow, icol).Value
                            End If
                            If Not IsEmpty(valc) And Not IsError(valc) Then
                                tarr(irow - istart + 1, icol) = valc
                            Else
                                tarr(irow - istart + 1, icol) = vbNullString
                            End If
                        Next
                    Next
                    c_size = UBound(constr_list)
                    constr_name = sheet.Cells(istart - 1, col_man_naen)
                    constr_index.Item(sheet_name & "_" & constr_name) = tarr
                    constr_list(c_size) = constr_name
                    ReDim Preserve constr_list(c_size + 1)
                    istart = iend + 2
                    flag = 0
                End If
            Next i
            constr_index.Item(sheet_name & "constr") = constr_list
        Next
        Set ReadConstr = constr_index
        materialbook.Close
        r = OutEnded()
    End If
End Function


