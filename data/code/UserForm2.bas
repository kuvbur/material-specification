VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   " "
   ClientHeight    =   11925
   ClientLeft      =   45
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
Public CodePath, MaterialPath, SortamentPath As String
Public lastsheet, lastconstrtype, lastconstr, lastfile, lastfilespec, lastfileadd, materialbook_index As Variant

Private Sub CommandButtonIns_Click()
    arr2paste = materialbook_index.Item(lastconstrtype & "_" & lastconstr)
    r = Paste2Sheet(arr2paste)
End Sub

Private Sub CommandButtonOTD_Click()
    suffix = "_вед"
    r = Spec_Select(lastfilespec, suffix)
End Sub

Private Sub CommandButtonPOL_Click()
    suffix = "_экспл"
    r = Spec_Select(lastfilespec, suffix)
End Sub

Private Sub CommandButtonSB_Click()
    suffix = "_грс"
    r = Spec_Select(lastfilespec, suffix)
End Sub

Private Sub CommandButtonShowS_Click()
    r = show_s()
End Sub

Private Sub UserForm_Initialize()
    MaterialPath = CheckPath(MaterialPatht.Text)
    SortamentPath = CheckPath(SortamentPatht.Text)
    CodePath = CheckPath(CodePatht.Text)
    Set materialbook_index = ReadConstr()
    FormRebild
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
    r = ManualAdd(lastfileadd)
End Sub

Private Sub CommandButtonAS_Click()
    sNameSheet = lastfilespec
    r = Spec_Select(lastfilespec, "")
End Sub

Private Sub CommandButtonExport_Click()
    r = SheetExport()
End Sub

Private Sub CommandButtonGR_Click()
    suffix = "_гр"
    r = Spec_Select(lastfilespec, suffix)
End Sub

Private Sub CommandButtonKM_Click()
    suffix = "_км"
    r = Spec_Select(lastfilespec, suffix)
End Sub

Private Sub CommandButtonKZH_Click()
    suffix = "_кж"
    r = Spec_Select(lastfilespec, suffix)
End Sub

Private Sub CommandButtonManPrep_Click()
    r = ManualCheck()
End Sub

Private Sub CommandButtonOBSH_Click()
    suffix = "_об"
    r = Spec_Select(lastfilespec, suffix)
End Sub

Private Sub CommandButtonUpdate_Click()
    sNameSheet = ActiveWorkbook.ActiveSheet.Name
    r = Spec_Select(sNameSheet, "")
End Sub

Private Sub FormatButton_Click()
    r = FormatTable()
End Sub

Sub FormRebild()
    calc_ver.Caption = macro_version
    com_ver.Caption = common_version
    man_ver.Caption = manual_version
    surf_ver.Caption = surf_version
    form_ver.Caption = "2.4"
    symb_diam = ChrW(8960)
    remat
End Sub

Private Sub HideButton_Click()
    r = SheetHideAll()
End Sub

Private Sub ListBoxFileSpec_Click()
    lastfilespec = ListBoxFileSpec.Value
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
    r = SheetIndex()
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
    For i = 1 To UBound(arr)
        objListBox.AddItem (arr(i))
    Next i
    ReList = True
End Function

Sub remat()
    lastconstrtype = materialbook_index.Item("sheet_list")(1)
    r = ReList(ListBoxTypeIns, materialbook_index.Item("sheet_list"))
    
    lastconstr = materialbook_index.Item(lastconstrtype & "constr")(1)
    r = ReList(ListBoxNaenIns, materialbook_index.Item(lastconstrtype & "constr"))
    
    listFile = GetListFile("*.txt")
    Dim listspec: ReDim listspec(1): n_man = 0
    Dim listadd: ReDim listadd(1): n_add = 0
    listsheet = GetListOfSheet(ThisWorkbook)
    For Each sheet In listsheet
        type_spec = SpecGetType(sheet)
        If type_spec = 7 Then
            n_man = n_man + 1
            ReDim Preserve listspec(n_man)
            listspec(n_man) = sheet
        End If
        If type_spec = 9 Then
            n_add = n_add + 1
            ReDim Preserve listadd(n_add)
            listadd(n_add) = sheet
        End If
    Next
    For i = 1 To UBound(listFile, 1)
        If ((listFile(i, 1) <> "Полы") * (listFile(i, 1) <> _
            "Отметки_перемычек") * (listFile(i, 1) <> "Типы_полов")) Then
            n_man = n_man + 1
            ReDim Preserve listspec(n_man)
            listspec(n_man) = listFile(i, 1)
            n_add = n_add + 1
            ReDim Preserve listadd(n_add)
            listadd(n_add) = listFile(i, 1)
        End If
    Next i
    r = ReList(ListBoxFileSpec, listspec)
    r = ReList(ListBoxName, listadd)
    lastfile = listspec(1)
    lastfilespec = listspec(1)
    lastfileadd = listadd(1)
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
    Set materialbook = GetObject(MaterialPath & "constr.xlsm")
    Set constr_index = CreateObject("Scripting.Dictionary")
    constr_index.comparemode = 1
    listsheet = GetListOfSheet(materialbook)
    constr_index.Item("sheet_list") = listsheet
    Dim constr_list: ReDim constr_list(1)
    Dim tarr: ReDim tarr(1, max_col_man)
    For Each sheet_name In listsheet
        ReDim constr_list(1)
        Set sheet = materialbook.Sheets(sheet_name)
        n_row = GetSizeSheet(sheet)(1)
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
                            tarr(irow - istart + 1, icol) = ""
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
    End If
End Function

