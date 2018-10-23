Attribute VB_Name = "update"
Option Compare Text
Option Base 1

#If Win64 Then
    Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
            (ByVal pCaller As LongLong, ByVal szURL As String, ByVal szFileName As String, _
             ByVal dwReserved As LongLong, ByVal lpfnCB As LongLong) As LongLong
    
#Else
    
    #If VBA7 Then
        Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
            (ByVal pCaller As LongPtr, ByVal szURL As String, ByVal szFileName As String, _
                ByVal dwReserved As Long, ByVal lpfnCB As LongPtr) As LongPtr
    #Else
        Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                                        (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, _
                                        ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    #End If
#End If

Function ExportAllMod() As Boolean
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.CodePath) Then
        MkDir (UserForm2.CodePath)
    End If
    r = ExportMod("UserForm2")
    r = ExportMod("calc")
    r = ExportMod("common")
End Function

Function ImportAllMod() As Boolean
    Ret_type = MsgBox("Будет произведена замена модулей. Продолжить?", vbYesNoCancel + vbQuestion, "Обновление")
    If Ret_type = 6 Then
        Wh_type = MsgBox("Читать из папки (Да) или из Гитхаба (Нет)", vbYesNo + vbQuestion, "Источник обновления")
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Select Case Wh_type
            Case 7
                pathtmp = "git\"
                If Not CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.CodePath & pathtmp) Then
                    MkDir (UserForm2.CodePath & pathtmp)
                End If
            Case 6
                pathtmp = ""
        End Select
        r = ImportMod("UserForm2", pathtmp)
        r = ImportMod("calc", pathtmp)
        r = ImportMod("common", pathtmp)
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        r = NameAllMod()
        Application.VBE.CommandBars.FindControl(id:=228).Execute
    End If
End Function

Function GetMacroGit(ByVal namemod As String, ByVal pathtmp As String) As Variant
    filename_url = "https://raw.githubusercontent.com/kuvbur/material-specification/master/data/code/" & namemod & ".bas"
    filename = UserForm2.CodePath & pathtmp & namemod & ".bas"
    r = (URLDownloadToFile(0, filename_url, filename, 0, 0) = 0)
    If r = False Then MsgBox ("Не удалось скачать модуль" & namemod)
    GetMacroGit = r
End Function

Function NameAllMod() As Boolean
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    r = CheckName("UserForm2")
    r = CheckName("calc")
    r = CheckName("common")
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function

Function ExportMod(ByVal namemod As String, Optional ByVal pathtmp As String = "") As Boolean
    out_ver = CheckOutVersion(namemod, pathtmp)
    in_ver = CheckInVersion(namemod)
    If in_ver >= out_ver And in_ver > 0 Then
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.CodePath & pathtmp) Then
            MkDir (UserForm2.CodePath & pathtmp)
        End If
        tdate = ""
        If pathtmp = "old\" Then tdate = "_" & CStr(in_ver) & "_" & Replace(Right(Str(DatePart("yyyy", Now)), 2) & Str(DatePart("m", Now)) & Str(DatePart("d", Now)), " ", "")
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Item(namemod).Export UserForm2.CodePath & pathtmp & namemod & tdate & ".bas"
    Else
        tdate = Replace(Right(Str(DatePart("yyyy", Now)), 2) & Str(DatePart("m", Now)) & Str(DatePart("d", Now)), " ", "")
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Item(namemod).Export UserForm2.CodePath & pathtmp & CStr(in_ver) & "_" & namemod & tdate & ".bas"
    End If
End Function

Function ImportMod(ByVal namemod As String, Optional ByVal pathtmp As String = "") As Boolean
    If GetMacroGit(namemod, pathtmp) = False Then Exit Function
    out_ver = CheckOutVersion(namemod, pathtmp)
    in_ver = CheckInVersion(namemod)
    If in_ver < out_ver And in_ver > 0 And out_ver > 0 Then
        r = ExportMod(namemod, "old\")
        MsgBox ("Модуль " & namemod & " обновлён до версии " & CStr(out_ver))
        r = LoadMod(namemod, pathtmp)
        r = DelMod(namemod)
    End If
End Function

Function CheckName(ByVal namemod As String)
    If IsModEx(namemod) Then Exit Function
    If IsModEx(namemod & "1") Then
        ThisWorkbook.VBProject.VBComponents(namemod & "1").Name = namemod
    End If
End Function

Function DelMod(ByVal namemod As String) As Boolean
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(namemod)
End Function

Function LoadMod(ByVal namemod As String, Optional ByVal pathtmp As String = "") As Boolean
    ThisWorkbook.VBProject.VBComponents.Import UserForm2.CodePath & pathtmp & namemod & ".bas"
End Function

Function IsModFileEx(ByVal namemod As String, Optional ByVal pathtmp As String = "") As Boolean
    On Error Resume Next
    IsModFileEx = CBool(Len(Dir$(UserForm2.CodePath & pathtmp & namemod & ".bas")))
End Function

Function IsModEx(ByVal namemod As String) As Boolean
    On Error Resume Next
    IsModEx = CBool(Len(ThisWorkbook.VBProject.VBComponents(namemod).Name))
End Function

Function CheckOutVersion(ByVal namemod As String, Optional ByVal pathtmp As String = "") As Double
    If Not IsModFileEx(namemod, pathtmp) And pathtmp <> "old\" Then
        MsgBox ("Файл модуля " & pathtmp & namemod & " не найден")
        CheckOutVersion = -1
        Exit Function
    End If
    On Error Resume Next
    Open UserForm2.CodePath & pathtmp & namemod & ".bas" For Input As #1
    s = Input(LOF(1), 1)
    Close #1
    f_str = "version As String ="
    n = InStr(s, f_str)
    ver = Right(Left(s, n + Len(f_str) + 4), 3)
    num_ver = ConvTxtNum(ver)
    If Not IsNumeric(num_ver) Then num_ver = 0
    CheckOutVersion = num_ver
End Function

Function CheckInVersion(ByVal namemod As String) As Double
    num_ver = 0
    Select Case namemod
        Case "calc"
            num_ver = ConvTxtNum(macro_version)
        Case "common"
            num_ver = ConvTxtNum(common_version)
        Case "UserForm2"
            num_ver = ConvTxtNum(UserForm2.form_ver.Caption)
    End Select
    If Not IsNumeric(num_ver) Then num_ver = 0
    CheckInVersion = num_ver
End Function

Function ConvTxtNum(ByVal x As Variant) As Variant
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
    ConvTxtNum = out
End Function
