Attribute VB_Name = "update"
Option Compare Text
Option Base 1

Public Const update_version As String = "4.03"

Public code_path As String
Public sortament_path As String
Public material_path As String

Function CheckVersion()
    r = set_path()
    If Not check_version Then check_version = read_ini_param("check_version")
    If Not Debug_mode Then Debug_mode = read_ini_param("Debug_mode")
    If InStr(code_path, "material-specification") <= 0 Then Debug_mode = False
    If Ping() And check_version Then
        If Download_Code() Then
            change_log = DownloadMod("changelog" & ".txt")
            common_local = 0
            calc_local = 0
            form1_local = 0
            form_local = 0
            update_local = 0
            
            msg_upd = vbNullString
            common_git = DownloadMod("common" & ".bas")
            If IsModEx("common") Then common_local = ConvTxt2Ver(common_version)
            If common_git > common_local And Not IsEmpty(common_local) And Not IsEmpty(common_git) Then msg_upd = msg_upd & "Загружена новая версия модуля common -" & CStr(common_git / 100) & vbNewLine

            calc_git = DownloadMod("calc" & ".bas")
            If IsModEx("calc") Then calc_local = ConvTxt2Ver(macro_version)
            If calc_git > calc_local And Not IsEmpty(calc_local) And Not IsEmpty(calc_git) Then msg_upd = msg_upd & "Загружена новая версия модуля calc -" & CStr(calc_git / 100) & vbNewLine
            
            r = DownloadMod("UserForm2.frx")
            form_git = DownloadMod("UserForm2" & ".frm")
            If IsModEx("UserForm2") Then form_local = ConvTxt2Ver(UserForm2.form_ver.Caption)
            If form_git > form_local And Not IsEmpty(form_local) And Not IsEmpty(form_git) Then msg_upd = msg_upd & "Загружена новая версия формы -" & CStr(form_git / 100) & vbNewLine
            
            r = DownloadMod("UserForm1.frx")
            form1_git = DownloadMod("UserForm1" & ".frm")
            If IsModEx("UserForm2") And IsModEx("UserForm1") Then form1_local = ConvTxt2Ver(UserForm2.form1_ver.Caption)
            If form1_git > form1_local And Not IsEmpty(form1_local) And Not IsEmpty(form1_git) Then msg_upd = msg_upd & "Загружена новая версия формы 2 -" & CStr(form_git / 100) & vbNewLine
    
            update_git = DownloadMod("update" & ".bas")
            update_local = ConvTxt2Ver(update_version)
            If update_git > update_local And Not IsEmpty(update_local) And Not IsEmpty(update_git) Then msg_upd = msg_upd & "Загружена новая версия модуля update -" & CStr(update_git / 100) & vbNewLine
            If Len(msg_upd) > 1 Then
                r = ExportAllMod()
                msg_upd = msg_upd & vbNewLine & change_log & vbNewLine & "Открыть справку с инструкцией?"
                intMessage = MsgBox(msg_upd, vbYesNo, "Доступно обновление")
                Set objShell = CreateObject("Wscript.Shell")
                If intMessage = vbYes Then objShell.Run ("https://docs.google.com/document/d/1bedvuS3quC37ivwVWzWDyfZSt_zxFPhGAQAQ38g3Ubo/edit#bookmark=id.5awyi0cjwmrs")
            End If
        End If
    End If
End Function

Function Ping() As Boolean
    Dim oPingResult As Variant
    For Each oPingResult In GetObject("winmgmts://./root/cimv2").ExecQuery _
        ("SELECT * FROM Win32_PingStatus WHERE Address = '" & "github.com" & "'")
        If IsObject(oPingResult) Then
            If oPingResult.StatusCode = 0 Then
                Ping = True
                Exit Function
            End If
        End If
    Next
End Function

Function set_path() As Boolean
    If IsModEx("UserForm2") Then
        code_path = UserForm2.CodePath
        sortament_path = UserForm2.SortamentPath
        material_path = UserForm2.MaterialPath
    Else
        code_path = ThisWorkbook.path & "\data\code\"
        sortament_path = ThisWorkbook.path & "\data\sort\"
        material_path = ThisWorkbook.path & "\data\mat\"
    End If
End Function

Function read_ini_param(ByVal paramname As String) As Boolean
    sIniFile = code_path & "setting.ini"
    If Not CBool(Len(Dir$(sIniFile))) Then r = Download_Settings()
    If CBool(Len(Dir$(sIniFile))) Then
        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        Set ts = FSO.OpenTextFile(sIniFile, 1, True): txt$ = ts.ReadAll: ts.Close
        Set ts = Nothing: Set FSO = Nothing
        txt = Trim$(txt): Err.Clear
        If InStr(txt, paramname) Then
            k = Split(txt, paramname)(1)
            k = Split(k, vbNewLine)(0)
            If InStr(k, "#") > 0 Then k = Split(k, "#")(0)
            k = Trim(LCase(k))
            If InStr(k, "true") > 0 Then
                read_ini_param = True
            Else
                read_ini_param = False
            End If
        End If
    Else
        read_ini_param = False
    End If
End Function

Function DownloadMod(ByVal namemod As String) As Variant
    gitdir = code_path & "from_git"
    If InStr(namemod, ".bas") > 0 Or InStr(namemod, ".txt") > 0 Or InStr(namemod, ".frm") > 0 Then
        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        Set ts = FSO.OpenTextFile(gitdir & "/" & namemod, 1, True): txt$ = ts.ReadAll: ts.Close
        Set ts = Nothing: Set FSO = Nothing
        txt = Trim$(txt): Err.Clear
    End If
    If InStr(namemod, ".bas") > 0 Or InStr(namemod, ".frm") > 0 Then
        seach_txt = "version As String ="
        For Each tRows In Split(txt, vbNewLine)
            If InStr(tRows, seach_txt) Then
                k = Mid$(tRows, InStr(tRows, seach_txt) + Len(seach_txt), 6)
                ver = Split(k, """")(1)
                Version_file = ConvTxt2Ver(ver)
            End If
        Next
    Else
        If InStr(namemod, ".txt") > 0 Then
            Version_file = txt
        Else
            Version_file = 0
        End If
    End If
    DownloadMod = Version_file
End Function

Function ExportAllMod() As Boolean
    pathtmp = "old\"
    If Debug_mode Then pathtmp = vbNullString
    If IsModEx("UserForm2") Then r = ExportMod("UserForm2", pathtmp, UserForm2.form_ver.Caption)
    If IsModEx("UserForm1") And IsModEx("UserForm2") Then r = ExportMod("UserForm1", pathtmp, UserForm2.form1_ver.Caption)
    If IsModEx("calc") Then r = ExportMod("calc", pathtmp, macro_version)
    If IsModEx("common") Then r = ExportMod("common", pathtmp, common_version)
    If IsModEx("update") Then r = ExportMod("update", pathtmp, update_version)
End Function

Function NameAllMod() As Boolean
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    r = CheckName("UserForm2")
    r = CheckName("UserForm1")
    r = CheckName("calc")
    r = CheckName("common")
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function

Function ModeType() As Boolean
    r = set_path()
    ismat = CreateObject("Scripting.FileSystemObject").FolderExists(material_path)
    issort = CreateObject("Scripting.FileSystemObject").FolderExists(sortament_path)
    iscode = CreateObject("Scripting.FileSystemObject").FolderExists(code_path)
    If ismat And issort And iscode Then
        read_only_mode = False
    Else
        read_only_mode = True
    End If
    ModeType = read_only_mode
End Function

Function Download_Sortament() As Boolean
    r = set_path()
    myURL = "https://raw.githubusercontent.com/kuvbur/material-specification/master/data/sort.zip"
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
    myURL = WinHttpReq.responseBody
    If WinHttpReq.Status = 200 Then
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\data") Then
            MkDir (ThisWorkbook.path & "\data")
        End If
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(sortament_path) Then
            MkDir (sortament_path)
        End If
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile ThisWorkbook.path & "\data\sort.zip", 2
        oStream.Close
        Set oApp = CreateObject("Shell.Application")
        For Each it In oApp.Namespace(ThisWorkbook.path & "\data\sort.zip").items: DoEvents: DoEvents: Next
        oApp.Namespace(ThisWorkbook.path & "\data").CopyHere oApp.Namespace(ThisWorkbook.path & "\data\sort.zip").items
        Download_Sortament = True
    Else
        Download_Sortament = False
    End If
End Function

Function Download_Code() As Boolean
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\data") Then
        On Error Resume Next
        MkDir (ThisWorkbook.path & "\data")
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(code_path) Then
        On Error Resume Next
        MkDir (code_path)
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(code_path & "\from_git") Then
        On Error Resume Next
        MkDir (code_path & "\from_git")
    End If
    code_filepath = ThisWorkbook.path & "\data\code"
    myURL = "https://raw.githubusercontent.com/kuvbur/material-specification/master/data/code/from_git.zip"
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
    myURL = WinHttpReq.responseBody
    If WinHttpReq.Status = 200 Then
        r = Delete_file(code_filepath & "\from_git.zip")
        r = Delete_file(code_filepath & "\from_git" & "\calc.bas")
        r = Delete_file(code_filepath & "\from_git" & "\update.bas")
        r = Delete_file(code_filepath & "\from_git" & "\common.bas")
        r = Delete_file(code_filepath & "\from_git" & "\changelog.txt")
        r = Delete_file(code_filepath & "\from_git" & "\UserForm1.frm")
        r = Delete_file(code_filepath & "\from_git" & "\UserForm1.frx")
        r = Delete_file(code_filepath & "\from_git" & "\UserForm1.bas")
        r = Delete_file(code_filepath & "\from_git" & "\UserForm2.frm")
        r = Delete_file(code_filepath & "\from_git" & "\UserForm2.frx")
        r = Delete_file(code_filepath & "\from_git" & "\UserForm2.bas")
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile code_filepath & "\from_git.zip", 2
        oStream.Close
        Set oApp = CreateObject("Shell.Application")
        For Each it In oApp.Namespace(code_filepath & "\from_git.zip").items: DoEvents: DoEvents: Next
        oApp.Namespace(code_filepath & "\from_git").CopyHere oApp.Namespace(code_filepath & "\from_git.zip").items
        Download_Code = True
    Else
        Download_Code = False
    End If
End Function

Function Delete_file(ByVal filepath As String) As Boolean
    On Error Resume Next
    If Dir(filepath) <> "" Then Kill filepath
End Function

Function Download_Settings() As Boolean
    myURL = "https://raw.githubusercontent.com/kuvbur/material-specification/master/data/code/setting.ini"
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
    myURL = WinHttpReq.responseBody
    If WinHttpReq.Status = 200 Then
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\data") Then
            MkDir (ThisWorkbook.path & "\data")
        End If
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(code_path) Then
            MkDir (code_path)
        End If
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile ThisWorkbook.path & "/data/code/setting.ini", 2
        oStream.Close
        Download_Settings = True
    Else
        Download_Settings = False
    End If
End Function

Function CreateFolders() As Boolean
    r = set_path()
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\import") Then
        On Error Resume Next
        MkDir (ThisWorkbook.path & "\import")
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\list") Then
        On Error Resume Next
        MkDir (ThisWorkbook.path & "\list")
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\data") Then
        On Error Resume Next
        MkDir (ThisWorkbook.path & "\data")
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(code_path) Then
        On Error Resume Next
        MkDir (code_path)
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(sortament_path) Then
        On Error Resume Next
        MkDir (sortament_path)
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(material_path) Then
        On Error Resume Next
        MkDir (material_path)
    End If
    If Download_Sortament() And Download_Settings() Then
        CreateFolders = True
    Else
        CreateFolders = False
    End If
End Function

Function ExportMod(ByVal namemod As String, Optional ByVal pathtmp As String = vbNullString, Optional ByVal in_ver As String = vbNullString) As Boolean
        tdate = vbNullString
        If pathtmp = "old\" Then
            tdate = "_" & CStr(in_ver) & "_" & Replace(Right$(str(DatePart("yyyy", Now)), 2) & str(DatePart("m", Now)) & str(DatePart("d", Now)), " ", vbNullString)
            If Not CreateObject("Scripting.FileSystemObject").FolderExists(code_path & pathtmp) Then
                MkDir (code_path & pathtmp)
            End If
        End If
        On Error Resume Next
        suffix = ".bas"
        If InStr(namemod, "form") > 0 Then suffix = ".frm"
        path = code_path & pathtmp & namemod & tdate & suffix
        ThisWorkbook.VBProject.VBComponents.Item(namemod).Export path
End Function

Function IsModFileEx(ByVal namemod As String, Optional ByVal pathtmp As String = vbNullString) As Boolean
    On Error Resume Next
    IsModFileEx = CBool(Len(Dir$(code_path & pathtmp & namemod & ".bas")))
End Function

Function IsModEx(ByVal namemod As String) As Boolean
    On Error Resume Next
    IsModEx = CBool(Len(ThisWorkbook.VBProject.VBComponents(namemod).Name))
End Function

Function ConvTxt2Ver(ByVal x As Variant) As Variant
    x = CStr(x)
    x = Replace(x, ".", vbNullString)
    x = Replace(x, ",", vbNullString)
    If Len(x) < 3 Then
        n_zero = 3 - Len(x)
        For i = 1 To n_zero
            x = x + "0"
        Next i
    End If
    If IsNumeric(x) Then
        out = CInt(x)
    Else
        out = Empty
    End If
    ConvTxt2Ver = out
End Function


