Attribute VB_Name = "update"
Option Compare Text
Option Base 1

Public Const update_version As String = "3.41"
Function CheckVersion()
    If Ping() And check_version Then
        Debug_mode = False
        r = ExportAllMod()
        change_log = DownloadMod("changelog" & ".txt")
        msg_upd = vbNullString
        common_git = DownloadMod("common" & ".bas")
        common_local = ConvTxtNum(common_version)
        If common_git > common_local Then msg_upd = msg_upd & "Загружена новая версия модуля common -" & CStr(common_git) & vbNewLine
        calc_git = DownloadMod("calc" & ".bas")
        calc_local = ConvTxtNum(macro_version)
        If calc_git > calc_local Then msg_upd = msg_upd & "Загружена новая версия модуля calc -" & CStr(calc_git) & vbNewLine
        r = DownloadMod("UserForm2.frx")
        form_git = DownloadMod("UserForm2" & ".bas")
        form_local = ConvTxtNum(UserForm2.form_ver.Caption)
        If form_git > form_local Then msg_upd = msg_upd & "Загружена новая версия формы -" & CStr(form_git) & vbNewLine
        update_git = DownloadMod("update" & ".bas")
        update_local = ConvTxtNum(update_version)
        If update_git > update_local Then msg_upd = msg_upd & "Загружена новая версия модуля update -" & CStr(update_git) & vbNewLine
        If Len(msg_upd) > 1 Then
            msg_upd = msg_upd & vbNewLine & change_log & vbNewLine & "Открыть справку с инструкцией?"
            intMessage = MsgBox(msg_upd, vbYesNo, "Доступно обновление")
            Set objShell = CreateObject("Wscript.Shell")
            If intMessage = vbYes Then objShell.Run ("https://docs.google.com/document/d/1bedvuS3quC37ivwVWzWDyfZSt_zxFPhGAQAQ38g3Ubo/edit#bookmark=id.5awyi0cjwmrs")
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

Function DownloadMod(ByVal namemod As String) As Variant
    Dim gitdir As String
    gitdir = UserForm2.CodePath & "from_git"
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(gitdir) Then
        MkDir (gitdir)
    End If
    Dim myURL As String
    myURL = "https://raw.githubusercontent.com/kuvbur/material-specification/master/data/code/" & namemod
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
    myURL = WinHttpReq.responseBody
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile gitdir & "/" & namemod, 2
        oStream.Close
    End If
    If InStr(namemod, ".bas") > 0 Or InStr(namemod, ".txt") > 0 Then
        On Error Resume Next
        Set fso = CreateObject("scripting.filesystemobject")
        Set ts = fso.OpenTextFile(gitdir & "/" & namemod, 1, True): txt$ = ts.ReadAll: ts.Close
        Set ts = Nothing: Set fso = Nothing
        txt = Trim(txt): Err.Clear
    End If
    If InStr(namemod, ".bas") > 0 Then
        seach_txt = "version As String ="
        For Each tRows In Split(txt, vbNewLine)
            If InStr(tRows, seach_txt) Then
                k = Mid(tRows, InStr(tRows, seach_txt) + Len(seach_txt), 6)
                Version_file = ConvTxtNum(Split(k, """")(1))
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
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.CodePath) Then
        MkDir (UserForm2.CodePath)
    End If
    pathtmp = "old\"
    If Debug_mode Then pathtmp = vbNullString
    r = ExportMod("UserForm2", pathtmp, UserForm2.form_ver.Caption)
    r = ExportMod("calc", pathtmp, macro_version)
    r = ExportMod("common", pathtmp, common_version)
    r = ExportMod("update", pathtmp, update_version)
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

Function ModeType() As Boolean
    ismat = CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.MaterialPath)
    issort = CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.SortamentPath)
    iscode = CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.CodePath)
    If ismat And issort And iscode Then
        read_only_mode = False
    Else
        read_only_mode = True
    End If
    ModeType = read_only_mode
End Function

Function Download_Sortament() As Boolean
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
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.SortamentPath) Then
            MkDir (UserForm2.SortamentPath)
        End If
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile ThisWorkbook.path & "\data\sort.zip", 2
        oStream.Close
        Set oApp = CreateObject("Shell.Application")
        For Each it In oApp.Namespace(ThisWorkbook.path & "\data\sort.zip").Items: DoEvents: DoEvents: Next
        oApp.Namespace(ThisWorkbook.path & "\data").CopyHere oApp.Namespace(ThisWorkbook.path & "\data\sort.zip").Items
        Download_Sortament = True
    Else
        Download_Sortament = False
    End If
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
        If Not CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.CodePath) Then
            MkDir (UserForm2.CodePath)
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
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\import") Then
        MkDir (ThisWorkbook.path & "\import")
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\list") Then
        MkDir (ThisWorkbook.path & "\list")
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(ThisWorkbook.path & "\data") Then
        MkDir (ThisWorkbook.path & "\data")
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.CodePath) Then
        MkDir (UserForm2.CodePath)
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.SortamentPath) Then
        MkDir (UserForm2.SortamentPath)
    End If
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.MaterialPath) Then
        MkDir (UserForm2.MaterialPath)
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
            tdate = "_" & CStr(in_ver) & "_" & Replace(Right(Str(DatePart("yyyy", Now)), 2) & Str(DatePart("m", Now)) & Str(DatePart("d", Now)), " ", vbNullString)
            If Not CreateObject("Scripting.FileSystemObject").FolderExists(UserForm2.CodePath & pathtmp) Then
                MkDir (UserForm2.CodePath & pathtmp)
            End If
        End If
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Item(namemod).Export UserForm2.CodePath & pathtmp & namemod & tdate & ".bas"
End Function

Function IsModFileEx(ByVal namemod As String, Optional ByVal pathtmp As String = vbNullString) As Boolean
    On Error Resume Next
    IsModFileEx = CBool(Len(Dir$(UserForm2.CodePath & pathtmp & namemod & ".bas")))
End Function

Function IsModEx(ByVal namemod As String) As Boolean
    On Error Resume Next
    IsModEx = CBool(Len(ThisWorkbook.VBProject.VBComponents(namemod).Name))
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

