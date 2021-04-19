VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Фильтрация элементов"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3645
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const form1_version As String = "4.01"
Public form1_sh_name As String
Private Sub CommandButton1_Click()
    r = write_option()
    Unload UserForm1
End Sub

Function write_option()
    UserForm2.Kzap.Text = UserForm1.Kzap.Text
    Set calc.this_sheet_option = OptionGetForm(form1_sh_name)
End Function

Private Sub UserForm_Activate()
    UserForm1.Kzap.Text = UserForm2.Kzap.Text
    If Not IsEmpty(calc.set_sheet_option) Then
        UserForm1.CheckBox_isarm.Value = calc.set_sheet_option.Item("isarm")
        UserForm1.CheckBox_isizd.Value = calc.set_sheet_option.Item("isizd")
        UserForm1.CheckBox_isprok.Value = calc.set_sheet_option.Item("isprok")
        UserForm1.CheckBox_ismat.Value = calc.set_sheet_option.Item("ismat")
        UserForm1.CheckBox_issubpos.Value = calc.set_sheet_option.Item("issubpos")
        UserForm1.TextBox_subpos_add.Value = OptionJoin(calc.set_sheet_option.Item("arr_subpos_add"))
        UserForm1.TextBox_subpos_del.Value = OptionJoin(calc.set_sheet_option.Item("arr_subpos_del"))
        UserForm1.TextBox_typeKM_add.Value = OptionJoin(calc.set_sheet_option.Item("arr_typeKM_add"))
        UserForm1.TextBox_typeKM_del.Value = OptionJoin(calc.set_sheet_option.Item("arr_typeKM_del"))
    Else
        Debug.Print "ERROR Словарь с настройками не загружен " & form1_sh_name
    End If
End Sub

Function show_form(ByVal nm As String)
    form1_sh_name = nm
    UserForm1.Show
End Function
