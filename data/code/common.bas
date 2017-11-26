Attribute VB_Name = "common"
Option Compare Text
Option Base 1

Public Const common_version As String = "3.1"

Public Function GetLeghtByID(id As String, table As Range, n_col_id As Integer, n_col_l As Integer) As Variant
    Sum_l = 0
    For i = 1 To table.Rows.Count
        If table(i, n_col_id) = id Then Sum_l = Sum_l + table(i, n_col_l)
    Next i
    GetLeghtByID = Sum_l
End Function

Public Function GetAreaList(razm As String) As Double
    ab = Split(razm, "x")
    If UBound(ab) = 0 Then ab = Split(razm, "õ")
    If UBound(ab) = 0 Then ab = Split(razm, "*")
    If UBound(ab) = 0 Then
        GetAreaList = 0
        Exit Function
    End If
    aa = ConvTxt2Num(ab(0))
    bb = ConvTxt2Num(ab(1))
    GetAreaList = aa * bb
End Function
