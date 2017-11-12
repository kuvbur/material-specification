Attribute VB_Name = "common"
Option Compare Text
Option Base 1

Public Const common_version As String = "3.0"

Public Function GetLeghtByID(id As String, table As Range, n_col_id As Integer, n_col_l As Integer) As Variant
    Sum_l = 0
    For i = 1 To table.Rows.Count
        If table(i, n_col_id) = id Then Sum_l = Sum_l + table(i, n_col_l)
    Next i
    GetLeghtByID = Sum_l
End Function
