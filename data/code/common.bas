Attribute VB_Name = "common"
Option Compare Text
Option Base 1
Public Const common_version As String = "3.5"

Public Function GetLeghtByID(id As String, table As Range, n_col_id As Integer, n_col_l As Integer) As Variant
    Sum_l = 0
    For i = 1 To table.Rows.Count
        If table(i, n_col_id) = id Then Sum_l = Sum_l + table(i, n_col_l)
    Next i
    GetLeghtByID = Sum_l
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



