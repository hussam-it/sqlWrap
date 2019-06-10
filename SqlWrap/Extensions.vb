'                                 بسم الله الرحمن الرحيم وبه نستعين
'                                 ---------------------------------
Imports System.Runtime.CompilerServices

Public Module Extensions
    Public Zero As Byte = 0
    Public Const EmptyString = ""

    <Extension>
    Public Function First(Rows As DataRowCollection) As DataRow
        If Rows.IsNotEmpty Then
            Return Rows(Zero)
        End If
        Return Nothing
    End Function

    <Extension>
    Public Function IsEmpty(Rows As DataRowCollection) As Boolean
        Return (Rows.Count = Zero)
    End Function

    <Extension>
    Public Function IsNotEmpty(Rows As DataRowCollection) As Boolean
        Return Not IsEmpty(Rows)
    End Function

    <Extension>
    Public Function First(Row As DataRow) As Object
        Return Row(Zero)
    End Function

    <Extension>
    Public Function GetValue(Row As DataRow, ColumnName As String, Optional DefaultValue As Object = Nothing, Optional Type As Type = Nothing) As Object
        Dim Result As Object = Nothing
        If Row.IsNull(ColumnName) Then
            If DefaultValue Is Nothing Then
                If Type Is Nothing Then
                    Result = String.Empty
                Else
                    Select Case Type
                        Case GetType(String)
                            Result = String.Empty
                        Case GetType(Short), GetType(Integer), GetType(Long)
                            Result = Zero
                        Case GetType(Date), GetType(DateTime)
                            Result = New DateTime(1900, 1, 1)
                    End Select
                End If
            Else
                Result = DefaultValue
            End If
        Else
            Result = Row(ColumnName)
        End If
        Return Result
    End Function

    <Extension>
    Public Function IsEmpty(Value As String) As Boolean
        Return String.IsNullOrEmpty(Value)
    End Function

    <Extension>
    Public Function IsNotEmpty(Value As String) As Boolean
        Return Not IsEmpty(Value)
    End Function

End Module
