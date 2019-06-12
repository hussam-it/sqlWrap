'                                 بسم الله الرحمن الرحيم وبه نستعين
'                                 ---------------------------------

Imports System.Data.SqlClient
Imports System.Text

Public Class Query
    Friend Sub New()

    End Sub

    Public Property Parent As Table = Nothing

    Public Property Items As New QItems
    Public Property Tables As New Tables
    Public Property DataTable As DataTable = Nothing

    Public Property CommandTimeOut As Short = 90
    Public Property StatmentString = String.Empty
    Public Property WhereString = String.Empty

    Private _Type As Types = Zero
    Public ReadOnly Property Type() As Types
        Get
            Return _Type
        End Get
    End Property

    Public ReadOnly Property Database() As Database
        Get
            Return Me.Parent.Database
        End Get
    End Property

    Public Enum Types
        SELECT_ = 1
        INSERT = 2
        UPDATE = 3
        DELETE = 4
    End Enum

    Public Enum RenderTypes
        WITH_NAME = 1
        WITHOUT_NAME = 2
    End Enum

    Public Sub Clear()
        Me.Items.Clear()
        Me.Tables.Clear()
        Me.DataTable = Nothing
    End Sub

    Public Function RunQuery() As DataTable
        _Type = Types.SELECT_
        Me.DataTable = RunQuery(Me.GetStatmentString(Types.SELECT_))

        Return Me.DataTable
    End Function

    Public Function RunQuery(ByVal StatmentString As String) As DataTable
        Dim dt As New DataTable

        Try
            If StatmentString.IsNotEmpty Then
                Me.StatmentString = StatmentString

                Dim da As New SqlDataAdapter
                da.SelectCommand = Me.Parent.Connection.CreateCommand
                da.SelectCommand.CommandText = StatmentString
                da.SelectCommand.CommandTimeout = Me.CommandTimeOut
                If Not Me.Database.Transaction Is Nothing Then
                    da.SelectCommand.Transaction = Me.Database.Transaction
                End If
                da.Fill(dt)

            End If
        Catch Dbex As SqlException
            Throw New Exception(Dbex.Message)
        End Try

        Return dt
    End Function

    Public Function RunNoneQuery(Type As Types) As Long
        Dim Result As String = String.Empty
        If Not Type = Types.SELECT_ Then
            _Type = Type
            Result = RunNoneQuery(Me.GetStatmentString(Type))
        End If
        Return Result
    End Function

    Public Function RunNoneQuery(ByVal StatmentString As String) As Long

        If StatmentString.IsNotEmpty Then
            Try
                Me.StatmentString = StatmentString

                Dim Cmd As SqlCommand = Nothing
                If Me.Database.Transaction Is Nothing Then
                    Cmd = New SqlCommand(StatmentString, Me.Parent.Connection)
                Else
                    Cmd = New SqlCommand(StatmentString, Me.Parent.Connection, Me.Database.Transaction)
                End If

                Cmd.CommandTimeout = Me.CommandTimeOut

                Return Cmd.ExecuteNonQuery()
            Catch Dbex As SqlException
                Throw New Exception(Dbex.Message)
            End Try
        End If

        Return Zero
    End Function

    Private Function GetFieldValue(FieldTable As String, FieldName As String, FieldValue As String, ByVal QItem As QItem) As String
        If Me.Parent.Connection.State = ConnectionState.Open Then

            Dim SQLstr As New StringBuilder
            SQLstr.Append("SELECT  data_type AS 'Data Type',character_maximum_length AS 'Max Length' ")
            SQLstr.Append("FROM information_schema.columns ")
            SQLstr.Append("WHERE table_name = '" & Me.GetTableName(FieldTable) & "' AND column_name = '" & FieldName & "'")

            With Me.RunQuery(SQLstr.ToString)
                If .Rows.IsNotEmpty Then
                    Select Case .Rows.First("Data Type").ToString.ToLower
                        Case "int", "bigint", "smallint", "tinyint", "real", "float", "money"
                            If IsNumeric(FieldValue) Then
                                Select Case .Rows.First("Data Type").ToString.ToLower
                                    Case "int"
                                        If Not (FieldValue >= Int32.MinValue And FieldValue <= Int32.MaxValue) Then
                                            FieldValue = Zero
                                        End If
                                    Case "bigint", "money"
                                        If Not (FieldValue >= Int64.MinValue And FieldValue <= Int64.MaxValue) Then
                                            FieldValue = Zero
                                        End If
                                    Case "smallint"
                                        If Not (FieldValue >= Int16.MinValue And FieldValue <= Int16.MaxValue) Then
                                            FieldValue = Zero
                                        End If
                                    Case "tinyint"
                                        If Not (FieldValue >= Byte.MinValue And FieldValue <= Byte.MaxValue) Then
                                            FieldValue = Zero
                                        End If
                                    Case "real"

                                    Case "float"

                                    Case "money"

                                End Select
                            Else
                                FieldValue = Zero
                            End If

                        Case "bit"
                            If FieldValue = 1 Then
                                FieldValue = "'True'"
                            Else
                                FieldValue = "'False'"
                            End If

                        Case "datetime", "datetime2", "smalldatetime"
                            If FieldValue.GetType.Name = "DateTime" Or FieldValue.GetType.Name = "Date" Then
                                FieldValue = "'" & CType(FieldValue, DateTime).ToString("yyyy-MM-ddTHH:mm:ss", Globalization.CultureInfo.CreateSpecificCulture("en")) & "'"
                            Else
                                FieldValue = "'" & FieldValue & "'"
                            End If

                        Case "char", "nchar", "varchar", "nvarchar", "text", "ntext"
                            If FieldValue.Length > .Rows.First("Max Length") Then
                                FieldValue = FieldValue.Substring(Zero, .Rows.First("Max Length"))
                            End If
                            FieldValue = FieldValue.Replace("'", "''")
                            If QItem.Type = QItem.Types.WHERE_LIKE Then
                                FieldValue = "N'%" & FieldValue & "%'"
                            Else
                                FieldValue = "N'" & FieldValue & "'"
                            End If
                        Case "uniqueidentifier"
                            FieldValue = FieldValue.Replace("'", String.Empty)
                            FieldValue = "'" & FieldValue & "'"

                    End Select

                End If
            End With
        End If
        Return FieldValue
    End Function

    Private Function RenderFieldValue(ByVal QItem As QItem, Optional ByVal RenderType As RenderTypes = RenderTypes.WITH_NAME) As String
        Dim FieldName As String = QItem.Name
        Dim FieldValue As String = QItem.Value
        Dim FieldTable As String = String.Empty
        Try
            If FieldName.Contains(".") Then
                Dim Split As String() = FieldName.Split(".")
                FieldName = Split(1).Trim
                FieldTable = Split(Zero).Trim
            Else
                FieldTable = Me.Parent.Name
            End If

            FieldValue = Me.GetFieldValue(FieldTable, FieldName, FieldValue, QItem)

            Select Case RenderType
                Case RenderTypes.WITH_NAME
                    If QItem.Type = QItem.Types.WHERE_LIKE Then
                        Return FieldTable & "." & FieldName & " LIKE " & FieldValue
                    Else
                        Return FieldTable & "." & FieldName & QItem.GetOperatorWord & FieldValue
                    End If
                Case RenderTypes.WITHOUT_NAME
                    Return FieldValue
            End Select

        Catch ex As Exception
            Throw
        End Try

        Return "''"
    End Function

    Public Function GetWhereString() As String
        Dim WhereString As String = String.Empty
        For Each QItem As QItem In Me.Items
            Select Case QItem.Type
                Case QItem.Types.WHERE
                    WhereString &= QItem.GetLogicWord & Me.RenderFieldValue(QItem)
                Case QItem.Types.WHERE_LIKE
                    WhereString &= QItem.GetLogicWord & Me.RenderFieldValue(QItem)
                Case QItem.Types.WHERE_IN
                    WhereString &= QItem.GetLogicWord & QItem.Name & " IN (" & QItem.Value & ") "
                Case QItem.Types.WHERE_IS_NULL
                    WhereString &= QItem.GetLogicWord & QItem.Name & " IS NULL "
                Case QItem.Types.WHERE_FREE
                    WhereString &= " (" & QItem.Value & ") "
            End Select
        Next

        RemoveFirst(WhereString, QItem.GetLogicWord(QItem.Logics.AND_))
        RemoveFirst(WhereString, QItem.GetLogicWord(QItem.Logics.OR_))

        Return WhereString
    End Function

    Public Function GetStatmentStringSelect() As String
        Dim StatmentString As String = String.Empty
        Dim WhereString As String = String.Empty

        Dim FieldsString As String = String.Empty
        Dim TopString As String = String.Empty
        Dim OrderString As String = String.Empty
        Dim GroupString As String = String.Empty
        Dim JoinString As String = String.Empty

        For Each QItem As QItem In Me.Items
            Select Case QItem.Type
                Case QItem.Types.FIELD
                    FieldsString &= QItem.Name & ","
                Case QItem.Types.TOP
                    TopString &= "TOP " & QItem.Value & " "
                Case QItem.Types.MAX
                    TopString &= "MAX (" & QItem.Value & ") "
                Case QItem.Types.ORDER_BY_ASC
                    OrderString &= QItem.Name & ","
                Case QItem.Types.ORDER_BY_DESC
                    OrderString &= QItem.Name & " DESC,"
                Case QItem.Types.GROUP_BY
                    GroupString &= QItem.Name & ","
                Case QItem.Types.JOIN_INNER
                    AddTableName(QItem.Name)
                    JoinString &= " INNER JOIN " & Me.Parent.DbSchema & "." & QItem.Name & " ON " & QItem.Value & " "
                Case QItem.Types.JOIN_LEFT
                    AddTableName(QItem.Name)
                    JoinString &= " LEFT JOIN " & Me.Parent.DbSchema & "." & QItem.Name & " ON " & QItem.Value & " "
                Case QItem.Types.JOIN_RIGHT
                    AddTableName(QItem.Name)
                    JoinString &= " RIGHT JOIN " & Me.Parent.DbSchema & "." & QItem.Name & " ON " & QItem.Value & " "
            End Select
        Next

        WhereString = GetWhereString()

        RemoveLast(FieldsString, ",")
        RemoveLast(OrderString, ",")
        RemoveLast(GroupString, ",")

        StatmentString = String.Format("SELECT {0} {1} FROM {2} {3} ", TopString,
                                                                  FieldsString,
                                                                  Me.Parent.DbSchema & "." & Me.Parent.Name & " " & Me.Parent.AliasName,
                                                                  JoinString)

        If WhereString.IsNotEmpty Then
            StatmentString &= "WHERE " & WhereString & " "
        End If

        If OrderString.IsNotEmpty Then
            StatmentString &= "ORDER BY " & OrderString & " "
        End If
        If GroupString.IsNotEmpty Then
            StatmentString &= "GROUP BY " & GroupString & " "
        End If

        Me.StatmentString = StatmentString
        Me.WhereString = WhereString

        Return StatmentString

    End Function

    Public Function GetStatmentStringInsert() As String
        Dim StatmentString As String = String.Empty

        StatmentString = "INSERT INTO " & Me.Parent.DbSchema & "." & Me.Parent.Name & " ("
        For Each QItem As QItem In Me.Items
            If QItem.Type = QItem.Types.FIELD Then
                StatmentString &= QItem.Name & ","
            End If
        Next

        RemoveLast(StatmentString, ",")

        StatmentString &= ") VALUES ("

        For Each QItem As QItem In Me.Items
            If QItem.Type = QItem.Types.FIELD Then
                StatmentString &= Me.RenderFieldValue(QItem, RenderTypes.WITHOUT_NAME) & ","
            End If
        Next

        RemoveLast(StatmentString, ",")

        StatmentString &= ") "

        Me.StatmentString = StatmentString
        Me.WhereString = String.Empty

        Return StatmentString

    End Function

    Public Function GetStatmentStringUpdate() As String
        Dim StatmentString As String = String.Empty
        Dim WhereString As String = String.Empty

        StatmentString = "UPDATE " & Me.Parent.DbSchema & "." & Me.Parent.Name & " SET "
        For Each QItem As QItem In Me.Items
            If QItem.Type = QItem.Types.FIELD Then
                StatmentString &= Me.RenderFieldValue(QItem) & ","
            End If
        Next

        RemoveLast(StatmentString, ",")

        WhereString = GetWhereString()

        StatmentString &= " WHERE " & WhereString

        Me.StatmentString = StatmentString
        Me.WhereString = WhereString

        Return StatmentString

    End Function

    Public Function GetStatmentStringDelete() As String
        Me.WhereString = GetWhereString()
        Me.StatmentString = "DELETE FROM " & Me.Parent.DbSchema & "." & Me.Parent.Name & " WHERE " & WhereString

        Return Me.StatmentString
    End Function

    Public Function GetStatmentString(Type As Types) As String
        Dim StatmentString As String = String.Empty
        Try
            If Me.Parent.Name.IsEmpty Then
                Throw New Exception("Empty Table Name")
            End If
            Select Case Type
                Case Query.Types.SELECT_
                    StatmentString = GetStatmentStringSelect()
                Case Query.Types.INSERT
                    StatmentString = GetStatmentStringInsert()
                Case Query.Types.UPDATE
                    StatmentString = GetStatmentStringUpdate()
                Case Query.Types.DELETE
                    StatmentString = GetStatmentStringDelete()
            End Select

            StatmentString = RefineSpaces(StatmentString)

        Catch ex As Exception
            Throw
        End Try

        Return StatmentString
    End Function


    Private Function GetTableName(TableName As String) As String
        If TableName = Me.Parent.AliasName Then
            TableName = Me.Parent.Name
        Else
            For Each Table As Table In Me.Tables
                If TableName = Table.AliasName Then
                    TableName = Table.Name
                    Exit For
                End If
            Next
        End If
        Return TableName
    End Function

    Private Sub AddTableName(Value As String)
        Dim Name As String = String.Empty
        Dim AliasName As String = String.Empty

        Value = RefineSpaces(Value)

        If Value.Trim.Contains(" ") Then
            Dim Split As String() = Value.Split(" ")
            Name = Split(Zero)
            If Value.ToLower.Contains(" as ") Then
                AliasName = Split(2)
            Else
                AliasName = Split(1)
            End If
        End If

        Me.Tables.Add(New Table With {.Name = Name, .AliasName = AliasName})
    End Sub

    Private Sub RemoveLast(ByRef Str As String, Value As String)
        If Str.EndsWith(Value) Then
            Str = Str.Substring(Zero, Str.Length - 1)
        End If
    End Sub

    Private Function RefineSpaces(ByVal Str As String) As String
        Dim Ind As Boolean = Str.Contains("  ")
        While Ind
            Str = Str.Replace("  ", " ")
            Ind = Str.Contains("  ")
        End While
        Return Str
    End Function
    Private Sub RemoveFirst(ByRef Str As String, Value As String)
        If Str.StartsWith(Value) Then
            Str = Str.Substring(Value.Length)
        End If
    End Sub

End Class

