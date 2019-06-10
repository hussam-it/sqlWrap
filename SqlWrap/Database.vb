'                                 بسم الله الرحمن الرحيم وبه نستعين
'                                 ---------------------------------
'-------------------------------------------------------------------------------------------
'- Title          : SqlWrap
'- Description    : Class for wraping the System.Data.SqlClient commands
'-                  and give an object oriented interface for dealing with database
'- Date           : 2019-06-06
'- Author         : hussam.it
'- About The Autor: Developer since 23 years
'-                  Individual
'-                  Main development language is VB.net
'-                  Very like the similarity of Life and programming
'-                  Specially this one:
'-                  https://notec1.wordpress.com/2017/04/12/life-programming-view/
'-------------------------------------------------------------------------------------------
Imports System.Data.SqlClient

Public Class Database
    Public Property CloseInd As Boolean = False
    Public Property ConnString = String.Empty
    Public Property Connection As SqlConnection = Nothing
    Public Property Transaction As SqlTransaction = Nothing

    Public ReadOnly Property Table As New Table

    Public ReadOnly Property Query() As Query
        Get
            Return Me.Table.Query
        End Get
    End Property

    Public Sub New()
        Me.Table.Parent = Me
    End Sub

    Public Sub New(ByVal ConnString As String)
        Me.New()
        Me.Connect(ConnString)
    End Sub

    Public Function Connect(Optional ByVal ConnString As String = EmptyString) As SqlConnection
        Try
            If String.IsNullOrEmpty(ConnString) Then
                ConnString = Me.ConnString
            End If
            If Not String.IsNullOrEmpty(ConnString) Then
                Me.Connection = New SqlConnection
                Me.Connection.ConnectionString = ConnString
                Me.Connection.Open()
                Me.ConnString = ConnString
                Return Me.Connection
            End If
        Catch Dbex As SqlException
            Throw New Exception(Dbex.Message)
        End Try

        Return Nothing
    End Function

    Public Shared Sub DisConnect(ByVal Database As Database)
        If Not Database Is Nothing Then
            Database.CloseInd = False
            If Not Database.Connection Is Nothing Then
                If Database.Connection.State = ConnectionState.Open Then
                    Database.Connection.Close()
                End If
            End If
            Database = Nothing
        End If
    End Sub

    Public Sub CommitTrans()
        If Not Me.Connection Is Nothing AndAlso Me.Connection.State = ConnectionState.Open Then
            Me.Transaction.Commit()
        End If
    End Sub

    Public Sub RollbackTrans()
        If Not Me.Connection Is Nothing AndAlso Me.Connection.State = ConnectionState.Open Then
            Me.Transaction.Rollback()
        End If
    End Sub

    Public Function GetFieldValue(TableName As String, FieldName As String, Conditions As QItems, Optional DefaultValue As Object = Nothing) As Object
        Dim Result As Object = Nothing
        Try
            Me.Table.Name = TableName

            Me.Table.Query.Items.Add(FieldName)
            Me.Table.Query.Items.InsertRange(Me.Table.Query.Items.Count, Conditions)

            With Me.Table.Query.RunQuery
                If .Rows.IsNotEmpty Then
                    Result = .Rows.First.GetValue(FieldName, DefaultValue)
                End If
            End With
        Catch ex As Exception
            Throw
        End Try
        Return Result
    End Function

    Public Function GetRows(TableName As String, Fields As List(Of String), Conditions As QItems) As DataRowCollection
        Dim Result As DataRowCollection = Nothing
        Try
            Me.Table.Name = TableName

            For Each Field As String In Fields
                Me.Table.Query.Items.Add(Field)
            Next

            Me.Table.Query.Items.InsertRange(Me.Table.Query.Items.Count, Conditions)

            Result = Me.Table.Query.RunQuery.Rows
        Catch ex As Exception
            Throw
        End Try
        Return Result
    End Function
End Class
