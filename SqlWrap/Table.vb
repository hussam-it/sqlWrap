'                                 بسم الله الرحمن الرحيم وبه نستعين
'                                 ---------------------------------

Imports System.Data.SqlClient

Public Class Table
    Public Property AliasName As String = String.Empty
    Public Property DbSchema As String = "dbo"
    Public Property Parent As Database = Nothing

    Private _Query As New Query
    Public ReadOnly Property Query() As Query
        Get
            Return _Query
        End Get
    End Property

    Private _Name As String
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
            Me.AliasName = String.Empty
            _Query = New Query With {.Parent = Me}
        End Set
    End Property

    Public ReadOnly Property Connection() As SqlConnection
        Get
            Return Me.Parent.Connection
        End Get
    End Property

    Public ReadOnly Property Database() As Database
        Get
            Return Me.Parent
        End Get
    End Property

    Friend Sub New()
        Me.Query.Parent = Me
    End Sub

End Class

Public Class Tables
    Inherits List(Of Table)

End Class