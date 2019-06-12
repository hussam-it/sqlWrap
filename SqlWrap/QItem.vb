Public Class QItems
    Inherits List(Of QItem)

    Public Overloads Function Add(Optional ByVal Name As String = EmptyString, Optional ByVal Value As String = EmptyString, Optional ByVal Type As QItem.Types = QItem.Types.FIELD, Optional ByVal Logic As QItem.Logics = QItem.Logics.AND_, Optional ByVal Operator_ As QItem.Operators = QItem.Operators.EQUAL) As QItems
        Me.Add(New QItem(Name, Value, Type, Logic, Operator_))
        Return Me
    End Function
End Class

Public Class QItem
    Public Property Name As String = String.Empty
    Public Property Value As String = String.Empty
    Public Property Type As Types = Zero
    Public Property Logic As Logics = Zero
    Public Property Operator_ As Operators = Zero

    Public Enum Logics
        AND_ = 1
        AND_NOT = 2
        OR_ = 3
        OR_NOT = 4
    End Enum

    Public Enum Operators
        EQUAL = 1
        GREATER_THAN = 2
        GREATER_THAN_EQ = 3
        LESS_THAN = 4
        LESS_THAN_EQ = 5
    End Enum

    Public Enum Types
        FIELD = 1
        TOP = 2
        MAX = 3
        WHERE = 4
        WHERE_IN = 5
        WHERE_LIKE = 6
        WHERE_IS_NULL = 7
        WHERE_FREE = 8
        GROUP_BY = 9
        JOIN_INNER = 10
        JOIN_LEFT = 11
        JOIN_RIGHT = 12
        ORDER_BY_ASC = 13
        ORDER_BY_DESC = 14
    End Enum

    Public Sub New()

    End Sub

    Public Sub New(ByVal Name As String, Optional ByVal Value As String = EmptyString, Optional ByVal Type As Types = Types.FIELD, Optional ByVal Logic As QItem.Logics = QItem.Logics.AND_, Optional ByVal Operator_ As QItem.Operators = QItem.Operators.EQUAL)
        Me.Name = Name
        Me.Value = Value
        Me.Type = Type
        Me.Logic = Logic
        Me.Operator_ = Operator_
    End Sub

    Public Shared Function GetLogicWord(Logic As Logics) As String
        Select Case Logic
            Case Logics.AND_
                Return " AND "
            Case Logics.AND_NOT
                Return " AND NOT "
            Case Logics.OR_
                Return " OR "
            Case Logics.OR_NOT
                Return " OR NOT "
        End Select

        Return String.Empty
    End Function

    Public Function GetLogicWord() As String
        Return GetLogicWord(Me.Logic)
    End Function

    Public Function GetOperatorWord(Operator_ As Operators) As String
        Select Case Operator_
            Case Operators.EQUAL
                Return " = "
            Case Operators.GREATER_THAN
                Return " > "
            Case Operators.GREATER_THAN_EQ
                Return " >= "
            Case Operators.LESS_THAN
                Return " < "
            Case Operators.LESS_THAN_EQ
                Return " <= "
        End Select

        Return String.Empty
    End Function

    Public Function GetOperatorWord() As String
        Return GetOperatorWord(Me.Operator_)
    End Function

End Class
