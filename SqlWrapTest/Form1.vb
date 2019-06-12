'                                 بسم الله الرحمن الرحيم وبه نستعين
'                                 ---------------------------------
Imports SqlWrap

Public Class Form1
    Public Property ConnString As String = "Data Source=.;Initial Catalog=Northwind;Integrated Security=True"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ComboBox1
            .Items.Add(New ListItem With {.Text = "Query Customer Table (GetRows)", .ID = 1})
            .Items.Add(New ListItem With {.Text = "Query Customer Table (GetFieldValue)", .ID = 2})
            .Items.Add(New ListItem With {.Text = "Query Customer Table (Normal)", .ID = 3})
            .Items.Add(New ListItem With {.Text = "Query Customer Table (Join with Orders)", .ID = 4})
            .Items.Add(New ListItem With {.Text = "Query Customer Table (Like)", .ID = 8})
            .Items.Add(New ListItem With {.Text = "Query Customer Table (Free)", .ID = 9})
            .Items.Add(New ListItem With {.Text = "Update Customer Table", .ID = 5})
            .Items.Add(New ListItem With {.Text = "Insert Into Customer Table", .ID = 6})
            .Items.Add(New ListItem With {.Text = "Delete From Customer Table", .ID = 7})
            .SelectedIndex = 0
        End With
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Select Case ComboBox1.SelectedItem.ID
            Case 1
                Operation_01()
            Case 2
                Operation_02()
            Case 3
                Operation_03()
            Case 4
                Operation_04()
            Case 5
                Operation_05()
            Case 6
                Operation_06()
            Case 7
                Operation_07()
            Case 8
                Operation_08()
            Case 9
                Operation_09()
        End Select
    End Sub

    Private Sub Operation_01()
        Dim Database As New Database(Me.ConnString)
        Try
            Dim Rows As DataRowCollection = Database.GetRows("Customers", New List(Of String) From {"CustomerID", "Fax"}, New QItems From {New QItem("city", "London", QItem.Types.WHERE)})

            Label1.Text = Database.Query.StatmentString
            Label3.ForeColor = Color.DarkGreen

            If Rows.IsEmpty Then
                DataGridView1.DataSource = Nothing
                Label3.Text = "Result: Nothing"
            Else
                DataGridView1.DataSource = Database.Query.DataTable
                Label3.Text = "Result: Done"
            End If
        Catch ex As Exception
            Label3.ForeColor = Color.Red
            Label3.Text = "Result: Error """ & ex.Message & """"
        Finally
            SqlWrap.Database.DisConnect(Database)
        End Try
    End Sub

    Private Sub Operation_02()
        Dim Database As New Database(Me.ConnString)
        Dim Result As Object = Nothing
        Try
            Result = Database.GetFieldValue("Customers", "Fax", New QItems From {New QItem("CustomerId", "BSBEV", QItem.Types.WHERE)})

            Label1.Text = Database.Query.StatmentString
            Label3.ForeColor = Color.DarkGreen
            DataGridView1.DataSource = Nothing

            If Result Is Nothing Then
                Label3.Text = "Result: Nothing"
            Else
                Label3.Text = "Result: " & Result
            End If
        Catch ex As Exception
            Label3.ForeColor = Color.Red
            Label3.Text = "Result: Error """ & ex.Message & """"
        Finally
            SqlWrap.Database.DisConnect(Database)
        End Try
    End Sub

    Private Sub Operation_03()
        Dim Database As New Database(Me.ConnString)
        Try
            Database.Table.Name = "Customers"
            Database.Table.AliasName = "Cust"
            With Database.Table.Query.Items
                .Add("Cust.CustomerID").Add("Cust.Fax")
                .Add("Cust.City", "London", QItem.Types.WHERE)
            End With

            With Database.Table.Query.RunQuery
                Label1.Text = Database.Query.StatmentString
                Label3.ForeColor = Color.DarkGreen

                If .Rows.IsEmpty Then
                    DataGridView1.DataSource = Nothing
                    Label3.Text = "Result: Nothing"
                Else
                    DataGridView1.DataSource = Database.Query.DataTable
                    Label3.Text = "Result: Done"
                End If
            End With

        Catch ex As Exception
            Label3.ForeColor = Color.Red
            Label3.Text = "Result: Error """ & ex.Message & """"
        Finally
            SqlWrap.Database.DisConnect(Database)
        End Try
    End Sub

    Private Sub Operation_04()
        Dim Database As New Database(Me.ConnString)
        Try
            Database.Table.Name = "Customers"
            Database.Table.AliasName = "cust"
            With Database.Table.Query.Items
                .Add("orders.OrderID ordId")
                .Add("orders.OrderDate")
                .Add("cust.*")
                .Add("Orders orders", "orders.CustomerID = cust.CustomerID", QItem.Types.JOIN_LEFT)
                .Add("cust.CustomerId", "BSBEV", QItem.Types.WHERE)
            End With

            With Database.Table.Query.RunQuery
                Label1.Text = Database.Query.StatmentString
                Label3.ForeColor = Color.DarkGreen

                If .Rows.IsEmpty Then
                    DataGridView1.DataSource = Nothing
                    Label3.Text = "Result: Nothing"
                Else
                    DataGridView1.DataSource = Database.Query.DataTable
                    Label3.Text = "Result: Done"
                End If
            End With

        Catch ex As Exception
            Label3.ForeColor = Color.Red
            Label3.Text = "Result: Error """ & ex.Message & """"
        Finally
            SqlWrap.Database.DisConnect(Database)
        End Try
    End Sub

    Private Sub Operation_05()
        Dim Database As New Database(Me.ConnString)
        Dim Result As Long = Nothing
        Try
            Database.Table.Name = "Customers"
            Database.Transaction = Database.Connection.BeginTransaction
            With Database.Query.Items
                .Add("Fax", "222-444")
                .Add("CustomerId", "BSBEV", QItem.Types.WHERE)
            End With

            Result = Database.Table.Query.RunNoneQuery(Query.Types.UPDATE)
            Database.CommitTrans()

            Label1.Text = Database.Query.StatmentString
            DataGridView1.DataSource = Nothing

            If Result = 0 Then
                Label3.ForeColor = Color.Red
                Label3.Text = "Result: Failed"
            Else
                Label3.ForeColor = Color.DarkGreen
                Label3.Text = "Result: Succeeded"
            End If
        Catch ex As Exception
            Label3.ForeColor = Color.Red
            Label3.Text = "Result: Error """ & ex.Message & """"
        Finally
            SqlWrap.Database.DisConnect(Database)
        End Try
    End Sub

    Private Sub Operation_06()
        Dim Database As New Database(Me.ConnString)
        Dim Result As Long = Nothing
        Try
            Database.Table.Name = "Region"
            Database.Transaction = Database.Connection.BeginTransaction
            With Database.Query.Items
                .Add("RegionID", "222")
                .Add("RegionDescription", "Test")
            End With
            Result = Database.Table.Query.RunNoneQuery(Query.Types.INSERT)
            Database.CommitTrans()

            Label1.Text = Database.Query.StatmentString
            DataGridView1.DataSource = Nothing

            If Result = 0 Then
                Label3.ForeColor = Color.Red
                Label3.Text = "Result: Failed"
            Else
                Label3.ForeColor = Color.DarkGreen
                Label3.Text = "Result: Succeeded"
            End If
        Catch ex As Exception
            Label3.ForeColor = Color.Red
            Label3.Text = "Result: Error """ & ex.Message & """"
        Finally
            SqlWrap.Database.DisConnect(Database)
        End Try
    End Sub

    Private Sub Operation_07()
        Dim Database As New Database(Me.ConnString)
        Dim Result As Long = Nothing
        Try
            Database.Table.Name = "Region"
            Database.Transaction = Database.Connection.BeginTransaction
            With Database.Query.Items
                .Add("RegionID", "222", QItem.Types.WHERE)
            End With
            Result = Database.Table.Query.RunNoneQuery(Query.Types.DELETE)
            Database.CommitTrans()

            Label1.Text = Database.Query.StatmentString
            DataGridView1.DataSource = Nothing

            If Result = 0 Then
                Label3.ForeColor = Color.Red
                Label3.Text = "Result: Failed"
            Else
                Label3.ForeColor = Color.DarkGreen
                Label3.Text = "Result: Succeeded"
            End If
        Catch ex As Exception
            Label3.ForeColor = Color.Red
            Label3.Text = "Result: Error """ & ex.Message & """"
        Finally
            SqlWrap.Database.DisConnect(Database)
        End Try
    End Sub

    Private Sub Operation_08()
        Dim Database As New Database(Me.ConnString)
        Try
            Database.Table.Name = "Customers"
            Database.Table.AliasName = "cust"
            With Database.Table.Query.Items
                .Add("cust.CustomerId")
                .Add("cust.CompanyName")
                .Add("cust.ContactName")
                .Add("cust.City")
                .Add("cust.City", "Lon", QItem.Types.WHERE_LIKE)
            End With

            With Database.Table.Query.RunQuery
                Label1.Text = Database.Query.StatmentString
                Label3.ForeColor = Color.DarkGreen

                If .Rows.IsEmpty Then
                    DataGridView1.DataSource = Nothing
                    Label3.Text = "Result: Nothing"
                Else
                    DataGridView1.DataSource = Database.Query.DataTable
                    Label3.Text = "Result: Done"
                End If
            End With

        Catch ex As Exception
            Label3.ForeColor = Color.Red
            Label3.Text = "Result: Error """ & ex.Message & """"
        Finally
            SqlWrap.Database.DisConnect(Database)
        End Try
    End Sub

    Private Sub Operation_09()
        Dim Database As New Database(Me.ConnString)
        Try
            Database.Table.Name = "Customers"
            Database.Table.AliasName = "cust"
            With Database.Table.Query.Items
                .Add("cust.*")
                .Add(, "City = N'Berlin'", QItem.Types.WHERE_FREE)
            End With

            With Database.Table.Query.RunQuery
                Label1.Text = Database.Query.StatmentString
                Label3.ForeColor = Color.DarkGreen

                If .Rows.IsEmpty Then
                    DataGridView1.DataSource = Nothing
                    Label3.Text = "Result: Nothing"
                Else
                    DataGridView1.DataSource = Database.Query.DataTable
                    Label3.Text = "Result: Done"
                End If
            End With

        Catch ex As Exception
            Label3.ForeColor = Color.Red
            Label3.Text = "Result: Error """ & ex.Message & """"
        Finally
            SqlWrap.Database.DisConnect(Database)
        End Try
    End Sub

End Class

Public Class ListItem
    Public Property Text As String = String.Empty
    Public Property ID As Short = 0

    Public Overrides Function ToString() As String
        Return Me.Text
    End Function
End Class