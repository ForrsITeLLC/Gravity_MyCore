Public Class cImport

    Dim Database As MyCore.Data.EasySql

    Public TableName As String = ""
    Public Data As DataTable
    Public Columns As New DataTable
    Public ReferenceTables As New Hashtable

    Public LastIncremented As Boolean = False
    Public LastInsertNo As String = ""
    Dim ColumnHashTable As New Hashtable

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
        Me.Columns.Columns.Add("Display Name", GetType(System.String))
        Me.Columns.Columns.Add("Data Property", GetType(System.String))
        Me.Columns.Columns.Add("Index", GetType(System.Int16))
        Me.Columns.Columns.Add("Reference Table", GetType(System.String))
        Me.Columns.Columns.Add("Required", GetType(System.Boolean))
        Me.Columns.Columns.Add("Default Value", GetType(System.String))
        Me.Columns.Columns.Add("Allow Null", GetType(System.Boolean))
        Me.Columns.Columns.Add("Ignore", GetType(System.Boolean))
        Me.Columns.Columns("Index").DefaultValue = -1
        Me.Columns.Columns("Required").DefaultValue = False
        Me.Columns.Columns("Allow Null").DefaultValue = False
        Me.Columns.Columns("Ignore").DefaultValue = False
    End Sub

    Public Function AddNewColumn(ByVal Name As String, ByVal Prop As String, Optional ByVal Required As Boolean = False, Optional ByVal RefTable As String = "", Optional ByVal DefaultValue As String = "", Optional ByVal AllowNull As Boolean = False) As DataRow
        Dim Row As DataRow = Me.Columns.NewRow
        Row.Item("Display Name") = Name
        Row.Item("Data Property") = Prop
        Row.Item("Index") = -1
        Row.Item("Reference Table") = RefTable
        Row.Item("Required") = Required
        Row.Item("Default Value") = DefaultValue
        Row.Item("Allow Null") = AllowNull
        Row.Item("Ignore") = False
        Me.Columns.Rows.Add(Row)
        Me.ColumnHashTable.Add(Prop, Me.Columns.Rows.Count - 1)
        Return Row
    End Function

    Public Function AddNewIgnoredColumn(ByVal Name As String, ByVal Prop As String, Optional ByVal Required As Boolean = False, Optional ByVal RefTable As String = "", Optional ByVal DefaultValue As String = "", Optional ByVal AllowNull As Boolean = False) As DataRow
        Dim Row As DataRow = Me.Columns.NewRow
        Row.Item("Display Name") = Name
        Row.Item("Data Property") = Prop
        Row.Item("Index") = -1
        Row.Item("Reference Table") = RefTable
        Row.Item("Required") = Required
        Row.Item("Default Value") = DefaultValue
        Row.Item("Allow Null") = AllowNull
        Row.Item("Ignore") = True
        Me.Columns.Rows.Add(Row)
        Me.ColumnHashTable.Add(Prop, Me.Columns.Rows.Count - 1)
        Return Row
    End Function

    Public Sub AddReferenceTable(ByVal TableName As String, Optional ByVal IdField As String = "id", Optional ByVal NameField As String = "name")
        If Not Me.ReferenceTables.ContainsKey(TableName) Then
            Dim ThisTable As New Hashtable
            Dim t As DataTable = Me.Database.GetAll("SELECT " & IdField & ", " & NameField & " FROM " & TableName)
            For Each r As DataRow In t.Rows
                If Not ThisTable.ContainsKey(r.Item(NameField)) Then
                    ThisTable.Add(r.Item(NameField), r.Item(IdField))
                End If
            Next
            Me.ReferenceTables.Add(TableName, ThisTable)
        Else
            Throw New Exception("Reference table already exists.")
        End If
    End Sub

    Public Function GetColIndex(ByVal Name As String) As Integer
        If Me.ColumnHashTable.ContainsKey(Name) Then
            Return Me.ColumnHashTable(Name)
        Else
            Return -1
        End If
    End Function

    Public Function GetDataColIndex(ByVal Name As String) As Integer
        Dim Index As Integer = Me.GetColIndex(Name)
        If Index >= 0 Then
            Return Me.Columns.Rows(Index).Item("Index")
        Else
            Return -1
        End If
    End Function

    Public Function GetDataValue(ByVal RowIndex As Integer, ByVal ColName As String) As String
        Dim Index As Integer = Me.GetDataColIndex(ColName)
        Dim ColIndex As Integer = Me.GetColIndex(ColName)
        If Index >= 0 Then
            If Me.Data.Rows(RowIndex).Item(Index) IsNot DBNull.Value Then
                Return Me.Data.Rows(RowIndex).Item(Index)
            ElseIf Me.Columns.Rows(ColIndex).Item("Allow Null") Then
                Return Nothing
            Else
                Return Me.Columns.Rows(ColIndex).Item("Default Value")
            End If
        ElseIf ColIndex >= 0 Then
            Return Me.Columns.Rows(ColIndex).Item("Default Value")
        Else
            Return Nothing
        End If
    End Function

    Public Function InsertSql(ByVal RowIndex As Integer) As String
        Me.LastIncremented = False
        Dim Sql As String = "INSERT INTO [" & Me.TableName & "] ("
        For Each Col As DataRow In Me.Columns.Rows
            If Not CType(Col.Item("Ignore"), Boolean) Then
                Sql &= "[" & Col.Item("Data Property").ToString.Replace("$", "") & "],"
            End If
        Next
        Sql = Sql.Substring(0, Sql.Length - 1)
        Sql &= ") VALUES ("
        For Each Col As DataRow In Me.Columns.Rows
            Dim DataProp As String = Col.Item("Data Property")
            If Not CType(Col.Item("Ignore"), Boolean) Then
                If Col.Item("Reference Table").ToString.Length = 0 Then
                    If DataProp.StartsWith("$") Then
                        If Col.Item("Index") >= 0 Then
                            Sql &= Me.Database.ToCurrency(Me.Database.Escape(Me.Data.Rows(RowIndex).Item(Col.Item("Index"))))
                        Else
                            Sql &= Me.Database.ToCurrency(Me.Database.Escape(Col.Item("Default Value")))
                        End If
                    Else
                        If Col.Item("Index") >= 0 Then
                            Sql &= Me.Database.Escape(Me.Data.Rows(RowIndex).Item(Col.Item("Index")))
                            If Col.Item("Display Name") = "Customer No" Then
                                Me.LastInsertNo = Me.Data.Rows(RowIndex).Item(Col.Item("Index"))
                            End If
                        ElseIf Col.Item("Display Name") = "Customer No" Then
                            Me.LastInsertNo = Me.GetNextNumber("company")
                            Sql &= Me.Database.Escape(Me.LastInsertNo)
                            Me.LastIncremented = True
                        Else
                            Sql &= Me.Database.Escape(Col.Item("Default Value"))
                        End If
                    End If
                Else
                    If Col.Item("Index") >= 0 Then
                        If Me.ReferenceTables.ContainsKey(Col.Item("Reference Table")) Then
                            Dim RefTable As Hashtable = Me.ReferenceTables(Col.Item("Reference Table"))
                            If RefTable.ContainsKey(Me.Data.Rows(RowIndex).Item(Col.Item("Index"))) Then
                                Sql &= Me.Database.Escape(RefTable(Me.Data.Rows(RowIndex).Item(Col.Item("Index"))))
                            Else
                                Sql &= Me.Database.Escape(Col.Item("Default Value"))
                            End If
                        Else
                            Sql &= Me.Database.Escape(Col.Item("Default Value"))
                        End If
                    Else
                        Sql &= Me.Database.Escape(Col.Item("Default Value"))
                    End If
                End If
                Sql &= ","
            End If
        Next
        Sql = Sql.Substring(0, Sql.Length - 1)
        Sql &= ")"
        Return Sql
    End Function

    Public Sub ImportDataFromFile(ByVal FileName As String, ByVal HeaderLine As Boolean, ByVal ColSep As String, ByVal RowSep As String, ByVal QuotesAroundFields As Boolean)
        Dim Content As String = My.Computer.FileSystem.ReadAllText(FileName)
        Dim Lines As String() = Content.Split(RowSep)
        Dim Count As Integer = 1
        Me.Data = New DataTable
        For Each Line As String In Lines
            If Line.Trim.Length > 0 Then
                Dim Fields As String() = Line.Split(ColSep)
                If Count = 1 Then
                    For i As Integer = 0 To Fields.Length - 1
                        Dim dc As DataColumn
                        If HeaderLine Then
                            Try
                                dc = New DataColumn(Fields(i), GetType(System.String))
                            Catch
                                dc = New DataColumn("Field " & (i + 1).ToString, GetType(System.String))
                            End Try
                        Else
                            dc = New DataColumn("Field " & (i + 1).ToString, GetType(System.String))
                        End If
                        Me.Data.Columns.Add(dc)
                    Next
                End If
                If Count > 1 Or Not HeaderLine Then
                    Dim NewRow As DataRow = Me.Data.NewRow
                    For i As Integer = 0 To Fields.Length - 1
                        If QuotesAroundFields Then
                            NewRow.Item(i) = Fields(i).Trim.Replace("""", "")
                        Else
                            NewRow.Item(i) = Fields(i).Trim
                        End If
                    Next
                    Me.Data.Rows.Add(NewRow)
                End If
                Count += 1
            End If
        Next
    End Sub

    Public Sub IncrementNextNumber(ByVal type As String)
        Me.Database.Execute("UPDATE next_number SET number=number+1 WHERE name='" & type & "'")
    End Sub

    Public Function GetNextNumber(ByVal type As String) As Integer
        Return Me.Database.GetOne("SELECT number FROM next_number WHERE name='" & Type & "'")
    End Function

    Public Function DataColumns() As DataTable
        Dim Table As New DataTable
        Table.Columns.Add("Name")
        Table.Columns.Add("Value")
        ' Add Blank Row
        Dim Blank As DataRow = Table.NewRow
        Blank.Item("Name") = ""
        Blank.Item("Value") = -1
        Table.Rows.Add(Blank)
        ' Add Actual Rows
        For Each Col As DataColumn In Me.Data.Columns
            Dim NewRow As DataRow = Table.NewRow
            NewRow.Item("Name") = Col.ColumnName
            NewRow.Item("Value") = Table.Rows.Count - 1
            Table.Rows.Add(NewRow)
        Next
        Return Table
    End Function


End Class
