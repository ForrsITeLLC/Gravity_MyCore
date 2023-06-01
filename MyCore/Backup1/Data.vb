Imports System.Data
Imports MySql.Data.MySqlClient

Namespace Data

    Public Class EasySql

        Dim _DbType As String = Nothing

        Dim OleDbCon As OleDb.OleDbConnection = Nothing
        Dim SqlDbCon As SqlClient.SqlConnection = Nothing
        Dim MySqlCon As New MySqlConnection
        Dim SqlLiteCon As SQLite.SQLiteConnection = Nothing

        Public LastQuery As Query
        Public Event ConnectionStringChanged(ByVal ConString As String)
        Public Event ConnectionTypeChanged(ByVal Type As String)

        Dim _LeaveConnectionOpen As Boolean = False
        Dim _IsConnectionOpen As Boolean = False
        Dim _IsDatabaseInitalized As Boolean = False

        ' Used for updating schema
        Dim TableHash As New Hashtable

        Public ReadOnly Property IsConnectionOpen() As Boolean
            Get
                Return Me._IsConnectionOpen
            End Get
        End Property

        Public ReadOnly Property IsDatabaseInitalized() As Boolean
            Get
                Return Me._IsDatabaseInitalized
            End Get
        End Property

        Public Property LeaveConnectionOpen() As Boolean
            Get
                Return Me._LeaveConnectionOpen
            End Get
            Set(ByVal LeaveOpen As Boolean)
                Me._LeaveConnectionOpen = LeaveOpen
                If Not LeaveOpen Then
                    Me.CloseConnection()
                End If
            End Set
        End Property

        Public ReadOnly Property DatabaseType() As String
            Get
                If Me._DbType <> Nothing Then
                    Return Me._DbType.ToLower
                Else
                    Return Nothing
                End If
            End Get
        End Property

        Public ReadOnly Property DatabaseConnection() As Object
            Get
                If Me.DatabaseType = "mssql" Then
                    Return Me.SqlDbCon
                ElseIf Me.DatabaseType = "mysql" Then
                    Return Me.MySqlCon
                ElseIf Me.DatabaseType = "sqlite" Then
                    Return Me.SqlLiteCon
                Else
                    Return Me.OleDbCon
                End If
            End Get
        End Property

        Public ReadOnly Property DatabaseConnectionString() As String
            Get
                Return Me.DatabaseConnection.ConnectionString
            End Get
        End Property

        Public Class Query

            Public Successful As Boolean
            Public ErrorMsg As String
            Public CommandText As String
            Public AffectedRows As Integer
            Public InsertId As Integer
            Public RowsReturned As Integer

            Public ReadOnly Property Exception() As Exception
                Get
                    If Me.ErrorMsg.Length > 0 Then
                        Return New Exception(Me.ErrorMsg)
                    Else
                        Return Nothing
                    End If
                End Get
            End Property

            Public Sub New(ByVal Success As Boolean, ByVal CommandText As String, Optional ByVal ErrorMsg As String = "")
                Me.Successful = Success
                Me.CommandText = CommandText
                Me.ErrorMsg = ErrorMsg
                Me.AffectedRows = Nothing
                Me.InsertId = Nothing
                Me.RowsReturned = Nothing
            End Sub

        End Class

        Public Sub New(ByVal strDbType As String)
            Me._DbType = strDbType
        End Sub

        Public Sub New(ByVal strConString As String, ByVal strType As String)
            Me.ChangeDb(strConString, strType)
        End Sub

        Public Sub ChangeDb(ByVal strCon As String, Optional ByVal strDbType As String = "")
            If strDbType.Length > 0 Then
                If Not strDbType.ToLower = Me.DatabaseType Then
                    Me._DbType = strDbType
                    RaiseEvent ConnectionTypeChanged(Me.DatabaseType)
                End If
            End If
            ' Kill all
            Me.CloseConnection()
            Me.OleDbCon = Nothing
            Me.MySqlCon = Nothing
            Me.SqlLiteCon = Nothing
            Me.SqlDbCon = Nothing
            Me._IsDatabaseInitalized = False
            ' Set new
            If strCon <> Nothing Then
                If Me.DatabaseType = "mssql" Then
                    Me.SqlDbCon = New SqlClient.SqlConnection(strCon)
                ElseIf Me.DatabaseType = "mysql" Then
                    Me.MySqlCon = New MySqlConnection(strCon)
                ElseIf Me.DatabaseType = "sqlite" Then
                    Me.SqlLiteCon = New SQLite.SQLiteConnection(strCon)
                Else
                    Me.OleDbCon = New OleDb.OleDbConnection(strCon)
                End If
                Me._IsDatabaseInitalized = True
            End If
            RaiseEvent ConnectionStringChanged(strCon)
        End Sub

        Public Function ConnectionState() As System.Data.ConnectionState
            Return Me.DatabaseConnection.State
        End Function

        Private Sub OpenConnection(Optional ByVal Retry As Boolean = True)
            If Me.IsDatabaseInitalized Then
                Try
                    If Not Me.IsConnectionOpen Then
                        Me.DatabaseConnection.Open()
                        Me._IsConnectionOpen = True
                    End If
                Catch ex As Exception
                    If Retry Then
                        ' Wait a second and then try to connect again, but do not retry again so as not to be in a loop
                        System.Threading.Thread.CurrentThread.Sleep(100)
                        Me.OpenConnection(False)
                    Else
                        Throw New Exception("Failed to open connection to database: " & ex.ToString)
                    End If
                End Try
            End If
        End Sub

        Private Sub CloseConnection(Optional ByVal Retry As Boolean = True)
            If Me.IsDatabaseInitalized Then
                Try
                    If Me.IsConnectionOpen Then
                        Me.DatabaseConnection.Close()
                        Me._IsConnectionOpen = False
                    End If
                Catch ex As Exception
                    If Retry Then
                        ' Wait a second and then try to connect again, but do not retry again so as not to be in a loop
                        System.Threading.Thread.CurrentThread.Sleep(100)
                        Me.CloseConnection(False)
                    Else
                        Throw New Exception("Failed to close connection to database: " & ex.ToString)
                    End If
                End Try
            End If
            Me._LeaveConnectionOpen = False
        End Sub

        Private Function GetCommand(ByVal Sql As String, Optional ByVal Type As CommandType = CommandType.Text, Optional ByVal Params As Collection = Nothing) As Object
            Dim Command As Object = Nothing
            If Me.DatabaseType = "mssql" Then
                Command = New SqlClient.SqlCommand(Sql, Me.SqlDbCon)
            ElseIf Me.DatabaseType = "mysql" Then
                Command = New MySqlCommand(Sql, Me.MySqlCon)
            ElseIf Me.DatabaseType = "sqlite" Then
                Command = New SQLite.SQLiteCommand(Sql, Me.SqlLiteCon)
            Else
                Command = New OleDb.OleDbCommand(Sql, Me.OleDbCon)
            End If
            Command.CommandType = Type
            If Not Params Is Nothing Then
                For Each Param As Param In Params
                    Command.Parameters.AddWithValue(Param.Name, Param.Value)
                Next
            End If
            Return Command
        End Function

        Private Function GetAdapter(ByVal Sql As String, Optional ByVal Type As CommandType = CommandType.Text, Optional ByVal Params As Collection = Nothing) As Object
            Dim Adapter As Object = Nothing
            If Me.DatabaseType = "mssql" Then
                Adapter = New SqlClient.SqlDataAdapter(Sql, Me.SqlDbCon)
            ElseIf Me.DatabaseType = "mysql" Then
                Adapter = New MySqlDataAdapter(Sql, Me.MySqlCon)
            ElseIf Me.DatabaseType = "sqlite" Then
                Adapter = New SQLite.SQLiteDataAdapter(Sql, Me.SqlLiteCon)
            Else
                Adapter = New OleDb.OleDbDataAdapter(Sql, Me.OleDbCon)
            End If
            Adapter.SelectCommand.CommandType = Type
            If Not Params Is Nothing Then
                For Each Param As Param In Params
                    Adapter.SelectCommand.Parameters.AddWithValue(Param.Name, Param.Value)
                Next
            End If
            Return Adapter
        End Function

        Private Function GetReader(ByVal Command As Object) As Object
            Return Command.ExecuteReader
        End Function

        Private Function GetInsertId() As Integer
            ' Connection should already be open
            Dim InsertId As Integer = Nothing
            Dim Sql As String = ""
            If Me.DatabaseType = "mssql" Then
                Sql = "SELECT SCOPE_IDENTITY()"
            ElseIf Me.DatabaseType = "mysql" Then
                Sql = "SELECT @@IDENTITY"
            ElseIf Me.DatabaseType = "sqlite" Then
                Sql = "SELECT @@IDENTITY"
            Else
                Sql = "SELECT @@IDENTITY"
            End If
            Dim Command As Object = Me.GetCommand(Sql)
            Try
                InsertId = Command.ExecuteScalar
            Catch
                InsertId = Nothing
            End Try
            Return InsertId
        End Function

        Public Function AddDays(ByVal NumDays As String, ByVal DateString As String) As String
            If Me.DatabaseType = "mysql" Then
                Return "ADDDATE(" & DateString & ", " & NumDays & ")"
            Else
                Return "DATEADD(day, " & NumDays & ", " & DateString & ")"
            End If
        End Function

        Public Function AddMonths(ByVal NumMonths As String, ByVal DateString As String) As String
            If Me.DatabaseType = "mysql" Then
                Return "DATE_ADD(" & DateString & ", INTERVAL " & NumMonths & " MONTH)"
            Else
                Return "DATEADD(month, " & NumMonths & ", " & DateString & ")"
            End If
        End Function

        Public Function DiffDays(ByVal Date1 As String, ByVal Date2 As String) As String
            If Me.DatabaseType = "mysql" Then
                Return "DATEDIFF(" & Date1 & ", " & Date2 & ")"
            Else
                Return "DATEDIFF(day, " & Date1 & ", " & Date2 & ")"
            End If
        End Function

        Public Function Timestamp() As String
            If Me.DatabaseType = "mysql" Then
                Return "NOW()"
            Else
                Return "GETDATE()"
            End If
        End Function

        Public Function ToCurrency(ByVal Value As String) As String
            If Me.DatabaseType = "mysql" Then
                Return Value
            Else
                Return "CONVERT(money, " & Value & ")"
            End If
        End Function

        Public Function Limit(ByVal Sql As String, ByVal Num As String) As String
            If Me.DatabaseType = "mysql" Then
                Return Sql & " LIMIT " & Num
            Else
                Dim Start As Integer = Sql.IndexOf("SELECT ") + 7
                Return Sql.Insert(Start, "TOP " & Num)
            End If
        End Function

        Public Function AgnosticizeQuery(ByVal Sql As String) As String
            If Me.DatabaseType = "mysql" Then
                ' Escape characters are different
                Sql = Sql.Replace("[", "`")
                Sql = Sql.Replace("]", "`")
                ' Isnull to ifnull
                Sql = Sql.Replace("ISNULL", "IFNULL")
                ' Top to limit
                If Sql.StartsWith("SELECT TOP ") Then
                    Dim Start As Integer = Sql.IndexOfAny("0123456789")
                    Dim Num As Integer = Sql.Substring(Start, Sql.IndexOf(" ", Start) - Start)
                    Sql = Sql.Replace("SELECT TOP " & Num, "").Trim
                    Sql = "SELECT " & Sql & " LIMIT " & Num
                End If
                ' Concat ... this is  prone to errors and wont' handle multiple... needs work
                If Sql.Contains(" (") Then
                    Dim Start As Integer = Sql.IndexOf(" (") + 2
                    Dim Len As Integer = Sql.IndexOf(")", Start) - Start
                    Dim Text As String = Sql.Substring(Start, Len)
                    ' Make sure it's not a subquery
                    Dim IsConcat As Boolean = True
                    If Text.StartsWith("SELECT ") Then
                        IsConcat = False
                    ElseIf Not Text.Contains(" + ") Then
                        IsConcat = False
                    ElseIf Text.Contains(" = ") Or Text.Contains(" LIKE ") Then
                        IsConcat = False
                    End If
                    If IsConcat Then
                        Sql = Sql.Remove(Start, Len)
                        Sql = Sql.Insert(Start, "CONCAT(" & Text.Replace(" + ", ", ") & ")")
                        ' MsgBox(Sql)
                    End If
                End If
            End If
            Return Sql
        End Function

        Public Function GetRow(ByVal SqlQuery As String, Optional ByVal Type As CommandType = CommandType.Text, Optional ByVal Params As Collection = Nothing) As DataRow
            Dim Row As DataRow = Nothing
            If Me.IsDatabaseInitalized Then
                SqlQuery = Me.AgnosticizeQuery(SqlQuery)
                Dim Table As New DataTable
                Dim Rows As Integer = 0
                Me.OpenConnection()
                Try
                    Dim Reader As Object = Me.GetReader(Me.GetCommand(SqlQuery, Type, Params))
                    Try
                        If Reader.HasRows Then
                            Reader.Read()
                            For i As Integer = 0 To Reader.FieldCount - 1
                                Table.Columns.Add(Reader.GetName(i), Reader.GetFieldType(i))
                            Next
                            Row = Table.NewRow
                            For i As Integer = 0 To Reader.FieldCount - 1
                                Try
                                    Row.Item(i) = Reader(i)
                                Catch ex As Exception
                                    Row.Item(i) = Nothing
                                End Try
                            Next
                            Rows = 1
                        Else
                            Row = Table.NewRow
                            Rows = 0
                        End If
                        Me.LastQuery = New Query(True, SqlQuery)
                        Me.LastQuery.RowsReturned = Rows
                    Catch ex As Exception
                        Me.LastQuery = New Query(False, SqlQuery, ex.ToString)
                        Me.LastQuery.RowsReturned = Rows
                    Finally
                        Reader.Close()
                    End Try
                Catch ex As Exception
                    Me.LastQuery = New Query(False, SqlQuery, ex.ToString)
                End Try
                If Not Me.LeaveConnectionOpen Then
                    Me.CloseConnection()
                End If
            End If
            Return Row
        End Function

        Public Function GetAll(ByVal SqlQuery As String, Optional ByVal Type As CommandType = CommandType.Text, Optional ByVal strTableName As String = "", Optional ByVal Params As Collection = Nothing) As DataTable
            Dim Table As New DataTable
            If Me.IsDatabaseInitalized Then
                SqlQuery = Me.AgnosticizeQuery(SqlQuery)
                Me.OpenConnection()
                Dim Adapter As Object = Me.GetAdapter(SqlQuery, Type, Params)
                Try
                    Adapter.Fill(Table)
                    Me.LastQuery = New Query(True, SqlQuery)
                    Me.LastQuery.RowsReturned = Table.Rows.Count
                Catch ex As Exception
                    Me.LastQuery = New Query(False, SqlQuery, ex.ToString)
                    Me.LastQuery.RowsReturned = 0
                End Try
                If Not Me.LeaveConnectionOpen Then
                    Me.CloseConnection()
                End If
                If strTableName.Length > 0 Then
                    Table.TableName = strTableName
                End If
            End If
            Return Table
        End Function

        Public Function GetOne(ByVal SqlQuery As String, Optional ByVal Type As CommandType = CommandType.Text, Optional ByVal Params As Collection = Nothing) As String
            Dim strReturn As String = Nothing
            If Me.IsDatabaseInitalized Then
                SqlQuery = Me.AgnosticizeQuery(SqlQuery)
                Dim Rows As Integer = 0
                Me.OpenConnection()
                Try
                    Dim Reader As Object = Me.GetReader(Me.GetCommand(SqlQuery, Type, Params))
                    Try
                        If Reader.HasRows Then
                            If Reader.Read() Then
                                strReturn = Reader.Item(0).ToString
                            End If
                            Rows = 1
                        End If
                        Me.LastQuery = New Query(True, SqlQuery)
                        Me.LastQuery.RowsReturned = Rows
                    Catch ex As Exception
                        Me.LastQuery = New Query(False, SqlQuery, ex.ToString)
                        Me.LastQuery.RowsReturned = Rows
                    Finally
                        Reader.Close()
                    End Try
                Catch ex As Exception
                    Me.LastQuery = New Query(False, SqlQuery, ex.ToString)
                End Try
                If Not Me.LeaveConnectionOpen Then
                    Me.CloseConnection()
                End If
            End If
            Return strReturn
        End Function

        Public Function GetOne(ByVal Sql As String, ByVal Type As CommandType, ByVal ParamString As String) As String
            If Me.IsDatabaseInitalized Then
                Dim Params As New Collection
                Dim ParamsSplit As String() = ParamString.Split("|")
                For i As Integer = 0 To ParamsSplit.Length - 1
                    Params.Add(New Param(ParamsSplit(i).Substring(0, ParamsSplit(i).IndexOf("=")), ParamsSplit(i).Substring(ParamsSplit(i).IndexOf("=") + 1)))
                Next
                Return Me.GetOne(Sql, Type, Params)
            Else
                Return Nothing
            End If
        End Function

        Public Function GetRowCount(ByVal SqlQuery As String, Optional ByVal Type As CommandType = CommandType.Text, Optional ByVal Params As Collection = Nothing) As Integer
            If Me.IsDatabaseInitalized Then
                SqlQuery = Me.AgnosticizeQuery(SqlQuery)
                Dim Count As Integer = 0
                Me.OpenConnection()
                Try
                    Dim Reader As Object = Me.GetReader(Me.GetCommand(SqlQuery, Type, Params))
                    Try
                        If Reader.HasRows Then
                            While Reader.Read()
                                Count += 1
                            End While
                        End If
                        Me.LastQuery = New Query(True, SqlQuery)
                    Catch ex As Exception
                        Me.LastQuery = New Query(False, SqlQuery, ex.ToString)
                    Finally
                        Reader.Close()
                    End Try
                Catch ex As Exception
                    Me.LastQuery = New Query(False, SqlQuery, ex.ToString)
                End Try
                If Not Me.LeaveConnectionOpen Then
                    Me.CloseConnection()
                End If
                Return Count
            Else
                Return Nothing
            End If
        End Function

        Public Function Execute(ByVal SqlQuery As String, Optional ByVal Type As CommandType = CommandType.Text, Optional ByVal Params As Collection = Nothing, Optional ByVal Insert As Boolean = False) As Integer
            If Me.IsDatabaseInitalized Then
                SqlQuery = Me.AgnosticizeQuery(SqlQuery)
                Dim Count As Integer = 0
                If SqlQuery.Trim.StartsWith("INSERT ") Then
                    Insert = True
                End If
                Me.OpenConnection()
                Try
                    Dim Reader As Object = Me.GetReader(Me.GetCommand(SqlQuery, Type, Params))
                    Count = Reader.RecordsAffected
                    Reader.Close()
                    Me.LastQuery = New Query(True, SqlQuery)
                    Me.LastQuery.AffectedRows = Count
                    If Insert Then
                        Try
                            Me.LastQuery.InsertId = Me.GetInsertId()
                        Catch
                            Me.LastQuery.InsertId = Nothing
                        End Try
                    End If
                Catch ex As Exception
                    Me.LastQuery = New Query(False, SqlQuery, ex.ToString)
                    Count = 0
                End Try
                If Not Me.LeaveConnectionOpen Then
                    Me.CloseConnection()
                End If
                Return Count
            Else
                Return Nothing
            End If
        End Function

        Public Function InsertAndReturnId(ByVal SqlQuery As String, Optional ByVal Type As CommandType = CommandType.Text, Optional ByVal Params As Collection = Nothing) As Integer
            If Me.IsDatabaseInitalized Then
                Me.Execute(SqlQuery, Type, Params, True)
                Return Me.LastQuery.InsertId
            Else
                Return Nothing
            End If
        End Function

        Public Function RowExists(ByVal strTable As String, ByVal strPK As String, ByVal strValue As String) As Boolean
            Dim str As String = "SELECT " & strPK & " FROM " & strTable & " WHERE " & strPK & "='" & strValue & "'"
            Return Me.RowExists(str)
        End Function

        Public Function RowExists(ByVal strTable As String, ByVal strPK As String, ByVal intValue As Integer) As Boolean
            Dim str As String = "SELECT " & strPK & " FROM " & strTable & " WHERE " & strPK & "=" & intValue
            Return Me.RowExists(str)
        End Function

        Public Function RowExists(ByVal strTable As String, ByVal strValue As String) As Boolean
            Dim str As String = "SELECT id FROM " & strTable & " WHERE id='" & strValue & "'"
            Return Me.RowExists(str)
        End Function

        Public Function RowExists(ByVal strTable As String, ByVal intValue As Integer) As Boolean
            Dim str As String = "SELECT id FROM " & strTable & " WHERE id=" & intValue
            Return Me.RowExists(str)
        End Function

        Public Function RowExists(ByVal strSql As String) As Boolean
            If Me.IsDatabaseInitalized Then
                Try
                    If Me.GetRowCount(strSql) > 0 Then
                        Return True
                    End If
                Catch
                    ' No action
                End Try
            End If
            Return False
        End Function

        'Public Function IsNewer(ByVal strTable As String, ByVal strField As String, ByVal intValue As Integer, ByVal dateLastUpdated As DateTime) As Boolean
        '    Dim str As String = "SELECT date_last_updated FROM " & strTable & " WHERE " & strField & "=" & intValue
        '    Return Me.IsNewer(str, dateLastUpdated)
        'End Function

        Public Function IsNewer(ByVal strTable As String, ByVal strField As String, ByVal Value As Object, ByVal dateLastUpdated As DateTime) As Boolean
            Dim sql As String
            If Value.GetType.ToString = "System.Int16" Or Value.GetType.ToString = "System.Int32" _
            Or Value.GetType.ToString = "System.Int64" Or Value.GetType.ToString = "System.Decimal" _
            Or Value.GetType.ToString = "System.Double" Or Value.GetType.ToString = "Integer" Then
                sql = "SELECT date_last_updated FROM " & strTable & " WHERE " & strField & "=" & Value
            Else
                sql = "SELECT date_last_updated FROM " & strTable & " WHERE " & strField & "='" & Value & "'"
            End If

            Return Me.IsNewer(sql, dateLastUpdated)
        End Function

        Public Function IsNewer(ByVal strTable As String, ByVal Value As Object, ByVal dateLastUpdated As DateTime) As Boolean
            Return Me.IsNewer(strTable, "id", Value, dateLastUpdated)
        End Function

        Public Function IsNewer(ByVal strSql As String, ByVal CompareDate As DateTime) As Boolean
            Dim strDate As String = Me.GetOne(strSql)
            Dim DateInDb As DateTime
            Try
                DateInDb = CType(strDate, DateTime)
            Catch
                Return True
            End Try
            If DateInDb < CompareDate Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Function InsertRow(ByVal strTable As String, ByVal Row As DataRow) As String
            Dim str As String
            Dim i As Integer
            str = "INSERT INTO [" & strTable & "] ("
            For i = 0 To Row.ItemArray.Length - 1
                str &= "[" & Row.Table.Columns(i).ColumnName.ToString & "]"
                str &= ", "
            Next
            str = str.Substring(0, str.Length - 2)
            str &= ")"
            str &= " VALUES ("
            For i = 0 To Row.ItemArray.Length - 1
                str &= Me.Escape(Row.Item(i)) & ", "
            Next
            ' Strip last comma
            str = str.Substring(0, str.Length - 2)
            str &= ")"
            Return Me.InsertAndReturnId(str)
        End Function

        Public Function UpdateRow(ByVal strTable As String, ByVal Where As Object, ByVal Row As DataRow) As Boolean
            Return Me.UpdateRow(strTable, "id", Where, Row)
        End Function

        Public Function UpdateRow(ByVal strTable As String, ByVal strPrimaryKey As String, ByVal PrimaryKeyValue As Object, ByVal Row As DataRow) As Boolean
            Dim str As String
            Dim i As Integer
            str = "UPDATE [" & strTable & "] SET "
            For i = 0 To Row.ItemArray.Length - 1
                str &= "[" & Row.Table.Columns(i).ColumnName.ToString & "]="
                str &= Me.Escape(Row.Item(i))
                str &= ", "
            Next
            str = str.Substring(0, str.Length - 2)
            str &= " WHERE " & strPrimaryKey & "=" & Me.Escape(PrimaryKeyValue)
            If Me.Execute(str) > 0 Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Function CopyRow(ByRef Row As DataRow) As DataRow
            Dim Table As DataTable = Row.Table
            Dim NewRow As DataRow = Table.NewRow
            For i As Integer = 0 To Table.Columns.Count - 1
                NewRow.Item(i) = Row.Item(i)
            Next
            Return NewRow
        End Function

        Public Sub RefreshTable(ByVal SqlQuery As String, ByRef Table As DataTable)
            ' Clear Table
            Table.Clear()
            Table.Rows.Clear()
            Table.Columns.Clear()
            ' Fill table
            Dim Adapter As Object = Me.GetAdapter(SqlQuery)
            Me.OpenConnection()
            Try
                Adapter.Fill(Table)
                Me.LastQuery = New Query(True, SqlQuery)
            Catch ex As Exception
                Me.LastQuery = New Query(False, SqlQuery, ex.ToString)
            End Try
            If Not Me.LeaveConnectionOpen Then
                Me.CloseConnection()
            End If
        End Sub

        Public Sub AddTable(ByRef Dataset As DataSet, ByVal TableName As String, ByVal Query As String)
            Dataset.Tables.Add(TableName)
            Me.RefreshTable(Query, Dataset.Tables(TableName))
        End Sub

        Public Function Escape(ByVal Value As Object) As String
            Try
                If Value Is DBNull.Value Then
                    Return "NULL"
                ElseIf Value Is Nothing Then
                    Return "NULL"
                ElseIf Value.GetType.ToString = "System.Int64" Or Value.GetType.ToString = "System.Int32" Or _
                Value.GetType.ToString = "System.Int16" Or Value.GetType.ToString = "System.Decimal" Or _
                Value.GetType.ToString = "System.Double" Or Value.GetType.ToString = "System.Byte" Then
                    Return Value.ToString
                ElseIf Value.GetType.ToString = "System.Boolean" Then
                    If Value = True Then
                        Return "1"
                    Else
                        Return "0"
                    End If
                ElseIf Value.GetType.ToString = "System.DateTime" Then
                    If Me.DatabaseType = "Access" Then
                        Dim d As DateTime = CType(Value, DateTime)
                        'Return "#" & d.ToString("MM/dd/yyyy HH:mm:ss") & "#"
                        Return "#" & d.ToString & "#"
                    ElseIf Me.DatabaseType = "mysql" Then
                        Dim d As DateTime = CType(Value, DateTime)
                        Return "'" & d.ToString("yyyy-MM-dd HH:mm:ss") & "'"
                    Else
                        Return "'" & Value.ToString.Replace("'", "''") & "'"
                    End If
                Else
                    Return "'" & Value.ToString.Replace("'", "''") & "'"
                End If
            Catch
                Return "''"
            End Try

        End Function

        Public Function FindRow(ByVal Table As DataTable, ByVal RowName As String, ByVal RowValue As String) As DataRow
            For i As Integer = 0 To Table.Rows.Count - 1
                If Table.Rows(i).Item(RowName) = RowValue Then
                    Return Table.Rows(i)
                End If
            Next
            Return Nothing
        End Function

        Public Function IndexExists(ByVal TableName As String, ByVal Name As String) As Boolean
            If Me.DatabaseType = "sqlite" Then
                Dim Table As DataTable = Me.GetAll("SELECT * FROM sqlite_master WHERE name = '" & Name & "' AND tbl_name='" & TableName & "' AND type = 'index'")
                If Table.Rows.Count > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return Nothing
            End If
        End Function

        Public Function ColumnExists(ByVal TableName As String, ByVal Name As String) As Boolean
            If Me.DatabaseType = "sqlite" Then
                Dim Table As DataTable = Me.GetAll("PRAGMA table_info('" & TableName & "')")
                Dim Rows As DataRow() = Table.Select("name='" & Name & "'")
                If Rows.Length > 0 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return Nothing
            End If
        End Function

        Public Function Tables(ByVal Name As String) As DatabaseTable
            Return Me.TableHash(Name)
        End Function

        Public Function Tables() As DatabaseTable()
            Dim Out(Me.TableHash.Count - 1) As DatabaseTable
            Dim i As Integer = 0
            For Each Table As DatabaseTable In Me.TableHash.Values
                Out(i) = Table
                i += 1
            Next
            Return Out
        End Function

        Public Function TableExists(ByVal TableName As String) As Boolean
            If Me.DatabaseType = "sqlite" Then
                Dim Table As DataTable = Me.GetAll("SELECT * FROM sqlite_master WHERE name = '" & TableName & "' AND type = 'table'")
                If Table.Rows.Count = 0 Then
                    Return False
                Else
                    Return True
                End If
            Else
                Return Nothing
            End If
        End Function

        Public Sub UpdateSchema(ByVal FileName As String)
            Dim Contents As String = My.Computer.FileSystem.ReadAllText(FileName)
            Dim Lines As String() = Contents.Split(ControlChars.CrLf)
            Dim TableName As String = ""
            Dim Mode As String = "Fields"
            For Each Line As String In Lines
                Dim Cells As String() = Line.Split(ControlChars.Tab)
                If Cells(0).Trim.Length > 0 Then
                    ' If there is something in the first cell it is the table name
                    TableName = Cells(0).Trim
                    Me.TableHash.Add(TableName, New DatabaseTable(Me, TableName))
                ElseIf Cells.Length > 1 And Line.Contains("**") Then
                    ' Look for **
                    If Cells(1).Contains("INDEXES") Then
                        Mode = "Indexes"
                    ElseIf (Cells(1).Contains("COLUMNS")) Then
                        Mode = "Fields"
                    End If
                ElseIf Mode = "Fields" And Cells.Length >= 3 Then
                    ' Store column values
                    Dim Col As New DatabaseColumn(Cells(1), Cells(2))
                    If Cells.Length >= 4 Then
                        If Cells(3).Contains("PRIMARY") Then
                            Col.IsPrimary = True
                        ElseIf Cells(3).Contains("NOT NULL") Then
                            Col.AllowNull = False
                            If Cells.Length >= 5 Then
                                Col.DefaultValue = Cells(4)
                            ElseIf Col.Type.Contains("INT") Or Col.Type = "NUMERIC" Or Col.Type = "DOUBLE" Then
                                Col.DefaultValue = "0"
                            Else
                                Col.DefaultValue = ""
                            End If
                        End If
                    End If
                    Me.Tables(TableName).Columns.Add(Col.Name, Col)
                ElseIf Mode = "Indexes" And Cells.Length >= 3 Then
                    ' Add indexes
                    Dim Index As New DatabaseIndex(TableName, Cells(1), Cells(2), Cells(3))
                    If Not Me.Tables(TableName).IndexExists(Index.Name) Then
                        Dim Sql As String = "CREATE "
                        If Index.Type = "UNIQUE" Then
                            Sql &= " UNIQUE"
                        End If
                        Sql &= " INDEX " & Index.Name & " ON " & TableName & "(" & Index.Fields & ")"
                        Me.Execute(Sql)
                    End If
                ElseIf Mode = "Fields" And Line.Trim.Length = 0 Then
                    If Me.Tables(TableName).Columns.Values.Count > 0 Then
                        If Me.TableExists(TableName) Then
                            ' Table exists, make sure we have all columns we need
                            For Each Col As DatabaseColumn In Me.Tables(TableName).Columns.Values
                                If Not Me.Tables(TableName).ColumnExists(Col.Name) Then
                                    Dim Sql As String = "ALTER TABLE " & TableName & " ADD COLUMN " & Col.Name & " " & Col.Type
                                    If Not Col.AllowNull Then
                                        Sql &= " NOT NULL"
                                        If Col.DefaultValue <> Nothing Then
                                            Sql &= " DEFAULT '" & Col.DefaultValue & "'"
                                        End If
                                        If Col.Collate.Length > 0 Then
                                            Sql &= " COLLATE " & Col.Collate
                                        End If
                                        If Col.CheckContraint.Length > 0 Then
                                            Sql &= " CHECK(" & Col.CheckContraint & ")"
                                        End If
                                    End If
                                    Me.Execute(Sql)
                                End If
                            Next
                        Else
                            ' Table does not exist, create it
                            Dim Sql As String = "CREATE TABLE " & TableName & " ("
                            Dim Count As Integer = 0
                            For Each Col As DatabaseColumn In Me.Tables(TableName).Columns.Values
                                If Count > 0 Then
                                    Sql &= ", "
                                End If
                                Sql &= Col.Name & " " & Col.Type
                                If Col.IsPrimary Then
                                    Sql &= " PRIMARY KEY"
                                End If
                                If Not Col.AllowNull Then
                                    Sql &= " NOT NULL"
                                    If Col.DefaultValue <> Nothing Then
                                        Sql &= " DEFAULT '" & Col.DefaultValue & "'"
                                    End If
                                End If
                                If Col.Collate.Length > 0 And Not Col.IsPrimary Then
                                    Sql &= " COLLATE " & Col.Collate
                                End If
                                If Col.CheckContraint.Length > 0 Then
                                    Sql &= " CHECK(" & Col.CheckContraint & ")"
                                End If
                                Count += 1
                            Next
                            Sql &= ")"
                            Me.Execute(Sql)
                            ' Clear values
                            Mode = ""
                        End If
                    End If
                End If
            Next
        End Sub

        Public Class DatabaseTable

            Friend db As EasySql
            Public name As String

            Public Columns As New Hashtable

            Public Sub New(ByVal con As EasySql, ByVal tbl As String)
                Me.db = con
                Me.name = tbl
            End Sub

            Public Function GetAll() As DataTable
                Return Me.db.GetAll("SELECT * FROM " & Me.name)
            End Function

            Public Function GetAll(ByVal Where As String) As DataTable
                Return Me.db.GetAll("SELECT * FROM " & Me.name & " WHERE " & Where)
            End Function

            Public Function GetAll(ByVal Fields As String, ByVal Where As String) As DataTable
                Return Me.db.GetAll("SELECT Fields FROM " & Me.name & " WHERE " & Where)
            End Function

            Public Function GetId(ByVal Field As String, ByVal Value As String) As Object
                Return Me.db.GetOne("SELECT id FROM " & Me.name & " WHERE " & Field & "=" & Me.db.Escape(Value))
            End Function

            Public Function GetId(ByVal Value As String) As Object
                Return Me.db.GetOne("SELECT id FROM " & Me.name & " WHERE name=" & Me.db.Escape(Value))
            End Function

            Public Function GetRow(ByVal Id As Integer) As DataRow
                Return Me.db.GetRow("SELECT * FROM " & Me.name & " WHERE id=" & Id)
            End Function

            Public Function GetOne(ByVal Field As String, ByVal Id As Integer) As String
                Return Me.db.GetOne("SELECT " & Field & " FROM " & Me.name & " WHERE id=" & Id)
            End Function

            Public Sub DeleteRow(ByVal Id As Integer)
                Me.db.Execute("DELETE FROM " & Me.name & " WHERE id=" & Id)
            End Sub

            Public Sub DeleteAll(ByVal Id As Integer)
                Me.db.Execute("DELETE FROM " & Me.name)
            End Sub

            Public Sub Insert(ByVal Fields As String, ByVal Values As String)
                Me.db.Execute("INSERT INTO " & Me.name & " (" & Fields & ") VALUES(" & Values & ")")
            End Sub

            Public Sub InsertValue(ByVal Field As String, ByVal Value As String)
                Me.db.Execute("INSERT INTO " & Me.name & " (" & Field & ") VALUES(" & Me.db.Escape(Value) & ")")
            End Sub

            Public Sub InsertRow(ByVal Row As DataRow, Optional ByVal AutoIncrementField As String = "")
                Dim Fields As String = ""
                Dim Values As String = ""
                For Each col As DataColumn In Row.Table.Columns
                    If col.ColumnName <> AutoIncrementField Then
                        If Me.Columns.ContainsKey(col.ColumnName) Then
                            Fields &= col.ColumnName & ", "
                            Values &= Me.db.Escape(Row.Item(col.ColumnName)) & ", "
                        End If
                    End If
                Next
                If Values.Length > 2 Then
                    Fields = Fields.Substring(0, Fields.Length - 2)
                    Values = Values.Substring(0, Values.Length - 2)
                    Dim Sql As String = "INSERT INTO " & Me.name & " (" & Fields
                    Sql &= ") VALUES (" & Values & ")"
                    Me.db.Execute(Sql)
                End If
            End Sub

            Public Sub Update(ByVal Id As Integer, ByVal Values As String)
                Me.db.Execute("UPDATE " & Me.name & " SET " & Values & " WHERE id=" & Id)
            End Sub

            Public Sub UpdateValue(ByVal Id As Integer, ByVal Field As String, ByVal Value As String)
                Me.db.Execute("UPDATE " & Me.name & " SET " & Field & "=" & Me.db.Escape(Value) & " WHERE id=" & Id)
            End Sub

            Public Sub UpdateRow(ByVal pk As String, ByVal Row As DataRow)
                Dim Values As String = ""
                For Each col As DataColumn In Row.Table.Columns
                    If col.ColumnName <> pk Then
                        If Me.Columns.ContainsKey(col.ColumnName) Then
                            Values &= col.ColumnName & "=" & Me.db.Escape(Row.Item(col.ColumnName)) & ", "
                        End If
                    End If
                Next
                If Values.Length > 2 Then
                    Values = Values.Substring(0, Values.Length - 2)
                    Dim Sql As String = "UPDATE " & Me.name & " SET " & Values
                    Sql &= " WHERE " & pk & "=" & Me.db.Escape(Row.Item(pk))
                    Me.db.Execute(Sql)
                End If
            End Sub

            Public Sub InsertOrUpdateRow(ByVal Row As DataRow, Optional ByVal pk As String = "id")
                If Me.RowExists(pk, Row) Then
                    Me.UpdateRow(pk, Row)
                Else
                    Me.InsertRow(Row)
                End If
            End Sub

            Public Sub InsertOrUpdateRowIfNewer(ByVal Row As DataRow, Optional ByVal pk As String = "id", Optional ByVal DateField As String = "date_last_updated")
                If Me.RowExists(pk, Row) Then
                    Dim Current As DateTime = Me.GetOne(DateField, Row.Item(pk))
                    Dim NewRow As DateTime = Row.Item(DateField)
                    If NewRow > Current Then
                        Me.UpdateRow(pk, Row)
                    End If
                Else
                    Me.InsertRow(Row)
                End If
            End Sub

            Public Function IndexExists(ByVal Name As String) As Boolean
                Dim Table As DataTable = Me.db.GetAll("SELECT * FROM sqlite_master WHERE name = '" & Name & "' AND tbl_name='" & Me.name & "' AND type = 'index'")
                If Table.Rows.Count > 0 Then
                    Return True
                Else
                    Return False
                End If
            End Function

            Public Function ColumnExists(ByVal Name As String) As Boolean
                Dim Table As DataTable = Me.db.GetAll("PRAGMA table_info('" & Me.name & "')")
                Dim Rows As DataRow() = Table.Select("name='" & Name & "'")
                If Rows.Length > 0 Then
                    Return True
                Else
                    Return False
                End If
            End Function

            Public Function RowExists(ByVal pk As String, ByVal row As DataRow) As Boolean
                Me.db.GetAll("SELECT " & pk & " FROM " & Me.name & " WHERE " & pk & "=" & Me.db.Escape(row.Item(pk)))
                If Me.db.LastQuery.RowsReturned > 0 Then
                    Return True
                Else
                    Return False
                End If
            End Function

            Public Function UpdateRowIfNewer(ByVal pk As String, ByVal Row As DataRow, Optional ByVal last_updated_row As String = "date_last_updated") As Boolean
                Dim Newer As Boolean = True
                Try
                    Dim Last As DateTime = Me.db.GetOne("SELECT " & last_updated_row & " FROM " & Me.name & " WHERE " & pk & "=" & Me.db.Escape(Row.Item(pk)))
                    If Row.Item(last_updated_row) < Last Then
                        Newer = False
                    End If
                Catch
                    Newer = True
                End Try
                If Newer Then
                    Me.UpdateRow(pk, Row)
                    Return True
                Else
                    Return False
                End If
            End Function

        End Class

        Public Class DatabaseColumn

            Public Name As String
            Public Type As String
            Public Length As Integer = 0
            Public DefaultValue As String = Nothing
            Public IsPrimary As Boolean = False
            Public AllowNull As Boolean = True
            Public CheckContraint As String = ""
            Public Collate As String = "NOCASE"

            Public Sub New(ByVal Name As String, ByVal Type As String, Optional ByVal pk As Boolean = False)
                Me.Name = Name
                Me.Type = Type
                Me.IsPrimary = pk
            End Sub

        End Class

        Public Class DatabaseIndex

            Public Name As String
            Public Type As String
            Public Table As String
            Public Fields As String

            Public Sub New(ByVal Table As String, ByVal Name As String, ByVal Type As String, ByVal Fields As String)
                Me.Name = Name
                Me.Type = Type
                Me.Table = Table
                Me.Fields = Fields
            End Sub

        End Class

    End Class

    Public Class Param

        Public Name As String
        Public Value As Object

        Public Sub New(ByVal Name As String, ByVal Value As Object)
            Me.Name = Name
            Me.Value = Value
        End Sub

    End Class

End Namespace