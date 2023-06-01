Public Class cCreditMemo

    Public Database As MyCore.Data.EasySql

    Public ID As Integer = Nothing
    Public MemoDate As Date = Nothing
    Public CustomerNo As String = ""
    Public InvoiceNo As String = ""
    Public Office As Integer = 0
    Public Notes As String = ""
    Public TaxGroupId As Integer = 0
    Public TaxAmount As Double = 0

    Public Sub New(ByVal db As MyCore.Data.EasySql)
        Me.Database = db
    End Sub

    Public Sub Open(ByVal ID As String)

    End Sub

    Public Sub Save()

    End Sub

    Public Sub IncrementNextNumber()
        Me.Database.Execute("UPDATE next_number SET number=number+1 WHERE name='invoice'")
    End Sub

    Public Function GetNextNumber() As Integer
        Return Me.Database.GetOne("SELECT number FROM next_number WHERE name='invoice'")
    End Function

End Class
