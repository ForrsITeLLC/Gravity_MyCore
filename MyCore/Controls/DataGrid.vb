Imports MyCore.DataGridTextBoxCombo

Public Class DataGrid


    Public Shared Function NewTextColumnStyle(ByVal strMap As String, ByVal strHeader As String, _
        ByVal intWidth As Integer, Optional ByVal blnReadOnly As Boolean = False) As Windows.Forms.DataGridTextBoxColumn
        Dim col As New Windows.Forms.DataGridTextBoxColumn
        col.MappingName = strMap
        col.HeaderText = strHeader
        col.Width = intWidth
        col.ReadOnly = blnReadOnly
        Return col
    End Function

    Public Shared Function NewComboColumnStyle(ByVal strMap As String, ByVal strHeader As String, ByVal intWidth As Integer, ByVal Options As String()) As DataGridComboBoxColumn
        Dim i As Integer
        Dim col As New DataGridComboBoxColumn
        col.MappingName = strMap
        col.HeaderText = strHeader
        col.Width = intWidth
        For i = 0 To Options.Length - 1
            col.ColumnComboBox.Items.Add(Options(i))
        Next
        Return col
    End Function


End Class