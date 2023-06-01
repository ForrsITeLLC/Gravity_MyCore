Option Strict Off
Option Explicit On 

Imports Microsoft.VisualBasic
Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace DataGridTextBoxCombo
    Public Class DataGridComboBoxColumn
        Inherits DataGridTextBoxColumn
        ' use the derived nokeyup combo to avoid tabbing problem
        Public WithEvents ColumnComboBox As NoKeyUpCombo

        Private WithEvents _source As CurrencyManager
        Private _rowNum As Integer
        Private _isEditing As Boolean

        'Fields
        'Constructors
        'Events
        'Methods
        Shared Sub New()
            'Warning: Implementation not found
        End Sub
        Public Sub New()
            MyBase.New()

             _source = Nothing
            _isEditing = False



            ColumnComboBox = New NoKeyUpCombo()
            AddHandler ColumnComboBox.Leave, New EventHandler(AddressOf LeaveComboBox)
            AddHandler ColumnComboBox.SelectionChangeCommitted, New EventHandler(AddressOf ComboStartEditing)

        End Sub
        Protected Overloads Overrides Sub Edit(ByVal source As CurrencyManager, ByVal rowNum As Integer, ByVal bounds As Rectangle, ByVal readOnly1 As Boolean, ByVal instantText As String, ByVal cellIsVisible As Boolean)

           
            MyBase.Edit(source, rowNum, bounds, readOnly1, instantText, cellIsVisible)
            _rowNum = rowNum
            _source = source
            ColumnComboBox.Parent = Me.TextBox.Parent
            ColumnComboBox.Location = Me.TextBox.Location
            ColumnComboBox.Size = New Size(Me.TextBox.Size.Width, ColumnComboBox.Size.Height)
            ColumnComboBox.Text = Me.TextBox.Text
            Me.TextBox.Visible = False
            ColumnComboBox.Visible = True
            ColumnComboBox.BringToFront()
            ColumnComboBox.Focus()

        End Sub
        Protected Overloads Overrides Function Commit(ByVal dataSource As CurrencyManager, ByVal rowNum As Integer) As Boolean

            If _isEditing Then
                _isEditing = False
                SetColumnValueAtRow(dataSource, rowNum, ColumnComboBox.Text)
            End If
            Return True

        End Function
        Private Sub ComboStartEditing(ByVal sender As Object, ByVal e As EventArgs)

            _isEditing = True
            MyBase.ColumnStartedEditing(sender)

        End Sub
        Private Sub LeaveComboBox(ByVal sender As Object, ByVal e As EventArgs)

            If _isEditing Then
                SetColumnValueAtRow(_source, _rowNum, ColumnComboBox.Text)
                _isEditing = False
                Invalidate()
            End If
            ColumnComboBox.Hide()

        End Sub
    End Class
End Namespace
