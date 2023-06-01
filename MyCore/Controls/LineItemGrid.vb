Imports MyCore

Public Class LineItemGrid

    Public Class Column

        Public Event ShowColumn(ByVal Col As Column)
        Public Event HideColumn(ByVal Col As Column)
        Public AllowOverride As Boolean = False

        Public Property DataColName() As String
            Get
                Return Me._DataColName
            End Get
            Set(ByVal value As String)
                Me._DataColName = value
            End Set
        End Property

        Public Property Header() As String
            Get
                Return Me._Header
            End Get
            Set(ByVal value As String)
                Me._Header = value
            End Set
        End Property

        Public Property Format() As String
            Get
                Return Me._Format
            End Get
            Set(ByVal value As String)
                Me._Format = value
            End Set
        End Property

        Public Property Width() As Integer
            Get
                Return Me._Width
            End Get
            Set(ByVal value As Integer)
                Me._Width = value
            End Set
        End Property

        Public Property VisiblePosition() As Integer
            Get
                Return Me._VisiblePosition
            End Get
            Set(ByVal value As Integer)
                Me._VisiblePosition = value
            End Set
        End Property

        Public Property Visible() As Boolean
            Get
                Return Me._Visible
            End Get
            Set(ByVal value As Boolean)
                Me._Visible = value
                If value Then
                    RaiseEvent ShowColumn(Me)
                Else
                    RaiseEvent HideColumn(Me)
                End If
            End Set
        End Property

        Dim _DataColName As String
        Dim _Header As String
        Dim _VisiblePosition As Integer
        Dim _Format As String = ""
        Dim _Width As Integer
        Dim _Visible As Boolean = False
        Public Exists As Boolean = False

        Public Sub New(ByVal name As String, Optional ByVal header As String = "", Optional ByVal Width As Integer = Nothing, Optional ByVal VisiblePosition As Integer = Nothing, Optional ByVal Format As String = Nothing)
            Me._DataColName = name
            If header.Length > 0 Then
                Me._Header = header
            Else
                Me._Header = name
            End If
            If Width <> Nothing Then
                Me._Width = Width
            End If
            If VisiblePosition <> Nothing Then
                Me._VisiblePosition = VisiblePosition
            End If
            If Format <> Nothing Then
                Me._Format = Format
            End If
        End Sub

        Public Sub Hide()
            Me.Visible = False
        End Sub

        Public Sub Show()
            Me.Visible = True
        End Sub

        Public Sub Show(ByVal datacol As String)
            Me.DataColName = datacol
            Me.Visible = True
        End Sub

        Public Sub Show(ByVal datacol As String, ByVal header As String)
            Me.DataColName = datacol
            Me.Header = header
            Me.Visible = True
        End Sub

        Public Sub Show(ByVal Width As Integer)
            Me.Width = Width
            Me.Visible = True
        End Sub

        Public Sub Show(ByVal Header As String, ByVal Width As Integer, ByVal VisiblePosition As Integer, Optional ByVal Format As String = Nothing)
            Me.Header = Header
            Me.Width = Width
            Me.VisiblePosition = VisiblePosition
            If Format <> Nothing Then
                Me._Format = Format
            End If
            Me.Visible = True
        End Sub

        Public Sub Show(ByVal Datacol As String, ByVal Header As String, ByVal Width As Integer, ByVal VisiblePosition As Integer, Optional ByVal Format As String = Nothing)
            Me.DataColName = Datacol
            Me.Header = Header
            Me.Width = Width
            Me.VisiblePosition = VisiblePosition
            If Format <> Nothing Then
                Me._Format = Format
            End If
            Me.Visible = True
        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
    End Class

    Dim Loaded As Boolean = False
    Dim Changes As Boolean = False

    Dim _FormType As Type = Type.NotSet

    Public Enum Type
        NotSet = 0
        PurchaseOrder = 1
        Invoice = 2
        SalesOrder = 3
        RentalOrder = 4
        ServiceOrder = 5
        Agreement = 6
        Quote = 7
        RMA = 8
        CreditMemo = 9
    End Enum

    Public Property FormType() As Type
        Get
            Return Me._FormType
        End Get
        Set(ByVal value As Type)
            Me._FormType = value
        End Set
    End Property

    ' Required Variables
    Public ParentWin As MyCore.Plugins.Host
    Public RateTableId As Integer
    Public IsTaxable As Boolean = True
    Dim _Columns(14) As Column
    Dim _LastCell As Infragistics.Win.UltraWinGrid.UltraGridCell = Nothing
    Dim _OriginalValue As Object = Nothing

    ' Columns
    Public Quantity As New Column("quantity", "Qty", 30, 0)
    Public PartNo As New Column("part_no", "Part No", 80, 1)
    Public Description As New Column("description", "Item Description", 160, 3)
    Public Options As New Column("options", "Options", 100, 5)
    Public SerialNo As New Column("serial_no", "Serial No", 90, 6)
    Public EquipmentID As New Column("equipment_id")
    Public Prime As New Column("prime")
    Public Taxable As New Column("taxable")
    Public TaxStatusID As New Column("tax_status_id", "Tax", 30, 7)
    Public StationID As New Column("station_id", "Stock Room", 60, 8)
    Public UnitPrice As New Column("unit_price", "Price", 50, 9, "c")
    Public Discount As New Column("discount", "Disc %", 30, 10)
    Public NetPrice As New Column("net_price", "Net", 50, 11, "c")
    Public Received As New Column("received", "Rcvd", 30, 12)
    Public ItemTypeID As New Column("item_type_id")

    Public Event UpdateGridTotal()
    Public Event PartNoChanged(ByVal sender As LineItemGrid, ByVal e As PartNoChangedEventArgs)

    Public Class PartNoChangedEventArgs
        Public PartNo As String
        Public ItemType As Integer
        Public RowIndex As Integer
        Public Cancel As Boolean = False
        Public Sub New(ByVal pn As String, ByVal type As Integer, ByVal i As Integer)
            Me.PartNo = pn
            Me.ItemType = type
            Me.RowIndex = i
        End Sub
    End Class

    Public Property LastCell() As Infragistics.Win.UltraWinGrid.UltraGridCell
        Get
            Return Me._LastCell
        End Get
        Set(ByVal value As Infragistics.Win.UltraWinGrid.UltraGridCell)
            Me._LastCell = value
        End Set
    End Property

    Public Property OriginalValue() As Object
        Get
            Return Me._OriginalValue
        End Get
        Set(ByVal value As Object)
            Me._OriginalValue = value
        End Set
    End Property

    Public Property DataSource() As DataTable
        Get
            Return Me.ugrdLineItems.DataSource
        End Get
        Set(ByVal value As DataTable)
            Me.ugrdLineItems.DataSource = value
            ' Mark that none exist
            For Each col As Column In Me._Columns
                col.Exists = False
            Next
            If value IsNot Nothing Then
                For Each Col As DataColumn In Me.DataSource.Columns
                    Try
                        Me.Columns(Col.ColumnName).Exists = True
                    Catch
                        Me.GridColumns(Col.ColumnName).Hidden = True
                    End Try
                Next
                Me.Grid.DisplayLayout.Bands(0).AddNew()
                Me.Grid.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ActivateCell, False, False)
                ' Need to format the rows
                For i As Integer = 0 To value.Rows.Count - 1
                    If value.Rows(i).Item(Me.PartNo.DataColName).ToString.Length > 0 Then
                        Dim pn As New cItemMaster(Me.ParentWin.Database)
                        Try
                            pn.Open(value.Rows(i).Item(Me.PartNo.DataColName))
                            Me.FormatRowByPartNumber(pn, i)
                        Catch
                            Me.NoPartNo(i)
                        End Try
                    Else
                        Me.NoPartNo(i)
                    End If
                Next
            End If
        End Set
    End Property

    Public Function Columns(ByVal DataColName As String) As Column
        For Each Col As Column In Me._Columns
            If Col.DataColName = DataColName Then
                Return Col
            End If
        Next
        Return Nothing
    End Function

    Public ReadOnly Property Grid() As Infragistics.Win.UltraWinGrid.UltraGrid
        Get
            Return Me.ugrdLineItems
        End Get
    End Property

    Public ReadOnly Property Rows() As Infragistics.Win.UltraWinGrid.RowsCollection
        Get
            Return Me.ugrdLineItems.Rows
        End Get
    End Property

    Public Function GetTotal(Optional ByVal Net As Boolean = False) As Double
        Dim Total As Double = 0
        Dim Price As String
        If Net Then
            Price = Me.NetPrice.DataColName
        Else
            Price = Me.UnitPrice.DataColName
        End If
        For Each Row As Infragistics.Win.UltraWinGrid.UltraGridRow In Me.Rows
            If Not Row.Hidden Then
                If Microsoft.VisualBasic.IsNumeric(Row.Cells(Price).Value) Then
                    Dim Amt As Double = 0
                    Try
                        Amt = Row.Cells(Me.Quantity.DataColName).Value * Row.Cells(Price).Value
                    Catch ex As Exception
                        Amt = Row.Cells(Price).Value
                    End Try
                    Total += Amt
                End If
            End If
        Next
        Return Total
    End Function

    Public Sub New()
        InitializeComponent()
        Me.Loaded = True
        ' Add to hash table
        Me._Columns(0) = Me.Quantity
        Me._Columns(1) = Me.PartNo
        Me._Columns(2) = Me.Description
        Me._Columns(3) = Me.SerialNo
        Me._Columns(4) = Me.EquipmentID
        Me._Columns(5) = Me.ItemTypeID
        Me._Columns(6) = Me.Prime
        Me._Columns(7) = Me.Taxable
        Me._Columns(8) = Me.TaxStatusID
        Me._Columns(9) = Me.StationID
        Me._Columns(10) = Me.UnitPrice
        Me._Columns(11) = Me.Discount
        Me._Columns(12) = Me.NetPrice
        Me._Columns(13) = Me.Options
        Me._Columns(14) = Me.Received
        ' On show
        AddHandler Me.Description.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.Discount.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.EquipmentID.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.ItemTypeID.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.NetPrice.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.Options.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.PartNo.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.Prime.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.Received.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.SerialNo.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.StationID.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.UnitPrice.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.Taxable.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.Quantity.ShowColumn, AddressOf Me.ColumnShown
        AddHandler Me.TaxStatusID.ShowColumn, AddressOf Me.ColumnShown
        ' On hide
        AddHandler Me.Description.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.Discount.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.EquipmentID.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.ItemTypeID.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.NetPrice.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.Options.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.PartNo.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.Prime.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.Received.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.SerialNo.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.StationID.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.UnitPrice.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.Taxable.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.Quantity.HideColumn, AddressOf Me.ColumnHidden
        AddHandler Me.TaxStatusID.HideColumn, AddressOf Me.ColumnHidden
        ' Other functions
        AddHandler Me.Grid.ClickCellButton, AddressOf Me.ClickCellButton
        AddHandler Me.Grid.KeyDown, AddressOf Me.KeyboardNavigation
    End Sub

    Private Sub ColumnShown(ByVal Col As Column)
        Me.Grid.DisplayLayout.Bands(0).Columns(Col.DataColName).Header.Caption = Col.Header
        Me.Grid.DisplayLayout.Bands(0).Columns(Col.DataColName).Width = Col.Width
        Me.Grid.DisplayLayout.Bands(0).Columns(Col.DataColName).Header.VisiblePosition = Col.VisiblePosition
        Me.Grid.DisplayLayout.Bands(0).Columns(Col.DataColName).Format = Col.Format
        ' More adjustments
        If Col.DataColName = Me.PartNo.DataColName Then
            If Me.Grid.DisplayLayout.Bands(0).Columns.Exists("find") Then
                Me.Grid.DisplayLayout.Bands(0).Columns("find").Hidden = False
            Else
                Me.Grid.DisplayLayout.Bands(0).Columns.Add("find", "")
            End If
            Me.Grid.DisplayLayout.Bands(0).Columns("find").Header.VisiblePosition = Me.PartNo.VisiblePosition + 1
            Me.Grid.DisplayLayout.Bands(0).Columns("find").CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly
            Me.Grid.DisplayLayout.Bands(0).Columns("find").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button
            Me.Grid.DisplayLayout.Bands(0).Columns("find").CellButtonAppearance.Image = My.Resources.Search
            Me.Grid.DisplayLayout.Bands(0).Columns("find").CellButtonAppearance.ImageHAlign = Infragistics.Win.HAlign.Center
            Me.Grid.DisplayLayout.Bands(0).Columns("find").ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
            Me.Grid.DisplayLayout.Bands(0).Columns("find").MaxWidth = 24
            Me.Grid.DisplayLayout.Bands(0).Columns("find").Hidden = False
        ElseIf Col.DataColName = Me.TaxStatusID.DataColName Then
            If Not Me.Grid.DisplayLayout.ValueLists.Exists("Tax Status") Then
                Me.Grid.DisplayLayout.ValueLists.Add("Tax Status")
                Dim Table As DataTable = Me.ParentWin.Database.GetAll("SELECT id, code, taxable FROM tax_status ORDER BY code")
                For Each Row As DataRow In Table.Rows
                    Dim Item As New Infragistics.Win.ValueListItem
                    Item.DisplayText = Row.Item("code")
                    Item.DataValue = Row.Item("id")
                    Item.Tag = Row.Item("id")
                    Me.Grid.DisplayLayout.ValueLists("Tax Status").ValueListItems.Add(Item)
                Next
            End If
            Me.Grid.DisplayLayout.Bands(0).Columns(Me.TaxStatusID.DataColName).Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            Me.Grid.DisplayLayout.Bands(0).Columns(Me.TaxStatusID.DataColName).ValueList = Me.Grid.DisplayLayout.ValueLists("Tax Status")
        ElseIf Col.DataColName = Me.StationID.DataColName Then
            If Not Me.Grid.DisplayLayout.ValueLists.Exists("Stock Rooms") Then
                Me.Grid.DisplayLayout.ValueLists.Add("Stock Rooms")
                Dim Table As DataTable = Me.ParentWin.Database.GetAll("SELECT id, name FROM station ORDER BY name")
                For Each Row As DataRow In Table.Rows
                    Dim Item As New Infragistics.Win.ValueListItem
                    Item.DisplayText = Row.Item("name")
                    Item.DataValue = Row.Item("id")
                    Item.Tag = Row.Item("id")
                    Me.Grid.DisplayLayout.ValueLists("Stock Rooms").ValueListItems.Add(Item)
                Next
            End If
            Me.Grid.DisplayLayout.Bands(0).Columns(Me.StationID.DataColName).Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            Me.Grid.DisplayLayout.Bands(0).Columns(Me.StationID.DataColName).ValueList = Me.Grid.DisplayLayout.ValueLists("Stock Rooms")
        End If
        ' Show it
        Me.Grid.DisplayLayout.Bands(0).Columns(Col.DataColName).Hidden = False
    End Sub

    Private Sub ColumnHidden(ByVal Col As Column)
        ' More adjustments
        If Col.DataColName = Me.PartNo.DataColName Then
            If Me.Grid.DisplayLayout.Bands(0).Columns.Exists("find") Then
                Me.Grid.DisplayLayout.Bands(0).Columns("find").Hidden = True
            End If
        End If
        ' Hide it
        Me.Grid.DisplayLayout.Bands(0).Columns(Col.DataColName).Hidden = True
    End Sub

    Public Function GridColumns(ByVal Name As String) As Infragistics.Win.UltraWinGrid.UltraGridColumn
        Return Me.Grid.DisplayLayout.Bands(0).Columns(Name)
    End Function

    Public Sub DisableEdit(ByVal Col As Column, Optional ByVal Index As Integer = -1)
        If Me.Columns(Col.DataColName).Exists Then
            If Index < 0 Then
                If Me.LastCell IsNot Nothing Then
                    Index = Me.LastCell.Row.Index
                End If
            End If
            Try
                Me.Grid.Rows(Index).Cells(Col.DataColName).Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled
                Me.Grid.Rows(Index).Cells(Col.DataColName).Hidden = True
            Catch
                ' Nothing
            End Try
        End If
    End Sub

    Public Sub EnableEdit(ByVal Col As Column, Optional ByVal Index As Integer = -1)
        If Me.Columns(Col.DataColName).Exists Then
            If Index < 0 Then
                If Me.LastCell IsNot Nothing Then
                    Index = Me.LastCell.Row.Index
                End If
            End If
            Try
                Me.Grid.Rows(Index).Cells(Col.DataColName).Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
                Me.Grid.Rows(Index).Cells(Col.DataColName).Hidden = False
            Catch
                ' Nothing
            End Try
        End If
    End Sub


    Public Sub HideCol(ByVal Col As Column, Optional ByVal Index As Integer = -1)
        If Me.Columns(Col.DataColName).Exists Then
            If Index < 0 Then
                If Me.LastCell IsNot Nothing Then
                    Index = Me.LastCell.Row.Index
                End If
            End If
            Try
                Me.Grid.Rows(Index).Cells(Col.DataColName).Hidden = True
            Catch
                ' Nothing
            End Try
        End If
    End Sub

    Public Sub ShowCol(ByVal Col As Column, Optional ByVal Index As Integer = -1)
        If Me.Columns(Col.DataColName).Exists Then
            If Index < 0 Then
                If Me.LastCell IsNot Nothing Then
                    Index = Me.LastCell.Row.Index
                End If
            End If
            Try
                Me.Grid.Rows(Index).Cells(Col.DataColName).Hidden = False
            Catch
                ' Nothing
            End Try
        End If
    End Sub

    Public Sub SetValue(ByVal ColName As String, ByVal Value As String, Optional ByVal Index As Integer = -1)
        If Me.Columns(ColName).Exists Then
            If Index < 0 Then
                If Me.LastCell IsNot Nothing Then
                    Index = Me.LastCell.Row.Index
                End If
            End If
            Try
                ' Fix invalid values
                If ColName = Me.TaxStatusID.DataColName Then
                    If Value = 0 Then
                        Value = 1
                    End If
                End If
                ' Set value
                Me.ugrdLineItems.Rows(Index).Cells(ColName).Value = Value
            Catch
                ' Nothing
            End Try
            Me.ugrdLineItems.UpdateData()
        End If
    End Sub

    Public Function GetValue(ByVal ColName As String, Optional ByVal DefaultVal As Object = "", Optional ByVal Index As Integer = -1) As Object
        If Me.Columns(ColName).Exists Then
            If Index < 0 Then
                If Me.LastCell IsNot Nothing Then
                    Index = Me.LastCell.Row.Index
                End If
            End If
            Dim Out As Object = Nothing
            Try
                Out = Me.Grid.Rows(Index).Cells(ColName).Value
            Catch ex As Exception
                Out = Nothing
            End Try
            If Out Is DBNull.Value Then
                Return DefaultVal
            ElseIf Out = Nothing Then
                Return DefaultVal
            Else
                Return Out
            End If
        Else
            Return DefaultVal
        End If
    End Function

    Private Function GetDiscountPercent() As Double
        If Me.Discount.Exists Then
            Dim List As Double = Me.GetValue(Me.UnitPrice.DataColName, 0)
            Dim Cost As Double = Me.GetValue(Me.NetPrice.DataColName, 0)
            If List <> 0 Then
                Return 100 - (Math.Round(Cost / List, 3) * 100)
            Else
                Return 0
            End If
        Else
            Return 0
        End If
    End Function

    Private Function GetNetPrice() As Double
        If Me.NetPrice.Exists Then
            Dim List As Double = Me.GetValue(Me.UnitPrice.DataColName, 0)
            Dim Discount As Double = Me.GetValue(Me.Discount.DataColName, 0)
            If Discount < 0 Then
                Return List
            ElseIf Discount > 100 Then
                Return 0
            Else
                Return Math.Round(List * (1 - (Discount / 100)), 2)
            End If
        Else
            Return 0
        End If
    End Function

    Private Sub ugrdLineItems_AfterCellActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugrdLineItems.AfterCellActivate
        Me.LastCell = Me.Grid.ActiveCell
        Me.OriginalValue = Me.Grid.ActiveCell.Value
    End Sub

    Private Sub ugrdLineItems_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugrdLineItems.AfterExitEditMode
        Dim Original As String = ""
        If Me.OriginalValue IsNot DBNull.Value Then
            Original = Me.OriginalValue
        End If
        Dim NewVal As String = ""
        If Me.LastCell.Value IsNot DBNull.Value Then
            NewVal = Me.LastCell.Value
        End If
        If NewVal <> Original Then
            Me.CellChange()
        End If
    End Sub

    Private Sub ugrdLineItems_CellListSelect(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugrdLineItems.CellListSelect
        If e.Cell.Column.Key = Me.SerialNo.DataColName Then
            ' == When serial number selected
            ' auto-select the stock room
            Dim List As Infragistics.Win.ValueList = e.Cell.ValueList
            Dim Tag As String = List.ValueListItems(e.Cell.ValueList.SelectedItemIndex).Tag.ToString
            Dim Station As Integer = Tag.Substring(Tag.IndexOf("|") + 1)
            Dim EquipId As Integer = Tag.Substring(0, Tag.IndexOf("|"))
            Me.SetValue(Me.StationID.DataColName, Station, e.Cell.Row.Index)
            Me.SetValue(Me.EquipmentID.DataColName, EquipId, e.Cell.Row.Index)
        ElseIf e.Cell.Column.Key = Me.StationID.DataColName Then
            ' == When a stock room is selected
            Dim index As Integer = e.Cell.ValueListResolved.SelectedItemIndex
            Dim List As Infragistics.Win.ValueList = Me.ugrdLineItems.DisplayLayout.ValueLists("Stock Rooms")
            Dim Item As Infragistics.Win.ValueListItem = List.ValueListItems(index)
            Dim StationId As Integer = Item.DataValue
            ' If this is prime
            If Not e.Cell.Row.Cells(Me.SerialNo.DataColName).Hidden Then
                If Not Me.SerialNo.AllowOverride Then
                    ' They've changed the location, so clear the serial
                    e.Cell.Row.Cells(Me.SerialNo.DataColName).Value = ""
                End If
                ' Get serial number list for this row
                Dim SerialList As Infragistics.Win.ValueList = e.Cell.Row.Cells(Me.SerialNo.DataColName).ValueList
                ' Loop through serials to see if any are at this stock room, if so select it
                For i As Integer = 0 To SerialList.ValueListItems.Count - 1
                    Dim Tag As String = SerialList.ValueListItems(i).Tag.ToString
                    Dim Station As Integer = Tag.Substring(Tag.IndexOf("|") + 1)
                    Dim EquipId As Integer = Tag.Substring(0, Tag.IndexOf("|"))
                    If Station = StationId Then
                        Me.SetValue(Me.SerialNo.DataColName, SerialList.ValueListItems(i).DataValue, e.Cell.Row.Index)
                        Me.SetValue(Me.EquipmentID.DataColName, EquipId, e.Cell.Row.Index)
                        Exit For
                    End If
                Next
            End If
        ElseIf e.Cell.Column.Key = Me.TaxStatusID.DataColName Then
            If Me.GetValue(Me.TaxStatusID.DataColName, 1, e.Cell.Row.Index) = 1 Then
                Me.SetValue(Me.Taxable.DataColName, True)
            Else
                Me.SetValue(Me.Taxable.DataColName, False)
            End If
            Me.UpdateTotal()
        End If
    End Sub

    Public Sub DeleteRow(ByVal i As Integer)
        Me.SetValue(Me.Quantity.DataColName, 0, i)
        Me.Rows(i).Hidden = True
    End Sub

    Private Sub ugrdLineItems_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles ugrdLineItems.BeforeRowsDeleted
        For Each Row As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows
            Me.DeleteRow(Row.Index)
        Next
        e.Cancel = True
    End Sub

    Private Sub ugrdLineItems_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugrdLineItems.CellChange
        If Me.Loaded Then
            Me.Changes = True
        End If
    End Sub

    Public Sub NoPartNo(ByVal Index As Integer)
        Me.HideCol(Me.Quantity, Index)
        Me.HideCol(Me.UnitPrice, Index)
        Me.HideCol(Me.Discount, Index)
        Me.HideCol(Me.NetPrice, Index)
        Me.HideCol(Me.SerialNo, Index)
        Me.HideCol(Me.StationID, Index)
        Me.HideCol(Me.TaxStatusID, Index)
        Me.SetValue(Me.PartNo.DataColName, "", Index)
        Me.SetValue(Me.UnitPrice.DataColName, 0, Index)
        Me.SetValue(Me.Quantity.DataColName, 1, Index)
    End Sub

    Public Sub SetPartNo(ByVal Part As cItemMaster, ByVal Index As Integer)
        Me.Loaded = False
        If Part IsNot Nothing Then
            If 1 = 1 Then
                ' If index is less than one, select current
                If Index < 0 Then
                    If Me.LastCell IsNot Nothing Then
                        Index = Me.LastCell.Row.Index
                    End If
                    ' If it's still less than 0, set it as the last row in the grid
                    If Index < 0 Then
                        Index = Me.Grid.Rows.Count - 1
                    End If
                End If
                ' Only bother if it is different than current
                Dim Row As Infragistics.Win.UltraWinGrid.UltraGridRow = Me.ugrdLineItems.Rows(Index)
                RaiseEvent PartNoChanged(Me, New PartNoChangedEventArgs(Part.PartNo, Part.ItemType, Index))
                ' Set part number value
                Me.SetValue(Me.PartNo.DataColName, Part.PartNo, Index)
                ' Show all
                Me.ShowCol(Me.Quantity, Index)
                Me.ShowCol(Me.UnitPrice, Index)
                Me.ShowCol(Me.Discount, Index)
                Me.ShowCol(Me.NetPrice, Index)
                Me.ShowCol(Me.SerialNo, Index)
                Me.ShowCol(Me.StationID, Index)
                Me.ShowCol(Me.TaxStatusID, Index)
                ' Form description
                Dim Desc As String = "%description%"
                If Me.FormType = Type.PurchaseOrder Or Me.FormType = Type.RMA Then
                    Desc = Me.ParentWin.SettingsGlobal.GetValue("Part No Description on PO", "%description%")
                ElseIf Me.FormType = Type.Invoice Or Me.FormType = Type.RMA Then
                    Desc = Me.ParentWin.SettingsGlobal.GetValue("Part No Description on Invoice", "%description%")
                Else
                    Desc = Me.ParentWin.SettingsGlobal.GetValue("Part No Description on Invoice", "%description%")
                End If
                Desc = Desc.Replace("%description%", Part.OrderDescription).Replace("%vendor_part_no%", Part.VendorPartNo)
                Desc = Desc.Replace("%notes%", Part.Notes)
                Desc = Desc.Replace("%manufacturer%", Part.Manufacturer).Replace("%model%", Part.Model)
                Me.SetValue(Me.Description.DataColName, Desc, Index)
                ' Set values that are always there
                Me.SetValue(Me.UnitPrice.DataColName, Part.GetPrice(Me.RateTableId), Index)
                Me.SetValue(Me.NetPrice.DataColName, Part.Cost, Index)
                ' If price is not set, set it at zero
                Me.SetValue(Me.UnitPrice.DataColName, Me.GetValue(Me.UnitPrice.DataColName, 0, Index), Index)
                Me.SetValue(Me.NetPrice.DataColName, Me.GetValue(Me.NetPrice.DataColName, 0, Index), Index)
                ' Discount
                Me.SetValue(Me.Discount.DataColName, Me.GetDiscountPercent, Index)
                ' Quantity
                If Me.GetValue(Me.Quantity.DataColName, Nothing, Index) = Nothing Then
                    Me.SetValue(Me.Quantity.DataColName, 1, Index)
                End If
                ' Others
                Me.SetValue(Me.Received.DataColName, 0, Index)
                Me.SetValue(Me.SerialNo.DataColName, "", Index)
                Me.SetValue(Me.EquipmentID.DataColName, 0, Index)
                Me.SetValue(Me.StationID.DataColName, 0, Index)
                ' Set tabxable
                If Me.IsTaxable Then
                    Me.SetValue(Me.TaxStatusID.DataColName, Part.TaxStatus, Index)
                    Me.SetValue(Me.Taxable.DataColName, Part.Taxable, Index)
                Else
                    Me.SetValue(Me.TaxStatusID.DataColName, 2, Index)
                    Me.SetValue(Me.Taxable.DataColName, False, Index)
                End If
                ' Adjust variable columns
                Me.FormatRowByPartNumber(Part, Index)
            End If
        Else
            Me.NoPartNo(Index)
        End If
        Me.UpdateTotal()
        Me.Loaded = True
    End Sub

    Private Sub FormatRowByPartNumber(ByVal Part As cItemMaster, ByVal Index As Integer)
        Me.SetValue(Me.ItemTypeID.DataColName, Part.ItemType, Index)
        Me.SetValue(Me.Prime.DataColName, Part.Prime, Index)
        If Part.ItemType > 1 Then
            Me.EnableEdit(Me.Quantity, Index)
            Me.DisableEdit(Me.SerialNo, Index)
            Me.DisableEdit(Me.StationID, Index)
        ElseIf Part.Prime And Me.SerialNo.Exists Then
            Me.DisableEdit(Me.Quantity, Index)
            Me.EnableEdit(Me.SerialNo, Index)
            Me.EnableEdit(Me.StationID, Index)
            If Me.SerialNo.Exists Then
                ' Get Serial numbers
                Dim Serials As DataTable = Part.SerialNumbersInStock
                ' Make value list
                Dim List As New Infragistics.Win.ValueList
                For Each Serial As DataRow In Serials.Rows
                    Dim Item As New Infragistics.Win.ValueListItem
                    Item.DataValue = Serial.Item("serial_no")
                    Item.DisplayText = Serial.Item("serial_no")
                    Item.Tag = Serial.Item("id") & "|" & Serial.Item("station_id")
                    List.ValueListItems.Add(Item)
                Next
                ' Format serial no field
                If Me.SerialNo.AllowOverride Then
                    Me.Rows(Index).Cells(Me.SerialNo.DataColName).Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDown
                Else
                    Me.Rows(Index).Cells(Me.SerialNo.DataColName).Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
                End If
                Me.Rows(Index).Cells(Me.SerialNo.DataColName).ValueList = List
            End If
        Else
            Me.EnableEdit(Me.Quantity, Index)
            Me.DisableEdit(Me.SerialNo, Index)
            Me.EnableEdit(Me.StationID, Index)
        End If
    End Sub

    Private Function MakeNumeric(ByVal val As String) As Double
        If Not Microsoft.VisualBasic.IsNumeric(val) Then
            Return 0
        Else
            Return val
        End If
    End Function

    Private Sub CellChange()
        If Me.Loaded Then
            If Not Me.LastCell Is Nothing Then
                ' Enforce valid input for certain fields (help prevent crazy null crap)
                Me.SetValue(Me.Quantity.DataColName, Me.MakeNumeric(Me.GetValue(Me.Quantity.DataColName, 0)))
                Me.SetValue(Me.UnitPrice.DataColName, Me.MakeNumeric(Me.GetValue(Me.UnitPrice.DataColName, 0)))
                ' Handle
                If Me.LastCell.Column.Key = Me.Quantity.DataColName Then
                    Me.ugrdLineItems.UpdateData()
                    Me.UpdateTotal()
                ElseIf Me.LastCell.Column.Key = Me.PartNo.DataColName Then
                    Dim PartNo As cItemMaster = Nothing
                    If Me.LastCell.Text.Trim.Length > 0 Then
                        Me.Loaded = False
                        PartNo = New cItemMaster(Me.ParentWin.Database)
                        Try
                            PartNo.Open(Me.LastCell.Value)
                        Catch ex As Exception
                            Dim Ask As New MyCore.Gravity.AskBox(Me.LastCell.Value & " was not found in the database. Do you want to create that part number?", "Part Number Not Found")
                            Dim Value As String = Me.LastCell.Text
                            If Ask.ButtonPress = MyCore.Gravity.Response.Yes Then
                                Dim Win As MyCore.Plugins.Window = Me.ParentWin.CallMethod(Me, Plugins.MethodType.NewItem, "Part No", Value, "Part No")
                                AddHandler Win.OnEvent, AddressOf Me.NewPartNoCreated
                                Win.Open()
                                Exit Sub
                            Else
                                PartNo = Nothing
                            End If
                        End Try
                    End If
                    Me.SetPartNo(PartNo, Me.LastCell.Row.Index)
                ElseIf Me.LastCell.Column.Key = Me.UnitPrice.DataColName Then
                    If Me.Discount.Exists And Me.NetPrice.Exists Then
                        Me.Loaded = False
                        Me.SetValue(Me.NetPrice.DataColName, Me.GetNetPrice)
                        Me.UpdateTotal()
                        Me.Loaded = True
                    End If
                ElseIf Me.LastCell.Column.Key = Me.Discount.DataColName Then
                    If Me.GetValue(Me.UnitPrice.DataColName).ToString.Length > 0 Then
                        Me.Loaded = False
                        Me.SetValue(Me.NetPrice.DataColName, Me.GetNetPrice)
                        Me.UpdateTotal()
                        Me.Loaded = True
                    End If
                ElseIf Me.LastCell.Column.Key = Me.NetPrice.DataColName Then
                    Me.Loaded = False
                    Me.SetValue(Me.Discount.DataColName, Me.GetDiscountPercent)
                    Me.Loaded = True
                    Me.UpdateTotal()
                    Me.Loaded = True
                ElseIf Me.LastCell.Column.Key = Me.TaxStatusID.DataColName Then
                    Me.Loaded = False
                    If Me.IsTaxable Then
                        If Me.GetValue(Me.TaxStatusID.DataColName, 0) > 0 Then
                            Dim Taxable As Boolean = Me.ParentWin.Database.GetOne("SELECT taxable FROM tax_status WHERE id=" & Me.LastCell.Value)
                            Me.SetValue(Me.Taxable.DataColName, Taxable)
                        Else
                            Me.SetValue(Me.Taxable.DataColName, True)
                            Me.SetValue(Me.TaxStatusID.DataColName, 1)
                        End If
                    Else
                        Me.SetValue(Me.Taxable.DataColName, False)
                        Me.SetValue(Me.TaxStatusID.DataColName, 2)
                    End If
                    Me.UpdateTotal()
                    Me.Loaded = True
                ElseIf Me.LastCell.Column.Key = Me.Description.DataColName Then
                    If Me.GetValue(Me.PartNo.DataColName).ToString.Length = 0 Then
                        Me.SetPartNo(Nothing, Me.LastCell.Row.Index)
                    End If
                End If
                Me.LastCell = Nothing
            End If
        End If
    End Sub

    Private Sub UpdateTotal()
        RaiseEvent UpdateGridTotal()
    End Sub

    Public Sub HideAll()
        ' Hide all
        For Each Col As Column In Me._Columns
            If Col.Exists Then
                Col.Visible = False
            End If
        Next
    End Sub

    Private Sub NewPartNoCreated(ByVal Win As MyCore.Plugins.Window, ByVal e As MyCore.Plugins.Window.EventInfo)
        Dim Index As Integer
        Try
            Index = Me.LastCell.Row.Index
        Catch ex As Exception
            Index = Me.Rows.Count - 1
        End Try
        Me.SetPartNo(e.Item1, Index)
    End Sub

    Private Sub ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)
        Me.Loaded = False
        Me.LastCell = e.Cell
        If e.Cell.Row.IsAddRow Then
            Me.Grid.DisplayLayout.Bands(0).AddNew()
            'Me.Grid.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ActivateCell, False, False)
        End If
        If e.Cell.Column.Key = "find" Then
            Dim Win As MyCore.Plugins.Window = Me.ParentWin.CallMethod(Me, Plugins.MethodType.Lookup, "Part No")
            AddHandler Win.OnEvent, AddressOf SelectPartNo
            Me.ParentWin.OpenWindow(Win)
        End If
        Me.Loaded = True
    End Sub

    Private Sub SelectPartNo(ByVal sender As MyCore.Plugins.Window, ByVal e As MyCore.Plugins.Window.EventInfo)
        Dim Index As Integer
        Try
            Index = Me.LastCell.Row.Index
        Catch ex As Exception
            Index = Me.Rows.Count - 1
        End Try
        Dim PartNo As New cItemMaster(Me.ParentWin.Database)
        PartNo.Open(e.Item1)
        Me.SetPartNo(PartNo, Index)
    End Sub

    Private Sub KeyboardNavigation(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyValue = Keys.Up Then
            Grid.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
            Grid.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.AboveCell, False, False)
            e.Handled = True
            Grid.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
        ElseIf e.KeyValue = Keys.Down Then
            Grid.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ExitEditMode, False, False)
            Grid.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.BelowCell, False, False)
            e.Handled = True
            Grid.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
        ElseIf Grid.ActiveRow Is Grid.GetRow(Infragistics.Win.UltraWinGrid.ChildRow.Last) Then
            If Not Grid.ActiveCell Is Nothing Then
                If Grid.ActiveCell.Column.Index = Grid.DisplayLayout.Bands(0).Columns.Count - 1 Then
                    If e.KeyCode = Keys.Tab Then
                        Grid.DisplayLayout.Bands(0).AddNew()
                        Grid.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ActivateCell, False, False)
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ugrdLineItems_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles ugrdLineItems.DoubleClickRow
        'If e.Row.Cells("quantity").Hidden Then
        '    Me.ParentWin.Ask.ShowDialog("Duplicate this line item?")
        '    If Me.ParentWin.Ask.ButtonPress = MyCore.Gravity.Response.Yes Then
        '        Me.ParentWin.Input.ShowDialog("How many of this item would you like to add?", "Duplicate", "1")
        '        If Me.ParentWin.Input.ButtonPress = MyCore.Gravity.Response.OK And Microsoft.VisualBasic.IsNumeric(Me.ParentWin.Input.Text) Then
        '            Dim num As Integer = Me.ParentWin.Input.Text
        '            If num > 1 Then
        '                num -= 1
        '                For i As Integer = 0 To num - 1
        '                    ' Add To DataSource
        '                    Dim r As DataRow = Me.ServiceOrder.LineItems.NewRow
        '                    r.Item("part_no") = e.Row.Cells("part_no").Value
        '                    r.Item("quantity") = 1
        '                    r.Item("price") = e.Row.Cells("price").Value
        '                    r.Item("tax_status_id") = e.Row.Cells("tax_status_id").Value
        '                    r.Item("taxable") = e.Row.Cells("taxable").Value
        '                    r.Item("description") = e.Row.Cells("description").Value
        '                    Me.ServiceOrder.LineItems.Rows.Add(r)
        '                    ' Format
        '                    Dim NewRow As Infragistics.Win.UltraWinGrid.UltraGridRow = Me.ugrdLineItems.Rows(Me.ugrdLineItems.Rows.Count - 1)
        '                    NewRow.Cells("serial_no").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
        '                    NewRow.Cells("serial_no").ValueList = e.Row.Cells("serial_no").ValueList
        '                    NewRow.Cells("quantity").Hidden = True
        '                Next
        '            End If
        '            Me.UpdatePricing()
        '        End If
        '    End If
        'End If
    End Sub

    Private Sub cmnuOpenPartNo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmnuOpenPartNo.Click
        Dim Opened As Boolean = False
        If Me.ugrdLineItems.ActiveRow IsNot Nothing And Me.PartNo.Visible Then
            If Not Me.ugrdLineItems.ActiveRow.IsAddRow Then
                If Me.ugrdLineItems.ActiveRow.Cells(Me.PartNo.DataColName).Value IsNot DBNull.Value Then
                    If Me.ugrdLineItems.ActiveRow.Cells(Me.PartNo.DataColName).Value.ToString.Length > 0 Then
                        Dim Win As MyCore.Plugins.Window = Me.ParentWin.CallMethod(Me, Plugins.MethodType.Open, "Part No", Me.ugrdLineItems.ActiveRow.Cells(Me.PartNo.DataColName).Value, "Part No")
                        Me.ParentWin.OpenWindow(Win)
                        Opened = True
                    End If
                End If
            End If
        End If
        If Not Opened Then
            Me.ParentWin.Info.ShowDialog("No part number selected.  Select the row with the part number you want to open first, and then open.")
        End If
    End Sub

    Private Sub ugrdLineItems_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugrdLineItems.InitializeLayout

    End Sub
End Class
