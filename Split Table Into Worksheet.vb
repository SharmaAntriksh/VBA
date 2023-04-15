Option Explicit

Sub SplitTableToWorksheets()
    
    Dim SalesTable As ListObject
    Dim Years() As Variant, Headers() As Variant
    Dim YearsDict As Scripting.Dictionary
    Dim x As Integer, CurrentYear As Integer
    Dim Item As Variant
    Dim ws As Worksheet
    
    Set SalesTable = wsSales.ListObjects("Sales")
    Headers = SalesTable.HeaderRowRange
    
    Years = SalesTable.ListColumns(2).DataBodyRange
    Set YearsDict = New Scripting.Dictionary
    
    ' Iterate each row in the Years Array and add it to the YearsDict
    For x = 1 To UBound(Years)
        CurrentYear = Years(x, 1)
        
        ' If the Key/Year already exists then don't add it
        If Not YearsDict.Exists(CurrentYear) Then
            YearsDict.Add CurrentYear, CurrentYear
        End If
    Next x
    
    Application.DisplayAlerts = False
    
    On Error Resume Next
    For Each Item In YearsDict.Items
        Set ws = Worksheets("" & Item & "")
        If Not ws Is Nothing Then ws.Delete
    Next Item
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    
    For Each Item In YearsDict.Items
        ' Apply the current year as filter to the SalesTable
        SalesTable.ListColumns(2).Range.AutoFilter Field:=2, Criteria1:=Item
        
        ' Add new worksheet, set the name as that of the current year and populate headers
        Set ws = Worksheets.Add
        ws.Name = Item
        ws.Range("A1:C1") = Headers
        
        ' Copy the filtered data in the Sales Table
        SalesTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
        
        'Paste it onto the Range A2 onwards.
        ws.Range("A2").PasteSpecial xlPasteValues
    Next Item
    
    SalesTable.AutoFilter.ShowAllData
    
End Sub


' Scenario where the table is to be split by specifying criteria on more than one column.

Sub SplitTableIntoWorksheets_TwoOrMoreColumns()
    
    Dim ProductsTable As ListObject
    Dim YearColor() As Variant
    Dim YearColorDict As Scripting.Dictionary
    Dim x As Integer
    Dim YearColorString As String
    Dim Item As Variant
    Dim YearColorArr As Variant
    Dim YearPart As String, ColorPart As String
    Dim ws As Worksheet
    Dim Headers() As Variant
    
    Set ProductsTable = wsSales.ListObjects("Products")
    Set YearColorDict = New Scripting.Dictionary
    YearColor = Application.Union(ProductsTable.ListColumns(1).DataBodyRange, ProductsTable.ListColumns(2).DataBodyRange)
    Headers = ProductsTable.HeaderRowRange
    
    For x = 1 To UBound(YearColor)
        YearColorString = YearColor(x, 1) & " " & YearColor(x, 2)
        If Not YearColorDict.Exists(YearColorString) Then
            YearColorDict.Add YearColorString, YearColorString
        End If
    Next x

    Application.DisplayAlerts = False
    
    On Error Resume Next
    For Each Item In YearColorDict.Items
        Set ws = Worksheets("" & Item & "")
        If Not ws Is Nothing Then ws.Delete
    Next Item
    On Error GoTo 0
    
    Application.DisplayAlerts = True
    
    For Each Item In YearColorDict
    
        YearColorArr = Split(Item, " ")
        YearPart = YearColorArr(0)
        ColorPart = YearColorArr(1)
        
        ProductsTable.ListColumns(1).Range.AutoFilter Field:=1, Criteria1:=YearPart
        ProductsTable.ListColumns(2).Range.AutoFilter Field:=2, Criteria1:=ColorPart
        
        Set ws = Worksheets.Add
        ws.Name = Item
        ws.Range("A1:C1") = Headers
        
        ProductsTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
        ws.Range("A2").PasteSpecial xlPasteValues
        
    Next Item
    
    ProductsTable.AutoFilter.ShowAllData
    
End Sub
