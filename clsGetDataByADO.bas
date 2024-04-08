Option Explicit

Dim MyConnection As ADODB.Connection
Dim MyConn As String
Dim NewShts As Long

Const c_NEW_SHTS As Long = 1

Private Sub Class_Initialize()
    NewShts = Application.SheetsInNewWorkbook
    With Application
        .Cursor = xlWait
        .DisplayAlerts = False
        .ScreenUpdating = False
        .SheetsInNewWorkbook = c_NEW_SHTS
    End With
    MyConn = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                    "Data Source=""" & ThisWorkbook.FullName & """;" & _
                    "Extended Properties=""Excel 12.0;HDR=Yes"""
    Set MyConnection = New ADODB.Connection
    MyConnection.Open MyConn
End Sub

Private Sub Class_Terminate()
    With Application
        .Cursor = xlDefault
        .DisplayAlerts = True
        .ScreenUpdating = True
        .SheetsInNewWorkbook = NewShts
    End With
    Call ClearObject(MyObject:=MyConnection)
    Call UpdateStatusBar
End Sub

Public Sub GetData(Sht As Worksheet, TableName As String, ColumnName As String)
    Dim arrColumn() As Variant
    arrColumn = WorksheetFunction.Transpose(WorksheetFunction.Unique(Sht.ListObjects(TableName).ListColumns(ColumnName).DataBodyRange))
    'Read Data
    Dim i As Long
    Dim ColumnItem As String
    For i = LBound(arrColumn) To UBound(arrColumn)
        ColumnItem = CStr(arrColumn(i))
        Call UpdateStatusBar("Creating workbook for " & ColumnItem)
        Call CreateWorkbook(Sht, TableName, ColumnName, ColumnItem)
    Next i
End Sub

Private Sub CreateWorkbook(Sht As Worksheet, TableName As String, ColumnName As String, ColumnValue As String)
    'Get data
    Dim arrData As Variant
    arrData = SQLFilter(Sht:=Sht, ColumnName:=ColumnName, ColumnValue:=ColumnValue)
    If IsEmpty(arrData) Then Exit Sub
    Dim arrHeaders As Variant
    arrHeaders = WorksheetFunction.Transpose(shtData.ListObjects(TableName).HeaderRowRange.Value2)
    'Create workbook
    Dim Wbk As Workbook
    Set Wbk = Workbooks.Add
    Wbk.SaveAs ThisWorkbook.Path & _
                Application.PathSeparator & _
                Format(Now, "yyyy-mm-dd") & _
                " " & ColumnValue & ".xlsx"
    Dim NewSht As Worksheet
    Set NewSht = Wbk.Sheets(1)
    With NewSht
        'Paste headers to sheet
        .Range("A1").Resize(1, UBound(arrHeaders)).Value = WorksheetFunction.Transpose(arrHeaders)
        'Paste to sheet
        .Range("A1").Offset(1).Resize(UBound(arrData, 1), UBound(arrData, 2) + 1).Value2 = arrData
        'Make table
        .ListObjects.Add(xlSrcRange, Range("$A$1:$I$" & UBound(arrData, 1)), , xlYes).Name = "Data"
        'Format table
        With .ListObjects("Data")
            'Format date columns
            .ListColumns(5).DataBodyRange.Resize(UBound(arrData, 1), 2).NumberFormat = "dd/MM/yyyy"
            'Do some formatting
            With .Range
                .Font.Name = "Arial"
                .Font.Size = 10
                .RowHeight = 18
                .VerticalAlignment = xlCenter
                .Columns.AutoFit
            End With
        End With
    End With
    Wbk.Close SaveChanges:=True
HandleExit:
    Exit Sub
End Sub

Private Function SQLFilter(Sht As Worksheet, ColumnName As String, ColumnValue As String) As Variant
    Dim MySQL As String
    MySQL = "SELECT * FROM [" & Sht.Name & "$] WHERE [" & ColumnName & "] = '" & ColumnValue & "'"
    Dim MyRS As New ADODB.Recordset
    MyRS.Open Source:=MySQL, ActiveConnection:=MyConn, CursorType:=adOpenKeyset
    'This should not trigger as all column values taken
    'from unique list of items in the required column
    If MyRS.RecordCount < 1 Then
        SQLFilter = Empty
        GoTo HandleExit
    End If
    Dim arrData As Variant
    arrData = MyRS.GetRows
    Dim arrFinalData As Variant
    ReDim arrFinalData(0 To UBound(arrData, 2), 0 To UBound(arrData, 1))
    Dim i As Long, j As Long
    For i = LBound(arrData, 2) To UBound(arrData, 2)
        For j = LBound(arrData, 1) To UBound(arrData, 1)
            arrFinalData(i, j) = arrData(j, i)
    Next j, i
    SQLFilter = arrFinalData
HandleExit:
    Call ClearObject(MyObject:=MyRS)
    Exit Function
End Function

Private Sub ClearObject(MyObject As Object)
    If Not (MyObject Is Nothing) Then
        If (MyObject.State And adStateOpen) = adStateOpen Then MyObject.Close
        Set MyObject = Nothing
    End If
End Sub

Private Sub UpdateStatusBar(Optional Status As Variant = False)
    Application.StatusBar = Status
End Sub
