# Using VBA & ADO to query worksheets 

A custom class module with procedures that perform the following 

- Create a unique list of values in the required table column
- Filter the table by each unique value
- Save the filtered data in a new workbook 

## Library Reference

A reference to a library required in the MS Excel VBE window. Choose Tools, References from the menu in the VBE and scroll down to the library below.

`Microsoft ActiveX Data Objects 6.1 Library`

## Setting up the custom class

Follow the steps below

1. Create a new class module in the VBE window
2. Rename the class to `clsGetDataByADO`
3. Paste the code in the page [here](clsGetDataByADO.bas)

## Explanation of Class Procedures

### Class Declarations

```vb
'ADO Connection
Dim MyConnection As ADODB.Connection

'The connection string used to by the previous objectvariable
Dim MyConn As String

'A variable used by the Class_Initialize event to record the number of new
'worksheets in a new workbook the current user has set in MS Excel Options
Dim NewShts As Long

'A constant variable used by the Class_Initialize event to set the number of new
'worksheets in a new workbook
Const c_NEW_SHTS As Long = 1
```

### Class_Initialize Event

This event is triggered when a new instance of the class is created. This event performs the following

- Records the default number of new worksheets the user has when creating a new workbook 
- Sets that initial application parameters
- Creates a connection string for use by the ADO Library
- Opens an ADO connection to the current workbook

```vb
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
```

### Class_Terminate Event

This event when the class is no longer required This event performs the following

- Resets the application parameters
- Closes and clears the ADO connection
- Resets the MS Excel Status Bar

```vb
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
```


### GetData Procedure

This is the procedure that is called from a normal code module after the class has been initialised. The three parameters required are

- Sht - The codename of the worksheet containing the required data table
- TableName - The name of the table on the worksheet to query
- ColumnName - The name of the column from which unique values will be extracted

Once the unique values 

```vb
Public Sub GetData(Sht As Worksheet, TableName As String, ColumnName As String)
    'Create variant array to store unique values from ColumnName
    Dim arrColumn() As Variant
    'Extract unique values from table column
    arrColumn = WorksheetFunction.Transpose(WorksheetFunction.Unique(Sht.ListObjects(TableName).ListColumns(ColumnName).DataBodyRange))
    'Read Data
    Dim i As Long
    Dim ColumnItem As String
    'Loop thru items in arrColumn array
    For i = LBound(arrColumn) To UBound(arrColumn)
        'Ensure value in position i in the arrColumn array is a string value
        ColumnItem = CStr(arrColumn(i))
        'Update the MS Excel status bar to inform the user which value is being exported
        Call UpdateStatusBar("Creating workbook for " & ColumnItem)
        'Call the procedure named CreateWorkbook
        Call CreateWorkbook(Sht, TableName, ColumnName, ColumnItem)
    'Loop to next item in arrColumn
    Next i
End Sub
``` 
