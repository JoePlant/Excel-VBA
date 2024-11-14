Attribute VB_Name = "modExportDataModelTables"
Option Explicit

'Constants
Private Const MAX_SHEET_NAME_LENGTH As Integer = 31
Private Const DEFAULT_SHEET_PREFIX As String = "Sheet"
Private Const PROGRESS_UPDATE_FREQUENCY As Long = 1000

'Type for storing table metadata
Private Type tableMetadata
    Name As String
    RowCount As Long
    columnCount As Integer
End Type

'Main export procedure
Public Sub ExportDataModelTables()
    Dim wb As Workbook
    Dim newWb As Workbook
    Dim mdl As Model
    Dim ws As Worksheet
    Dim oCnn As Object
    Dim oRS As Object
    Dim tbl As ModelTable
    Dim tableMetadata() As tableMetadata
    Dim tableCount As Integer
    
    On Error GoTo ErrorHandler
    
    'Initialize application settings
    SetApplicationSettings False
    
    'Get active workbook and model
    Set wb = ActiveWorkbook
    Set mdl = wb.Model
    
    'Validate input
    If Not ValidateModel(mdl) Then Exit Sub
    
    'Create new workbook
    Set newWb = Workbooks.Add
    
    'Initialize connection
    If Not InitializeConnection(mdl, oCnn, oRS) Then GoTo Cleanup
    
    'Initialize metadata array
    tableCount = mdl.ModelTables.Count
    ReDim tableMetadata(1 To tableCount)
    
    'Process each table
    Dim i As Integer
    i = 1
    For Each tbl In mdl.ModelTables
        If ProcessTable(tbl, newWb, oCnn, oRS, tableMetadata(i)) Then
            i = i + 1
        End If
    Next tbl
    
    'Clean up default sheets
    RemoveDefaultSheets newWb
    
    'Create summary sheet
    CreateSummarySheet newWb, tableMetadata, i - 1
    
    'Success message
    MsgBox "Export complete! Created " & newWb.Worksheets.Count - 1 & " data tables." & vbNewLine & _
           "Check the Summary sheet for details."
    
Cleanup:
    CleanupResources oCnn, oRS
    SetApplicationSettings True
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    LogError "An error occurred: " & Err.Description
    MsgBox "An error occurred. Please check the immediate window for details.", vbCritical
    GoTo Cleanup
End Sub

'Initialize ADO connection
Private Function InitializeConnection(mdl As Model, oCnn As Object, oRS As Object) As Boolean
    On Error GoTo ConnError
    
    Set oCnn = mdl.DataModelConnection.ModelConnection.ADOConnection
    Set oRS = CreateObject("ADODB.RecordSet")
    InitializeConnection = True
    Exit Function
    
ConnError:
    LogError "Failed to initialize connection: " & Err.Description
    InitializeConnection = False
End Function

'Process single table
Private Function ProcessTable(tbl As ModelTable, newWb As Workbook, oCnn As Object, oRS As Object, metadata As tableMetadata) As Boolean
    Dim ws As Worksheet
    Dim sSQL As String
    Dim i As Long
    
    On Error GoTo TableError
    
    Application.StatusBar = "Processing table " & tbl.Name & "..."
    Debug.Print "Processing table: " & tbl.Name
    
    'Construct SQL query
    sSQL = "SELECT * FROM $" & tbl.Name & ".$" & tbl.Name
    
    'Open recordset
    oRS.Open sSQL, oCnn
    
    'Create and name worksheet
    Set ws = newWb.Worksheets.Add
    ws.Name = GetValidSheetName(newWb, tbl.Name)
    
    'Store metadata
    metadata.Name = ws.Name
    metadata.columnCount = oRS.Fields.Count
    
    'Write headers
    WriteHeaders ws, oRS
    
    'Write data
    If Not oRS.EOF Then
        ws.Range("A2").CopyFromRecordset oRS
        metadata.RowCount = ws.UsedRange.Rows.Count - 1
    Else
        metadata.RowCount = 0
    End If
    
    'Format worksheet
    FormatWorksheet ws, oRS.Fields.Count
    
    'Close recordset
    oRS.Close
    
    ProcessTable = True
    Exit Function
    
TableError:
    LogError "Error processing table " & tbl.Name & ": " & Err.Description
    If Not oRS Is Nothing Then
        If oRS.State = 1 Then oRS.Close
    End If
    ProcessTable = False
End Function

'Write headers to worksheet
Private Sub WriteHeaders(ws As Worksheet, oRS As Object)
    Dim i As Integer
    For i = 0 To oRS.Fields.Count - 1
        ws.Cells(1, i + 1).Value = oRS.Fields(i).Name
    Next i
End Sub

'Format worksheet
Private Sub FormatWorksheet(ws As Worksheet, columnCount As Integer)
    With ws
        'Format headers
        .Range("A1").Resize(1, columnCount).Font.Bold = True
        
        'AutoFit columns
        .UsedRange.Columns.AutoFit
        
        'Add filters
        .Range("A1").Resize(1, columnCount).AutoFilter
        
        'Freeze top row
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
        
        'Go back to A1
        .Range("A1").Select
    End With
End Sub

'Create summary sheet
Private Sub CreateSummarySheet(wb As Workbook, tableMetadata() As tableMetadata, tableCount As Integer)
    Dim ws As Worksheet
    Dim i As Integer
    
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "Summary"
    
    With ws
        'Headers
        .Range("A1:C1") = Array("Table Name", "Row Count", "Column Count")
        .Range("A1:C1").Font.Bold = True
        
        'Data
        For i = 1 To tableCount
            .Cells(i + 1, 1).Value = tableMetadata(i).Name
            .Cells(i + 1, 2).Value = tableMetadata(i).RowCount
            .Cells(i + 1, 3).Value = tableMetadata(i).columnCount
        Next i
        
        'Formatting
        .UsedRange.Columns.AutoFit
        .Range("A1").AutoFilter
    End With
    
    'Activate summary
    ws.Activate
    ws.Range("A1").Select
End Sub

'Get valid sheet name
Private Function GetValidSheetName(wb As Workbook, proposedName As String) As String
    Dim result As String
    Dim i As Integer
    
    'Remove invalid characters
    result = RemoveInvalidCharacters(proposedName)
    
    'Truncate if necessary
    result = Left(result, MAX_SHEET_NAME_LENGTH)
    
    'Check for duplicates
    i = 1
    While SheetExists(wb, result)
        result = Left(RemoveInvalidCharacters(proposedName), MAX_SHEET_NAME_LENGTH - 3) & "(" & i & ")"
        i = i + 1
    Wend
    
    GetValidSheetName = result
End Function

'Remove invalid characters from sheet name
Private Function RemoveInvalidCharacters(str As String) As String
    Dim result As String
    result = str
    
    'Remove invalid characters
    result = Replace(result, "\", "")
    result = Replace(result, "/", "")
    result = Replace(result, "?", "")
    result = Replace(result, "*", "")
    result = Replace(result, "[", "")
    result = Replace(result, "]", "")
    result = Replace(result, ":", "")
    
    RemoveInvalidCharacters = result
End Function

'Check if sheet exists
Private Function SheetExists(wb As Workbook, sheetName As String) As Boolean
    Dim ws As Worksheet
    
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

'Remove default sheets
Private Sub RemoveDefaultSheets(wb As Workbook)
    Dim ws As Worksheet
    
    Application.DisplayAlerts = False
    For Each ws In wb.Worksheets
        If ws.Name Like "Sheet*" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True
End Sub

'Validate model
Private Function ValidateModel(mdl As Model) As Boolean
    If mdl Is Nothing Then
        MsgBox "No data model found in active workbook.", vbExclamation
        ValidateModel = False
        Exit Function
    End If
    
    If mdl.ModelTables.Count = 0 Then
        MsgBox "No tables found in data model.", vbExclamation
        ValidateModel = False
        Exit Function
    End If
    
    ValidateModel = True
End Function

'Set application settings
Private Sub SetApplicationSettings(restore As Boolean)
    With Application
        .ScreenUpdating = restore
        .EnableEvents = restore
        .Calculation = IIf(restore, xlCalculationAutomatic, xlCalculationManual)
    End With
End Sub

'Cleanup resources
Private Sub CleanupResources(oCnn As Object, oRS As Object)
    If Not oRS Is Nothing Then
        If oRS.State = 1 Then oRS.Close
        Set oRS = Nothing
    End If
    
    If Not oCnn Is Nothing Then
        If oCnn.State = 1 Then oCnn.Close
        Set oCnn = Nothing
    End If
End Sub

'Log error
Private Sub LogError(message As String)
    Debug.Print Format(Now, "yyyy-mm-dd hh:mm:ss") & " - ERROR: " & message
End Sub

