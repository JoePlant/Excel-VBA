Attribute VB_Name = "modExportDataModelTables"
Option Explicit

'Constants
Private Const MAX_SHEET_NAME_LENGTH As Integer = 31
Private Const DEFAULT_SHEET_PREFIX As String = "Sheet"
Private Const MAX_RECORDS_PER_BATCH As Long = 10000

'Type for storing table metadata
Private Type tableMetadata
    Name As String
    RowCount As Long
    RecordCount As Long
    columnCount As Integer
    ActualRowCount As Long
End Type

'Main export procedure
Public Sub ExportDataModelTables()
    Dim wb As Workbook
    Dim newWb As Workbook
    Dim mdl As Model
    Dim tableMetadata() As tableMetadata
    Dim tableCount As Integer
    Dim oCnn As Object
    Dim oRS As Object
    
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
    ProcessAllTables mdl, newWb, oCnn, oRS, tableMetadata, tableCount
    
    'Clean up default sheets and create summary
    RemoveDefaultSheets newWb
    CreateEnhancedSummarySheet newWb, tableMetadata, tableCount
    
    'Display completion message
    ShowCompletionMessage newWb, tableMetadata, tableCount
    
Cleanup:
    CleanupResources oCnn, oRS
    SetApplicationSettings True
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    LogError "An error occurred in ExportDataModelTables: " & Err.Description
    MsgBox "An error occurred. Please check the immediate window for details.", vbCritical
    GoTo Cleanup
End Sub

'Initialize ADO connection with optimized settings
Private Function InitializeConnection(mdl As Model, oCnn As Object, oRS As Object) As Boolean
    On Error GoTo ConnError
    
    'Set up connection
    Set oCnn = mdl.DataModelConnection.ModelConnection.ADOConnection
    With oCnn
        .CommandTimeout = 0 'Prevent timeout for large datasets
    End With
    
    'Set up recordset with optimal properties
    Set oRS = CreateObject("ADODB.RecordSet")
    With oRS
        .CursorLocation = 3 'adUseClient
        .CursorType = 1    'adOpenKeyset
        .LockType = 1      'adLockReadOnly
    End With
    
    InitializeConnection = True
    Exit Function
    
ConnError:
    LogError "Failed to initialize connection: " & Err.Description
    InitializeConnection = False
End Function

'Process all tables in the model
Private Sub ProcessAllTables(mdl As Model, newWb As Workbook, oCnn As Object, oRS As Object, _
                           tableMetadata() As tableMetadata, ByRef tableCount As Integer)
    Dim tbl As ModelTable
    Dim i As Integer
    
    i = 1
    For Each tbl In mdl.ModelTables
        If ProcessTable(tbl, newWb, oCnn, oRS, tableMetadata(i)) Then
            Debug.Print "Completed processing table: " & tbl.Name
            Debug.Print "Expected records: " & tbl.RecordCount
            Debug.Print "Actual records: " & tableMetadata(i).ActualRowCount
            i = i + 1
        End If
    Next tbl
    
    tableCount = i - 1
End Sub

'Process single table with batch processing and verification
Private Function ProcessTable(tbl As ModelTable, newWb As Workbook, oCnn As Object, oRS As Object, _
                            metadata As tableMetadata) As Boolean
    Dim ws As Worksheet
    Dim sSQL As String
    Dim currentRow As Long
    Dim recordsProcessed As Long
    Dim batchSize As Long
    
    On Error GoTo TableError
    
    'Update status and start processing
    Application.StatusBar = "Processing table " & tbl.Name & "..."
    Debug.Print "Starting to process table: " & tbl.Name
    
    'Set up worksheet
    sSQL = "SELECT * FROM $" & tbl.Name & ".$" & tbl.Name
    oRS.Open sSQL, oCnn
    
    Set ws = newWb.Worksheets.Add
    ws.Name = GetValidSheetName(newWb, tbl.Name)
    
    'Initialize metadata
    InitializeTableMetadata metadata, ws.Name, tbl.RecordCount, oRS.Fields.Count
    
    'Write headers
    WriteHeaders ws, oRS
    
    'Process data in batches
    currentRow = 2 'Start after headers
    recordsProcessed = 0
    
    Do While Not oRS.EOF
        'Calculate batch size
        batchSize = GetBatchSize(metadata.RecordCount, recordsProcessed)
        
        'Copy batch
        ws.Range("A" & currentRow).CopyFromRecordset oRS, batchSize
        
        'Update progress
        recordsProcessed = recordsProcessed + batchSize
        currentRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        UpdateProgress tbl.Name, recordsProcessed, metadata.RecordCount
    Loop
    
    'Finalize processing
    FinalizeTableProcessing ws, metadata
    
    'Verify counts
    If metadata.ActualRowCount <> metadata.RecordCount Then
        LogError "Warning: " & tbl.Name & " has count mismatch. Expected: " & _
                metadata.RecordCount & ", Actual: " & metadata.ActualRowCount
    End If
    
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

'Initialize table metadata
Private Sub InitializeTableMetadata(metadata As tableMetadata, Name As String, _
                                  RecordCount As Long, columnCount As Integer)
    With metadata
        .Name = Name
        .columnCount = columnCount
        .RecordCount = RecordCount
        .ActualRowCount = 0
    End With
End Sub

'Get batch size for processing
Private Function GetBatchSize(totalRecords As Long, processedRecords As Long) As Long
    If totalRecords - processedRecords > MAX_RECORDS_PER_BATCH Then
        GetBatchSize = MAX_RECORDS_PER_BATCH
    Else
        GetBatchSize = totalRecords - processedRecords
    End If
End Function

'Update progress indicator
Private Sub UpdateProgress(tableName As String, processed As Long, total As Long)
    Application.StatusBar = "Processing " & tableName & "... " & _
                          Format(processed / total, "0%")
    DoEvents
End Sub

'Finalize table processing
Private Sub FinalizeTableProcessing(ws As Worksheet, ByRef metadata As tableMetadata)
    metadata.ActualRowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row - 1
    metadata.RowCount = metadata.ActualRowCount
    FormatWorksheet ws, metadata.columnCount
End Sub

'Write headers to worksheet
Private Sub WriteHeaders(ws As Worksheet, oRS As Object)
    Dim i As Integer
    For i = 0 To oRS.Fields.Count - 1
        ws.Cells(1, i + 1).Value = oRS.Fields(i).Name
    Next i
End Sub

'Format worksheet with proper styling
Private Sub FormatWorksheet(ws As Worksheet, columnCount As Integer)
    With ws
        'Format headers
        With .Range("A1").Resize(1, columnCount)
            .Font.Bold = True
            .Interior.Color = RGB(240, 240, 240)
        End With
        
        'AutoFit columns and add filters
        .UsedRange.Columns.AutoFit
        .Range("A1").Resize(1, columnCount).AutoFilter
        
        'Freeze panes and reset selection
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
        .Range("A1").Select
    End With
End Sub

'Create enhanced summary sheet with discrepancy highlighting
Private Sub CreateEnhancedSummarySheet(wb As Workbook, tableMetadata() As tableMetadata, tableCount As Integer)
    Dim ws As Worksheet
    Dim i As Integer
    
    Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    ws.Name = "Summary"
    
    With ws
        'Headers
        .Range("A1:E1") = Array("Table Name", "Expected Records", "Actual Records", "Difference", "Columns")
        .Range("A1:E1").Font.Bold = True
        
        'Data
        For i = 1 To tableCount
            With .Cells(i + 1, 1).Resize(1, 5)
                .Cells(1) = tableMetadata(i).Name
                .Cells(2) = tableMetadata(i).RecordCount
                .Cells(3) = tableMetadata(i).ActualRowCount
                .Cells(4).Formula = "=C" & (i + 1) & "-B" & (i + 1)
                .Cells(5) = tableMetadata(i).columnCount
                
                'Highlight discrepancies
                If tableMetadata(i).RecordCount <> tableMetadata(i).ActualRowCount Then
                    .Interior.Color = RGB(255, 255, 200)
                End If
            End With
        Next i
        
        'Formatting and timestamp
        .UsedRange.Columns.AutoFit
        .Range("A1").AutoFilter
        
        .Cells(tableCount + 3, 1).Value = "Export Date:"
        With .Cells(tableCount + 3, 2)
            .Value = Now
            .NumberFormat = "yyyy-mm-dd hh:mm:ss"
        End With
    End With
End Sub

'Show completion message
Private Sub ShowCompletionMessage(wb As Workbook, tableMetadata() As tableMetadata, tableCount As Integer)
    Dim msg As String
    
    msg = "Export complete! Created " & wb.Worksheets.Count - 1 & " data tables." & vbNewLine & _
          "Check the Summary sheet for details."
    
    If HasDiscrepancies(tableMetadata, tableCount) Then
        msg = msg & vbNewLine & vbNewLine & _
              "WARNING: Some tables showed record count discrepancies. " & _
              "See the Summary sheet for details."
    End If
    
    MsgBox msg
End Sub

'Get valid sheet name handling Excel's limitations
Private Function GetValidSheetName(wb As Workbook, proposedName As String) As String
    Dim result As String
    Dim i As Integer
    
    'Remove invalid characters and handle length
    result = RemoveInvalidCharacters(proposedName)
    result = Left(result, MAX_SHEET_NAME_LENGTH)
    
    'Handle special cases
    If result = "History" Then result = "_History"
    
    'Handle duplicates
    i = 1
    While SheetExists(wb, result)
        result = Left(RemoveInvalidCharacters(proposedName), MAX_SHEET_NAME_LENGTH - Len(CStr(i)) - 2) & "(" & i & ")"
        i = i + 1
    Wend
    
    GetValidSheetName = result
End Function

'Remove invalid characters from sheet name
Private Function RemoveInvalidCharacters(str As String) As String
    Dim result As String
    result = str
    
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
        If ws.Name Like DEFAULT_SHEET_PREFIX & "*" Then ws.Delete
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
        .DisplayAlerts = restore
    End With
End Sub

'Cleanup resources
Private Sub CleanupResources(oCnn As Object, oRS As Object)
    On Error Resume Next
    
    If Not oRS Is Nothing Then
        If oRS.State = 1 Then oRS.Close
        Set oRS = Nothing
    End If
    
    If Not oCnn Is Nothing Then
        If oCnn.State = 1 Then oCnn.Close
        Set oCnn = Nothing
    End If
    
    On Error GoTo 0
End Sub

'Helper function to check for discrepancies
Private Function HasDiscrepancies(tableMetadata() As tableMetadata, tableCount As Integer) As Boolean
    Dim i As Integer
    For i = 1 To tableCount
        If tableMetadata(i).RecordCount <> tableMetadata(i).ActualRowCount Then
            HasDiscrepancies = True
            Exit Function
        End If
    Next i
    HasDiscrepancies = False
End Function

'Log error
Private Sub LogError(message As String)
    Debug.Print Format(Now, "yyyy-mm-dd hh:mm:ss") & " - ERROR: " & message
End Sub

