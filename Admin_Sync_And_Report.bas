' ===============================================================================
' Subroutine   : LoadDataToAdminSheet
' Description  : Fetches today's attendance data from SQL database and loads it into the Admin sheet.
' Module       : Standard Module
' ===============================================================================
Sub LoadDataToAdminSheet()
    Dim conn As Object, rs As Object
    Dim ConString As String, sql As String
    Dim ws As Worksheet
    Dim i As Integer
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    ' Step 1: Safely select the Admin worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Admin")
    On Error GoTo 0
    
    If ws Is Nothing Then Set ws = ThisWorkbook.ActiveSheet
    ws.Cells.Clear
    
    ' Step 2: Set up SQL Server connection string (using dummy placeholders for security)
    ConString = "Provider=SQLOLEDB;" & _
                "Data Source=YOUR_SERVER_NAME;" & _
                "Initial Catalog=YOUR_DATABASE_NAME;" & _
                "Integrated Security=SSPI;"
    
    ' Step 3: Open database connection
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 10
    conn.Open ConString
    
    ' Step 4: Execute SQL query to fetch today's attendance records
    Set rs = CreateObject("ADODB.Recordset")
    sql = "SELECT * FROM Attendance " & _
          "WHERE CAST(EntryDate AS DATE) = CAST(GETDATE() AS DATE) " & _
          "ORDER BY BranchName ASC"
    
    rs.Open sql, conn
        
    ' Step 5: If data exists, write column headers and paste records into sheet
    If Not rs.EOF Then
        For i = 0 To rs.Fields.Count - 1
            ws.Cells(1, i + 1).Value = rs.Fields(i).Name
        Next i
        ws.Range("A2").CopyFromRecordset rs
        ws.UsedRange.Columns.AutoFit
        MsgBox "Today's attendance data loaded successfully into the Admin sheet!", vbInformation, "Sync Complete"
    Else
        MsgBox "No data found today", vbExclamation, "Information"
    End If

CleanExit:
    ' Properly close connection and clean up memory objects
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Failed to load data from database: " & Err.Description, vbCritical, "Data Extraction Error"
    Resume CleanExit
End Sub


' ===============================================================================
' Subroutine   : CleanData
' Description  : Formats attendance table, adds serial numbers, sets page layout,
'                exports report as landscape PDF to Desktop, and clears sheet.
' Module       : Module1 (Report Generation & Export)
' ===============================================================================
Sub CleanData()
    Dim ws As Worksheet
    Dim lastrow As Long, lastcol As Long
    Dim x As Long
    Dim fso As Object
    Dim folderPath As String
    Dim pdfpath As String
    Dim fileName As String
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    ' Step 1: Target the Admin worksheet safely
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Admin")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets(1)
    End If
    
    ' Step 2: Find total rows and columns of data
    lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastcol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Step 3: Apply cell borders and header formatting (Yellow fill with bold text)
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastrow, lastcol))
        .Borders.LineStyle = xlContinuous
    End With
    
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastcol))
        .Interior.Color = vbYellow
        .Font.Bold = True
    End With
    
    ' Step 4: Insert serial numbers in Column A
    ws.Cells(1, 1).Value = "S.No"
    For x = 2 To lastrow
        ws.Cells(x, 1).Value = x - 1
    Next x
    
    ws.Columns.AutoFit
    
    ' Step 5: Create PDF output folder on Desktop if it doesn't exist
    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = fso.Buildpath(Environ("userprofile"), "Desktop\Attendance PDF")
    
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder (folderPath)
    End If
    
    ' Generate unique filename using date and time stamp
    fileName = "Attendance_Report_" & Format(Now, "YYYY-MM-DD_hh-mm-ss") & ".pdf"
    pdfpath = fso.Buildpath(folderPath, fileName)
    
    ' Step 6: Configure page setup for landscape A3 PDF export
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA3
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
    
    ' Step 7: Export worksheet directly to PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
                          Filename:=pdfpath, _
                          Quality:=xlQualityStandard, _
                          IncludeDocProperties:=True, _
                          IgnorePrintAreas:=False, _
                          OpenAfterPublish:=True
                           
    ' Step 8: Clear temporary data from sheet after export
    ws.Cells.Clear

    Application.ScreenUpdating = True
    Set fso = Nothing
    MsgBox "Attendance report generated, exported, and sheet cleared successfully!", vbInformation, "Export Completed"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error during PDF export: " & Err.Description, vbCritical, "Export Error"
End Sub