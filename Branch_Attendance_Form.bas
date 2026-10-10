' ===============================================================================
' MODULE      : Standard Module (Module1)
' PURPOSE     : Stores global variables used across the workbook
' ===============================================================================
Option Explicit

' Stores the logged-in branch name after successful authentication
Public loggedbranch As String


' ===============================================================================
' MODULE      : ThisWorkbook (Object Module)
' PURPOSE     : Handles workbook open event and branch user login
' ===============================================================================
Private Sub workbook_open()
    Dim conn As Object, rs As Object
    Dim constring As String, sql As String
    Dim userid As String, password As String
    
    ' Step 1: Ask user for login details
    userid = Trim(InputBox("Enter a userid"))
    password = Trim(InputBox("Enter a password"))
    
    ' Check if user left input boxes empty
    If userid = "" Or password = "" Then
        MsgBox "Enter proper userid and password", vbExclamation
        GoTo closefile
    End If
    
    On Error GoTo errorhandler
    Application.ScreenUpdating = False
    
    ' Step 2: Create ADO objects for database connection
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Connection string for SQL Server database
    constring = "Driver={ODBC Driver 18 for SQL Server};" & _
                "Server=Your IP Address,1433;" & _
                "Database=Your DataBase Name;" & _
                "Uid=SQL UserID;" & _
                "Pwd=SQL Password;" & _
                "Encrypt=yes;" & _
                "TrustServerCertificate=yes;"
    
    conn.ConnectionTimeout = 15
    conn.Open constring
    
    ' Escape single quotes to prevent basic SQL errors
    userid = Replace(userid, "'", "''")
    password = Replace(password, "'", "''")
    
    ' Step 3: Check if user exists in database
    sql = "SELECT BranchName FROM branchlogin WHERE UserID = '" & userid & "' AND Password = '" & password & "'"
    
    Set rs = conn.Execute(sql)
    
    ' If login details match, show UserForm
    If Not rs.EOF Then
        loggedbranch = rs.Fields("BranchName").Value
        MsgBox "Welcome! " & loggedbranch, vbInformation, "Success"
        UserForm1.Show
    Else
        MsgBox "Invalid userid and password", vbCritical, "Login Failed"
        GoTo closefile
    End If

cleanup:
    ' Close connection safely
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Application.ScreenUpdating = True
    Exit Sub

closefile:
    ' Close connection and close workbook without saving
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
    Application.ScreenUpdating = True
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub

errorhandler:
    ' Display error message if connection fails
    MsgBox "Reason for error: " & Err.Description, vbCritical, "Error"
    Resume cleanup
End Sub


' ===============================================================================
' MODULE      : UserForm1 Code
' PURPOSE     : Manages attendance input validation and submission
' ===============================================================================

' Set logged branch name in text box when form opens
Private Sub UserForm_Initialize()
    Me.txtBranch.Value = loggedbranch
    Me.txtBranch.Enabled = False
End Sub


' Validate and submit attendance data to SQL database
Private Sub btnSubmit_Click()
    Dim conn As Object, cmd As Object
    Dim constring As String, sql As String, temp As String
    Dim x As Long, standard As Long, present As Long
    Dim txtcon As Object

    standard = 0
    present = 0
    
    ' Step 1: Loop through all 20 textboxes to check data
    For x = 1 To 20
        temp = Trim(Me.Controls("txt" & x).Text)
        
        ' Check if input is empty or non-numeric
        If temp = "" Or Not IsNumeric(temp) Then
            MsgBox "Invalid data entered at TextBox " & x, vbExclamation, "Validation Error"
            Me.Controls("txt" & x).SetFocus
            GoTo closefile
        End If
        
        ' Odd textboxes are Standard counts, Even textboxes are Present counts
        If x Mod 2 <> 0 Then
            standard = standard + Val(temp)
        Else
            present = present + Val(temp)
        End If
    Next x
    
    ' Logical check: Present students cannot be more than total standard count
    If present > standard Then
        MsgBox "Present count (" & present & ") cannot be higher than Standard count (" & standard & ")!", vbCritical, "Logic Warning"
        GoTo closefile
    End If
    
    On Error GoTo errorhandler
    Application.ScreenUpdating = False
    
    ' Step 2: Prepare SQL Connection and Command
    Set conn = CreateObject("ADODB.Connection")
    Set cmd = CreateObject("ADODB.Command")
    
    constring = "Driver={ODBC Driver 18 for SQL Server};" & _
                "Server=Your IP Address,1433;" & _
                "Database=Your DataBase Name;" & _
                "Uid=SQL User ID;" & _
                "Pwd=SQL password;" & _
                "Encrypt=yes;" & _
                "TrustServerCertificate=yes;"
                
    ' Parameterized query for safe data insertion
    sql = "INSERT INTO Attendance(BranchName, Class1_Standard, Class1_Present, Class2_Standard,Class2_Present,Class3_Standard,Class3_Present, " & _
          " Class4_Standard,Class4_Present,Class5_Standard,Class5_Present,Class6_Standard,Class6_Present,Class7_Standard,Class7_Present, " & _
          "Class8_Standard,Class8_Present,Class9_Standard,Class9_Present,Class10_Standard,Class10_Present) " & _
          "values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) "
          
    conn.ConnectionTimeout = 10
    conn.Open constring
    
    ' Start transaction to ensure full data safety
    conn.BeginTrans
    
    cmd.ActiveConnection = conn
    cmd.CommandType = 1
    cmd.CommandText = sql

    ' Append Branch Name parameter
    cmd.Parameters.Append cmd.Createparameter("@Branch", 200, 1, 50, Me.txtBranch.Text)
    
    ' Append all 20 textbox values as parameters
    For x = 1 To 20
        Set txtcon = Me.Controls("txt" & x)
        cmd.Parameters.Append cmd.Createparameter("@txt" & x, 3, 1, 0, Val(txtcon.Text))
    Next x
    
    ' Execute query and commit transaction
    cmd.Execute
    conn.CommitTrans
    
    MsgBox "Data submitted successfully!", vbInformation, "Success"

cleanup:
    ' Close connection and reset objects
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Set cmd = Nothing
    Set txtcon = Nothing
    Application.ScreenUpdating = True
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub
    
closefile:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Set cmd = Nothing
    Set txtcon = Nothing
    Application.ScreenUpdating = True
    ThisWorkbook.Close SaveChanges:=False
    Exit Sub
    
errorhandler:
    ' Undo database changes if an error occurs during execution
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.RollbackTrans
    End If
    MsgBox "Reason for error: " & Err.Description, vbCritical, "Execution Error"
    Resume cleanup
End Sub