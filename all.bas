Option Explicit

Sub Initialize()
    Call CreateDashboard
End Sub

Sub SetupDashboardLayout()
    Dim wsDash As Worksheet
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0
    
    ' Check if dashboard exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Dashboard")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Please create dashboard first!", vbExclamation
        Exit Sub
    End If
    
        
    If wsDash Is Nothing Then
        Set wsDash = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        wsDash.Name = "Dashboard"
    End If
    
    ' Clear existing content
    wsDash.Cells.Clear
    
    ' Create headers
    wsDash.Range("A1").Value = "Well List"
    wsDash.Range("C1").Value = "Select Well"
    wsDash.Range("E1").Value = "Gross Oil"
    wsDash.Range("F1").Value = "Net Oil"
    wsDash.Range("G1").Value = "Water Cut"
    wsDash.Range("H1").Value = "GOR"
    
    ' Format headers
    With wsDash.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
        .HorizontalAlignment = xlCenter
    End With
    
    ' Set column widths
    wsDash.Columns("A:A").ColumnWidth = 20
    wsDash.Columns("C:C").ColumnWidth = 15
    wsDash.Columns("E:H").ColumnWidth = 12
    
    ' Add border to input area
    wsDash.Range("A2:H2").Borders.Weight = xlThin
End Sub





Sub RefreshAll()
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    Debug.Print "Step 1: Checking/creating Dashboard"
    If Not sheetExists("Dashboard") Then CreateDashboard
    Debug.Print "Step 2: Importing data"
    Import_Well_Data
    Debug.Print "Step 3: Building selector"
    Build_Well_Selector
    Debug.Print "Step 4: Calculating KPIs"
    Calculate_KPIs
    Debug.Print "Step 5: Updating charts"
    Update_Charts
    Debug.Print "Step 6: Exporting PDF"
    Export_To_PDF
    Debug.Print "Step 7: Archiving file"
    Archive_File

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    MsgBox "Dashboard refreshed successfully!", vbInformation
    Exit Sub

ErrHandler:
    Debug.Print "Error at " & Err.Source & ": " & Err.Number & " - " & Err.Description
    MsgBox "Error " & Err.Number & " in " & Err.Source & ":" & vbCrLf & Err.Description, vbCritical
    Resume CleanExit
End Sub



'------------------------------------------------
' Import and transform data with dynamic filename
'------------------------------------------------
Sub Import_Well_Data()
    Dim wsRaw    As Worksheet
    Dim tempWB   As Workbook
    Dim lo       As ListObject
    Dim fileName As String, filePath As String

    Debug.Print "? Import_Well_Data start"

    ' 1) Pastikan sheet RawData ada
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets("RawData")
    On Error GoTo 0
    If wsRaw Is Nothing Then
        MsgBox "Error: Sheet 'RawData' tidak ditemukan!", vbCritical
        Exit Sub
    End If
    Debug.Print "   • Found sheet RawData"

    ' 2) Bersihkan sheet
    wsRaw.Cells.Clear
    Debug.Print "   • RawData sheet cleared"

    ' 3) Tentukan path file CSV (kemarin)
    fileName = Format(Date - 1, "yyyymmdd") & "_WITSML.csv"
    filePath = "C:\WellDashboard\Raw\" & fileName
    Debug.Print "   • Looking for file: " & filePath

    If Dir(filePath) = "" Then
        MsgBox "Error: File data tidak ditemukan:" & vbCrLf & filePath, vbExclamation
        Exit Sub
    End If
    Debug.Print "   • File found, opening CSV"

    ' 4) Buka CSV, copy data ke RawData, lalu tutup
    Set tempWB = Workbooks.Open(filePath, ReadOnly:=True)
    Debug.Print "   • CSV opened"
    tempWB.Sheets(1).UsedRange.Copy wsRaw.Range("A1")
    tempWB.Close SaveChanges:=False
    Debug.Print "   • CSV data copied into RawData"

    ' 5) Hapus tabel lama (jika ada), lalu buat ListObject baru
    On Error Resume Next
    For Each lo In wsRaw.ListObjects: lo.Delete: Next lo
    On Error GoTo 0
    Debug.Print "   • Old tables deleted"

    Set lo = wsRaw.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=wsRaw.UsedRange, _
        XlListObjectHasHeaders:=xlYes)
    lo.Name = "WellData"
    Debug.Print "   • New table 'WellData' created"

    ' 6) Rename kolom jika perlu
    On Error Resume Next
    lo.ListColumns("Gross_Oil_bbl").Name = "Gross_Oil"
    lo.ListColumns("Water_bbl").Name = "Water"
    lo.ListColumns("Gas_Mscf").Name = "Gas"
    On Error GoTo 0
    Debug.Print "   • Columns renamed (if applicable)"

    Debug.Print "? Import_Well_Data complete"
End Sub




'==============================
' Build Dropdown
'==============================

' Perbaikan pada Build_Well_Selector

Sub Build_Well_Selector()
    Dim wsDash  As Worksheet
    Dim wsData  As Worksheet
    Dim lo       As ListObject
    Dim dict     As Object
    Dim cell     As Variant   ' ? must be Variant, not Range
    Dim i        As Long
    Dim lastRow As Long

    Debug.Print "? Build_Well_Selector start"

    ' 1) Sheets
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsData = ThisWorkbook.Sheets("RawData")

    ' 2) Table
    Set lo = wsData.ListObjects("WellData")

    ' 3) Unique wells
    Set dict = CreateObject("Scripting.Dictionary")
    For Each cell In lo.ListColumns("Well_ID").DataBodyRange.Value
        If Len(cell) > 0 Then dict(cell) = 1
    Next cell
    Debug.Print "   Unique wells collected:"; dict.count

    ' 4) Write to A2:A
    wsDash.Range("A2:A100").ClearContents
    i = 2
    For Each cell In dict.Keys
        wsDash.Cells(i, 1).Value = cell
        i = i + 1
    Next cell
    lastRow = i - 1
    Debug.Print "   Wells written to A2:A" & lastRow

    ' 5) Data Validation C2
    With wsDash.Range("C2")
        On Error Resume Next: .Validation.Delete: On Error GoTo 0
        If dict.count > 0 Then
            .Validation.Add Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Formula1:="=$A$2:$A$" & lastRow
            .Validation.IgnoreBlank = True
            .Validation.InCellDropdown = True
            Debug.Print "   Validation set"
        Else
            .Value = "No wells"
        End If
    End With

    Debug.Print "? Build_Well_Selector complete"
End Sub


'------------------------------------------------
' Calculate KPIs
'------------------------------------------------
Sub Calculate_KPIs()
    Dim wsDash As Worksheet
    Dim wsData As Worksheet
    Dim lo As ListObject
    Dim selectedWell As String
    Dim rng As Range
    Dim totalOil As Double, totalWater As Double, totalGas As Double
    Dim count As Long
    
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsData = ThisWorkbook.Sheets("RawData")
    selectedWell = wsDash.Range("C2").Value
    
    ' Reset values
    wsDash.Range("E2:H2").ClearContents
    
    On Error Resume Next
    Set lo = wsData.ListObjects("WellData")
    On Error GoTo 0
    
    If Not lo Is Nothing And selectedWell <> "No wells found" Then
        Set rng = lo.ListColumns("Well_ID").DataBodyRange
        
        ' Gross Oil
        totalOil = Application.WorksheetFunction.SumIfs( _
            lo.ListColumns("Gross_Oil").DataBodyRange, rng, selectedWell)
        
        ' Water
        totalWater = Application.WorksheetFunction.SumIfs( _
            lo.ListColumns("Water").DataBodyRange, rng, selectedWell)
        
        ' Gas
        totalGas = Application.WorksheetFunction.SumIfs( _
            lo.ListColumns("Gas").DataBodyRange, rng, selectedWell)
        
        ' Count records for average calculations
        count = Application.WorksheetFunction.CountIfs(rng, selectedWell)
        
        ' Set values
        wsDash.Range("E2").Value = totalOil
        wsDash.Range("F2").Value = totalOil * (1 - totalWater / (totalOil + totalWater))
        
        ' Water Cut
        If totalOil > 0 Then
            wsDash.Range("G2").Value = totalWater / (totalOil + totalWater)
        Else
            wsDash.Range("G2").Value = 0
        End If
        
        ' GOR
        If totalOil > 0 And count > 0 Then
            wsDash.Range("H2").Value = totalGas / totalOil
        Else
            wsDash.Range("H2").Value = 0
        End If
    End If
    
    ' Cleanup
    Set rng = Nothing
    Set lo = Nothing
    Set wsData = Nothing
    Set wsDash = Nothing
End Sub

'------------------------------------------------
' Update rolling 30-day charts
'------------------------------------------------
Sub Update_Charts()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    
    ' Remove existing chart
    On Error Resume Next
    wsDash.ChartObjects("ProdChart").Delete
    On Error GoTo 0
    
    ' Only create chart if we have data
    If wsDash.Range("E2").Value > 0 Then
        Dim cht As ChartObject
        Set cht = wsDash.ChartObjects.Add(Left:=150, Width:=400, Top:=100, Height:=250)
        cht.Name = "ProdChart"
        
        With cht.Chart
            .ChartType = xlColumnClustered
            .SetSourceData Source:=wsDash.Range("E1:H2")
            .HasTitle = True
            .ChartTitle.Text = "Production KPIs"
            .Axes(xlCategory).CategoryNames = wsDash.Range("E1:H1").Value
            .ApplyLayout 3 ' Layout with title and axis labels
        End With
    End If
    
    Set cht = Nothing
    Set wsDash = Nothing
End Sub

'------------------------------------------------
' Export to PDF
'------------------------------------------------

Sub Export_To_PDF()
    Dim reportPath As String
    Dim fso As Object
    Dim folderPath As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Dashboard")
    folderPath = "C:\WellDashboard\Reports\"
    reportPath = folderPath & "ProductionReport_" & Format(Date, "yyyymmdd") & ".pdf"
    
    ' Create folder if needed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    With ws.PageSetup
        .Orientation = xlLandscape            ' Set orientation to landscape
        .Zoom = False                         ' Disable automatic zoom
        .FitToPagesWide = 1                   ' Scale to fit width on one page
        .FitToPagesTall = False               ' No limit on pages tall
    End With
    
    ' Export to PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=reportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False
    
    ' Verify creation
    If fso.FileExists(reportPath) Then
        MsgBox "Report saved to:" & vbCrLf & reportPath, vbInformation
    Else
        MsgBox "Failed to create PDF report", vbExclamation
    End If
    
    Set fso = Nothing
    Set ws = Nothing
End Sub



Sub Export_To_PDF_portrait()
    Dim reportPath As String
    Dim fso As Object
    Dim folderPath As String
    
    folderPath = "C:\WellDashboard\Reports\"
    reportPath = folderPath & "ProductionReport_" & Format(Date, "yyyymmdd") & ".pdf"
    
    ' Create folder if needed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    ' Export
    ThisWorkbook.Sheets("Dashboard").ExportAsFixedFormat _
        Type:=xlTypePDF, _
        fileName:=reportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False
    
    ' Verify creation
    If fso.FileExists(reportPath) Then
        MsgBox "Report saved to:" & vbCrLf & reportPath, vbInformation
    Else
        MsgBox "Failed to create PDF report", vbExclamation
    End If
    
    Set fso = Nothing
End Sub

'------------------------------------------------
' Archive processed file
'------------------------------------------------
Sub Archive_File()
    Dim fso As Object
    Dim srcPath As String
    Dim destPath As String
    Dim folderPath As String
    
    srcPath = "C:\WellDashboard\Raw\" & Format(Date - 1, "yyyymmdd") & "_WITSML.csv"
    folderPath = "C:\WellDashboard\Archive\"
    destPath = folderPath & Format(Date - 1, "yyyymmdd") & "_WITSML.csv"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Create archive folder if needed
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    
    If fso.FileExists(srcPath) Then
        If fso.FileExists(destPath) Then fso.DeleteFile destPath
        fso.MoveFile Source:=srcPath, Destination:=destPath
    End If
    
    Set fso = Nothing
End Sub

Function sheetExists(sheetName As String) As Boolean
    On Error Resume Next
    sheetExists = (ThisWorkbook.Sheets(sheetName).Name <> "")
    On Error GoTo 0
End Function

