Sub CreateDashboard()
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    Dim i As Integer
    
    ' Check if sheet already exists
    sheetExists = False
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = "Dashboard" Then
            sheetExists = True
            Exit For
        End If
    Next ws
    
    ' Delete existing Dashboard if user confirms
    If sheetExists Then
        If MsgBox("Dashboard already exists. Create new one?", vbYesNo + vbQuestion) = vbYes Then
            Application.DisplayAlerts = False
            ThisWorkbook.Sheets("Dashboard").Delete
            Application.DisplayAlerts = True
        Else
            Exit Sub
        End If
    End If
    
    ' Add new sheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
    ws.Name = "Dashboard"
    
    ' Setup layout
    With ws
        ' Create headers
        .Range("A1").Value = "Well List"
        .Range("C1").Value = "Select Well"
        .Range("E1").Value = "Gross Oil"
        .Range("F1").Value = "Net Oil"
        .Range("G1").Value = "Water Cut"
        .Range("H1").Value = "GOR"
        
        ' Create units row
        .Range("E2").Value = "(bbl)"
        .Range("F2").Value = "(bbl)"
        .Range("G2").Value = "(%)"
        .Range("H2").Value = "(Mscf/bbl)"
        
        ' Format headers
        With .Range("A1:H1")
            .Font.Bold = True
            .Interior.Color = RGB(191, 191, 191)  ' Light gray
            .HorizontalAlignment = xlCenter
        End With
        
        ' Format units row
        With .Range("E2:H2")
            .Font.Size = 9
            .Font.Italic = True
            .HorizontalAlignment = xlCenter
        End With
        
        ' Set column widths
        .Columns("A:A").ColumnWidth = 20
        .Columns("B:B").ColumnWidth = 5   ' Spacer
        .Columns("C:C").ColumnWidth = 15
        .Columns("D:D").ColumnWidth = 5   ' Spacer
        .Columns("E:E").ColumnWidth = 12
        .Columns("F:F").ColumnWidth = 12
        .Columns("G:G").ColumnWidth = 12
        .Columns("H:H").ColumnWidth = 12
        
        ' Add borders
        With .Range("A3:H100") ' Well list area
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlInsideHorizontal).Weight = xlHairline
        End With
        
        With .Range("E3:H3") ' KPI value area
            .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).Weight = xlThin
        End With
        
        ' Add instructions
        .Range("A10").Value = "INSTRUCTIONS:"
        .Range("A11").Value = "1. Place daily CSV in C:\WellDashboard\Raw\"
        .Range("A12").Value = "2. Click 'Refresh' button"
        .Range("A13").Value = "3. Select well from dropdown"
        .Range("A14").Value = "4. Report will auto-generate"
        
        ' Format instructions
        With .Range("A10")
            .Font.Bold = True
            .Font.Size = 12
        End With
        With .Range("A11:A14")
            .Font.Size = 10
            .Font.Color = RGB(100, 100, 100)
        End With
        
        ' Add button placeholder
        .Range("C10").Value = "REFRESH BUTTON HERE"
        .Range("C10").Interior.Color = RGB(146, 208, 80)  ' Green
        .Range("C10").Font.Bold = True
        .Range("C10").HorizontalAlignment = xlCenter
    End With
    
    ' Add actual button programmatically
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Range("C10").Left, ws.Range("C10").Top, _
                            ws.Range("C10").Width, ws.Range("C10").Height)
    With btn
        .Caption = "Refresh Data"
        .OnAction = "RefreshAll"  ' Assign to main macro
        .Name = "btnRefresh"
    End With
    
    ' Select cell C3 for dropdown
    ws.Range("C3").Select
    
    ' Create named ranges for easier reference
    ThisWorkbook.Names.Add Name:="SelectedWell", RefersTo:=ws.Range("C3")
    ThisWorkbook.Names.Add Name:="WellList", RefersTo:=ws.Range("A4:A100")
    
    MsgBox "Dashboard created successfully!" & vbCrLf & _
           "Remember to add the Refresh button macro.", vbInformation
End Sub




