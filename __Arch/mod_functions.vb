Option Explicit


Public Function sum_range(my_range As Range) As Double
    
    Dim cell As Range
    
    sum_range = 0
    
    For Each cell In my_range
        sum_range = sum_range + cell.Value
    Next
    
End Function

Sub PrintLines()

    ActiveSheet.DisplayPageBreaks = Not ActiveSheet.DisplayPageBreaks
    
End Sub

Sub CheckReferences()
' Check for possible missing or erroneous links in
' formulas and list possible errors in a summary sheet

  Dim iSh           As Integer
  Dim sShName       As String
  Dim c             As Range
  Dim rng           As Range
  Dim i             As Integer
  Dim j             As Integer
  Dim sChr          As String
  Dim addr          As String
  Dim sFormula      As String
  Dim scVal         As String
  Dim lNewRow       As Long
  Dim vHeaders      As Variant

  vHeaders = Array("Sheet Name", "Cell", "Cell Value", "Formula")
  'check if 'Summary' worksheet is in workbook
  'and if so, delete it
  With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .Calculation = xlCalculationManual
  End With

  For i = 1 To Worksheets.Count
    If Worksheets(i).Name = "Summary" Then
      Worksheets(i).Delete
    End If
  Next i

  iSh = Worksheets.Count

  'create a new summary sheet
    Sheets.Add After:=Sheets(iSh)
    Sheets(Sheets.Count).Name = "Summary"
  With Sheets("Summary")
    Range("A1:D1") = vHeaders
  End With
  lNewRow = 2

  ' this will not work if the sheet is protected,
  ' assume that sheet should not be changed; so ignore it
  On Error Resume Next

  For i = 1 To iSh
    sShName = Worksheets(i).Name
    Application.Goto Sheets(sShName).Cells(1, 1)
    Set rng = Cells.SpecialCells(xlCellTypeFormulas, 23)

    For Each c In rng
      addr = c.Address
      sFormula = c.Formula
      scVal = c.Text

      For j = 1 To Len(c.Formula)
        sChr = Mid(c.Formula, j, 1)

        If sChr = "[" Or sChr = "!" Or IsError(c) Then
          'write values to summary sheet
          With Sheets("Summary")
            .Cells(lNewRow, 1) = sShName
            .Cells(lNewRow, 2) = addr
            .Cells(lNewRow, 3) = scVal
            .Cells(lNewRow, 4) = "'" & sFormula
          End With
          lNewRow = lNewRow + 1
          Exit For
        End If
      Next j
    Next c
  Next i


' housekeeping
  With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlCalculationAutomatic
  End With

' tidy up
  Sheets("Summary").Select
  Columns("A:D").EntireColumn.AutoFit
  Range("A1:D1").Font.Bold = True
  Range("A2").Select
  
End Sub

Public Sub print_array(my_array As Variant)
    Dim counter As Integer
    
    For counter = LBound(my_array) To UBound(my_array)
        Debug.Print counter & " --> " & my_array(counter)
    Next counter
    
End Sub

Public Function get_last_day_of_month(my_date As Date) As Date
    get_last_day_of_month = DateSerial(Year(my_date), Month(my_date) + 1, 0)
End Function

Public Function add_months(my_date As Date, i_month As Integer) As Date
    
    add_months = get_last_day_of_month(DateAdd("m", i_month, my_date))

End Function

Public Sub ShowMeTheNames()
    Dim i As Integer
    
    For i = 1 To ActiveWorkbook.Names().Count
        
        Debug.Print vbCrLf & ActiveWorkbook.Names(i).Name
        Debug.Print ActiveWorkbook.Names(i).RefersTo
    Next i
    
End Sub

Public Sub Normal()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    Application.DisplayStatusBar = True
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Public Function last_row_with_data(ByVal lng_column_number As Long, shCurrent As Variant) As Long
    
    last_row_with_data = shCurrent.Cells(Rows.Count, lng_column_number).End(xlUp).Row
    
End Function

Sub CopyValues(rngSource As Range, rngTarget As Range)
 
    rngTarget.Resize(rngSource.Rows.Count, rngSource.Columns.Count).Value = rngSource.Value
 
End Sub

