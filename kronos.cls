Option Explicit

Public Function KronosMonday(labelToFind As String, Optional required As Boolean = True) As String
  KronosMonday = Kronos_GetNameFromLabel(labelToFind, 1, required)
End Function

Public Function KronosTuesday(labelToFind As String, Optional required As Boolean = True) As String
  KronosTuesday = Kronos_GetNameFromLabel(labelToFind, 2, required)
End Function

Public Function KronosWednesday(labelToFind As String, Optional required As Boolean = True) As String
  KronosWednesday = Kronos_GetNameFromLabel(labelToFind, 3, required)
End Function

Public Function KronosThursday(labelToFind As String, Optional required As Boolean = True) As String
  KronosThursday = Kronos_GetNameFromLabel(labelToFind, 4, required)
End Function

Public Function KronosFriday(labelToFind As String, Optional required As Boolean = True) As String
  KronosFriday = Kronos_GetNameFromLabel(labelToFind, 5, required)
End Function

Public Function KronosSaturday(labelToFind As String, Optional required As Boolean = True) As String
  KronosSaturday = Kronos_GetNameFromLabel(labelToFind, 6, required)
End Function

Public Function KronosSunday(labelToFind As String, Optional required As Boolean = True) As String
  KronosSunday = Kronos_GetNameFromLabel(labelToFind, 7, required)
End Function

Private Function Kronos_GetNameFromLabel(labelToFind As String, dayOfWeek As Integer, Optional required As Boolean = True) As String
  ' Function looks for a label in a pasted Kronos staff report and returns the name of the person that matches the shift label
  
  Dim labelColumn As String  ' Specify which column in the pasted report the labels are located
  Dim rangeSize As Long      ' Contains the last row in the Kronos report to establish the size of the data range
  Dim searchCell As Range    ' Variable serach cell
  Dim rawData As String
  Dim searchData As String
  Dim searchResult As String
  Dim i As Variant
  
  'Set day of week column for searching
  Select Case dayOfWeek
    Case 1
      'Monday
      labelColumn = Range("MondayColumn").Value 'Labels are in column J
    Case 2
      'Tuesday
      labelColumn = Range("TuesdayColumn").Value 'Labels are in column N
    Case 3
      'Wednesday
      labelColumn = Range("WednesdayColumn").Value 'Labels are in column R
    Case 4
      'Thursday
      labelColumn = Range("ThursdayColumn").Value 'Labels are in column U
    Case 5
      'Friday
      labelColumn = Range("FridayColumn").Value 'Labels are in column X
    Case 6
      'Saturday
      labelColumn = Range("SaturdayColumn").Value 'Labels are in column AB
    Case 7
      'Sunday
      labelColumn = Range("SundayColumn").Value 'Labels are in column AF
    Case Else
      'Error
      Kronos_GetNameFromLabel = "Invalid day of week specified.  Must be from 1 - 7."
  End Select
  
  'Establish how big the range of names is
  rangeSize = Worksheets("Kronos Data").Cells(Rows.Count, 1).End(xlUp).Row ' Establishes how many rows of data are in Column 1 (A)
  
  
  ' Loop through label column looking for a match
  On Error GoTo ErrHand
  For i = 1 To rangeSize
  
    ' rawData is the interated column containing the shift labels
    rawData = LCase(Worksheets("Kronos Data").Range(labelColumn & CStr(i)).Value)
    searchData = LCase(labelToFind)
  
    If InStr(rawData, searchData) > 0 Then  'If the search label matches the current cell
      
      'Label has been matched.  Get name from column A on the same row and assign the value to searchResult variable
      searchResult = Worksheets("Kronos Data").Range("A" & CStr(i)).Value
      
      If searchResult = "" Then
        'Search result returned a blank cell, likely as a result of a multi line entry in the kronos export
        Kronos_GetNameFromLabel = "<Not Found>"
      Else
        'Return the staffname from column A
        Kronos_GetNameFromLabel = searchResult
      End If
      Exit Function 'Break point
    End If
  Next i
  
  If required = False Then
    'Return no value if this is passed as a not required shift
    Kronos_GetNameFromLabel = ""
  Else
    'This handles spare staff shifts by not returning anything if they are not found
    'Search string has not been found.  Handle not found output
    If InStr(labelToFind, " ") > 0 Then
      Kronos_GetNameFromLabel = ""  ' No match was found but the shift was expected to be site staff.  Return empty string.
    Else
      Kronos_GetNameFromLabel = "<Not Found>" ' No match was found for a standard shift, return N/F
    End If
  End If
  
  Exit Function 'Break point
  
ErrHand:
  Kronos_GetNameFromLabel = Err.Description
End Function

Public Function getSpareStaff(dayOfWeek As Integer) As String
  ' Function looks for a label in a pasted Kronos staff report and returns the name of the person that matches the shift label
  
  Dim labelColumn As String  ' Specify which column in the pasted report the labels are located
  Dim nameOffset As Integer  ' Specific how many columns offset the names are from the labels, negative integers to the left, poitive integers to the right
  Dim rangeSize As Long      ' Contains the last row in the Kronos report to establish the size of the data range
  Dim searchCell As Range    ' Variable serach cell
  Dim rawData As String
  Dim searchData As String
  Dim result As String
  Dim cell As Variant
  Dim i As Variant
  
  'Set day of week column for searching
  Select Case dayOfWeek
    Case 1
      'Monday
      labelColumn = Range("MondayColumn").Value 'Labels are in column J
    Case 2
      'Tuesday
      labelColumn = Range("TuesdayColumn").Value 'Labels are in column N
    Case 3
      'Wednesday
      labelColumn = Range("WednesdayColumn").Value 'Labels are in column R
    Case 4
      'Thursday
      labelColumn = Range("ThursdayColumn").Value 'Labels are in column U
    Case 5
      'Friday
      labelColumn = Range("FridayColumn").Value 'Labels are in column X
    Case 6
      'Saturday
      labelColumn = Range("SaturdayColumn").Value 'Labels are in column AB
    Case 7
      'Sunday
      labelColumn = Range("SundayColumn").Value 'Labels are in column AF
    Case Else
      'Error
      getSpareStaff = "Invalid day of week specified.  Must be from 1 - 7."
  End Select
  
  rangeSize = Worksheets("Kronos Data").Cells(Rows.Count, 1).End(xlUp).Row ' Establishes how many rows of data are in Column 1
  nameOffset = -9 'Names are in Column A -9 columns to the left
  
  'Loop through label column looking for a match
  'On Error GoTo errhand
  For Each cell In Worksheets("Kronos Data").Range(labelColumn & "1:" & labelColumn & rangeSize).Cells
    rawData = Worksheets("Kronos Data").Range(labelColumn & cell.Row).Value
    
    If InStr(rawData, " ") > 0 Then
      result = result & Worksheets("Kronos Data").Range("A" & cell.Row).Value & " - " & cell.Value & vbNewLine
    End If
  Next
  
  getSpareStaff = result  ' Return list of non essential staff
  Exit Function
  
ErrHand:
  getSpareStaff = Err.Description
End Function

Public Function KronosNameLookup(name As String, dayOfWeek As Integer) As String
  Dim labelColumn As String  ' Specify which column in the pasted report the labels are located
  Dim rangeSize As Long      ' Contains the last row in the Kronos report to establish the size of the data range
  Dim searchCell As Range    ' Variable serach cell
  Dim rawData As String
  Dim searchData As String
  Dim searchResult As String
  Dim i As Variant
  
  'Set day of week column for searching
  Select Case dayOfWeek
    Case 1
      'Monday
      labelColumn = Range("MondayColumn").Value 'Labels are in column J
    Case 2
      'Tuesday
      labelColumn = Range("TuesdayColumn").Value 'Labels are in column N
    Case 3
      'Wednesday
      labelColumn = Range("WednesdayColumn").Value 'Labels are in column R
    Case 4
      'Thursday
      labelColumn = Range("ThursdayColumn").Value 'Labels are in column U
    Case 5
      'Friday
      labelColumn = Range("FridayColumn").Value 'Labels are in column X
    Case 6
      'Saturday
      labelColumn = Range("SaturdayColumn").Value 'Labels are in column AB
    Case 7
      'Sunday
      labelColumn = Range("SundayColumn").Value 'Labels are in column AF
    Case Else
      'Error
      KronosNameLookup = "Invalid day of week specified.  Must be from 1 - 7."
  End Select
  
  'Establish how big the range of names is
  rangeSize = Worksheets("Kronos Data").Cells(Rows.Count, 1).End(xlUp).Row ' Establishes how many rows of data are in Column 1 (A)
  
  
  ' Loop through label column looking for a match
  On Error GoTo ErrHand
  For i = 1 To rangeSize
    rawData = Worksheets("Kronos Data").Range("A" & CStr(i)).Value
    searchData = name
  
    If rawData = name Then  'If the search label matches the current cell
      searchResult = Worksheets("Kronos Data").Range(labelColumn & CStr(i)).Value 'Label has been found, get name and return the function
      If searchResult <> "" And InStr(searchResult, "LVE") <= 0 And InStr(searchResult, "NW") <= 0 Then
        'Search result returned a blank cell, likely as a result of a multi line entry in the kronos export
        KronosNameLookup = searchData
      Else
        If InStr(searchResult, "LVE") > 0 Then
          KronosNameLookup = "Leave"
        Else
          If InStr(searchResult, "NW") > 0 Then
            KronosNameLookup = "NW"
          Else
            If searchResult = "" Then
              KronosNameLookup = "RDO"
            End If
          End If
        End If
      End If
      Exit Function
    End If
  Next i
  
  KronosNameLookup = "<Not Found>"  ' Return an empty string if no match is found
  Exit Function
  
ErrHand:
  KronosNameLookup = Err.Description
End Function

Public Sub PasteKronosData()
    Dim ws As Worksheet
    
    Set ws = Worksheets("Kronos Data")
    ws.Range("A1").PasteSpecial xlPasteValues
    
    Call fixReportFormat
    Call refreshAllSheets
End Sub

Private Sub fixReportFormat()
  Dim cell As Range
  Dim i As Integer
  Dim ws As Worksheet
  Dim day As Integer
  
  Set ws = Worksheets("Kronos Data")

  For i = 1 To 100
    If ws.Cells(10, i).Value <> "" Then
        day = Weekday(ws.Cells(10, i).Value, vbMonday)
        
        Select Case day
            Case 1
                Range("MondayColumn").Value = Split(ws.Cells(10, i).Address, "$")(1)
            Case 2
                Range("TuesdayColumn").Value = Split(ws.Cells(10, i).Address, "$")(1)
            Case 3
                Range("WednesdayColumn").Value = Split(ws.Cells(10, i).Address, "$")(1)
            Case 4
                Range("ThursdayColumn").Value = Split(ws.Cells(10, i).Address, "$")(1)
            Case 5
                Range("FridayColumn").Value = Split(ws.Cells(10, i).Address, "$")(1)
            Case 6
                Range("SaturdayColumn").Value = Split(ws.Cells(10, i).Address, "$")(1)
            Case 7
                Range("SundayColumn").Value = Split(ws.Cells(10, i).Address, "$")(1)
            Case Else
                Debug.Print (ws.Cells(10, i).Address)
        End Select
    End If
  Next i
End Sub

Public Function SCOLookup(dayOfWeek As Integer, sco1 As String, sco2 As String, Optional sco3 As String) As String
  Dim labelColumn As String  ' Specify which column in the pasted report the labels are located
  Dim rangeSize As Long      ' Contains the last row in the Kronos report to establish the size of the data range
  Dim searchCell As Range    ' Variable serach cell
  Dim rawData As String
  Dim searchData As String
  Dim searchResult As String
  Dim sco1shift As String, sco2shift As String, sco3shift As String
  Dim i As Variant
  
  'Set day of week column for searching
  Select Case dayOfWeek
    Case 1
      'Monday
      labelColumn = Range("MondayColumn").Value
    Case 2
      'Tuesday
      labelColumn = Range("TuesdayColumn").Value
    Case 3
      'Wednesday
      labelColumn = Range("WednesdayColumn").Value
    Case 4
      'Thursday
      labelColumn = Range("ThursdayColumn").Value
    Case 5
      'Friday
      labelColumn = Range("FridayColumn").Value
    Case 6
      'Saturday
      labelColumn = Range("SaturdayColumn").Value
    Case 7
      'Sunday
      labelColumn = Range("SundayColumn").Value
    Case Else
      'Error
      SCOLookup = "Invalid day of week specified.  Must be from 1 - 7."
  End Select
  
  'Establish how big the range of names is
  rangeSize = Worksheets("Kronos Data").Cells(Rows.Count, 1).End(xlUp).Row ' Establishes how many rows of data are in Column 1 (A)
  
  ' Get shift for SCO1
  On Error GoTo ErrHand
  For i = 1 To rangeSize
    rawData = LCase(Worksheets("Kronos Data").Range("A" & CStr(i)).Value)
    searchData = LCase(sco1)
    
    If rawData = searchData Then  'If the search label matches the current cell
      sco1shift = Worksheets("Kronos Data").Range(labelColumn & CStr(i)).Value 'Label has been found, get name and return the function
      Exit For
    End If
  Next i
  
  ' Get shift for SCO2
  For i = 1 To rangeSize
    rawData = LCase(Worksheets("Kronos Data").Range("A" & CStr(i)).Value)
    searchData = LCase(sco2)
  
    If rawData = searchData Then  'If the search label matches the current cell
      sco2shift = Worksheets("Kronos Data").Range(labelColumn & CStr(i)).Value 'Label has been found, get name and return the function
      Exit For
    End If
  Next i
  
  ' Get shift for SCO3 if it was passed in the first place
  If sco3 <> "" Then
    For i = 1 To rangeSize
      rawData = LCase(Worksheets("Kronos Data").Range("A" & CStr(i)).Value)
      searchData = LCase(sco1)
    
      If rawData = searchData Then  'If the search label matches the current cell
        sco3shift = Worksheets("Kronos Data").Range(labelColumn & CStr(i)).Value 'Label has been found, get name and return the function
        Exit For
      End If
    Next i
  End If

  Debug.Print (sco1 & " - " & sco1shift)
  Debug.Print (sco2 & " - " & sco2shift)
  Debug.Print (sco3 & " - " & sco3shift)

  If InStr(sco1shift, "0700-1900") > 0 Or InStr(sco1shift, "0800-1800") > 0 And InStr(sco1shift, " ") = 0 Then
    SCOLookup = sco1
    Exit Function
  End If
  
  If InStr(sco2shift, "0700-1900") > 0 Or InStr(sco2shift, "0800-1800") > 0 And InStr(sco2shift, " ") = 0 Then
    SCOLookup = sco2
    Exit Function
  End If
  
  If sco3 <> "" Then
    If InStr(sco3shift, "0700-1900") > 0 Or InStr(sco3shift, "0800-1800") > 0 And InStr(sco3shift, " ") = 0 Then
      SCOLookup = sco3
      Exit Function
    End If
  End If
  
  SCOLookup = "<Not Found>"
  
  Exit Function
  
ErrHand:
  SCOLookup = Err.Description
End Function

Public Sub clearKronosData()
  'Clears all data from the Kronos Data sheet.  Keyboard shortcut Ctrl+Alt+F12
  Worksheets("Kronos Data").Cells.Clear
  Call refreshAllSheets
  MsgBox "Kronos Data has been deleted", vbInformation
End Sub

Private Sub refreshAllSheets()
  Dim ws As Worksheet
  
  For Each ws In ThisWorkbook.Worksheets
    ws.EnableCalculation = False
    ws.EnableCalculation = True
    ws.Calculate
  Next ws
End Sub

