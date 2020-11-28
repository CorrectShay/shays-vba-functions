Sub ImportIntranetHeader()
  Dim objIE As Object
  Dim msg As String
  Set objIE = New InternetExplorerMedium
  
  objIE.Visible = False
  
  'On Error Resume Next
  objIE.Navigate 'Website URL as string here
  
  Do
    DoEvents
  Loop Until objIE.ReadyState = 4
  
  'Get code from tags here and assign to variables
  
  Set tables = objIE.Document.getelementsbytagname("TABLE")
  
  For t = 0 To (tables.Length - 1)

    For r = 0 To (tables(t).Rows.Length - 1)
        For c = 0 To (tables(t).Rows(r).Cells.Length - 1)
            ThisWorkbook.Worksheets(1).Cells(r + 1, c + 1) = tables(t).Rows(r).Cells(c).innerText
        Next c
    Next r
  Next t
  
  'objIE.Quit
  MsgBox msg
  Set objIE = Nothing
End Sub
