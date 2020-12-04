Function IsMember(strDomain As String, strGroup As String, strMember As String) As Boolean
  ' This fucntion returns a boolean based on if the user (strMember) is a member of an active directory group (strGroup).
  ' It requires the current domain which can be obtained with the GetCurrentDomain function.
  
  Dim grp As Object
  Dim strPath As String

  strPath = "WinNT://" & strDomain & "/"
  Set grp = GetObject(strPath & strGroup & ",group")
  IsMember = grp.IsMember(strPath & strMember)
End Function

Function GetCurrentUser() As String
    ' Returns the current Windows username as a string
    
    GetCurrentUser = Environ("USERNAME")
End Function

Function GetCurrentDomain() As String
    ' Returns the current Windows domain as a string
    
    GetCurrentDomain = Environ("USERDOMAIN")
End Function
