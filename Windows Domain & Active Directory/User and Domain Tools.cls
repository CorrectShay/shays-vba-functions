Function IsMember(strDomain As String, strGroup As String, strMember As String) As Boolean
  Dim grp As Object
  Dim strPath As String

  strPath = "WinNT://" & strDomain & "/"
  Set grp = GetObject(strPath & strGroup & ",group")
  IsMember = grp.IsMember(strPath & strMember)
End Function

Function GetCurrentUser() As String
    GetCurrentUser = Environ("USERNAME")
End Function

Function GetCurrentDomain() As String
    GetCurrentDomain = Environ("USERDOMAIN")
End Function
