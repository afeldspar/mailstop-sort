'$ Input to this function: WScript.Arguments
'$ Output:  An array of arrays;
'$               each sub-array starts with a string indicating role
'$               and each following entry lists the files holding that role

Public Function fileRoleAssign (objArgs)

  Dim filenameArray()
  ReDim Preserve filenameArray(objArgs.Count - 1)

  Dim filenamesString
  filenamesString = ""
  
  For I = 0 To (objArgs.Count - 1)
    filenameArray(I) = objArgs(I)
    filenamesString = filenamesString & (I + 1) & ": " & objArgs(I) & vbCrLf & vbCrlf
  Next
  
  WScript.Echo filenamesString       ' testing-only code
  
  title = "Select file of user data"
  message = "The following files were dropped on the script:" & vbCrlf & vbCrlf & filenamesString
  message = message & vbCrlf & vbCrlf & "Please select one by number"
  defaultValue = ""
  
  Dim myValue
  myValue = InputBox(message, title, defaultValue)
  
End Function
