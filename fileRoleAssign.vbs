'$ Input to this function: WScript.Arguments
'$ Output:  An array of arrays;
'$               each sub-array starts with a string indicating role
'$               and each following entry lists the files holding that role

Public Function fileRoleAssign (objArgs)

  '$ We'll start by hard-coding the limit of files at 8.
  '$ Any problems we run into with that are FAR in the future.
  Dim filenameArray(8)
  Dim filenamesString
  filenamesString = ""
  
  For I = 0 To (objArgs.Count - 1)
    filenameArray(I) = objArgs(I)
    filenamesString = filenamesString & I & ": " & objArgs(I) & vbCrLf
  Next
  
  WScript.Echo filenamesString
End Function
