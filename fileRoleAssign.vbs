'$ Input to this function: WScript.Arguments
'$ Output:  An array of arrays;
'$               each sub-array starts with a string indicating role
'$               and each following entry lists the files holding that role

Public Function fileRoleAssign (objArgs)
  For I = 0 To objArgs.Count - 1
    WScript.Echo "**" & objArgs(I)
  Next
End Function
