
Function TylkoCyfry(s As String) As String
With CreateObject("vbscript.regexp")
    .Pattern = "\D"
    .Global = True
    OnlyDigits = .Replace(s, "")
End With
End Function
