Attribute VB_Name = "Module1"
Private Function NameNow(ByRef wbName, Optional addExt As String) As String

Dim timeStamp As String

timeStamp = CStr(Format(Now(), "yyyy.mmm.dd-hh.mm.ss"))

If wbName <> "" Then
    If Right(wbName, 1) = "." Then
    wbName = wbName & "-" & timeStamp & "." & addExt
    Else
        If InStrRev(wbName, ".") = 0 Then
            wbName = wbName & "-" & timeStamp & "." & addExt
        Else
            wbName = Left(wbName, InStrRev(wbName, ".") - 1) & "-" & timeStamp & "." & Right(wbName, Len(wbName) - InStrRev(wbName, "."))
        End If
    End If
Else
    GoTo ErrHan
End If

ErrExit:
   Exit Function

ErrHan:
   MsgBox "Error, nothing to rename."

End Function
Sub testNameNow()
Attribute testNameNow.VB_ProcData.VB_Invoke_Func = "e\n14"

Dim wbName As String
Dim addExt As String

addExt = Range("A3").Value
wbName = Range("A1").Value
Call NameNow(wbName, addExt)
Range("A1").Value = wbName

End Sub
