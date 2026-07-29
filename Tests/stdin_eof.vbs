' stdin_eof.vbs

' Run BOTH with stdin at end-of-file, which is the ONLY case that can error:
'   cscript //nologo stdin_eof.vbs < nul > vbs_stdin.txt
'   directive.exe stdin_eof.ds < nul > ds_stdin.txt
' Interactively (double-click, or just running it and typing) there is no EOF,
' so none of these raise -- that is why your original Read(1) pause never failed.

On Error Resume Next
Dim v

Err.Clear : v = WScript.StdIn.AtEndOfStream : Rep "AtEndOfStream", v
Err.Clear : v = WScript.StdIn.Read(1)       : Rep "Read(1)", v
Err.Clear : v = WScript.StdIn.ReadLine      : Rep "ReadLine", v
Err.Clear : v = WScript.StdIn.ReadAll       : Rep "ReadAll", v
Err.Clear : WScript.StdIn.Skip 1            : Rep "Skip(1)", "(no return)"
Err.Clear : WScript.StdIn.SkipLine          : Rep "SkipLine", "(no return)"
Err.Clear : v = WScript.StdIn.Line          : Rep "Line", v
Err.Clear : v = WScript.StdIn.Column        : Rep "Column", v
Say "=== END ==="

Sub Rep(label, v)
    Dim pad : pad = label
    Do While Len(pad) < 18
        pad = pad & " "
    Loop
    If Err.Number <> 0 Then
        Say pad & "ERR " & Err.Number
    Else
        Say pad & "= [" & CStr(v) & "]"
    End If
    Err.Clear
End Sub

Sub Say(s)
    WScript.Echo s
End Sub
