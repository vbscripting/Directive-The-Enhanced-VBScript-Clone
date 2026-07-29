' withdot_b.vbs

' The workarounds, kept separate so a parse error in the bare form cannot hide them.
Dim d : Set d = CreateObject("Scripting.Dictionary")
d.Add "k", "v"
Dim n
With d
    WScript.Echo "parenthesised:"
    WScript.Echo (.Count)
    n = .Count
    WScript.Echo "assignment: " & n
    WScript.Echo "concatenated: " & .Count
End With
WScript.Echo "=== END B ==="
