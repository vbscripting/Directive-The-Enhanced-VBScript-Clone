' withdot_a.vbs

' Does a BARE leading-dot argument to a paren-less call parse at all?
' If this file prints nothing, it is a compile error in that dialect.
Dim d : Set d = CreateObject("Scripting.Dictionary")
d.Add "k", "v"
With d
    WScript.Echo "bare leading dot:"
    WScript.Echo .Count
End With
WScript.Echo "=== END A ==="
