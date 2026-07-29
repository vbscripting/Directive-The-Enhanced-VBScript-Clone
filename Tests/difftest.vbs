' difftest.vbs - VBScript half of the differential test (the reference).
' Run:  cscript //nologo difftest.vbs > vbs.txt

' Every line prints:  <case padded to 30> <TypeName [value]>  or  ERR <number>
' Run BOTH, then diff. Any differing line is a place Directive does not mirror
' VBScript -- send the diff back.

On Error Resume Next
Dim r

Say "=== argument validation (should raise error 5) ==="
Err.Clear : r = Mid("abc", 0, 2)          : Chk "Mid(""abc"",0,2)", r
Err.Clear : r = Mid("abc", 1, -1)         : Chk "Mid(""abc"",1,-1)", r
Err.Clear : r = Mid("abc", 99)            : Chk "Mid(""abc"",99)", r
Err.Clear : r = Left("abc", -1)           : Chk "Left(""abc"",-1)", r
Err.Clear : r = Right("abc", -1)          : Chk "Right(""abc"",-1)", r
Err.Clear : r = Right("abc", 100)         : Chk "Right(""abc"",100)", r
Err.Clear : r = InStr(0, "abc", "b")      : Chk "InStr(0,""abc"",""b"")", r
Err.Clear : r = InStr(1, "abc", "b")      : Chk "InStr(1,""abc"",""b"")", r
Err.Clear : r = InStrRev("abc", "b", 0)   : Chk "InStrRev(""abc"",""b"",0)", r
Err.Clear : r = InStrRev("abcb", "b")     : Chk "InStrRev(""abcb"",""b"")", r
Err.Clear : r = Replace("aaa","a","b",0)  : Chk "Replace(...,start=0)", r
Err.Clear : r = Replace("aaa","","b")     : Chk "Replace(...,find="""")", r
Err.Clear : r = String(5, "")             : Chk "String(5,"""")", r
Err.Clear : r = String(3, 65)             : Chk "String(3,65)", r
Err.Clear : r = String(-1, "x")           : Chk "String(-1,""x"")", r
Err.Clear : r = Space(-1)                 : Chk "Space(-1)", r
Err.Clear : r = Log(0)                    : Chk "Log(0)", r
Err.Clear : r = Log(-1)                   : Chk "Log(-1)", r
Err.Clear : r = Sqr(-1)                   : Chk "Sqr(-1)", r

Say ""
Say "=== overflow / division ==="
Err.Clear : r = 2 ^ 10                    : Chk "2^10", r
Err.Clear : r = 2 ^ 10000                 : Chk "2^10000", r
Err.Clear : r = 0 ^ -1                    : Chk "0^-1", r
Err.Clear : r = 1 / 0                     : Chk "1/0", r
Err.Clear : r = 1 Mod 0                   : Chk "1 Mod 0", r
Err.Clear : r = CInt(99999999999.0)       : Chk "CInt(1e11)", r
Err.Clear : r = CLng(2147483648.0)        : Chk "CLng(2^31)", r

Say ""
Say "=== rounding / conversion ==="
Err.Clear : r = CInt(2.5)                 : Chk "CInt(2.5)", r
Err.Clear : r = CInt(3.5)                 : Chk "CInt(3.5)", r
Err.Clear : r = CInt(-2.5)                : Chk "CInt(-2.5)", r
Err.Clear : r = Round(2.5)                : Chk "Round(2.5)", r
Err.Clear : r = Round(-2.5)               : Chk "Round(-2.5)", r
Err.Clear : r = Int(-2.5)                 : Chk "Int(-2.5)", r
Err.Clear : r = Fix(-2.5)                 : Chk "Fix(-2.5)", r
Err.Clear : r = 7 \ 2                     : Chk "7 \ 2", r
Err.Clear : r = -7 \ 2                    : Chk "-7 \ 2", r
Err.Clear : r = 7.6 Mod 2                 : Chk "7.6 Mod 2", r
Err.Clear : r = CInt("12abc")             : Chk "CInt(""12abc"")", r
Err.Clear : r = Hex(-1)                   : Chk "Hex(-1)", r
Err.Clear : r = Oct(-1)                   : Chk "Oct(-1)", r

Say ""
Say "=== Empty / Null semantics ==="
Err.Clear : r = Empty + 1                 : Chk "Empty+1", r
Err.Clear : r = Empty & "x"               : Chk "Empty&""x""", r
Err.Clear : r = Null + 1                  : Chk "Null+1", r
Err.Clear : r = Null & "x"                : Chk "Null&""x""", r
Err.Clear : r = (Empty = 0)               : Chk "Empty=0", r
Err.Clear : r = (Empty = "")              : Chk "Empty=""""", r
Err.Clear : r = VarType(Nothing)          : Chk "VarType(Nothing)", r
Err.Clear : r = TypeName(Nothing)         : Chk "TypeName(Nothing)", r

Say ""
Say "=== arrays ==="
Err.Clear : r = UBound(Split("", ","))    : Chk "UBound(Split("""","",""))", r
Err.Clear : r = UBound(Split("a,,b", ",")): Chk "UBound(Split(""a,,b""))", r
Err.Clear : r = UBound(Array())           : Chk "UBound(Array())", r
Err.Clear : r = Join(Array(), "-")        : Chk "Join(Array(),""-"")", r
Err.Clear : r = UBound("nope")            : Chk "UBound(""nope"")", r

Say ""
Say "=== Dictionary: does reading a MISSING key create it? ==="
Dim d : Set d = CreateObject("Scripting.Dictionary")
d.Add "real", 1
Say Pad("Count after one Add", 30) & "= " & d.Count
Err.Clear : r = d("ghost")                : Chk "read d(""ghost"")", r
Say Pad("Count after that read", 30) & "= " & d.Count
Say Pad("Exists(""ghost"")", 30) & "= " & CStr(d.Exists("ghost"))
Err.Clear : d.Remove "nothere"            : Chk "Remove missing key", 0

Say ""
Say "=== RegExp ==="
Dim re : Set re = New RegExp
re.Pattern = "([unclosed"
Err.Clear : r = re.Test("x")              : Chk "invalid pattern .Test", r
re.Pattern = "\d+"
Err.Clear : r = re.Replace("a1b22", "#")  : Chk "Replace \d+ (Global=False)", r

Say ""
Say "=== END ==="

Function Pad(s, n)
    If Len(s) >= n Then Pad = s Else Pad = s & Space(n - Len(s))
End Function

Function Fmt(v)
    If IsObject(v) Then
        If v Is Nothing Then Fmt = "Nothing" Else Fmt = "obj:" & TypeName(v)
    ElseIf IsNull(v) Then
        Fmt = "Null"
    ElseIf IsEmpty(v) Then
        Fmt = "Empty"
    ElseIf IsArray(v) Then
        Fmt = "array ubound=" & UBound(v)
    Else
        Fmt = TypeName(v) & " [" & CStr(v) & "]"
    End If
End Function

Sub Chk(label, v)
    If Err.Number <> 0 Then
        Say Pad(label, 30) & "ERR " & Err.Number
    Else
        Say Pad(label, 30) & "= " & Fmt(v)
    End If
    Err.Clear
End Sub

Sub Say(s)
    WScript.Echo s
End Sub
