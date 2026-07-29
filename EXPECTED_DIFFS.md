# Known-intentional differences from VBScript

Directive is verified against real `vbscript.dll` by running each `difftest*.vbs`
under `cscript` and the matching `.ds` under `directive.exe`, then diffing. Every
line is expected to match **except** the ones below, which are deliberate. Anything
else that differs is a bug worth reporting.

Run them all:

```
cscript //nologo difftest.vbs   > vbs.txt   & directive.exe difftest.ds   > ds.txt
cscript //nologo difftest2.vbs  > vbs2.txt  & directive.exe difftest2.ds  > ds2.txt
cscript //nologo difftest3.vbs  > vbs3.txt  & directive.exe difftest3.ds  > ds3.txt
cscript //nologo stdin_eof.vbs  < nul > vbs_stdin.txt & directive.exe stdin_eof.ds < nul > ds_stdin.txt
cscript //nologo withdot_a.vbs  > vbs_wa.txt & directive.exe withdot_a.ds > ds_wa.txt
cscript //nologo withdot_b.vbs  > vbs_wb.txt & directive.exe withdot_b.ds > ds_wb.txt
```

## 1. Whole numbers above 2,147,483,647 are exact (difftest2, difftest3)

VBScript's numeric ladder is Integer -> Long -> Double, so a whole number past
2^31 becomes a Double and stops holding whole numbers past 2^53. Directive adds a
LongLong rung: Integer -> Long -> **LongLong** -> Double.

| Expression | VBScript | Directive |
|---|---|---|
| `TypeName(2147483648)` | `Double` | `LongLong` |
| `TypeName(2000000000 + 2000000000)` | `Double` | `LongLong` |
| `TypeName(2147483647 + 1)` | `Double` | `LongLong` |
| `TypeName(3000000000 * 3)` | `Double` | `LongLong` |
| `CStr(1000000000000000000)` | `1E+18` | `1000000000000000000` |

The numeric values are identical or more accurate; `9007199254740992 + 1` is
`9007199254740993` here and `9.00719925474099E+15` in VBScript. Past 64 bits
Directive promotes to Double, exactly as VBScript does past 32.

**The one thing to watch:** the last row. A whole number needing 16 or more digits
now prints its digits instead of scientific notation. If a script writes such a
value to a file another program parses, either format it explicitly or use
`CDbl(x)` to get VBScript's rendering.

## 2. Integer operators work past the Long ceiling (difftest4)

VBScript converts the operands of `\`, `Mod`, `And`/`Or`/`Xor`/`Not`, `Hex` and
`Oct` to a Long, so anything above 2,147,483,647 raises **error 6 (Overflow)**.
With the LongLong rung those values are real integers, so Directive computes them:

| Expression | VBScript | Directive |
|---|---|---|
| `Hex(2147483648)` | error 6 | `80000000` |
| `Hex(4000000000)` | error 6 | `EE6B2800` |
| `Oct(4000000000)` | error 6 | `35632624000` |
| `4000000000 \ 3` | error 6 | `1333333333` |
| `4000000000 Mod 3` | error 6 | `1` |
| `4000000000 And 255` | error 6 | `0` |
| `1000000000000000000 \ 1` | error 6 | `1000000000000000000` |

This follows from the rung rather than being a separate decision: a type you can
hold but cannot divide or mask would be incoherent. It cannot break a ported
script either — any script hitting these already failed under VBScript.

## 3. COM objects keep their parent alive (Illustrator)

VBScript releases the intermediate objects of an expression when the statement
ends. For a host whose accessors are thin wrappers over a parent — Illustrator's
`CharacterAttributes` over a `TextRange`, for instance — that leaves a dangling
reference, and the next line can take the host process down with it. This script
**crashes Illustrator** under VBScript:

```vbs
Set appRef = CreateObject("Illustrator.Application")
Set docRef = appRef.ActiveDocument
Set txt    = docRef.Layers("pictogram").TextFrames("tekst")
Set ca     = txt.TextRange.CharacterAttributes   ' TextRange released here
ca.Size    = 12                                  ' ...and used here
```

Directive gives every COM object a strong reference to the object it came from, so
lifetime is structural and survives across statements. The same script runs
correctly. Confirmed against real Illustrator.

Mirroring VBScript here would mean reproducing a crash, so this divergence is
deliberate and permanent.

## 4. ChrB is not implemented (difftest5)

`ChrB` returns a single raw byte — `LenB(ChrB(65))` is 1, not 2. Directive holds
strings as UTF-8, where half of a character is not a representable value, so `ChrB`
raises error 500 rather than returning something quietly wrong.

`LenB`, `AscB` and `InStrB` are exact, since they return numbers. `LeftB`, `MidB`
and `RightB` are exact for even byte counts and raise error 5 for odd ones, which
would split a character. In practice text is whole characters, so even counts are
the normal case.

## 5. Non-ASCII characters survive (difftest3)

`FormatCurrency(1234.5)` on a euro locale gives `? 1.234,50` under cscript and
`€ 1.234,50` under Directive. cscript downconverts the output to the console
codepage and loses the character; Directive is UTF-8 throughout, so it is preserved
even when redirected to a file. Directive is the correct one here.

## 6. Platform limits outside Windows

On Linux and macOS there is no LCID to consult, so Directive uses a built-in
locale table (1033, 2057, 2067, 1043, 1031, 1036, 1040, 3082) defaulting to 1033.
Month and day names are English, times are 24-hour, and the long date form is
`March 4, 2020`. On Windows all of these come from the OS — `GetLocaleInfoW`,
`GetDateFormatW`, `GetTimeFormatW` — and match VBScript exactly, so comparisons
must be run **on Windows**.
