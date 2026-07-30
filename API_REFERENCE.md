# Directive — Complete API Reference

A checklist of everything the interpreter exposes: every built-in function with
its parameters, all statements, operators, constants, and objects. Function and
member names were extracted directly from the source, so this matches what the
engine actually implements.

**Legend**
- `[arg]` = optional argument
- ✔ tested = exercised by the regression/edge suites on this platform
- 🪟 Windows-only = compiles and links, runs only on real Windows (raises a clean
  error on non-Windows). Items marked 🪟 have **not** been runtime-tested.

---

## 1. Running scripts

```
directive script.ds          # run a file (console build)
directive script.ds a b c    # run a file, passing arguments (Directive.Arguments)
directive -e "directive.echo 1+1"   # run an inline snippet
directivew script.ds         # GUI build: Echo/Input via message boxes
```

---

## 2. Built-in functions (101)

### Strings
| Function | Parameters | Notes |
|---|---|---|
| `Len` | `(string \| var)` | length; `Len(array)` not supported — use `UBound` |
| `Left` | `(string, count)` | |
| `Right` | `(string, count)` | |
| `Mid` | `(string, start, [length])` | 1-based; also a **statement** (see §4) |
| `InStr` | `([start,] haystack, needle, [compare])` | 1-based; 0 if not found |
| `InStrRev` | `(haystack, needle, [start, [compare]])` | |
| `Replace` | `(string, find, repl, [start, [count, [compare]]])` | |
| `Split` | `(string, [delim, [count, [compare]]])` | ✔ `count`/`compare` supported |
| `Join` | `(array, [delim])` | |
| `UCase` / `LCase` | `(string)` | |
| `Trim` / `LTrim` / `RTrim` | `(string)` | |
| `Space` | `(count)` | |
| `String` | `(count, char)` | char may be code or 1-char string |
| `StrReverse` | `(string)` | |
| `StrComp` | `(a, b, [compare])` | returns -1/0/1 |
| `Filter` | `(array, match, [include, [compare]])` | include=False excludes matches |
| `Asc` / `AscW` | `(string)` | |
| `Chr` / `ChrW` | `(code)` | |

### Conversion
| Function | Parameters | Notes |
|---|---|---|
| `CInt` | `(expr)` | banker's rounding |
| `CLng` | `(expr)` | banker's rounding |
| `CByte` | `(expr)` | 0–255 |
| `CDbl` / `CSng` | `(expr)` | |
| `CCur` | `(expr)` | currency (stored as double) |
| `CStr` | `(expr)` | |
| `CBool` | `(expr)` | |
| `CDate` | `(expr)` | string/number → date serial |
| `Hex` | `(number)` | 32-bit two's complement for negatives |
| `Oct` | `(number)` | |
| `Int` | `(number)` | floor (toward −∞) |
| `Fix` | `(number)` | truncate (toward 0) |
| `Round` | `(number, [digits])` | banker's (round-half-to-even) |

### Math
| Function | Parameters |
|---|---|
| `Abs` | `(number)` |
| `Sgn` | `(number)` |
| `Sqr` | `(number)` |
| `Atn` / `Cos` / `Sin` / `Tan` | `(number)` — radians |
| `Exp` / `Log` | `(number)` — natural |
| `Rnd` | `([seed])` |
| `Randomize` | `([seed])` — statement-style |
| `RGB` | `(r, g, b)` → 0xBBGGRR long |

### Dates
| Function | Parameters | Notes |
|---|---|---|
| `Now` / `Date` / `Time` | `()` | |
| `Year` / `Month` / `Day` | `(date)` | |
| `Hour` / `Minute` / `Second` | `(date)` | |
| `Weekday` | `(date, [firstDayOfWeek])` | |
| `MonthName` | `(n, [abbrev])` | |
| `WeekdayName` | `(n, [abbrev, [firstDayOfWeek]])` | |
| `DateSerial` | `(year, month, day)` | |
| `TimeSerial` | `(hour, minute, second)` | |
| `DateValue` | `(expr)` | ✔ date part only |
| `TimeValue` | `(expr)` | ✔ time-of-day only |
| `DateAdd` | `(interval, number, date)` | intervals `yyyy,m,d,ww,h,n,s` |
| `DateDiff` | `(interval, date1, date2, [fdow, [fwoy]])` | |
| `DatePart` | `(interval, date, [fdow, [fwoy]])` | |

### Formatting
| Function | Parameters |
|---|---|
| `FormatNumber` | `(number, [decimals, [leadingDigit, [parensNeg, [groupDigits]]]])` |
| `FormatCurrency` | `(number, [decimals, …])` |
| `FormatPercent` | `(number, [decimals, …])` |
| `FormatDateTime` | `(date, [namedFormat])` — 0 General,1 LongDate,2 ShortDate,3 LongTime,4 ShortTime |

### Arrays
| Function | Parameters | Notes |
|---|---|---|
| `Array` | `(v1, v2, …)` | build a variant array |
| `UBound` | `(array, [dimension])` | |
| `LBound` | `(array, [dimension])` | always 0 |
| `Erase` | `(array)` | **statement** (see §4) |

### Type inspection
| Function | Parameters |
|---|---|
| `IsArray` / `IsDate` / `IsEmpty` / `IsNull` / `IsNumeric` / `IsObject` | `(expr)` |
| `TypeName` | `(expr)` → `"Long"`,`"Double"`,`"String"`,`"Boolean"`,`"Empty"`,`"Null"`,`"Object"`,class name, etc. |
| `VarType` | `(expr)` → numeric VbVarType code |

### Evaluation & references
| Function | Parameters | Notes |
|---|---|---|
| `Eval` | `(string)` | evaluate an expression |
| `Execute` | `(string)` | run statements in local scope |
| `ExecuteGlobal` | `(string)` | run statements in global scope |
| `GetRef` | `(procName)` | returns a callable reference |

### Engine / locale
| Function | Parameters |
|---|---|
| `ScriptEngine` | `()` → `"Directive"` |
| `ScriptEngineMajorVersion` / `ScriptEngineMinorVersion` / `ScriptEngineBuildVersion` | `()` |
| `GetLocale` / `SetLocale` | `([locale])` |

### Dialogs
| Function | Parameters | Notes |
|---|---|---|
| `MsgBox` | `(prompt, [buttons, [title]])` | console prints; GUI shows a box |
| `InputBox` | `(prompt, [title, [default]])` | |

> Clipboard access moved from the flat `ClipPut`/`ClipGet` functions to the
> `New Clipboard` object (`.Set` / `.Get` / `.Clear`); see §6.10.

### COM (Windows)
| Function | Parameters | Notes |
|---|---|---|
| `CreateObject` | `(progID)` | 🪟 real COM/ActiveX; err 429 elsewhere |
| `GetObject` | `([path, [progID]])` | 🪟 |

---

### Strings and Unicode

Strings are held as UTF-8 but **measured in UTF-16 code units**, which is what
VBScript counts. So `Len` returns characters rather than bytes, and `Left`, `Mid`,
`Right`, `InStr`, `InStrRev` and `StrReverse` all index by character — a euro sign
counts as 1, an emoji (surrogate pair) as 2:

```vbs
Len("h" & ChrW(233) & "llo")   ' 5, not 6
Left("caf" & ChrW(233), 4)     ' the whole word, not a half-character
StrReverse("caf" & ChrW(233))  ' reverses characters, never bytes
UCase("caf" & ChrW(233))       ' CAFE with a capital acute
```

`UCase`/`LCase` fold code points across ASCII, the Latin-1 supplement, Latin
Extended-A, Greek and Cyrillic; other scripts pass through unchanged.

`Chr` takes an ANSI code 0-255 (bytes 0x80-0x9F follow Windows-1252, so `Chr(147)`
is a left double quote) while `ChrW` takes a UTF-16 code unit up to 65535. `Asc`
reports the ANSI code of the first character and `AscW` its UTF-16 code unit. All
four emit and accept UTF-8 correctly.

One limit: indexing into the middle of a surrogate pair cannot be represented in
UTF-8, so `Mid(pair, 2, 1)` yields an empty string where VBScript would hand back a
lone low surrogate.

### Numeric types

Directive reproduces VBScript's numeric subtypes exactly, then adds one of its own.
The first four are what a ported script sees; `LongLong` is opt-in.

| TypeName | VarType | Range | Produced by |
|---|---|---|---|
| `Integer` | 2 | -32,768 .. 32,767 | literals that fit, `CInt`, `Asc`, `VarType`, arithmetic on Integers |
| `Long` | 3 | -2,147,483,648 .. 2,147,483,647 | larger literals, `CLng`, `Len`, `InStr`, `UBound` |
| `Byte` | 17 | 0 .. 255 | `CByte` |
| `Double` | 5 | ±1.7e308, exact to 2^53 | decimal literals, `/`, `^`, `CDbl`, integer overflow |
| **`LongLong`** | **20** | **-9,223,372,036,854,775,808 .. 9,223,372,036,854,775,807** | **automatically, once a value exceeds Long** |

The subtype is observable, so getting it right matters:

```vbs
TypeName(1)        ' Integer      -- not Long
TypeName(32768)    ' Long
Hex(-1)            ' "FFFF"       -- -1 is an Integer, so 16-bit
Hex(CLng(-1))      ' "FFFFFFFF"   -- explicitly a Long, so 32-bit
TypeName(7 \ 2)    ' Integer      -- both operands 16-bit
TypeName(7.6 Mod 2)' Long         -- a Double operand promotes the result
```

#### LongLong — exact 64-bit integers

VBScript's ladder is Integer -> Long -> Double, so a whole number past 2^31 becomes
a Double and stops holding whole numbers past 2^53. Directive adds one more rung to
that same ladder, **automatically**:

```
Integer (16-bit) -> Long (32-bit) -> LongLong (64-bit) -> Double
```

```vbs
TypeName(2147483648)        ' LongLong   (VBScript: Double)
9007199254740992 + 1        ' 9007199254740993   exact
CDbl("9007199254740992") + 1 ' 9.00719925474099E+15   what VBScript gives
3000000000 * 3000000000     ' 9000000000000000000
```

Arithmetic widens but never wraps: past 64 bits the result promotes to a Double,
exactly as VBScript promotes past 32. `CLngLng` converts explicitly and raises
error 6 beyond range; `dsLongLong` (20) is the `VarType` constant.

Every VBScript rule around it is untouched — `7 \ 2` is `Integer`, `7.6 Mod 2` is
`Long`, `/` is always `Double`, and a string operand gives `Double`. The one visible
consequence is that a whole number needing 16+ digits now prints its digits rather
than `1E+18`; see `tests/EXPECTED_DIFFS.md`.

#### Strict conversions

`CInt`, `CLng`, `CByte`, `CDbl` and `CLngLng` reject a partly-numeric string with
error 13, as VBScript does — `CInt("12abc")` is a type mismatch, not `12`.

### Dates

`Date` is a real subtype (`TypeName` `"Date"`, `VarType` 7), not a bare number, so
`IsDate` works and concatenation formats rather than leaking a serial:

```vbs
Directive.Echo "Started: " & Now      ' 4/03/2020 13:05:00   -- not 43894,54
TypeName(DateSerial(2020, 3, 4))      ' Date
IsDate(Now)                           ' True
```

`Now`, `Date`, `Time`, `CDate`, `DateSerial`, `TimeSerial`, `DateValue`,
`TimeValue` and `DateAdd` all return a `Date`. Arithmetic follows VBScript: a Date
plus or minus a number is still a Date, while the difference of two Dates is a
plain number of days. `Year`/`Month`/`Day`/`Hour`/`Minute`/`Second`/`Weekday`
return `Integer`. `IsNumeric(Now)` is False.

A Date with no time part prints as a date, one with no date part as a time, and
otherwise as both -- all in the active locale's format.

### Locale

VBScript formats and parses numbers and dates through the Windows locale, and
Directive mirrors that. On Windows the OS is asked directly, so output matches
VBScript on the same machine without configuration. On a comma-decimal locale
such as nl-BE:

| Expression | en-US (1033) | nl-BE (2067) |
|---|---|---|
| `CStr(2.5)` | `2.5` | `2,5` |
| `CDbl("1,5")` | error 13 | `1,5` |
| `CDbl("1.5")` | `1.5` | `15` — the dot is a thousands separator |
| `CDbl("1.000,5")` | error 13 | `1000,5` |
| `FormatNumber(1234.5, 2)` | `1,234.50` | `1.234,50` |
| `FormatCurrency(1234.5)` | `$1,234.50` | `€ 1.234,50` |
| `CDate("1/2/2020")` | 2 January | 1 February |

Parsing follows VBScript's rule exactly: the thousands separator is removed, then
the decimal separator becomes a point. Slash dates follow the locale's field order.

`GetLocale()` returns the active LCID; `SetLocale(lcid)` changes it and returns the
previous one, and `SetLocale(0)` restores the machine default. Use
`SetLocale(1033)` to pin invariant output for files another program has to read:

```vbs
SetLocale 1033                       ' force a dot, whatever the machine
fh.WriteLine CStr(price)
```

On Linux and macOS there is no LCID to consult, so Directive uses a built-in table
(1033, 2057, 2067, 1043, 1031, 1036, 1040, 3082) defaulting to 1033, and formats
times as 24-hour. Windows delegates to the OS and is therefore exact.

## 3. Operators
Arithmetic `+  -  *  /  \ (int div)  Mod  ^ (power)` ·
Concatenation `&` (always string; `+` concatenates only when **both** sides are
strings, else adds) · Comparison `=  <>  <  >  <=  >=` ·
Object identity `Is` · Logical/bitwise `And  Or  Not  Xor  Eqv  Imp`.

---

## 4. Statements & language constructs
- `Option Explicit`
- `Dim` / `Public` / `Private` / `ReDim [Preserve]` / `Const`
- `Set x = …` (object assignment), plain `x = …` (value assignment)
- `If … Then … [ElseIf …] [Else] … End If` (and single-line `If … Then …`)
- `For … To … [Step …] … Next`
- `For Each … In … … Next`
- `Do [While|Until] … Loop` and `Do … Loop [While|Until]`
- `While … Wend`
- `Select Case … Case … [Case Else] End Select` (value lists, `Is` comparisons, `To` ranges)
- `With obj … End With` (nestable)
- `Sub` / `Function` with `ByRef` / `ByVal`; `Exit Sub`/`Exit Function`
- `Call proc(args)` and paren-less `proc args`
- `Class … End Class` with `Public`/`Private` members, `Property Get/Let/Set`,
  `Public Default` member, `Me`, `Class_Initialize`, `Class_Terminate`; `Exit Property`
- `Mid(target, start[, length]) = value` (in-place string replacement statement)
- `Erase array`
- `On Error Resume Next` / `On Error Goto 0`
- `Randomize [seed]`
- `#…#` date literals, e.g. `#2020-03-15#`, `#1/15/2020 3:45 PM#`, `#13:45:00#`
- Native `Include "path"` (top-level; splices another file at parse time)
- `[bracket identifiers]` — any text in `[ ]` is a literal name (variable, member,
  `Sub`/`Function`), even a keyword, spaces, or a leading digit: `Dim [hello world]`,
  `[End] = 1`, `obj.[Next]`. Case-insensitive like all identifiers. Mainly for COM
  interop with members whose names collide with keywords.

---

## 5. Constants
Constants use the Directive-branded `ds*` prefix (there is no `vb*` prefix —
porting a classic script is a straight `vb`→`ds` rename):

**Whitespace/strings:** `dsCr` `dsLf` `dsCrLf` `dsNewLine` `dsTab` `dsNullString`
`dsNullChar` `dsBack` `dsFormFeed` `dsVerticalTab`
**Booleans/compare:** `dsTrue` `dsFalse` `dsBinaryCompare` `dsTextCompare`
**VarType:** `dsEmpty` `dsNull` `dsInteger` `dsLong` `dsSingle` `dsDouble`
`dsCurrency` `dsDate` `dsString` `dsObject` `dsError` `dsBoolean` `dsVariant`
`dsByte` `dsArray`
**MsgBox buttons/icons/returns:** `dsOKOnly` `dsOKCancel` `dsAbortRetryIgnore`
`dsYesNoCancel` `dsYesNo` `dsRetryCancel` `dsCritical` `dsQuestion` `dsExclamation`
`dsInformation` `dsOK` `dsCancel` `dsAbort` `dsRetry` `dsIgnore` `dsYes` `dsNo`
**Weekdays:** `dsSunday`…`dsSaturday` `dsUseSystemDayOfWeek` `dsFirstJan1`
**Date formats:** `dsGeneralDate` `dsLongDate` `dsShortDate` `dsLongTime` `dsShortTime`
**Errors:** `dsObjectError`

---

## 6. Objects

### 6.1 `Directive` — host object (always available)
| Member | Signature | Notes |
|---|---|---|
| `Echo` | `Directive.Echo a, b, …` | space-joined; objects use their default member |
| `Arguments` (alias `Args`) | `Directive.Arguments` | command-line args collection (see below) |
| `StdOut` / `StdErr` / `StdIn` | stream objects with `.Write`/`.WriteLine`/`.ReadLine`/`.ReadAll` | |
| `Sleep` | `Directive.Sleep ms` | |
| `Quit` | `Directive.Quit [code]` | |
| `ScriptName` / `ScriptFullName` / `ScriptPath` (alias `ScriptDir`) | path info | |

**Echo / StdOut behavior** depends on the build: in the
console build (`directive.exe`) `Echo` and the `StdOut`/`StdIn`/`StdErr` streams go
to the terminal; in the GUI build (`directivew.exe`) `Echo` writes to a parent
console when launched from one, otherwise pops up a message box, and the streams
write to the console only when one is attached. `MsgBox`/`InputBox` are dialogs in
both builds on Windows (stdio on other platforms).

**`Directive.Arguments`** — a 0-based collection of the arguments passed after the
script on the command line (the host object is always named `Directive`):

| Member | Signature |
|---|---|
| `Count` (alias `Length`) | `.Count` |
| `Item` (default) | `.Item(i)` / `Directive.Arguments(i)` — 0-based |
| For Each | iterates the argument strings |

```
directive myscript.ds one two "three with spaces" file.txt
```
```vb
For Each a In Directive.Arguments
    directive.echo a
Next
```
On Windows, associate the `.ds` extension with `directive.exe "%1" %*`
(or `directivew.exe "%1" %*`) so that dropping files onto a script passes their
paths through as arguments.

### 6.2 `Dictionary` — `New Dictionary`  ✔ tested
Keys are **typed**: numeric `1` ≠ string `"1"`; `1` and `1.0` coincide; objects
key by identity.

| Member | Signature |
|---|---|
| `Add` | `.Add key, item` |
| `Item` (default) | `.Item(key)` / `d(key)` — get or set |
| `Exists` | `.Exists(key)` |
| `Remove` | `.Remove key` |
| `RemoveAll` | `.RemoveAll` |
| `Keys` / `Items` | `.Keys` / `.Items` → array |
| `Count` | `.Count` |
| `CompareMode` | read/write (0 binary) |
| For Each | iterates keys |

### 6.3 `List` — `New List`  ✔ tested
A dynamic list (not part of the classic language).

| Member | Signature |
|---|---|
| `Add` / `Push` / `Append` | `.Add value` |
| `Insert` | `.Insert index, value` |
| `Remove` | `.Remove value` (first match) |
| `RemoveAt` | `.RemoveAt index` |
| `Item` (default) | `.Item(index)` / `l(index)` — get or set |
| `Count` | `.Count` |
| `Clear` | `.Clear` |
| `Contains` | `.Contains(value)` |
| `IndexOf` | `.IndexOf(value)` |
| `Reverse` | `.Reverse` |
| `Sort` | `.Sort` |
| `Swap` | `.Swap i, j` |
| `First` / `Last` | `.First` / `.Last` |
| `ToArray` | `.ToArray` → array |
| For Each | iterates items |

### 6.4 `RegExp` — `New RegExp`  ✔ tested
| Member | Signature |
|---|---|
| `Pattern` | read/write |
| `Global` / `IgnoreCase` / `MultiLine` | read/write booleans |
| `Test` | `.Test(string)` → bool |
| `Execute` | `.Execute(string)` → Matches collection |
| `Replace` | `.Replace(string, replacement)` |

**Match** (from `Execute`): `.Value` (default), `.FirstIndex`, `.Length`,
`.SubMatches([i])`, `.Count`. Iterable with For Each.

### 6.5 `FileSystem` — `New FileSystem`  ✔ tested
| Member | Signature |
|---|---|
| `FileExists` / `FolderExists` | `(path)` |
| `GetFile` / `GetFolder` | `(path)` → File/Folder |
| `CreateFolder` | `(path)` |
| `CreateTextFile` | `(path, [overwrite, [unicode]])` → TextStream |
| `OpenTextFile` | `(path, [iomode, [create, [format]]])` → TextStream |
| `DeleteFile` / `DeleteFolder` | `(path)` |
| `CopyFile` | `(src, dest)` |
| `GetSpecialFolder` | `(n)` |

**File / Folder** objects: `.Name`, `.Path`, `.ShortName`, `.Size`, `.Type`,
`.Attributes`, `.DateCreated`, `.DateLastModified`, `.DateLastAccessed`,
`.Drive`, `.ParentFolder`, `.IsRootFolder`, `.Copy`, `.Move`, `.Delete`,
`.OpenAsTextStream([iomode, [format]])` (File). Folder also: `.Files`,
`.SubFolders`, `.CreateTextFile`, `.CreateFolder`, `.FileExists`, `.FolderExists`,
`.GetFolder`.

**TextStream:** `.Write`, `.WriteLine`, `.ReadLine`, `.ReadAll`,
`.AtEndOfStream`, `.Close`.

**Encoding (Unicode).** Text files default to **UTF-8** (no BOM). The `unicode`
argument to `CreateTextFile` (`True`) and the `format` argument to `OpenTextFile`
/ `OpenAsTextStream` (`-1` = `TristateTrue`/Unicode, `0` = `TristateFalse`/ASCII,
`-2` = `TristateUseDefault`) select **UTF-16LE with a BOM**, matching what such tooling
is commonly called "Unicode". On **read**, a byte-order mark is auto-detected and
decoded regardless of the argument — UTF-8 BOM is stripped, and UTF-16 **LE or
BE** is decoded to UTF-8 (surrogate pairs, e.g. emoji, handled); `format = -1`
forces UTF-16LE for a BOM-less file. UTF-16 output uses `\r\n` line endings.
Runtime-verified on Linux and, via the Windows binary, under Wine.

### 6.6 `Sound` — `New Sound`  ✔ tested (WAV combining cross-platform; playback 🪟 Windows-only)
Two roles on one object. **WAV combining** (Create/Append/Duration/Close)
concatenates same-format PCM WAV files and is pure file I/O — works everywhere
(runtime-verified on Linux and under Wine). **Playback** (Play/Stop/Beep) uses
`winmm` and is Windows-only (raises 429 elsewhere).
| Member | Signature | Notes |
|---|---|---|
| `Create` | `.Create(path)` | open an output WAV for combining |
| `Append` (alias `Add`) | `.Append(wavPath)` | first file fixes the format; mismatches raise |
| `Duration` | `.Duration` → seconds written | |
| `Close` | `.Close` | patches the header and closes |
| `Play` | `.Play(file, [wait])` | 🪟 wait=True blocks until done |
| `Stop` | `.Stop` | 🪟 |
| `Beep` | `.Beep([frequency, [duration]])` | 🪟 |

### 6.7 `Mouse` — `New Mouse`  🪟 Windows-only
| Member | Signature |
|---|---|
| `Move` | `.Move x, y` |
| `Click` | `.Click([button, [x, [y, [count]]]])` — button `"left"`/`"right"`/`"middle"` |
| `Down` / `Up` | `.Down([button])` / `.Up([button])` |
| `Wheel` | `.Wheel("up"\|"down", [clicks])` |
| `GetPos` | `.GetPos` → position |
| `X` / `Y` | current coordinates |

### 6.8 `Screen` — `New Screen`  🪟 Windows-only
| Member | Signature |
|---|---|
| `Width` / `Height` | screen dimensions |
| `PixelColor` (alias `GetPixel`) | `.PixelColor(x, y)` → 0xRRGGBB |

### 6.9 `Keyboard` — `New Keyboard`  🪟 Windows-only
Synthetic keyboard input (symmetric with `Mouse`). Compile-verified; keystroke
delivery needs a real desktop session, so it isn't runtime-tested here.
| Member | Signature | Notes |
|---|---|---|
| `Send` (alias `SendKeys`) | `.Send(keys, [raw])` | SendKeys mini-language (below); `raw`=True/1 sends `keys` literally, no `{}`/`+^%()` parsing — like AutoIt's `Send($keys, 1)`, equivalent to `.Type` |
| `Type` | `.Type(text)` | literal text, no special-key parsing |
| `Down` / `Up` | `.Down(key)` / `.Up(key)` | hold/release a key by name or character |

**SendKeys syntax:** literal characters type themselves; `{NAME}` presses a named
key (`{ENTER}` `{TAB}` `{ESC}` `{BACKSPACE}` `{DEL}` `{HOME}` `{END}` `{PGUP}`
`{PGDN}` `{UP}` `{DOWN}` `{LEFT}` `{RIGHT}` `{F1}`–`{F24}` `{INSERT}` `{SPACE}` …);
`{KEY N}` repeats (`{LEFT 5}`); modifiers `+` `^` `%` = Shift/Ctrl/Alt applied to
the next key or a `(grouped)` run; `~` = Enter; `{+}` `{^}` `{%}` `{(}` `{)}`
`{{}` `{}}` are the literal symbols.

### 6.10 `Clipboard` — `New Clipboard`  🪟 Windows-only
Text clipboard access (replaces the old flat `ClipPut`/`ClipGet`). Runtime-verified
under Wine.
| Member | Signature |
|---|---|
| `Set` (aliases `SetText`, `PutText`) | `.Set(text)` → success |
| `Get` (alias `GetText`) | `.Get` → clipboard text |
| `Clear` | `.Clear` — empties the clipboard |

### 6.11 `Shell` — `New Shell`  ✔ tested (native replacement for the classic Windows shell)
A COM-free replacement for the classic Windows shell, so scripts keep working if the
Windows Script Host objects go away. Process/environment features are
cross-platform; registry, special-folder, shortcut, and window features are
Windows-only (they raise error 429 on other platforms). Runtime-verified on Linux
and, via the Windows binary, under Wine — including the registry round-trip,
clipboard, and shortcut creation.

| Member | Signature | Notes |
|---|---|---|
| `Run` | `.Run(command, [windowStyle], [waitOnReturn])` | returns exit code (0 if not waiting); `windowStyle` ignored off-Windows |
| `Exec` | `.Exec(command)` | returns an exec object (below); captures stdout |
| `Environment` | `.Environment([type])` | env collection: `env(name)` reads/writes, `.Count`, For Each yields `NAME=VALUE` |
| `ExpandEnvironmentStrings` | `.ExpandEnvironmentStrings(s)` | replaces `%VAR%` |
| `CurrentDirectory` | read/write property | |
| `Popup` | `.Popup(text, [seconds, [title, [type]]])` | message box; returns button code (`seconds` auto-close only honored where supported) |
| `RegRead` | `.RegRead(key)` | 🪟 Windows-only |
| `RegWrite` | `.RegWrite(key, value, [type])` | 🪟 `type` = `REG_SZ` / `REG_EXPAND_SZ` / `REG_DWORD` |
| `RegDelete` | `.RegDelete(key)` | 🪟 trailing `\` deletes the key |
| `SpecialFolders` | `.SpecialFolders(name)` | 🪟 e.g. `Desktop`, `AppData`, `Programs`, `Windows` |
| `CreateShortcut` | `.CreateShortcut(path)` → shortcut object | 🪟 set `.TargetPath`/`.Arguments`/`.WorkingDirectory`/`.IconLocation`/`.Description`, then `.Save` |
| `AppActivate` | `.AppActivate(title)` → success | 🪟 focus a top-level window by (partial) title |

**Exec result object:** `.StdOut` (readable TextStream: `.ReadAll`, `.ReadLine`,
`.AtEndOfStream`), `.ExitCode`, `.Status` (1 = complete). (`.StdErr` is present
but not captured separately.)

```vb
Dim sh : Set sh = New Shell
Dim e  : Set e  = sh.Exec("git status")
Directive.Echo e.StdOut.ReadAll
Directive.Echo "PATH = " & sh.Environment("PROCESS")("PATH")
```

---

### Console — `Set c = New Console`

Console colours, cursor and single-key input. **Not a VBScript feature**: VBScript
has no console API at all, which is why batch scripts shell out to `color` and
`cls`. Cross-platform — the Win32 console API on Windows (honoured by conhost and
Windows Terminal alike), ANSI escapes on Linux/macOS. All escapes are suppressed
when output is redirected, so `directive x.ds > log.txt` yields clean text.
Coordinates are **0-based**, like `Directive.Arguments` and `List`.

| Member | Type | Description |
|---|---|---|
| `Write text[, fg[, bg]]` | method | Write text. The optional colours apply to this call only and are restored afterwards. |
| `WriteLine text[, fg[, bg]]` | method | As `Write`, plus a newline. |
| `Foreground` | Integer, r/w | Persistent foreground colour, `-1` = console default. |
| `Background` | Integer, r/w | Persistent background colour, `-1` = console default. |
| `Reset` | method | Restore the console's own default colours. |
| `Cls` | method | Clear the screen and home the cursor. |
| `ClearLine` | method | Blank the current line and return to column 0. |
| `Width` / `Height` | Integer, r | Visible console size in **characters** (`Screen.Width` is display pixels). Falls back to 80x25. |
| `MoveTo x, y` | method | Move the cursor (0-based). |
| `CursorX` / `CursorY` | Integer, r/w | Cursor position; returns `-1` when it cannot be determined. |
| `CursorVisible` | Boolean, r/w | Show or hide the caret. |
| `Title` | String, r/w | Console window title. On POSIX, reading returns the last value set. |
| `Attached` | Boolean, r | Whether a console is present at all. |
| `ReadKey` | String, r | Wait for one key, no Enter needed. Returns `""` if input is not a console. |
| `Pause [prompt]` | method | Wait for any key (the batch `PAUSE`). |
| `KeepOpen [prompt]` | method | Wait **only** when this script owns the console. |

`ReadKey` returns exactly the names `Keyboard.Send` accepts, so a key can be read
and replayed:

| Returned | Meaning |
|---|---|
| `"a"`, `"€"` | A printable character, UTF-8 (multi-byte characters arrive whole) |
| `"{UP}"` | `{ENTER}` `{ESC}` `{TAB}` `{BACKSPACE}` `{DELETE}` `{INSERT}` `{HOME}` `{END}` `{PGUP}` `{PGDN}` `{UP}` `{DOWN}` `{LEFT}` `{RIGHT}` `{F1}`..`{F24}` |
| `"^a"` | Ctrl+A, matching `Send`'s `^` prefix |

**Telling the builds apart.** `Directive.FullName` is the full path of the running
interpreter, `Directive.Path` its folder and `Directive.Name` the host name. Testing
`FullName` is how a script detects which build is running it:

```vbs
If InStr(LCase(Directive.FullName), "directivew") > 0 Then   ' the GUI build
```

Note this is not the same test as `Console.Attached`, which is also False when a
console run merely has its output redirected.

**In the windowed build.** Under `directivew.exe` a double-clicked
script has no console, so `Attached` is False, `Write`/`WriteLine` go nowhere and
the colour and cursor calls are no-ops — while `Directive.Echo` becomes a popup.
Guard console work with `If Console.Attached Then`, as `examples/17_console.ds`
does. `Pause`, `KeepOpen` and `ReadKey` are safe either way: they require stdin to
be a console, so they return immediately instead of hanging.

**`Pause` vs `KeepOpen`.** `Pause` always waits while a console is attached.
`KeepOpen` waits only when the console will vanish on exit — detected via
`GetConsoleProcessList() == 1` on Windows, so a double-clicked or dragged-onto
script stays readable while running the same script from an existing prompt does
not pause. On POSIX `KeepOpen` is deliberately a no-op, because a terminal
emulator's window belongs to your shell, not to the script.

Both are gated on stdin actually being a console, so a piped, redirected or
scheduled run can never block:

```vbs
Dim c : Set c = New Console
c.Cls
c.WriteLine " BUILD PASSED ", dsBrightWhite, dsGreen
c.Write "status: " : c.WriteLine "OK", dsBrightGreen
c.MoveTo 0, 10
Dim k : k = c.ReadKey            ' "{F5}", "a", "^c", ...
c.KeepOpen                       ' the .ds equivalent of batch PAUSE
```

### TextWindow — `Set tw = New TextWindow`

A Unicode console replacement: a scrollable output pane on top, an editable
multiline input pane below, and a Submit button. It runs on **its own thread with
its own message loop**, so it stays responsive while the script is busy and output
appears as it is written. Windows only.

The window appears on the first write. Setting `Visible = False` beforehand
suppresses that.

| Output | |
|---|---|
| `Write(text, ...)` | Append without a newline. Takes several arguments. |
| `WriteLine([text, ...])` | Append with a newline. No argument writes a blank line. |
| `Text` | Get or set the whole output pane. Default member. |
| `Clear` | Empty the output pane. |

A lone line feed or carriage return becomes a proper line break, so
`tw.WriteLine "a" & dsLf & "b"` shows two lines instead of a box character.

| Input | |
|---|---|
| `ReadLine` | Block until the user submits; return one line at a time. |
| `ReadAll` | Block until the user submits; return the whole submission. |
| `Prompt(message)` | `Write(message)` then `ReadLine`. |
| `InputText` | Get or set the input pane. |
| `ClearInput` | Empty it and discard anything unread. |
| `HasInput` | True if a submitted line is waiting. Never blocks. |

The user submits with **Submit** or **Ctrl+Enter**; plain Enter inserts a newline.
`SubmitOnEnter = True` swaps that round (Enter submits, Shift+Enter inserts).

| Window | |
|---|---|
| `Show`, `Hide`, `Close`, `Activate` | |
| `WaitForClose` | Block until the user closes it. |
| `Visible`, `IsClosed` | |
| `Title`, `Width`, `Height` | |
| `FontName`, `FontSize` | Consolas 10 by default; any Unicode font works. |
| `WordWrap` | Default True. Changing it rebuilds the pane, keeping the text. |
| `ShowInput` | Default True. False gives an output-only log window. |
| `EchoInput` | Default True — submitted text is copied into the output as a transcript. |
| `SubmitOnEnter` | Default False. |

`ReadLine`, `ReadAll` and `Prompt` raise a catchable error rather than hanging if
the user closes the window; anything already submitted is returned first, and
`IsClosed` lets you poll instead. Blocking calls pump messages, so nothing is
starved.

Closing the window does not destroy the object. The transcript is kept, so `Text`
still returns it afterwards and writing continues to accumulate; `Show` reopens the
window with everything restored.

```vbs
Dim tw : Set tw = New TextWindow
tw.SubmitOnEnter = True
Do
    cmd = tw.Prompt("> ")
    If tw.IsClosed Or LCase(cmd) = "quit" Then Exit Do
    tw.WriteLine "you said: " & cmd
Loop
```

See `examples/20_textwindow.ds`.

### Console colour constants

Same numbering as cmd.exe's `color` command and the Win32 attribute bits
(`1`=blue, `2`=green, `4`=red, `+8`=bright) — so `dsBlue` is 1 and `dsRed` is 4,
**not** the ANSI order. `dsYellow` (6) is the dark/brown yellow; `dsBrightYellow`
(14) is the vivid one.

| 0-7 | 8-15 |
|---|---|
| `dsBlack` 0, `dsBlue` 1, `dsGreen` 2, `dsCyan` 3, `dsRed` 4, `dsMagenta` 5, `dsYellow` 6, `dsWhite` 7 | `dsGray` 8, `dsBrightBlue` 9, `dsBrightGreen` 10, `dsBrightCyan` 11, `dsBrightRed` 12, `dsBrightMagenta` 13, `dsBrightYellow` 14, `dsBrightWhite` 15 |

### StdIn — full TextStream surface

`Directive.StdIn` now implements the whole VBScript `TextStream` read surface.
`Read`, `ReadAll`, `ReadLine` and `SkipLine` all raise **error 62
("Input past end of file")** at EOF, exactly as VBScript does — so guard them
with `AtEndOfStream`, or use `Console.KeepOpen`, which cannot hang.

| Member | Description |
|---|---|
| `Read(n)` | `n` characters. Error 62 at EOF. |
| `ReadLine` | One line without its terminator. Error 62 at EOF. |
| `ReadAll` | Everything remaining. Error 62 at EOF. |
| `Skip(n)` / `SkipLine` | Discard `n` characters / one line. |
| `AtEndOfStream` / `AtEndOfLine` | Boolean. Note `AtEndOfStream` blocks on a live console until a key is available. |
| `Line` / `Column` | Current 1-based position. |
| `Close` | Accepted, no-op. |

## 7. Known limitations (see README for the full list)
- Not a security sandbox — as capable as any classic script host (file access, COM, OS
  automation, Include). Don't run untrusted scripts.
- No distinct Date subtype: dates are OLE serial doubles, so `TypeName(#…#)` is
  `"Double"` and `IsDate(#…#)` is `False` (the literal is already a number).
- `Private` class members are **not enforced** for external access (they behave
  as accessible). They work correctly inside the class.
- A leading-dot argument to a paren-less call (`Directive.Echo .Member` inside a
  `With`) parses as member chaining. Workaround: parenthesize —
  `Directive.Echo(.Member)`.
- Reference cycles leak and skip `Class_Terminate` (same as COM; no GC).
- `Rnd` is not bit-identical to any legacy generator.
- `Mouse`/`Screen`/`Sound`/`Clipboard`/`CreateObject` are Windows-only and, in
  this build environment, are compile/link-verified but not runtime-tested.
