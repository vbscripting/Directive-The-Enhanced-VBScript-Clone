# Directive-The-Enhanced-VBScript-Clone

VBScript will be retired and eliminated from future versions of Windows,
I wanna build a better alternative!

It runs the familiar VBScript-style language,
but as a single self-contained native binary,
cross-platform, with no dependency on the Windows Script Host or COM.

Directive is a small dynamic scripting language, one C++17 file (`directive.cpp`,
no external dependencies) that compiles to a native executable.
It stays source-compatible with the classic language where it can,
while giving it a cleaner, more explicit object model of its own.

It keeps the parts of the classic language that are pleasant to write — `Dim`,
`If/Select`, `For/Do`, `Sub/Function`, `Class`, `On Error`/`Err` — and gives it a
cleaner, more explicit object model:

- **`New` creates native engine objects** (`Dictionary`, `List`, `RegExp`,
  `FileSystem`, `Sound`, `Shell`, …). They work on every platform.
- **`CreateObject` is reserved for real COM/ActiveX** automation on Windows.
- The host object is **`Directive`**: `Directive.Echo`,
  `Directive.StdOut.Write`, `Directive.StdIn.ReadLine`, `Directive.Quit`.

## Object model

| Task | Directive |
|------|-----------|
| dictionary / map | `New Dictionary` |
| file & folder access, text streams | `New FileSystem` |
| regular expressions | `New RegExp` |
| shell: run / exec / environment / registry | `New Shell` |
| dynamic list | `New List` |
| desktop automation | `New Mouse`, `New Screen`, `New Keyboard`, `New Clipboard`, `New Sound` |
| real COM / ActiveX (Windows only) | `CreateObject("Excel.Application")`, … |
| console output | `Directive.Echo` |
| standard streams | `Directive.StdOut` / `Directive.StdIn` |
| exit the script | `Directive.Quit` |

So `New` = in-process, cross-platform, engine-provided; `CreateObject` = the
operating system's COM, Windows-only. The split keeps the two worlds from blurring.
Each engine object has **one canonical name** (no aliases), and functions are
grouped by responsibility: `Dictionary`, `List`, `RegExp`, `FileSystem`, `Sound`
(playback + WAV combining), `Mouse`, `Screen`, `Keyboard`, `Clipboard`, `Shell`.

## Build

One file, no build system.

**Console build** (console-style: `Directive.Echo` and the `StdOut`/`StdIn`
streams use the terminal). This is the portable default and the only mode on
Linux/macOS.

```sh
# Linux / macOS
g++ -std=c++17 -O2 directive.cpp -o directive

# Windows, MSVC (x64 Native Tools prompt)
cl /std:c++17 /EHsc /O2 directive.cpp /Fe:directive.exe /link /STACK:16777216

# Windows, MinGW-w64 — link statically so the .exe is self-contained (no DLLs to ship)
g++ -std=c++17 -O2 -static -static-libgcc -static-libstdc++ -Wl,--stack,16777216 directive.cpp -o directive.exe -lole32 -loleaut32 -luuid -lwinmm -lgdi32 -luser32 -ladvapi32 -lshell32
```

**GUI build** (GUI-style: `MsgBox`, `InputBox`, and `Directive.Echo` show real
windows). Add `-DDIRECTIVE_GUI` and a windowed subsystem:

```sh
# Windows, MSVC
cl /std:c++17 /EHsc /O2 /DDIRECTIVE_GUI directive.cpp /Fe:directivew.exe /link /SUBSYSTEM:WINDOWS /STACK:16777216

# Windows, MinGW-w64 (static)
g++ -std=c++17 -O2 -DDIRECTIVE_GUI -static -static-libgcc -static-libstdc++ -Wl,--stack,16777216 directive.cpp -o directivew.exe -mwindows -lole32 -loleaut32 -luuid -lwinmm -lgdi32 -luser32 -ladvapi32 -lshell32
```

> **MinGW users: use `-static`.** Without it the `.exe` depends on
> `libstdc++-6.dll` / `libgcc_s_seh-1.dll`, which won't exist on a clean machine
> and the program will fail to start. Static linking bundles everything into one
> file. (MSVC links its runtime differently and doesn't need this.)

The GUI build also calls `AttachConsole` at startup, so if you launch it from a
`cmd` window `Directive.StdOut` still writes there; double-clicked, it just shows
dialogs.

## Run

```sh
./directive script.ds          # run a file
./directive script.ds a b c    # run a file, passing arguments
./directive -e "Directive.Echo 6*7"   # inline
./directive -                         # read from stdin
```

There are two builds, a console one and a windowed one:

- **`directive.exe` — console.** `Directive.Echo` and the `Directive.StdOut` /
  `StdIn` / `StdErr` streams go to the terminal. `MsgBox` and `InputBox` still pop
  up real dialogs — they are language functions, not host-specific ones. Use this
  for terminals, pipes, and automation; the shell waits for it.
- **`directivew.exe` — windowed.** Double-click it, or associate `.ds` with it:
  `Directive.Echo` pops up a message box, and `MsgBox` / `InputBox` are dialogs.
  Two things it does better than the classic windowed host — launch it **from a
  console** and `Directive.StdOut` still writes to that console rather than failing
  with an invalid-handle error, and `Directive.Echo` writes to the console instead
  of popping up. (It's a GUI-subsystem binary, so the shell does **not**
  wait for it; for clean, ordered console output prefer `directive.exe`.)

| | `directive.exe` | `directivew.exe` |
|---|---|---|
| `Directive.Echo` | console | console if launched from one, else popup |
| `Directive.StdOut` / `StdIn` / `StdErr` | console | console when attached (else discarded) |
| `MsgBox` / `InputBox` | dialog | dialog |

### Command-line arguments

Anything after the script path is exposed to the script through
`Directive.Arguments`, a 0-based collection (the host is always named
`Directive`):

```vb
Directive.Echo "count: " & Directive.Arguments.Count
Directive.Echo "first: " & Directive.Arguments(0)
Dim f
For Each f In Directive.Arguments
    Directive.Echo f
Next
```

To make **drag-and-drop** work on Windows, associate the `.ds` extension
with `directive.exe "%1" %*` (or `directivew.exe "%1" %*`). The `%*` forwards the
dropped file paths, so dragging files onto a script delivers their paths as
`Directive.Arguments`. See `examples/13_arguments.ds`.

## Error messages

Errors are reported in the same shape as the Windows Script Host dialog — script
path, line, column, message, an HRESULT-style code, and a source:

```
Script: /path/to/script.ds
Line:   3
Char:   1
Error:  Division by zero
Code:   800A000B
Source: Directive runtime error
```

(The fields are tab-aligned into a column, like the Windows Script Host dialog.)

The `Code` is the the classic form form `0x800A0000 + errorNumber`, so it lines up
with the numbers you already know: type mismatch `800A000D` (13), division by zero
`800A000B` (11), subscript out of range `800A0009` (9), object required `800A01A8`
(424), undefined variable `800A01F4` (500). Compilation errors carry
`Source: Directive compilation error` and point at the offending token; runtime
errors point at the statement being executed. In the console build this prints to
stderr; in the GUI build it appears in a message box.

## The `List` object

`New List` is a genuine doubly-linked list:

```vb
Dim L : Set L = New List
L.Add "alpha"
L.Insert 1, "beta"          ' by index
L.Swap 0, 1
L.Remove "beta"             ' by value; returns True/False
L.RemoveAt 0                ' by index
x = L(0)                    ' Item(index), 0-based; also L(0) = "new"
L.Reverse
L.Sort
For Each item In L : Directive.Echo item : Next
```

Members: `Add`/`Append`/`Push`, `Insert(i, v)`, `Remove(v)`, `RemoveAt(i)`,
`Item(i)` (get/set, default), `Count`, `Clear`, `Contains(v)`, `IndexOf(v)`,
`Swap(i, j)`, `Reverse`, `Sort`, `First`, `Last`, `ToArray`, and `For Each`.

## The `Directive` host object

- `Directive.Echo a, b, …` — writes a line (console: stdout; GUI: a message box)
- `Directive.StdOut` / `Directive.StdErr` — `.Write`, `.WriteLine`, `.WriteBlankLines(n)`
- `Directive.StdIn` — `.ReadLine`, `.Read(n)`, `.ReadAll`, `.AtEndOfStream`
- `Directive.Quit [code]`, `Directive.Sleep ms`
- `Directive.ScriptName` — the script's file name; `Directive.ScriptFullName` — its
  absolute path; `Directive.ScriptPath` (alias `ScriptDir`) — the folder containing
  it. Handy for locating files relative to the script (`-e`/stdin scripts report
  the current directory).

## Include

A native `Include` pulls another file's definitions into the current script:

```vb
Include "lib/strutil.ds"     ' path is relative to THIS file's folder
directive.echo TitleCase("hello world")
```

Unlike the old `ExecuteGlobal`-plus-read-the-file trick, `Include` happens at
**parse time**, so everything the included file defines (Const, Function, Sub,
Class) is visible everywhere in the including file — including *above* the
`Include` line. Details:

- **Paths resolve relative to the including file's directory** (nested includes
  resolve relative to their own location), so a library can `Include` its
  siblings without caring where the top-level script lives.
- **Each file is included at most once.** Diamond includes (A pulls in B and C,
  both of which pull in D) and circular includes (A ↔ B) are handled by an
  include guard — no duplicate-definition errors, no infinite loops.
- **Top level only** — `Include` inside a `Sub`/`Function`/`Class` is a parse
  error. Parse errors inside an included file name the file.

## File / Folder object model

`New FileSystem` gives you the full object model, not just loose functions:

```vb
Dim fso : Set fso = New FileSystem
Dim folder : Set folder = fso.GetFolder("demo")
Directive.Echo folder.Name & " has " & folder.Files.Count & " files"
Dim f
For Each f In folder.Files
    Directive.Echo f.Name & " — " & f.Size & " bytes, modified " & f.DateLastModified
Next
Dim doc : Set doc = fso.GetFile("demo\notes.txt")
doc.Copy "demo\backup.txt"
Set ts = doc.OpenAsTextStream(1) : Directive.Echo ts.ReadLine : ts.Close
```

- **FileSystem**: `FileExists`, `FolderExists`, `CreateFolder`, `DeleteFile`,
  `DeleteFolder`, `CopyFile`, `CopyFolder`, `MoveFile`, `MoveFolder`,
  `CreateTextFile`, `OpenTextFile`, `GetFile`, `GetFolder`, `GetSpecialFolder`,
  `GetBaseName`, `GetFileName`, `GetExtensionName`, `GetParentFolderName`,
  `BuildPath`, `GetAbsolutePathName`, `GetTempName`.
- **Folder**: `Path`, `Name`, `ParentFolder`, `Drive`, `Size`, `Type`,
  `Attributes`, `DateLastModified` (etc.), `IsRootFolder`, `Files`, `SubFolders`,
  `Delete`, `Copy`, `Move`, `CreateTextFile`.
- **File**: `Path`, `Name`, `ParentFolder`, `Drive`, `Size`, `Type`,
  `Attributes`, `DateLastModified` (etc.), `Delete`, `Copy`, `Move`,
  `OpenAsTextStream`.
- **Files / SubFolders**: `Count`, `Item(name)`, and `For Each`.

On non-Windows, `\` in paths is treated as a separator so ordinary scripts work.

**Unicode text files.** Text files default to **UTF-8** (no BOM). Pass the
`unicode` argument to `CreateTextFile(path, overwrite, True)` — or `format = -1`
to `OpenTextFile` / `OpenAsTextStream` — to write **UTF-16LE with a BOM**, which
is what the classic form calls "Unicode". Reading auto-detects the BOM and decodes
UTF-8, UTF-16LE, or UTF-16BE either way (emoji and other astral characters
included), so you usually don't pass a format at all. See
`examples/15_unicode.ds`.

## Automation & multimedia

AutoIt-style desktop automation in single-purpose objects, plus a cross-platform
WAV combiner folded into `Sound`.

```vb
' --- Mouse (Windows) ---
Dim m : Set m = New Mouse
m.Move 400, 300              ' instant; add a 3rd arg for smooth motion, e.g. m.Move 400,300,5
m.Click "left"               ' Click([button],[x],[y],[count]); button = left/right/middle
m.Click "left", 400, 300, 2  ' move there, then double-click
m.Down "left" : m.Up "left"  ' press/release separately (drag)
m.Wheel "down", 3            ' scroll
directive.echo "cursor at " & m.X & "," & m.Y   ' also m.GetPos -> [x, y]

' --- Screen (Windows) ---
Dim s : Set s = New Screen
directive.echo Hex(s.PixelColor(100, 100))   ' 0xRRGGBB, like AutoIt PixelGetColor
directive.echo s.Width & "x" & s.Height

' --- Keyboard (Windows) — symmetric with Mouse ---
Dim kb : Set kb = New Keyboard
kb.Send "Hello{ENTER}"        ' SendKeys syntax: {ENTER} {TAB} {F5} {LEFT 5} ...
kb.Send "^a"                  ' Ctrl+A   (+ = Shift, ^ = Ctrl, % = Alt)
kb.Send "+(abc)"              ' hold Shift over the whole group -> ABC
kb.Send "50% off {not a key}", 1   ' raw mode (2nd arg=True/1): typed literally,
                                    ' like AutoIt's Send($keys, 1) -- same as kb.Type
kb.Type "verbatim text"       ' literal text, no special-key parsing
kb.Down "shift" : kb.Up "shift"   ' hold/release a key by name or character

' --- Clipboard (Windows) — an object now (was ClipPut/ClipGet) ---
Dim clip : Set clip = New Clipboard
clip.Set "hello from Directive"       ' returns True on success
directive.echo clip.Get               ' reads clipboard text back
clip.Clear                            ' empties the clipboard

' --- Sound playback (Windows) ---
Dim snd : Set snd = New Sound
snd.Play "chime.wav"          ' async by default (returns immediately)
snd.Play "chime.wav", True    ' True = synchronous (waits until finished)
snd.Stop                      ' stop async playback
snd.Beep 1000, 250            ' frequency Hz, duration ms

' --- Sound WAV combining: concatenate WAV files into one (all platforms) ---
Dim w : Set w = New Sound
w.Create "out.wav"
w.Append "a.wav"              ' first file sets the format
w.Append "b.wav"              ' must match channels / sample-rate / bit-depth
directive.echo "combined length: " & w.Duration & " sec"
w.Close                        ' writes the RIFF/WAVE header and closes
```

The WAV side of `Sound` streams each input's PCM straight into the output and
patches the header on `Close` — direct concatenation of same-format uncompressed
PCM only, so mismatched files raise an error rather than silently corrupting the
output.

> **Platform note.** `Mouse`, `Screen`, `Keyboard`, `Clipboard`, and `Sound`
> *playback* (`Play`/`Stop`/`Beep`) are **Windows-only** (Win32 input, GDI, and
> `winmm`); on Linux/macOS they raise a clear "requires the Windows build" error.
> `Sound`'s WAV combining (`Create`/`Append`/`Duration`/`Close`) is pure file I/O
> and works everywhere. See the verification note under *Known limitations*.

## Shell — a COM-free `the classic Windows shell`

`New Shell` is a native replacement for `the classic Windows shell`, so scripts that shell
out, read the environment, or touch the registry keep working even if Microsoft
retires the Windows Script Host objects. The process and environment features are
cross-platform; registry and special-folder lookups are Windows-only.

```vb
Dim sh : Set sh = New Shell

' Run a program; waitOnReturn = True blocks and returns the exit code
Dim rc : rc = sh.Run("notepad.exe", 1, True)

' Exec: run a command and capture its output (cross-platform)
Dim e : Set e = sh.Exec("git status")
Do Until e.StdOut.AtEndOfStream
    directive.echo e.StdOut.ReadLine
Loop
directive.echo "exit code: " & e.ExitCode

' Environment + expansion (cross-platform)
directive.echo sh.Environment("PROCESS")("PATH")
directive.echo sh.ExpandEnvironmentStrings("Temp is %TEMP%")
directive.echo sh.CurrentDirectory

' Registry + special folders (Windows only; raise 429 elsewhere)
sh.RegWrite "HKCU\Software\MyApp\Setting", "on", "REG_SZ"
directive.echo sh.RegRead("HKCU\Software\MyApp\Setting")
directive.echo sh.SpecialFolders("Desktop")

' Create a desktop shortcut, then focus a window (Windows only)
Dim lnk : Set lnk = sh.CreateShortcut(sh.SpecialFolders("Desktop") & "\MyApp.lnk")
lnk.TargetPath = "C:\Windows\System32\notepad.exe"
lnk.Description = "Launch Notepad"
lnk.Save
If sh.AppActivate("Notepad") Then directive.echo "focused Notepad"

' Popup: a message box that returns the button code
Dim answer : answer = sh.Popup("Continue?", 0, "Confirm", 4)   ' 4 = dsYesNo
```

`Run` takes `(command, [windowStyle], [waitOnReturn])`; `windowStyle` is ignored
off-Windows. `Exec` returns an object with `.StdOut` (a readable stream:
`.ReadAll` / `.ReadLine` / `.AtEndOfStream`), `.ExitCode`, and `.Status`.
`Environment([type])` is a collection — `env(name)` reads or (via assignment)
writes a variable, `.Count` counts them, and `For Each` yields `NAME=VALUE`.
`CreateShortcut(path)` returns a shortcut object whose `.TargetPath`,
`.Arguments`, `.WorkingDirectory`, `.IconLocation`, and `.Description` you set
before calling `.Save`; `AppActivate(title)` focuses a top-level window by
(partial) title.

> **Platform note.** `Run`, `Exec`, `Environment`, `ExpandEnvironmentStrings`,
> `CurrentDirectory`, and `Popup` work everywhere. `RegRead`/`RegWrite`/
> `RegDelete` (via `advapi32`), `SpecialFolders` (via `shell32`), `CreateShortcut`
> (via `IShellLink`), and `AppActivate` are Windows-only and raise error 429 on
> other platforms. The cross-platform parts are tested on Linux; the registry
> round-trip, `SpecialFolders`, `CreateShortcut`, and the `Clipboard` object were
> also exercised through the Windows binary under Wine. (Only keystroke delivery,
> real window focus, and audio output can't be tested here — no desktop/audio.)

## Console control — colours, cursor, keys

VBScript has no console API whatsoever, which is exactly why batch files reach for
`color` and `cls`. Directive adds a `Console` object that works on every platform:
the Win32 console API on Windows (honoured by both conhost and Windows Terminal),
ANSI escapes on Linux and macOS. When output is redirected every escape is
suppressed, so `directive report.ds > report.txt` still produces clean text.

```vbs
Dim c : Set c = New Console
c.Cls
c.WriteLine " BUILD PASSED ", dsBrightWhite, dsGreen   ' one-shot colour
c.Foreground = dsBrightCyan                            ' or persistent
c.WriteLine "17 tests, 0 failures"
c.Reset

c.CursorVisible = False                                ' self-overwriting progress
Dim p
For p = 0 To 100 Step 25
    c.ClearLine
    c.Write "  working: " & p & "%"
    Directive.Sleep 100
Next
c.CursorVisible = True

c.KeepOpen                                             ' batch PAUSE equivalent
```

Colours use cmd.exe's `color` numbering (`dsBlue`=1, `dsRed`=4, `+8` bright), so
`dsYellow` is the brown one and `dsBrightYellow` is vivid. See
`examples/17_console.ds` for the full tour and API_REFERENCE.md for every member.

### Keeping a double-clicked window open

The batch `PAUSE` problem: a double-clicked script gets a console that closes the
instant it finishes. Two members cover it, and neither can hang an automated run
because both require stdin to actually be a console:

| | Waits when... |
|---|---|
| `Console.Pause` | always, while a console is attached (the batch `PAUSE`) |
| `Console.KeepOpen` | **only** when this script owns the console — i.e. it was double-clicked or dragged onto, so the window would vanish. Running from an existing prompt, piping, and the GUI build are all unaffected. |

Ownership is detected with `GetConsoleProcessList() == 1`. On Linux and macOS
`KeepOpen` is deliberately a no-op, since the terminal window belongs to your
shell rather than to the script. Every example in `examples/` ends with it.

`Console.ReadKey` reads a single key with no Enter, returning the same names
`Keyboard.Send` accepts (`"a"`, `"{UP}"`, `"^a"`) so a key can be read and replayed.

## Bigger numbers than VBScript can hold

VBScript types a whole number as the narrowest that fits — Integer, then Long —
then falls back to a Double, which silently stops holding whole numbers past 2^53.
Directive adds one more rung to that same ladder, automatically:

```
Integer (16-bit) -> Long (32-bit) -> LongLong (64-bit) -> Double
```

```vbs
Directive.Echo 9007199254740992 + 1          ' 9007199254740993      exact
Directive.Echo CDbl("9007199254740992") + 1  ' 9.00719925474099E+15  what VBScript gives
Directive.Echo 3000000000 * 3000000000       ' 9000000000000000000
```

Arithmetic widens but never wraps — past 64 bits it promotes to Double, exactly as
VBScript promotes past 32 — and every surrounding VBScript rule is untouched.
`CLngLng` converts explicitly. See `examples/19_bignumbers.ds`, the numeric-type
table in API_REFERENCE.md, and `tests/EXPECTED_DIFFS.md` for the one visible
consequence.

## Locale and dates behave like VBScript

VBScript formats numbers and dates through the Windows locale; Directive asks the
same OS, so on your machine `CStr(2.5)` gives `2,5` if that is what VBScript gives,
and `CDbl("1.5")` is 15 where the dot is a thousands separator. `Date` is a real
subtype too, so `"Started: " & Now` prints a date rather than a serial number.

`SetLocale(1033)` pins invariant formatting when a file has to be machine-readable,
and `SetLocale(0)` restores the machine default. See the locale and date tables in
API_REFERENCE.md.

## A window for text, when a message box will not do

`New TextWindow` is a resizable window with a multiline edit control and OK/Cancel.
It fills the gap between `MsgBox` (too small, cannot be copied from comfortably) and
writing a temporary file just to read something:

```vbs
Dim w : Set w = New TextWindow
w.Title = "Results"
w.WriteLine "anything, including caf" & ChrW(233) & " and 100" & ChrW(8364)
w.ReadOnly = True
w.Show                                  ' user can scroll, select, copy
```

Leave `ReadOnly` off and it becomes a large InputBox — prefill `Text`, and `Show`
returns True for OK or False for Cancel, with the edited `Text` read back after.
It works in both builds and, being a real window, is immune to the console
codepage. Windows only. See `examples/20_textwindow.ds`.

## Language coverage

Everything dynamic you'd expect: `Option Explicit`, `Dim`/`Public`/`Private`,
`Const`, all operators (`Mod`, `\`, `^`, `&`, `Is`, `And/Or/Not/Xor/Eqv/Imp`),
`If/ElseIf/Else` (block and single-line), `For/Next` (+`Step`), `For Each`,
`Do/Loop` (all forms), `While/Wend`, `Select Case` (lists, `Is`, `To` ranges),
`With`, `Exit`, `Sub`/`Function` with `ByRef`/`ByVal`, `Class` with
`Property Get/Let/Set`, `Public Default` members, `Me`, `Class_Initialize`/
`Class_Terminate`, `#...#` date literals, `On Error Resume Next`/`Err`, and
~90 built-in functions (strings, conversion, math, dates, arrays, type checks,
plus `Eval`, `GetRef`, `Erase`, `MonthName`/`WeekdayName`, `FormatNumber`/
`FormatCurrency`/`FormatPercent`, `DatePart`, `DateValue`/`TimeValue`,
`ScriptEngine`; clipboard access is the `New Clipboard` object).
Constants use the Directive-branded **`ds*`** prefix: `dsCrLf`, `dsTab`,
`dsNewLine`, `dsString`, `dsYesNo`, `dsInformation`, `dsTextCompare`, and so on.
There is no `vb*` prefix — porting a classic script is a straight `vb`→`ds`
rename on constant names (`vbCrLf` → `dsCrLf`).

It also supports **`[bracket identifiers]`** — any text in `[ ]` is a literal name
(even a keyword, spaces, or a leading digit), e.g. `Dim [hello world]`, `[End] = 1`,
`obj.[Next]`. Handy for COM members whose names collide with keywords.

## Examples

See `examples/`. Highlights: `05_list.ds`, `06_dictionary.ds`,
`07_files.ds` (File/Folder objects), `09_console_io.ds`,
`11_automation.ds` (Mouse/Screen/Keyboard/Clipboard/Sound),
`12_include.ds` (native `Include` pulling in `lib/strutil.ds`),
`13_arguments.ds` (command-line args), `14_shell.ds` (the `Shell` object),
`15_unicode.ds` (UTF-8 / UTF-16 text files),
`16_bracket_identifiers.ds` (`[bracket]` names).

```sh
./directive examples/05_list.ds
```

## Extending

- **New built-in function** → add a lambda in `Interpreter::registerBuiltins()`.
- **New `New`-able object** → implement the `IObject` interface and add a case in
  `Interpreter::evalNew()`.

Pipeline is `Lexer` → recursive-descent `Parser` → tree-walking `Interpreter`, all
in one file.

## Known limitations

- **Objects — `Set` vs plain assignment.** `Set x = obj` is the explicit form.
  Directive is *lenient* about a missing `Set`: plain `x = obj` also stores the
  object reference, whereas a stricter language would try to read the object's
  default property. So correct `Set`-using code behaves identically; only code
  that deliberately relied on the stricter rule differs. (Not a bug — a superset.)
- **Arrays copy on assignment** — `b = a` gives `b` its own copy, so changing
  `b(0)` doesn't touch `a(0)`. The copy is shallow: if elements are objects, the
  references are copied (the objects aren't cloned).
- **`Class_Terminate` runs deterministically** — when you `Set x = Nothing`,
  reassign the last reference, a local goes out of scope at end of a `Sub`/
  `Function`, or the script ends. The one case where it can't run is an unbroken
  **reference cycle** (objects holding `Set` references to each other): the cycle
  is never collected, so it neither terminates nor frees — the same standard reference-counting behavior.
  Break cycles (`Set a.other = Nothing`) if you need the terminators.
- Dates are stored as OLE serial doubles. `#...#` date literals, `DateValue`/
  `TimeValue`, and all the date functions work, but because there's no distinct
  Date subtype, `TypeName` of a date reports `Double` and `IsDate(#...#)` is
  `False` (the literal is already a number). `Rnd` isn't bit-identical to any legacy
  generator.
- File/Folder `DateCreated`/`DateLastAccessed` report the last-modified time
  (portable `std::filesystem` exposes only that); `Attributes` are approximated on
  non-Windows.
- **`Private` class members aren't access-controlled.** They work correctly
  inside the class, but reading one from outside returns `Empty` rather than
  raising an error the way a stricter language would. Treat `Private` as a
  convention here, not an enforced boundary.
- **Leading-dot argument to a paren-less call.** Because the lexer discards
  whitespace, `Directive.Echo .Member` inside a `With` parses as member chaining
  (`Directive.Echo.Member`) and fails. Parenthesize the argument instead:
  `Directive.Echo(.Member)`.
- **Verification status of the automation objects.** `Sound`'s WAV combining and
  the `Clipboard` object are tested (natively where applicable and through the
  Windows `.exe` under Wine); `Shell.CreateShortcut` is also Wine-verified (it
  produces a real `.lnk`). `Mouse`, `Screen`, `Keyboard`, `Sound` *playback*, and
  `Shell.AppActivate` are **compile- and link-verified for Windows but not fully
  runtime-tested** — the build environment has no display, mouse, or audio device,
  so keystroke delivery, real window focus, and audio can't be exercised here.
  Give those a real-Windows smoke test before relying on them.

## Safety & robustness (read before trusting it with real work)

Directive is a capable hobby interpreter, tested with the bundled examples, a set
of targeted stress tests, and an AddressSanitizer/UBSan/LeakSanitizer pass. It has
**not** been fuzzed at scale or security-audited to production-runtime standards.
Concretely:

- **It is not a security sandbox.** It deliberately exposes powerful
  capabilities — `FileSystem` (read/write/delete files), `CreateObject` (arbitrary
  COM on Windows, including shells), OS automation (`Mouse`/`Screen`/`Keyboard`/`Clipboard`),
  and `Include` (reads arbitrary files). A `.ds` script can do anything the
  user running it can do, so **treat an untrusted script exactly like an untrusted
  `.exe` — don't run it.** (Directive isn't wired into browsers, Office, or email,
  so it isn't an ambient attack surface; but the language itself is just as
  capable as any general-purpose script host.)
- **Reference cycles leak.** Objects are reference-counted; if two objects hold
  `Set` references to each other (`a.other = b : b.other = a`) and you drop all
  outside references, that cluster won't be freed until the process exits. This is a standard
  reference-counting limitation. Harmless for short-lived scripts (the OS
  reclaims everything on exit); avoid unbroken cycles in long-running processes.
- **Recursion is bounded** at ~2000 nested calls; deeper raises a catchable
  *Out of stack space* (code `800A001C`) instead of crashing. The shipped Windows
  builds also link with a 16 MB stack. Normal recursion is unaffected.
- **Use-after-close is safe** (writing to a closed `TextStream` is a no-op, not a
  crash), and deeply nested `Select`/`For`/`Do`/`If`/`With` and class bodies behave
  correctly.
- Bottom line: fine for personal automation, build scripts, and trusted internal
  tooling — the niches where you'd have reached for `.vbs`. Don't put it in front of
  untrusted input or use it as a security boundary.

## License

MIT.
