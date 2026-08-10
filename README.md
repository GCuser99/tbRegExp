# RegExp for twinBASIC

A drop-in replacement for VBScript's regular-expressions object model (`VBScript.RegExp` /
*Microsoft VBScript Regular Expressions 5.5*), written for [twinBASIC](https://twinbasic.com).
It exposes the familiar `RegExp` / `Match` / `MatchCollection` / `SubMatches` objects, so
existing VBScript-style regex code works unchanged — while also supporting a number of
modern regex features that `VBScript.RegExp` never had.

It is **pure twinBASIC**: no COM reference, no external DLL, no runtime dependency. The
regular-expressions engine underneath is [sihlfall's vba-regex](https://github.com/sihlfall/vba-regex),
an excellent ECMAScript-compatible matcher; this project is the compatibility wrapper around it.

---

## Why this exists

`VBScript.RegExp` has been the de-facto regex object in the VBA/VB6 world for two decades,
but it arrives via the (now-deprecated) VBScript runtime. Microsoft has since built a native
`RegExp` into the **VBA** library (Office version 2508, 2025) so VBA users no longer need the
VBScript reference — but that is VBA, not twinBASIC.

**twinBASIC has no built-in regular-expressions support today.** Its author, Wayne Phillips,
has indicated native regex is on the roadmap, but not before v1.0. Until then, tB developers
either have to depend on the deprecated `VBScript.RegExp` COM object or roll their own.

This library fills that gap now: a self-contained, dependency-free, VBScript-compatible
`RegExp` you can drop into a twinBASIC project today, with no reliance on VBScript and no
external components to ship.

---

## Credits

The heavy lifting — the actual regular-expressions compilation and matching — is
**sihlfall's vba-regex** engine:

> https://github.com/sihlfall/vba-regex

Please star and support that project. The engine module used here (`StaticRegex`) is sihlfall's
work, carrying a few small, self-contained performance tweaks (clearly documented); the original base is
available in sihlfall's [aio folder](https://github.com/sihlfall/vba-regex/tree/master/aio).

---

## Install

The library comes in three forms — pick whichever fits how you work:

| Form | Use it when | Updates |
|-|-|-|
|**twinBASIC package**|You want a referenced `.twinpack`, no copied code|Automatic via the Package Server|
|**Single-file drop-in**|You want one `.twin` in your project, no reference|Manual|
|**ActiveX DLL**|You're calling from VBA, twinBASIC or another COM host|Re-install the DLL (Inno script included)|

The package can be referenced through twinBASIC's Package Server. In the IDE, click **References -> Available Packages**. Scroll to find the entry and click the checkbox on the left of the entry. Then click **Apply Changes** button in lower-right.

<img src="https://github.com/GCuser99/tbRegExp/blob/main/Images/package install.png" alt="Package" width=90%>

---

## Quick start

```vba
Dim re As New RegExp
re.Pattern = "(\d{4})-(\d{2})-(\d{2})"
re.Global = True

Dim m As Match
For Each m In re.Execute("2025-08-26 and 2026-07-29")
    Debug.Print m.Value, m.SubMatches(0), m.SubMatches(1), m.SubMatches(2)
Next m
```

The object model mirrors `VBScript.RegExp` exactly, including zero-based indexing:

|Object|Members|
|-|-|
|`RegExp`|`Pattern`, `Global`, `IgnoreCase`, `MultiLine`, `Test()`, `Execute()`, `Replace()`|
|`Match`|`Value`, `FirstIndex`, `Length`, `SubMatches`|
|`MatchCollection`|`Count`, `Item()`, `For Each`|
|`SubMatches`|`Count`, `Item()`, `For Each`|

Existing code written against `VBScript.RegExp` should run as-is — just change
`CreateObject("VBScript.RegExp")` / the *Microsoft VBScript Regular Expressions 5.5*
reference to `New RegExp`.

---

## Beyond VBScript.RegExp

Because the engine is a modern ECMAScript-style matcher, this library supports several
things `VBScript.RegExp` (and the new native VBA `RegExp`, which is the same legacy engine)
cannot do:

**`DotAll` — let `.` match newlines.** A `DotAll` property (the ECMAScript `s` flag):

```vba
re.Pattern = "start.*end"
re.DotAll = True                                  ' now '.' spans line breaks
Debug.Print re.Test("start" & vbCrLf & "end")     ' -> True
re.DotAll = False                                 ' default: '.' stops at a line break
Debug.Print re.Test("start" & vbCrLf & "end")     ' -> False
```

**`Split` — split a string on a pattern.** A method `VBScript.RegExp` never provided:

```vba
re.Global = True
re.Pattern = "\s*,\s*"
Dim parts As Collection
Set parts = re.Split("a , b,c ,  d")   ' -> "a","b","c","d"
```

**Named capture groups** — `(?<name>...)`, retrievable by name from `SubMatches`:

```vba
re.Pattern = "(?<year>\d{4})-(?<month>\d{2})"
Dim m As Match
Set m = re.Execute("2026-07").Item(0)
Debug.Print m.SubMatches("year"), m.SubMatches("month")   ' by name
Debug.Print m.SubMatches(0),      m.SubMatches(1)          ' still by index too
```

**Named references in Replace** — use `$<name>` in the replacement string, not just `$1`:

```vba
re.Pattern = "(?<y>\d{4})-(?<m>\d{2})"
Debug.Print re.Replace("2026-07", "$<m>/$<y>")   ' -> 07/2026
```

**Lookbehind** — `(?<=...)` and `(?<!...)`, in addition to lookahead:

```vba
re.Pattern = "(?<=\$)\d+"    ' digits preceded by a literal '$'
Debug.Print re.Execute("$123").Item(0).Value  ' -> 123
```

**Inline mode modifiers** — both the whole-group form and the scoped form, for `i`/`m`/`s`:

```vba
' (?i) - case-insensitive from this point on
re.Pattern = "(?i)hello"
Debug.Print re.Test("HELLO there")            ' -> True

' (?s:...) - DotAll for the group only, so '.' spans the line break
re.Pattern = "foo(?s:.*)bar"
Debug.Print re.Test("foo" & vbCrLf & "bar")   ' -> True

' (?i:...) - case-insensitive for 'abc' only; 'DEF' stays case-sensitive
re.Pattern = "(?i:abc)DEF"
Debug.Print re.Test("ABCDEF")                 ' -> True
Debug.Print re.Test("ABCdef")                 ' -> False  ('def' is not 'DEF')
```

**A `Flags` convenience property** — set several options from one string:

```vba
re.Flags = "gis"   ' Global + IgnoreCase + DotAll  (g m i s)
```
---

## Performance

A summary: the underlying **matching engine is competitive with the native
`VBScript.RegExp` engine — and often faster** on the raw match. Where this library costs
more than `VBScript.RegExp`, the overhead is concentrated in specific pattern classes and
in building result objects, not in matching itself.

**For typical VBA/tB workloads — validation and extraction on modest strings — the
difference is negligible in absolute terms (microseconds).** The relative multiples below
look larger than they matter in practice, because they are measured on deliberately large
(200 KB) inputs to expose the costs; on everyday inputs both engines finish effectively
instantly.

> **These figures are provisional.** They were captured using a compiled build with no optimization.
> They will be refreshed once LLVM optimization is available in twinBASIC.

*Lower is better; **1.0× = parity** with`VBScript.RegExp`.*

|Operation|Example|Relative to `VBScript.RegExp`|Notes|
|-|-|-:|-|
|`Test`, short input|validate a phone/date|~1–2×|At/near parity|
|`Execute` (global), no groups|`\w+` over large text|~1–1.5×|At/near parity|
|`Execute` (global), with groups|`(\d{4})-(\d{2})-(\d{2})`|~3–5×|Cost is per-match result objects|
|`Replace`, sparse single-char class|`\d` → `#` over large text|~20–30×|Many small operations over a big string|
|`Test`, long non-matching scan|rare literal, large text|~15–25×|Dominated by scanning cost|
|Dense-first-set scan|`[\w.]+@…`, wide alternation|~7–16×|Every position is a real match attempt|

**Where it genuinely matters:** very large inputs combined with dense-scanning patterns, and
pathological catastrophic-backtracking patterns (which are bounded by an internal step limit
rather than allowed to run away). 

**Where it doesn't:** the common case of anchored/validating
patterns and extraction over normal-sized text, which is at or near parity.

If raw throughput on huge inputs with complex patterns is your priority and a runtime
dependency is acceptable, a [wrapper around a native engine such as .NET's](https://github.com/6DiegoDiego9/VBA-dotNET-regex) will probably beat a pure-tB matcher — that is the trade this library deliberately does *not* make, in favor of being
self-contained and portable.

---

## Compatibility notes and limitations

Unicode support matches `VBScript.RegExp`: literal BMP characters compare correctly and
case-insensitive matching folds BMP characters (including accented Latin, Greek, Cyrillic).
Not supported (as with `VBScript.RegExp`): `\p{...}` Unicode property classes, and
supplementary/astral characters above U+FFFF are handled as surrogate pairs rather than
single code points. `\w`, `\d`, and `\b` are ASCII, as in `VBScript.RegExp`.

Other constructs not currently supported: named backreferences (`\k<name>` — numbered
backreferences like `\1` do work), atomic groups / possessive quantifiers, `\A`/`\Z` anchors,
and conditionals.

> These limitations are shared with the engine this library replaces, and with the native
> VBA `RegExp` added in Office 2508 (which is the same legacy engine).

---

## Acknowledgements

* **sihlfall** — for [vba-regex](https://github.com/sihlfall/vba-regex), the engine that makes this possible.
* **fafalone** — for [WinDevLib](https://github.com/fafalone/WinDevLib), source of standardized Win32 API declares.
* **Wayne Phillips** — for [twinBASIC](https://twinbasic.com).

---

## License

MIT License. Copyright (c) 2025–2026, GCUser99.

The bundled engine (`StaticRegex`) is a minimally edited version of [sihlfall's vba-regex](https://github.com/sihlfall/vba-regex); MIT License.
