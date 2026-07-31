# tbFileSysTools

A modern replacement for Scripting Runtime's `FileSystemObject`, written in twinBASIC.

It does everything FSO does, plus many things FSO cannot: **full text-encoding support**, **files larger than 2 GB**, **line-ending detection and normalization**, and **long paths beyond the legacy 260-character limit**.

```vba
' Read a UTF-8 file with a BOM, a UTF-16LE file, and a Shift-JIS file.
' You don't have to know which is which.
Dim text As String
text = TextFileToString("data.txt")     ' encoding auto-detected
```

---

## Comparison with Scripting FSO

`Scripting.FileSystemObject` shipped in 1996 and shows it:

|Feature|FSO|tbFileSysTools|
|-|-|-|
|Text encodings|ASCII and UTF-16 only|Any Windows code page, with BOM handling and auto-detection|
|Files > 2 GB|Reads them as **empty**, silently|Streams them, in both directions|
|Appending to a > 2 GB file|Not possible|Works|
|Line endings|No support|Detect, preserve, normalize|
|Paths > 260 chars|Fails, and `FileExists` **returns False** on a file that exists|Supported transparently|
|Junction in a folder tree|`Folder.Size` double-counts, or recurses forever|Skipped|
|Unreadable folder|Reports it as empty|Raises|

Everything above is measured - see **Verification** section below.

---

## Install

The library comes in three forms — pick whichever fits how you work:

| Form | Use it when | Updates |
|-|-|-|
|**twinBASIC package**|You want a referenced `.twinpack`, no copied code|Automatic via the package server|
|**Single-file drop-in**|You want one `.twin` in your project, no reference|Manual|
|**ActiveX DLL**|You're calling from VBA, VBScript, or another COM host|Re-register the DLL|

> **Note:** You don't have to clone this repo to use the package. In your own project, go to **Project → References → Available Packages** and check **File System Tools** - twinBASIC pulls it from the package server. Clone the repo only if you want to build the DLL, modify the source, or host your own local package.

---

## Two ways in

The library has one implementation and two pathways to it.

### FileSystemTools — the Standard Module

Preferred for twinBASIC code. Call it directly; there's no object to create. Potentially smaller compile footprint than the object version.

```vba
If FileExists(path) Then
    Debug.Print GetFile(path).Size
End If
```

### FileSystemObject — the Class Module

A drop-in replacement for the Scripting Runtime object. Use it for hosts that need an object, or when porting existing FSO code that you'd rather not rewrite.

```vba
Dim fso As New FileSystemObject
Debug.Print fso.GetFile(path).Size
```

The class is a one-line delegation to the module for every member — same behaviour, same defaults. It uses FSO's argument names (`fileSpec`, `folderSpec`) so named arguments in ported code keep working.

> **Note:** the class is deliberately named `FileSystemObject`, the same as Scripting's. In a project referencing both, qualify it — `tbFileSysTools.FileSystemObject` — drop the Scripting Runtime reference, or move tbFileSysTools position in the References dialog ABOVE Scripting Runtime.

---

## Text encodings

The headline feature. FSO can read ASCII and UTF-16. This reads anything Windows has a code page for.

```vba
' Auto-detect: BOM first, then UTF-16/32 and UTF-8 heuristics, then system ANSI.
Dim s As String
s = TextFileToString("mystery.txt")

' Or be explicit. A BOM in the file still wins.
s = TextFileToString("legacy.txt", encGB18030)

' Write with a BOM by choosing a BOM-bearing encoding.
StringToTextFile s, "out.txt", encUtf8Bom

' What encoding IS this file?
Debug.Print GetFileEncoding("mystery.txt")   ' e.g. encUtf16Bom
```

Supported: UTF-8, UTF-16 LE/BE, UTF-32 LE/BE, UTF-7, GB2312, GB18030, Big5, Latin-1, Latin-9, US-ASCII, system ANSI — each with and without a BOM (if applicable) — plus any other code page installed on the machine.

**Auto-detection on Read.** Defaults to auto-detect using a reliable heuristical algorithm but user can specify a code-page if known.

### Line endings

```vba
Debug.Print GetFileLineEnding("script.sh")   ' nlUnix

' Rewrite in place: UTF-8, CRLF, attributes preserved. Skips the write if the
' file is already in that form.
NormalizeTextFile "script.sh", encUtf8, nlWindows
```

Appending to an existing file adopts **that file's** newline style, so you can't accidentally turn a clean file into a mixed one.

---

## Large files

FSO cannot read a file larger than 2 GB. It doesn't error — it returns an empty string. This library streams them.

```vba
Dim ts As TextStream
Set ts = OpenTextFile("huge.log", ForReading)
Do Until ts.AtEndOfStream
    ProcessLine ts.ReadLine()       ' no size limit
Loop
ts.Close

' Appending to a 4 GB file also works.
Set ts = OpenTextFile("huge.log", ForAppending)
ts.WriteLine "another line"
ts.Close
```

`ReadLine`, `Read`, `Write` and append are **unbounded**. `ReadAll` is capped at 2 GB by the size of a VB `String` and raises a clear error rather than misbehaving — use `ReadLine` for anything larger.

Multi-byte encodings are handled correctly across chunk boundaries, including surrogate pairs split by a 64 KB read. This is verified against 54 million lines of mixed 1-, 2-, 3- and 4-byte characters (see below).

---

## Long paths

Windows' legacy `MAX_PATH` limit is 260 characters. `Scripting.FileSystemObject` is bound by it - on a path longer than 260 characters, FSO's `FileExists` returns False for a file that exists, `GetFile` and `OpenTextFile` raise "path not found".

This library supports long paths transparently. Reading, writing, creating, copying, moving, deleting, enumerating, normalizing and merging all work well past 260 characters — no prefix, no flag, no special call. You pass a normal path; the library canonicalizes it and applies the `\\?\` prefix to the underlying Win32 calls when needed.

```vba
' A 400-plus-character path is just a path.
Dim deep As String
deep = "C:\...\a\very\deeply\nested\...\structure\notes.txt"   ' > 260 chars
StringToTextFile "hello", deep
Debug.Print FileExists(deep)                   ' True
Debug.Print TextFileToString(deep)             ' hello
```

A few members stay bound to 260 characters, because the specific Win32 APIs behind them don't honor the `\\?\` prefix. Each degrades or raises rather than returning a wrong answer:

|Member|On a > 260 path|Why|
|-|-|-|
|`GetFileType` / `File.Type`|Falls back to the lexical `"<EXT> File"`|`SHGetFileInfoW` (shell API) rejects `\\?\`|
|`GetFileVersion` / `File.Version`|Returns `""`, as for a file with no version resource|Version-resource APIs don't honor `\\?\`|
|`GetRelativePath`|Falls back to the absolute target path|`PathRelativePathToW` is capped at `MAX_PATH`|
|`File.ShortPath`|Returns the full path unchanged|8.3 shortening is a legacy-`MAX_PATH` mechanism|
|Wildcard patterns in `CopyFile` / `MoveFile` / `DeleteFile` / `CopyFolder` / `MoveFolder` / `DeleteFolder` / `GetFilePaths`|The pattern is 260-bound; matched items are handled normally|A wildcard pattern can't be safely `\\?\`-prefixed|
|`SetCurrentDir`|Raises; the CWD is left unchanged|`SetCurrentDirectoryW` loses the limit only under the process-wide long-path opt-in, which a library can't guarantee|

> **Note:** tbFileSysTools does not require the Windows long-path opt-in
> (`LongPathsEnabled` + `longPathAware`). Long paths work regardless of registry
> or manifest configuration, and behave identically with the opt-in enabled.

---

## Objects

`File`, `Folder`, `Drive` and their collections work as they do in FSO.

```vba
Dim f As Folder
Set f = GetFolder("C:\\Projects")

Debug.Print f.Size                          ' total bytes, subtree
For Each fl In f.Files
    Debug.Print fl.Name, fl.Size, fl.DateLastModified
Next
```

`File` and `Folder` objects are **live**: every property read re-stats the path, so values are never stale — and a deleted file raises rather than reporting stale values. The trade-off is that each property read costs a round trip, so hoist values out of tight loops.

`Folder.Files` and `Folder.SubFolders` return a **snapshot** of the membership. The objects inside are live; the list is not.

---

## Deviations from FSO

Parity is the goal, but not at any price. Each of these was checked against the `Scripting.FileSystemObject`, and broken deliberately:

**`Folder.Size` does not follow directory reparse points.** FSO does. A junction pointing into its own subtree makes FSO double-count (measured: 2500 vs 1500), and a junction pointing at an ancestor makes it recurse until it dies. This library reports what physically lives in the tree.

**An unreadable folder raises.** FSO silently reports it as empty. A size or file list that quietly omits a subtree that can't be read is worse than an error.

**`FileAttribute.Volume` and `.Alias` are not provided.** `Volume` has no Win32 equivalent (use `Drive.VolumeName`). `Alias` is `FILE\_ATTRIBUTE\_REPARSE\_POINT` under a misleading name — and collides with VBA's `vbAlias` (64 vs 1024). Use `ReparsePoint`.

---

## Beyond FSO...

The following members are either added, or their function significantly improved.

|Member|Description|
|-|-|
|`CreateTextFile`|Create any format - not just ascii/utf16|
|`OpenTextFile`|Format auto-detection or user-specified code page|
|`GetFileEncoding`|Detect a file's encoding|
|`GetFileLineEnding`|Detect CRLF / LF / CR / mixed|
|`MergeTextFiles`|Merges two text files, normalizing the encoding to the first|
|`NormalizeTextFile`|Rewrite encoding + newlines in place, idempotently|
|`GetFilePaths`|Enumerate with a wildcard, recursion, hidden/system filters|
|`GetRelativePath`|Path from A to B|
|`CleanFileName`|Strip reserved characters and device names|
|`Rename`|In-place rename|
|`GetFileType`|Shell type description ("Text Document")|
|`GetCurrentDir` / `SetCurrentDir`|Gets/Sets the current directory or folder.|
|`GetSpecialFolder`|25 known folders, including FSO's 3|
|`ReadStream` / `WriteStream`|Raw bytes, Unicode-safe|
|`TextFileToString` / `StringToTextFile`|Whole-file text I/O|
|`TextFileToArray` / `ArrayToTextFile`|Whole-file line I/O|
|`File.Encoding`|Detect a file's encoding|
|`File.HasAttribute`|Determines if an attribute is set|
|`File.LineEnding`|Detect CRLF / LF / CR / mixed|
|`File.Normalize`|Rewrite encoding + newlines in place, idempotently|
|`File.OpenAsTextStream`|Format auto-detection or user-specified codepage|
|`File.SetAttribute`|Sets a single attribute|
|`File.ToStream`|Reads file to byte array|
|`File.ToString`|Reads file to string|
|`File.Version`|Gets the file version string|
|`Folder.HasAttribute`|Determines if an attribute is set|
|`Folder.SetAttribute`|Sets a single attribute|
|`TextStream.IsStreaming`|Whether byte-streaming or buffered access is supported|
|`TextStream.Encoding`|Returns a file's encoding|

---

## Errors

Error numbers follow FSO convention, so existing handlers keep working:

|||
|-|-|
|**5**|Invalid argument, or content isn't decodable text|
|**52**|Bad file name|
|**53**|File not found|
|**58**|File already exists|
|**61**|Disk full|
|**68**|Device unavailable|
|**70**|Permission denied|
|**71**|Disk not ready|
|**76**|Path not found|

---

## Verification

The large text file, encoding and long-path claims above were measured through comprehensive testing. The test modules can  be found in the **Tests** folder.

|Claim|Evidence|
|-|-|
|Streams > 2 GB text file|4.5 GB file written and read back block-by-block; every block verified in place|
|Append to a > 2 GB text file is correct and non-destructive|Head, size and tail all verified after appending to a 2.4 GB file|
|Multi-byte encodings survive chunk boundaries|54 M lines of UTF-16LE and UTF-8 containing 1-, 2-, 3- and 4-byte characters and surrogate pairs; line length chosen so chunk boundaries sweep every character position. Zero errors|
|Line/column tracking is exact|14 CR/LF edge cases, plus 35 M lines end to end|
|Long paths work end to end|>400-character folder tree built by the library's own `CreateFolder`; write, read, normalize and in-place merge round-trips all pass, including a temp name that crosses 260 during an atomic swap|
|FSO is 260-bound and misreports|Oracle test: on a >260 file, the real `Scripting.FileSystemObject` returns `FileExists = False` and raises 53/76 from `GetFile` / `OpenTextFile`|
|FSO parity|Differential tests against the real `Scripting.FileSystemObject`|

---

## Structure

| File | Description |
|-|-|
|`FileSystemTools`|The API where all the logic lives|
|`FileSystemObject`|COM-creatable thin wrapper class over the above|
|`TextStream`|Streaming text reader/writer|
|`TextCodec`|Encoding detection, encode/decode, BOM handling|
|`FSTShared`|Shared file/folder/drive procs|
|`WinAPI`|WinDevLib's Win32 declarations - not needed if referenced to WinDevLib|
|`File`, `Folder`, `Drive`|Objects|
|`Files`, `Folders`, `Drives`|Collections|

---

## License

MIT © 2026 GCUser99
