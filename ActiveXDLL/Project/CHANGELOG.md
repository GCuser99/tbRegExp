# Change Log

[v1.7.0.0, 12 Oct 2026]

- Pattern coverage: closed the escape gaps reported by RegExpress [@Viorel, GitHub]
   - "[\f]" is a form feed; it used to match the letter "f"
   - "\cX" outside a character class gives the right control character ("\cJ" is vbLf); it was wrong for every letter
   - "\CX" is a control escape too, since VBScript's "\c" is case-insensitive
   - an unrecognised letter escape is now the letter itself, as in VBScript ("\q" matches "q"), instead of raising
   - Fixed IsValid reporting True for an out-of-range backreference ("\101", or "\2" with one group); these are now rejected when the Pattern is set, instead of raising from Test/Execute
   - legacy octal escapes work outside a character class too: "\101" is "A" when the pattern has no group 101, as in VBScript; an in-range backreference like "(a)\1" still wins
   - "\012"-style octal, and "\0" no longer swallows the character after it ("\0a")
   
[v1.6.0.0, 12 Oct 2026]

 - Fixed zero-width matches being reported twice when the engine had to skip forward to reach them [@KubaO, tbDocs]
   - affected Execute, Replace and Split with patterns that match empty at only some positions - multiline "^" and "$", "\b", and similar

[v1.5.0.0, 12 Oct 2026]

 - Fixed VBScript compatibility for '.' wrt to vbCr line terminator which should be accepted [@Viorel, GitHub]
 - Fixed VBScript compatibility for "$" substitution in Replace [@KubaO, tbDocs]
    - "$" that starts no substitution item is now copied as text, as VBScript does (was: error 3031)
    - "$n" now reads at most two digits with fallback, so "$100" with 2 groups is group 1 then "00" (was: empty string)
    - "$0", and "$<" with no closing ">", no longer raise raw errors
 - Performance: removed every Collection from the result path
    - MatchCollection and SubMatches are array-backed; For Each on both is served by a real IEnumVARIANT
    - named-capture metadata is cached per compiled Pattern instead of rebuilt on every Execute call
    - Compiled Execute 4–7× cheaper per call; For Each over SubMatches ~24×; per-match and Test/Replace/Split unchanged

[v1.4.0.0, 10 Oct 2026]

 - Bug fix - ^ inside a lookbehind succeeds one position too late [@Viorel, GitHub]
 - Bug fix - \b inside a lookbehind is wrong at the first two boundaries
 - Replaced As-New auto-instancing declares with explicit form

[v1.3.0.0, 19 Sept 2026]

 - Fixed VBScript compatibility: unescaped ] and } are now accepted as literals (e.g. "\[.*]")
 
[v1.2.0.0, 10 Aug 2026]

 - Improved error handling
 - Added error handling tests GitHub repo Tests

[v1.1.0.0, 10 Aug 2026]

 - Widened the allowed characters in named capture groups from a-zA-Z to a-zA-Z0-9_
 - Added illustrative test suite to GitHub repo: https://github.com/GCuser99/tbRegExp/blob/main/Tests/GitHubTests.twin
 
[v1.0.0.6, 07 Aug 2026]

 - Removed commented-out dead code

[v1.0.0.5, 07 Aug 2026]

 - Initial Public release
