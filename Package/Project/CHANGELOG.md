# Change Log

[v1.9.0.0, 14 Oct 2026]

 - The SearchAborted property has been removed (v1.8). A search that exhausts its step budget now 
   raises instead of returning a value. The flag only existed so a caller could detect an answer the engine
   had not actually established — with nothing silently returned there is nothing left to query.
 - A search that runs out of budget raises error vbObjectError + 3016. Running out of budget means the 
   answer is unknown, and reporting a result the search never reached is a silent wrong answer. The message
   names the budget, the input length and, where raising the limit would plausibly help, a value to try.
 - The step budget now scales with the SQUARE of the input length, because ordinary backtracking does.
   "\d+x" over n digits with no "x" costs about 2.25 × n² steps, so a budget linear in n loses for large
   enough n whatever constant is chosen — only the crossover moves. "\d+x" over 5,000 digits used to give up
   and report no match; it now completes. [@KubaO, tbDocs]
 - New SearchStepLimit property. 0 (the default) means the automatic budget. Raise it when a genuinely
   quadratic pattern on a long input exceeds the cap; the abandonment error message suggests a value. Nothing
   can cover a pattern whose cost grows faster than the square of the input, and the message says so.
 - Pattern coverage: closed the remaining gaps found by the tbRegExp test suite in the twinBASIC
   documentation [@KubaO, tbDocs]
      - "a{2,1}" is now rejected when the Pattern is set.
      - "{" is literal text when it starts no valid quantifier, as in VBScript: "a{", "a{x}", "a{}", "a{1,2,3}"
         and "\p{L}" all compile and match the braces as text. A brace that does form a quantifier with unusable
         numbers still raises.
      - "a{,2}" is text rather than "{0,2}" — ECMAScript requires digits before the comma, so "a{,2}" no longer nmatches "aa"
      -  a "<digits>" escape is a backreference only if that group has already been OPENED at that point in the
         pattern, and a legacy octal escape otherwise. "\1(a)" is ChrW(1) followed by "a", not a backreference to
         a group declared later
      -  "$~" in a replacement string is copied as text instead of being dropped
 - IsValid, LastError and LastErrorNumber no longer leave Err populated after an invalid pattern. [@KubaO, tbDocs]

[v1.8.0.0, 13 Oct 2026]

 - Long matches are no longer truncated. The engine’s step limit was a fixed 10,000. The budget now scales
   with the length of the string being searched, so a match is bounded by the input rather than by a constant.
 - Large searches are much faster. That same fixed limit was also a performance bug — a failing scan restarted
   at every position and burned the full allowance at each one. A 9,500-character search that took ~13 seconds
   now takes ~0.02 seconds.
 - New SearchAborted property (read-only). True if the last Test, Execute, Replace or Split gave up before 
   finishing its search; results may be truncated or missing, and nothing is raised. Only a catastrophically
   backtracking pattern such as (a+)+b can reach it, and because the budget is capped such a pattern now fails
   in seconds instead of hanging the host the way VBScript.RegExp does.
 - Replace allocates its result buffer up front and skips an empty copy between adjacent matches (~3% faster);
   capture arrays are reused across the matches of a call rather than reallocated per match.

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
