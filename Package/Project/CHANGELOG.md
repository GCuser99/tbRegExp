# Change Log

[v1.4.0.0, 10 Oct 2026]

 - Bug fix - ^ inside a lookbehind succeeds one position too late [Viorel]
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
