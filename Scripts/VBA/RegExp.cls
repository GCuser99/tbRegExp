VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "This class exposes a drop-in replacement for VBSript's regular expression object model."
'@ModuleDescription "This class exposes a drop-in replacement for VBSript's regular expression object model."
'@Exposed
Option Explicit
' ==========================================================================
' RegExp version 1.00
'
' A drop-in replacement for VBSript's regular expression object model 
' by wrapping sihlfall's excellent vba-regex (https://github.com/sihlfall/vba-regex)
'
' The dependency for this class is the StaticRegex module. For how to obtain the 
' latest copy, please see following link:
' https://github.com/sihlfall/vba-regex/tree/master/aio
'
' Contact Info: https://github.com/GCUser99
' ==========================================================================
' MIT License
'
' Copyright (c) 2025, GCUser99
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
' ==========================================================================   
Private Enum BytecodeDescriptionConstant
    BYTECODE_IDX_MAX_PROPER_CAPTURE_SLOT = 0
    BYTECODE_IDX_N_IDENTIFIERS = 1
    BYTECODE_IDX_CASE_INSENSITIVE_INDICATOR = 2
    BYTECODE_IDENTIFIER_MAP_BEGIN = 3
    BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE = 3
    BYTECODE_IDENTIFIER_MAP_ENTRY_START_IN_PATTERN = 0
    BYTECODE_IDENTIFIER_MAP_ENTRY_LENGTH_IN_PATTERN = 1
    BYTECODE_IDENTIFIER_MAP_ENTRY_ID = 2
    RE_QUANTIFIER_INFINITE = &H7FFFFFFF
End Enum

Private mGlobal As Boolean
Private mMultiLine As Boolean
Private mIgnoreCase As Boolean
Private mPattern As String
Private mDotAll As Boolean
Private mIsInitialized As Boolean

Private regex As StaticRegex.RegexTy

'@Description("Sets or Gets the options in use: g-global; m-multiline; i-ignore case; s-dot all. Example: 'gm' sets Global and MultiLine to 'True', IgnoreCase and DotAll to 'False'.")
Public Property Let Flags(ByVal pFlags As String)
Attribute SameSite.VB_Description = "Sets or Gets the options in use: g-global; m-multiline; i-ignore case; s-dot all. Example: 'gm' sets Global and MultiLine to 'True', IgnoreCase and DotAll to 'False'."
    'this "convenience property was borrowed from 
    'https://github.com/sihlfall/vba-regex/blob/master/stdvba_regex3/stdRegex3.cls
    'note that if a flag indicator is not present in the string, then it is set to False
    'user setting Flags value to vbNullString effectively sets all flags to False (default state)
    pFlags = LCase(pFlags)
    mGlobal = CBool(InStr(pFlags, "g"))
    mMultiLine = CBool(InStr(pFlags, "m"))
    IgnoreCase = CBool(InStr(pFlags, "i")) 'let property to handle isInitialized flag 
    mDotAll = CBool(InStr(pFlags, "s"))
End Property
Public Property Get Flags() As String
    If mGlobal Then Flags = Flags & "g"
    If mMultiLine Then Flags = Flags & "m"
    If mIgnoreCase Then Flags = Flags & "i"
    If mDotAll Then Flags = Flags & "s"
End Property

'@Description("If 'True', all matches will be returned. If 'False', only the first match will be returned.")
Public Property Let GlobalMatch(ByVal pGlobal As Boolean)
Attribute SameSite.VB_Description = "If 'True', all matches will be returned. If 'False', only the first match will be returned."
    mGlobal = pGlobal
End Property
Public Property Get GlobalMatch() As Boolean
    GlobalMatch = mGlobal
End Property

'@Description("If 'True', '^' pattern will match the beginning of a line after a new line character and '$' will match the end of a line before a new line character. If 'False', '^' will match the beginning of the entire string and '$' will match the end of the entire string.")
Public Property Let MultiLine(ByVal pMultiLine As Boolean)
Attribute SameSite.VB_Description = "If 'True', '^' pattern will match the beginning of a line after a new line character and '$' will match the end of a line before a new line character. If 'False', '^' will match the beginning of the entire string and '$' will match the end of the entire string."
    mMultiLine = pMultiLine
End Property
Public Property Get MultiLine() As Boolean
    MultiLine = mMultiLine
End Property

'@Description("If 'True', '.' pattern will match across new line charactors.")
Public Property Let DotAll(ByVal pDotAll As Boolean)
Attribute SameSite.VB_Description = "If 'True', '.' pattern will match across new line charactors."
    mDotAll = pDotAll
End Property
Public Property Get DotAll() As Boolean
    DotAll = mDotAll
End Property

'@Description("If 'True', matches will be case-insensitive. If 'False', matches will be case-sensitive.")
Public Property Let IgnoreCase(ByVal pIgnoreCase As Boolean)
Attribute SameSite.VB_Description = "If 'True', matches will be case-insensitive. If 'False', matches will be case-sensitive."
    If pIgnoreCase <> mIgnoreCase Then
        mIgnoreCase = pIgnoreCase
        mIsInitialized = False
    End If
End Property
Public Property Get IgnoreCase() As Boolean
    IgnoreCase = mIgnoreCase
End Property

'@Description("The regular expression pattern.")
Public Property Let Pattern(ByVal pPattern As String)
Attribute SameSite.VB_Description = "The regular expression pattern."
    If pPattern <> mPattern Then
        mPattern = pPattern
        mIsInitialized = False
    End If
End Property
Public Property Get Pattern() As String
    Pattern = mPattern
End Property

'@Description("Returns a MatchCollection. Uses the Pattern property to search the sourceString.")
Public Function Execute(ByVal sourceString As String) As MatchCollection
Attribute SameSite.VB_Description = "Returns a MatchCollection. Uses the Pattern property to search the sourceString."
    Dim matcherState As StaticRegex.MatcherStateTy
    Dim oMatches As MatchCollection
    Dim oMatch As Match
    Dim i As Long
    Dim submatchString As String
    Dim matchStart As Long, matchLength As Long
    Dim colNamedKeyMap As New Collection
    
    If Not mIsInitialized Then
        StaticRegex.InitializeRegex regex, mPattern, mIgnoreCase
        mIsInitialized = True
    End If

    StaticRegex.InitializeMatcherState matcherState, Not mGlobal, mMultiLine, mDotAll

    Dim colNames As New Collection
    Dim colEntryPoints As New Collection
    
    getNamedCapturesInfo colNames, colEntryPoints
    
    Set oMatches = New MatchCollection
    
    Do While StaticRegex.MatchNext(matcherState, regex, sourceString)
        
        With matcherState.captures
            Set oMatch = New Match
            matchStart = .entireMatch.start
            matchLength = .entireMatch.Length
            If matchStart > 0 Then
                oMatch.Value = Mid$(sourceString, matchStart, matchLength)
                oMatch.FirstIndex = matchStart - 1 'to match VBScript_RegExp_55 zero-based
                oMatch.Length = matchLength
            Else
                oMatch.Value = vbNullString
            End If
            
            If .nNumberedCaptures > 0 Then
                Dim keymap() As String
                ReDim keymap(1 To .nNumberedCaptures)
            
                If .nNamedCaptures > 0 Then
                    'build a key-position map into the numbered capture collection
                    Dim referenced As Long
                    For i = 1 To colNames.Count
                        referenced = .namedCaptures(colEntryPoints(i))
                        If referenced <> -1 Then keymap(referenced) = colNames(i)
                    Next i
                End If
            
                'add captures to submatch collection - note that numbered captures includes any named captures too
                For i = 0 To .nNumberedCaptures - 1
                    matchStart = .numberedCaptures(i).start
                    matchLength = .numberedCaptures(i).Length
                    If matchStart > 0 Then
                        submatchString = Mid$(sourceString, matchStart, matchLength)
                    Else
                        submatchString = vbNullString
                    End If
                    If keymap(i + 1) = vbNullString Then
                        oMatch.SubMatches.Add submatchString
                    Else
                        'we have a named capture - use name as submatch key
                        oMatch.SubMatches.Add submatchString, keymap(i + 1)
                    End If
                Next i
            End If

            oMatches.Add oMatch
        End With
    Loop

    Set Execute = oMatches
End Function

'@Description("Returns 'True' if Pattern finds a match in sourceString, otherwise returns 'False'.")
Public Function Test(ByVal sourceString As String) As Boolean
Attribute SameSite.VB_Description = "Returns 'True' if Pattern finds a match in sourceString, otherwise returns 'False'."
    If Not mIsInitialized Then
        StaticRegex.InitializeRegex regex, mPattern, mIgnoreCase
        mIsInitialized = True
    End If
    Test = StaticRegex.Test(regex, sourceString, mMultiLine, mDotAll)
End Function

'@Description("Replaces matches with replaceString and returns the String with replacements made.")
Public Function Replace(ByVal sourceString As String, ByVal replaceString As String) As String
Attribute SameSite.VB_Description = "Replaces matches with replaceString and returns the String with replacements made."
    If Not mIsInitialized Then
        StaticRegex.InitializeRegex regex, mPattern, mIgnoreCase
        mIsInitialized = True
    End If
    Replace = StaticRegex.Replace(regex, replaceString, sourceString, Not mGlobal, mMultiLine, mDotAll)
End Function

'@Description("Splits an input string into an array of substrings at the positions defined by a regular expression match.")
Public Function Split(ByVal sourceString As String) As Collection
Attribute SameSite.VB_Description = "Splits an input string into an array of substrings at the positions defined by a regular expression match."
    If Not mIsInitialized Then
        StaticRegex.InitializeRegex regex, mPattern, mIgnoreCase
        mIsInitialized = True
    End If
    Set Split = StaticRegex.SplitByRegex(regex, sourceString, Not mGlobal, mMultiLine, mDotAll)
End Function

Private Sub getNamedCapturesInfo(ByRef colNames As Collection, ByRef colEntryPoints As Collection)
    'modified from https://github.com/sihlfall/vba-regex/blob/master/stdvba_regex3/stdRegex3.cls
    Dim identifierMapEnd As Long
    Dim currentEntry As Long
    Dim identifier As String
    
    identifierMapEnd = BYTECODE_IDENTIFIER_MAP_BEGIN + _
        BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE * regex.bytecode(BYTECODE_IDX_N_IDENTIFIERS)
    currentEntry = BYTECODE_IDENTIFIER_MAP_BEGIN

    Do While currentEntry < identifierMapEnd
        identifier = Mid$( _
            regex.pattern, _
            regex.bytecode(currentEntry + BYTECODE_IDENTIFIER_MAP_ENTRY_START_IN_PATTERN), _
            regex.bytecode(currentEntry + BYTECODE_IDENTIFIER_MAP_ENTRY_LENGTH_IN_PATTERN) _
        )
        colNames.Add(identifier)
        colEntryPoints.Add(regex.bytecode(currentEntry + BYTECODE_IDENTIFIER_MAP_ENTRY_ID))
        currentEntry = currentEntry + BYTECODE_IDENTIFIER_MAP_ENTRY_SIZE
    Loop
End Sub

Private Sub Class_Initialize()
    mGlobal = False 'StaticRegex defaults to True
    mMultiLine = False
    mIgnoreCase = False
    mDotAll = False
End Sub


