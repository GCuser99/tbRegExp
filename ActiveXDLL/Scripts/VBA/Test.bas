Attribute VB_Name = "Test"
Sub test_vba_regex()
    Dim re As New RegExp
    Dim mc As MatchCollection
    Dim m As Match
    Dim source As String
    
    source = "On Jul-4-1776, independence was declared. On Apr-30-1789, George Washington became the first president."
    
    re.pattern = "(?<month>\w{3})-(?<day>\d{1,2})-(?<year>\d{4})"
    re.GlobalMatch = True
    're.Flags = "g"
    
    Debug.Assert re.Test(source)
    
    Set mc = re.Execute(source)
    
    Debug.Assert mc(0).Value = "Jul-4-1776"
    Debug.Assert mc(1).Value = "Apr-30-1789"
    
    Debug.Assert mc(0).SubMatches(0) = "Jul"
    Debug.Assert mc(0).SubMatches(1) = "4"
    Debug.Assert mc(0).SubMatches(2) = "1776"
    
    For Each m In mc
        Debug.Assert m.SubMatches("month") = m.SubMatches(0)
        Debug.Assert m.SubMatches("day") = m.SubMatches(1)
        Debug.Assert m.SubMatches("year") = m.SubMatches(2)
    Next m
    
    Debug.Assert re.Replace(source, "$<month>") = "On Jul, independence was declared. On Apr, George Washington became the first president."
    Debug.Assert re.Replace(source, "$1") = "On Jul, independence was declared. On Apr, George Washington became the first president."
    
    re.pattern = "On \w{3}-\d{1,2}-\d{4}, "
    Dim c As Collection
    Set c = re.Split(source)
    Debug.Assert c.Count = 3

    Debug.Assert c(1) = vbNullString
    Debug.Assert c(2) = "independence was declared. "
    Debug.Assert c(3) = "George Washington became the first president."
    
    re.pattern = "On (\w{3}-\d{1,2}-\d{4}), "
    Set c = re.Split(source)
    Debug.Assert c.Count = 5

    Debug.Assert c(1) = vbNullString
    Debug.Assert c(2) = "Jul-4-1776"
    Debug.Assert c(3) = "independence was declared. "
    Debug.Assert c(4) = "Apr-30-1789"
    Debug.Assert c(5) = "George Washington became the first president."
    
    Debug.Print "tests completed!"
End Sub

Sub test_vbs()
    Dim re As Object 'VBScript_RegExp_55.RegExp
    Dim mc As Object 'VBScript_RegExp_55.MatchCollection
    Dim m As Object 'VBScript_RegExp_55.Match
    Dim source As String
    
    Set re = CreateObject("VBScript.RegExp")
    'Set re = New VBScript_RegExp_55.RegExp
    
    source = "On Jul-4-1776, independence was declared. On Apr-30-1789, George Washington became the first president."
    
    re.pattern = "(\w{3})-(\d{1,2})-(\d{4})"
    re.Global = True
    
    Debug.Assert re.Test(source)
    
    Set mc = re.Execute(source)
    
    Debug.Assert mc(0).Value = "Jul-4-1776"
    Debug.Assert mc(1).Value = "Apr-30-1789"
    
    Debug.Assert mc(0).SubMatches(0) = "Jul"
    Debug.Assert mc(0).SubMatches(1) = "4"
    Debug.Assert mc(0).SubMatches(2) = "1776"
    
    Debug.Assert re.Replace(source, "$1") = "On Jul, independence was declared. On Apr, George Washington became the first president."
    
    Debug.Print "tests completed!"
End Sub

