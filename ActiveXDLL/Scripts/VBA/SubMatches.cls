VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class provides the means to access SubMatch (Group) values."
'@ModuleDescription "This class provides the means to access SubMatch (Group) values."
'@Exposed
Option Explicit

Private colSubMatches As Collection

Friend Sub Add(ByVal val As String, Optional ByVal key As String)
    If key <> vbNullString Then
        colSubMatches.Add val, key
    Else
        colSubMatches.Add val
    End If
End Sub

'@Description("Returns the number of SubMatch objects found in the source string.")
Public Property Get Count() As Long
Attribute SameSite.VB_Description = "Returns the number of SubMatch objects found in the source string."
   Count = colSubMatches.Count
End Property

'@Description("Returns the SubMatch string associated with the provided index. Named groups are allowed.")
'@DefaultMember
Public Property Get Item(ByVal pItemIndex As Variant) As String
Attribute SameSite.VB_Description = "Returns the SubMatch string associated with the provided index. Named groups are allowed."
Attribute Item.VB_UserMemId = 0
    If Not IsNumeric(pItemIndex) Then
        If colSubMatches.Exists(pItemIndex) Then
            Item = colSubMatches.Item(pItemIndex)
        Else
            'all we know here is that either the key does not exists or
            'it does exists and has an empty value - so we must exit without an error...
        End If
    Else
        'VBS is zero-based, but Collection is 1-based
        If pItemIndex + 1 >= 1 And pItemIndex + 1 <= colSubMatches.Count Then
            Item = colSubMatches.Item(pItemIndex + 1) 'to match VBScript_RegExp_55 zero-based
        Else
            Err.Raise vbObjectError + 9, , "Requested item not found in SubMatches collection"
        End If
    End If
End Property

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colSubMatches.[_NewEnum]
End Function

Private Sub Class_Terminate()
    Set colSubMatches = Nothing
End Sub

Private Sub Class_Initialize()
    Set colSubMatches = New Collection
End Sub


