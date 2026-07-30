VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MatchCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class represents a collection of match objects."
'@ModuleDescription "This class represents a collection of match objects."
'@Exposed
Option Explicit

Private colMatches As Collection

Friend Sub Add(obj As Match)
    colMatches.Add obj
End Sub

'@Description("Returns the number of Match objects found in the source string.")
Public Property Get Count() As Long
Attribute SameSite.VB_Description = "Returns the number of Match objects found in the source string."
   Count = colMatches.Count
End Property

'@Description("Returns the Match object associated with the provided index")
'@DefaultMember
Public Property Get Item(ByVal pItemIndex As Variant) As Match
Attribute SameSite.VB_Description = "Returns the Match object associated with the provided index"
Attribute Item.VB_UserMemId = 0
    'VBS is zero-based, but Collection is 1-based
    If pItemIndex + 1 >= 1 And pItemIndex + 1 <= colMatches.Count Then
        Set Item = colMatches.Item(pItemIndex + 1)
    Else
        Err.Raise vbObjectError + 9, , "Requested item not found in Matches collection"
    End If
End Property

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colMatches.[_NewEnum]
End Function

Private Sub Class_Terminate()
    Set colMatches = Nothing
End Sub

Private Sub Class_Initialize()
    Set colMatches = New Collection
End Sub


