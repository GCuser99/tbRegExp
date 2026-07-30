VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Match"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class represents the results from a single regular expression match."
'@ModuleDescription "This class represents the results from a single regular expression match."
'@Exposed
Option Explicit

Private oSubMatches As SubMatches
Private mValue As String
Private mFirstIndex As Long
Private mLength As Long

'@Description("Returns the SubMatches collection of submatch (Group) strings.")
Public Property Get SubMatches() As SubMatches
Attribute SameSite.VB_Description = "Returns the SubMatches collection of submatch (Group) strings."
    Set SubMatches = oSubMatches
End Property

'@Description("Returns the entire matched string")
'@DefaultMember
Public Property Get Value() As String
Attribute SameSite.VB_Description = "Returns the entire matched string"
Attribute Item.VB_UserMemId = 0
    Value = mValue
End Property
Friend Property Let Value(ByVal pValue As String)
    mValue = pValue
End Property

'@Description("Returns the character index position associated with the start of the match string. Note that the index is zero-based.")
Public Property Get FirstIndex() As Long
Attribute SameSite.VB_Description = "Returns the character index position associated with the start of the match string. Note that the index is zero-based."
    FirstIndex = mFirstIndex
End Property
Friend Property Let FirstIndex(ByVal pFirstIndex As Long)
    mFirstIndex = pFirstIndex
End Property

'@Description("Returns the character length of the match string.")
Public Property Get Length() As Long
Attribute SameSite.VB_Description = "Returns the character length of the match string."
    Length = mLength
End Property
Friend Property Let Length(ByVal pLength As Long)
    mLength = pLength
End Property

Private Sub Class_Initialize()
    Set oSubMatches = New SubMatches
End Sub

Private Sub Class_Terminate()
    Set oSubMatches = Nothing
End Sub


