# This script parses and reformats the all-in-one twinBASIC RegExp wrapper into
# corresponding four VBA class modules (RegExp, MatchCollection, Match, and SubMatches)
 
$inFilePath = Join-Path -Path $PSScriptRoot -ChildPath ".\twinBASIC\RegExp.twin"
$outFolderPath = ".\VBA"

$nl="`r`n"

if (-not (Test-Path -Path $outFolderPath)) {
    New-Item -Path $outFolderPath -ItemType Directory > $null
}

# Read in the twinBASIC all-in-one file data
$aio = Get-Content -Path $inFilePath -Raw

# Replace class headers with VBA version
$oldBlock = '\[COMCreatable\(False\)\]' + $nl
$oldBlock += '\[PredeclaredId\((True|False)\)\]' + $nl
$oldBlock += '\[Hidden\(True\)\]' + $nl
$oldBlock += '\[Description\((.+)\)\]' + $nl
$oldBlock += 'Private Class (\S+)'
$newBlock = 'VERSION 1.0 CLASS' + $nl
$newBlock += 'BEGIN' + $nl
$newBlock += '  MultiUse = -1  ''True' + $nl
$newBlock += 'END' + $nl
$newBlock += 'Attribute VB_Name = "$3"' + $nl
$newBlock += 'Attribute VB_GlobalNameSpace = False' + $nl
$newBlock += 'Attribute VB_Creatable = False' + $nl
$newBlock += 'Attribute VB_PredeclaredId = $1' + $nl
$newBlock += 'Attribute VB_Exposed = True' + $nl
$newBlock += 'Attribute VB_Description = $2' + $nl
$newBlock += '''@ModuleDescription $2' + $nl
$newBlock += '''@Exposed'
$aio = $aio -replace $oldBlock, $newBlock

# Remove extra white space at beginning of each line
$aio = $aio -replace '(?m)^    ', ''

# Convert Descriptions
$aio = $aio -replace '\[Description(.+)\]', '''@Description$1'

# Convert DefaultMember designations
$aio = $aio -replace '\[DefaultMember\]', '''@DefaultMember'

# Convert Enumerator designations
$aio = $aio -replace '\[Enumerator\]', '''@Enumerator'

# Remove Hidden Designations
$aio = $aio -replace "\[Hidden\]`r`n", ""

# Remove parenthesis around Err.Raise function
$aio = $aio -replace 'Err\.Raise\((.+)\)', 'Err.Raise $1'

# Insert NewEnum Attributes
$oldBlock = 'Public Function NewEnum\(\) As IUnknown' + $nl
$oldBlock += '    Set NewEnum = (.+)\.\[_NewEnum\]' + $nl
$oldBlock += 'End Function'
$newBlock =  'Public Function NewEnum() As IUnknown' + $nl
$newBlock += 'Attribute NewEnum.VB_UserMemId = -4' + $nl
$newBlock += '    Set NewEnum = $1.[_NewEnum]' + $nl
$newBlock += 'End Function'
$aio = $aio -replace $oldBlock, $newBlock

# Insert Default Attributes for SubMatches and MatchCollection classes
$oldBlock =  'Public Property Get Item\(ByVal pItemIndex As Variant\) As (.+)' + $nl
$newBlock =  'Public Property Get Item(ByVal pItemIndex As Variant) As $1' + $nl
$newBlock += 'Attribute Item.VB_UserMemId = 0' + $nl
$aio = $aio -replace $oldBlock, $newBlock

# Insert Default Attribute for RegExp Class
$oldBlock =  'Public Property Get Value\(\) As String' + $nl
$newBlock =  'Public Property Get Value() As String' + $nl
$newBlock += 'Attribute Item.VB_UserMemId = 0' + $nl
$aio = $aio -replace $oldBlock, $newBlock

# Change Global Property to GlobalMatch Property (Global is keyword in VBA)
$oldBlock = 'Public Property Let Global\(ByVal pGlobal As Boolean\)'
$newBlock =  'Public Property Let GlobalMatch(ByVal pGlobal As Boolean)'
$aio = $aio -replace $oldBlock, $newBlock

$oldBlock = 'Public Property Get Global\(\) As Boolean'
$newBlock =  'Public Property Get GlobalMatch() As Boolean'
$aio = $aio -replace $oldBlock, $newBlock

$oldBlock = 'Global = mGlobal'
$newBlock =  'GlobalMatch = mGlobal'
$aio = $aio -replace $oldBlock, $newBlock

# Add method description attributes
$oldBlock = "('@Description\((.+)\)`r`n.*(`r`n)*Public.+`r`n)"
$newBlock =  '$1Attribute SameSite.VB_Description = $2' + $nl
$aio = $aio -replace $oldBlock, $newBlock

# Split the aio into seperate classes
$classes = $aio -split '(?m)^End Class'

# Remove the lead nl's
$classes[0] = $classes[0] -replace '(?m)\s+VERSION 1.0 CLASS', 'VERSION 1.0 CLASS'
$classes[1] = $classes[1] -replace '(?m)\s+VERSION 1.0 CLASS', 'VERSION 1.0 CLASS'
$classes[2] = $classes[2] -replace '(?m)\s+VERSION 1.0 CLASS', 'VERSION 1.0 CLASS'
$classes[3] = $classes[3] -replace '(?m)\s+VERSION 1.0 CLASS', 'VERSION 1.0 CLASS'

# Output class files
Set-Content -Path (Join-Path -Path $outFolderPath -ChildPath 'RegExp.cls') -Value $classes[0] > $null
Set-Content -Path (Join-Path -Path $outFolderPath -ChildPath 'Match.cls') -Value $classes[1] > $null
Set-Content -Path (Join-Path -Path $outFolderPath -ChildPath 'MatchCollection.cls') -Value $classes[2] > $null
Set-Content -Path (Join-Path -Path $outFolderPath -ChildPath 'SubMatches.cls') -Value $classes[3] > $null