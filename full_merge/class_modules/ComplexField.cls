VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComplexField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String
Private pParent As ComplexField
Private pCell As Range
Private pNumSubFields As Long
Private pSubFields() As ComplexField
Private pComplexityLevel As Integer
Private pPath As String
Private pLowestFields As Dictionary
Private ws As Worksheet

Public Sub init(sheet As Worksheet, c As String, Optional level As Integer, Optional parent As ComplexField)
  Set ws = sheet
  Set pParent = parent
  Set pCell = ws.Range(c)
  Set pLowestFields = New Dictionary
  
  pName = Trim(eraseSPs(eraseEOLs(s:=pCell.value)))
  pPath = pName
  pNumSubFields = 0
  pComplexityLevel = level
  
  Me.buildSubFields level:=pComplexityLevel, initial:=Me
  
  If pComplexityLevel = 0 Then
    Dim pf As PrimitiveField
    Set pf = PrimitiveFieldFactory.CreatePrimitiveField(addr:=Me.address, path:=Me.path, ws:=sheet)
    pLowestFields.Add Key:=pf.path, Item:=pf
  Else
    Me.collectLowestFields level:=pComplexityLevel, initial:=Me, acc:=Me
  End If
End Sub

Public Sub addChild(child As ComplexField)
  ReDim Preserve pSubFields(pNumSubFields)
  Set pSubFields(pNumSubFields) = child
  
  pNumSubFields = pNumSubFields + 1
  
  child.setParent p:=Me
End Sub

Public Sub setParent(p As ComplexField)
  Set pParent = p
  pPath = p.path & "/" & pPath
End Sub

Public Function hasChildren() As Boolean
  hasChildren = pNumSubFields <> 0
End Function

'-------------------------------------------------'
'-----------------LowestFields--------------------'
'-------------------------------------------------'
Public Sub collectLowestFields(level As Integer, initial As ComplexField, acc As ComplexField)
  If level > 1 Then
    For i = 0 To UBound(acc.subFields)
      acc.subFields(i).collectLowestFields level:=(level - 1), initial:=initial, acc:=(acc.subFields(i))
    Next
  End If
  If level = 1 Then
    For i = 0 To UBound(acc.subFields)
      Dim sf As ComplexField
      Set sf = acc.subFields(i)
      
      initial.addLowestFields lf:=sf
    Next
  End If
End Sub

Public Sub addLowestFields(lf As ComplexField)
  Dim pf As PrimitiveField
  Set pf = PrimitiveFieldFactory.CreatePrimitiveField(addr:=lf.address, path:=lf.path, ws:=ws)
  
  pLowestFields.Add Key:=pf.path, Item:=pf
End Sub

'-------------------------------------------------'
'--------------------SubFields--------------------'
'-------------------------------------------------'
Public Sub buildSubFields(level As Integer, initial As ComplexField)
  If level > 0 Then
    initial.combineSubFields init:=initial
    If initial.numSubFields <> 0 Then
      For i = 0 To UBound(initial.subFields)
        initial.subFields(i).buildSubFields (level - 1), (initial.subFields(i))
      Next
    End If
  End If
End Sub

Public Sub combineSubFields(init As ComplexField)
  subCols = combineSubCols(init.address, ws)
  
  For i = 0 To UBound(subCols)
    Dim subCol As ComplexField
    Set subCol = subCols(i)
    init.addChild subCol
  Next i
End Sub
'-------------------------------------------------'


'-------------------------------------------------'
'---------------------Cells-----------------------'
'-------------------------------------------------'

Public Function setValByPath(path As String, row As Long, v As String)
  Dim letter As String
  letter = Split(Cells(1, pCell.Column).address, "$")(1)
  
  Dim subfield As PrimitiveField
  Set subfield = pLowestFields.Item(path)
  subfield.setValueAt newVal:=v, row:=row
End Function

Public Function getCellByPath(path As String, row As Long) As Range
  Dim letter As String
  letter = Split(Cells(1, pCell.Column).address, "$")(1)
  
  Dim subfield As PrimitiveField
  Set subfield = pLowestFields.Item(path)
  Set getCellByPath = subfield.getCellAt(row:=row)
End Function

Public Function getValueByPath(path As String, row As Long) As String
  Dim letter As String
  letter = Split(Cells(1, pCell.Column).address, "$")(1)
  
  Dim subfield As PrimitiveField
  Set subfield = pLowestFields.Item(path)
  getValueByPath = subfield.getValueAt(row:=row)
End Function

'-------------------------------------------------'
'--------------------Getters----------------------'
'-------------------------------------------------'
Public Property Get subFields() As Variant
    subFields = pSubFields
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Get address() As String
    address = pCell.address
End Property

Public Property Get parent() As ComplexField
  Set parent = pParent
End Property

Public Property Get path() As String
  path = pPath
End Property

Public Property Get numSubFields() As Long
    numSubFields = pNumSubFields
End Property

Public Property Get lowestFields() As Dictionary
  Set lowestFields = pLowestFields
End Property

