VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrimitiveField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String
Private pCell As Range
Private pPath As String
Private ws As Worksheet

Public Sub init(address As String, path As String, Worksheet As Worksheet)
  Set ws = Worksheet
  Set pCell = ws.Range(address)
  
  pName = pCell.value
  pPath = path
End Sub

Public Sub setValueAt(newVal As String, row As Long)
  Dim letter As String
  letter = Split(ws.Cells(1, pCell.Column).address, "$")(1)
  ws.Range(letter & row).value = newVal
End Sub

Public Function getCellAt(row As Long) As Range
  Dim letter As String
  letter = Split(ws.Cells(1, pCell.Column).address, "$")(1)
  
  Set getCellAt = ws.Range(letter & row)
End Function

Public Function getValueAt(row As Long) As String
  Dim letter As String
  letter = Split(Cells(1, pCell.Column).address, "$")(1)
  
  getValueAt = ws.Range(letter & row).value
End Function

Public Property Get name() As String
  name = pName
End Property

Public Property Get cell() As Range
  Set cell = pCell
End Property

Public Property Get address() As String
  address = pCell.address
End Property

Public Property Get path() As String
  path = pPath
End Property
