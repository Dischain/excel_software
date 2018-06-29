VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PrimitiveRow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String
Private pCell As Range
Private ws As Worksheet
Private pSigns As Dictionary

Public Sub init(sheet As Worksheet, c As String, Optional signs As Variant)
  Set ws = sheet
  Set pCell = ws.Range(c)
  
  If Not IsMissing(signs) Then
    t = signs
    Set pSigns = createSignsDict(t)
  End If
  
  pName = Trim(eraseTrailingPeriod(eraseSPs(eraseEOLs(s:=pCell.Value))))
End Sub

Public Property Get signs() As Dictionary
  Set signs = pSigns
End Property

Public Property Get row() As Long
  row = pCell.row
End Property

Public Property Get name() As String
  name = pName
End Property

Public Property Get address() As String
  address = pCell.address
End Property

' Строит хэш-таблицу из признаков, где ключ - имя признака,
' значение - значение признака в данной строке
' param sgns Array[String] - массив адресов колонок с признаками
Private Function createSignsDict(sgns As Variant) As Dictionary
  Dim dict As New Dictionary
  For Each s In sgns
    letter = Split(ws.Cells(1, Range(s).Column).address, "$")(1)
    r = Me.row()
    
    k = ws.Range(s).Value
    v = ws.Range(letter & r).Value
    dict.Add Key:=k, Item:=v
  Next
  Set createSignsDict = dict
End Function
