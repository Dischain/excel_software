Attribute VB_Name = "PrimitiveRowFactory"
Public Function CreatePrimitiveRow(sheet As Worksheet, cell As String, Optional sgns As Variant) As PrimitiveRow
  Dim pr_obj As PrimitiveRow
  Set pr_obj = New PrimitiveRow
  
  pr_obj.init sheet:=sheet, c:=cell, signs:=sgns
  
  Set CreatePrimitiveRow = pr_obj
End Function
