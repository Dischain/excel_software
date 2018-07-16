Attribute VB_Name = "ComplexFieldFactory"
Public Function CreateComplexField(sheet As Worksheet, cell As String, Optional l As Integer, Optional parent As ComplexField) As ComplexField
  Dim cf_obj As ComplexField
  Set cf_obj = New ComplexField
  
  cf_obj.init sheet:=sheet, c:=cell, level:=l, parent:=parent
  
  Set CreateComplexField = cf_obj
End Function
