Attribute VB_Name = "PrimitiveFieldFactory"
Public Function CreatePrimitiveField(addr As String, path As String, ws As Worksheet)
  Dim pf_obj As PrimitiveField
  Set pf_obj = New PrimitiveField
  
  pf_obj.init address:=addr, path:=path, Worksheet:=ws
  
  Set CreatePrimitiveField = pf_obj
End Function
