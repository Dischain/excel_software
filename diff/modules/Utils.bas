Attribute VB_Name = "Utils"
Public Function createRowMap(addr As String, ws As Worksheet, Optional signs As Variant) As Dictionary
  Dim rowMap As Dictionary
  Set rowMap = New Dictionary
  
  For Each r In ws.Range(addr)
    If r.Value <> "" And Not containsEscapeWords(r.Value) Then
      Dim row As PrimitiveRow
      Set row = PrimitiveRowFactory.CreatePrimitiveRow(ws, (r.address), signs)
      
      rowMap.Add Key:=row.name, Item:=row
    End If
  Next
  
  Set createRowMap = rowMap
End Function

Public Function escapeWords() As Variant
  escapeWords = Array("Министерство", "Дирекция", "Объекты", "Модернизация", "Служба", "Государственный комитет", "Управление")
End Function

Public Function diffRows(inRowMap As Dictionary, outRowMap As Dictionary) As Dictionary
  Dim newRows As New Dictionary
  Dim deletedRows As New Dictionary
  Dim result As New Dictionary

  For Each inr In inRowMap.Items
    If Not outRowMap.Exists(inr.name) Then
      newRows.Add Key:=inr.name, Item:=inr
    End If
  Next

  For Each outr In outRowMap.Items
    If Not inRowMap.Exists(outr.name) Then
      deletedRows.Add Key:=outr.name, Item:=outr
    End If
  Next

  result.Add Key:="new", Item:=newRows
  result.Add Key:="deleted", Item:=deletedRows
  
  Set diffRows = result
End Function

Public Function buildOutRowsHTWithParents(rows As Dictionary, ws As Worksheet, execColor As Long, supervColor As Long) As Dictionary
  Dim result As New Dictionary
  For Each row In rows.Items
    supervAddr = getParentByColor((supervColor), row.address, ws)
    supervName = ws.Range(supervAddr).Value
    execAddr = getParentByColor((execColor), row.address, ws)
    execName = ws.Range(execAddr).Value
    execLastRow = getLastRowByColors(execAddr, ws, execColor, supervColor)
  Next
End Function

Public Function getLastRowByColors(addr As String, ws As Worksheet, sep1 As Long, sep2 As Long) As String
  Dim tempCell As Range
  Set tempCell = ws.Range(addr).Offset(1)
  printArea = ws.PageSetup.printArea
    
  While tempCell.Interior.Color <> sep1 And tempCell.Interior.Color <> sep2 And isRowAtRange(tempCell.row, (printArea))
    Set tempCell = tempCell.Offset(1)
  Wend
  
  getLastRowByColors = tempCell.Offset(-1).address
End Function

Public Function getParentByColor(parentColor As Long, addr As String, ws As Worksheet) As String
  Dim tempCell As Range
  Set tempCell = ws.Range(addr)
  
  While tempCell.Interior.Color <> parentColor
    Set tempCell = tempCell.Offset(-1)
  Wend
  
  getParentByColor = tempCell.address
End Function

