Attribute VB_Name = "Utils"
Public Function combineSubCols(addr As String, ws As Worksheet) As Variant
  Dim parentColl, currentSub, parentSibling As Range
  Dim length As Integer
  Dim combined() As ComplexField
  
  Set parentColl = ws.Range(addr)
  Set parentSibling = parentColl.Offset(0, 1)
  Set currentSub = parentColl.Offset(1, 0)
  
  length = 0
  
  Do While currentSub.Column <> parentSibling.Column
    Dim c As ComplexField
    Set c = ComplexFieldFactory.CreateComplexField(ws, currentSub.address)
    
    ReDim Preserve combined(length)
    Set combined(length) = c
    length = length + 1
    Set currentSub = currentSub.Offset(0, 1)
  Loop
  
  combineSubCols = combined
End Function

Public Function createFieldMap(addr As String, complLevel As Integer, ws As Worksheet) As Dictionary
  Dim fieldMap As Dictionary
  Set fieldMap = New Dictionary
  
  For Each f In ws.Range(addr)
    If f.value <> "" Then
      Dim field As ComplexField
      Set field = ComplexFieldFactory.CreateComplexField(ws, (f.address), l:=complLevel)
      
      fieldMap.Add Key:=field.name, Item:=field
    End If
  Next
  
  Set createFieldMap = fieldMap
End Function

Public Function createRowMap(addr As String, ws As Worksheet, Optional signs As Variant) As Dictionary
  Dim rowMap As Dictionary
  Set rowMap = New Dictionary
  
  For Each r In ws.Range(addr)
    If r.value <> "" And Not containsEscapeWords(r.value) Then
      Dim row As PrimitiveRow
      Set row = PrimitiveRowFactory.CreatePrimitiveRow(ws, (r.address), signs)
      
      rowMap.Add Key:=row.name, Item:=row
    End If
  Next
  
  Set createRowMap = rowMap
End Function

' Выполняет слияние множества строк с предварительной проверкой совпадения по именам строк, без учета доп. признаков
Public Function mergeRows(inRowMap As Dictionary, outRowMap As Dictionary, inFieldMap As Dictionary, outFieldMap As Dictionary) As Dictionary
  Dim notMatched As Dictionary
  Set notMatched = New Dictionary

  For Each inRow In inRowMap.Items
    If outRowMap.Exists(inRow.name) Then
      Dim outRow As PrimitiveRow
      Set outRow = outRowMap.Item(inRow.name)
      mergeSingleRow inRow:=inRow.row, outRow:=outRow.row, inFieldMap:=inFieldMap, outFieldMap:=outFieldMap
    Else
      notMatched.Add Key:=inRow.name, Item:=inRow
    End If
  Next

  Set mergeRows = notMatched
End Function

' Выполняет слияние множества строк с предварительной проверкой совпадения по именам строк, с учетом доп. признаков
Public Function mergeRowsWithSigns(inRowMap As Dictionary, outRowMap As Dictionary, inFieldMap As Dictionary, outFieldMap As Dictionary) As Dictionary
  Dim notMatched As Dictionary
  Set notMatched = New Dictionary

  For Each inRow In inRowMap.Items
    If outRowMap.Exists(inRow.name) Then
      Dim outRow As PrimitiveRow
      Set outRow = outRowMap.Item(inRow.name)
      
      Dim inRowSigns As New Dictionary
      Dim outRowSigns As New Dictionary
      Set inRowSigns = inRow.signs
      Set outRowSigns = outRow.signs
      
      If dictEquals(inRowSigns, outRowSigns) Then
        Debug.Print ("eq")
        mergeSingleRow inRow:=inRow.row, outRow:=outRow.row, inFieldMap:=inFieldMap, outFieldMap:=outFieldMap
      Else
        Debug.Print ("not eq")
        notMatched.Add Key:=inRow.name, Item:=inRow
      End If
    Else
      notMatched.Add Key:=inRow.name, Item:=inRow
    End If
  Next

  Set mergeRowsWithSigns = notMatched
End Function

' Выполняет слияние двух строк путем проверки всех полей
Public Sub mergeSingleRow(inRow As Long, outRow As Long, inFieldMap As Dictionary, outFieldMap As Dictionary)
  For Each inField In inFieldMap.Items
    If outFieldMap.Exists(Key:=inField.name) Then
      Dim outField As ComplexField
      Set outField = outFieldMap.Item(inField.name)
      
      Dim inFieldLowestFields As New Dictionary
      Set inFieldLowestFields = inField.lowestFields
      Dim outFieldLowestFields As New Dictionary
      Set outFieldLowestFields = outField.lowestFields
      
      For Each inLF In inFieldLowestFields.Items
        If outFieldLowestFields.Exists(Key:=inLF.path) Then
          Dim outLF As PrimitiveField
          Set outLF = outFieldLowestFields.Item(Key:=inLF.path)
          
          inVal = inLF.getValueAt(inRow)
          outLF.setValueAt (inVal), (outRow)
        End If
      Next
    End If
  Next
End Sub

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
    supervName = ws.Range(supervAddr).value
    execAddr = getParentByColor((execColor), row.address, ws)
    execName = ws.Range(execAddr).value
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

