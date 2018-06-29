Attribute VB_Name = "NewObjs"
Public Function newObjsDict(newRows As Dictionary, executiveColor As Long, inWS As Worksheet) As Dictionary
  Dim newRowsTable As New Dictionary
  For Each newRow In newRows.Items
    Dim executiveAddr, executiveName As String
    executiveAddr = getParentByColor((executiveColor), newRow.address, inWS)
    executiveName = Trim(eraseTrailingPeriod(eraseSPs(eraseEOLs(inWS.Range(executiveAddr).Value))))
    
    If Not newRowsTable.Exists(executiveName) Then
      Dim executiveDict As Dictionary
      Set executiveDict = New Dictionary
      Dim objectsDict As Dictionary
      Set objectsDict = New Dictionary
     
      executiveDict.Add Key:="address", Item:=executiveAddr
      executiveDict.Add Key:="objs", Item:=objectsDict

      newRowsTable.Add Key:=executiveName, Item:=executiveDict
    End If
  
    Dim execs, objsDict As New Dictionary
    Set execs = newRowsTable.Item(Key:=executiveName)
    Set objsDict = execs.Item(Key:="objs")
    objsDict.Add Key:=newRow.name, Item:=newRow
    execs.Remove ("objs")
    
    execs.Add Key:="objs", Item:=objsDict
    newRowsTable.Remove (executiveName)
    newRowsTable.Add Key:=executiveName, Item:=execs
  Next
  
  Set newObjsDict = newRowsTable
End Function

Public Sub printNewObjs(newObjs As Dictionary)
  For Each ename In newObjs.Keys
    Dim edata As Dictionary
    Set edata = newObjs.Item(ename)
  
    Debug.Print ("Ответственный исполнитель: " & edata.Item("address") & " " & ename)
  
    For Each eobj In edata.Item("objs").Items
      Debug.Print ("--- " & eobj.address & " " & eobj.name)
    Next
  Next
End Sub

Public Function printNewObjsWS(newObjs As Dictionary, ws As Worksheet, i As Integer, startColumn As String, startRow As String) As Variant
  ws.Range(startColumn & startRow).Value = "Добавлено"
  For Each ename In newObjs.Keys
    r = startRow + i
    Dim edata As Dictionary
    Set edata = newObjs.Item(ename)
    
    ws.Range(startColumn & r).Value = "Ответственный исполнитель: " & edata.Item("address") & " " & ename
    i = i + 1
    For Each eobj In edata.Item("objs").Items
      r = startRow + i
      ws.Range(startColumn & r).Value = eobj.address
      ws.Range(startColumn & r).Offset(0, 1).Value = eobj.name
      i = i + 1
    Next
  Next
  printNewObjsWS = i
End Function

Public Function printDeletedObjsWS(newObjs As Dictionary, ws As Worksheet, i As Integer, startColumn As String, startRow As String) As Variant
  ws.Range(startColumn & startRow).Value = "Удалено"
  For Each ename In newObjs.Keys
    r = startRow + i
    Dim edata As Dictionary
    Set edata = newObjs.Item(ename)
    
    ws.Range(startColumn & r).Value = "Ответственный исполнитель: " & edata.Item("address") & " " & ename
    i = i + 1
    For Each eobj In edata.Item("objs").Items
      r = startRow + i
      ws.Range(startColumn & r).Value = eobj.address
      ws.Range(startColumn & r).Offset(0, 1).Value = eobj.name
      i = i + 1
    Next
  Next
  printDeletedObjsWS = i
End Function
