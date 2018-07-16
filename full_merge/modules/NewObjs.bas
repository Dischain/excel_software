Attribute VB_Name = "NewObjs"
Public Function newObjsDict(newRows As Dictionary, executiveColor As Long, inWS As Worksheet) As Dictionary
  Dim newRowsTable As New Dictionary
  
  For Each newRow In newRows.Items
    Dim executiveAddr, executiveName As String
    executiveAddr = getParentByColor((executiveColor), newRow.address, inWS)
    executiveName = Trim(eraseTrailingPeriod(eraseSPs(eraseEOLs(inWS.Range(executiveAddr).value))))
    
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

Public Function selectRedactedObjsByColor(startRow As Integer, endRow As Integer, ws As Worksheet, dataWS As Worksheet) As Dictionary
  Dim result As New Dictionary
  currentExec = ""
  For Each cell In ws.Range("B" & startRow & ":" & "B" & endRow)
    If cell.Interior.Color = 65535 And cell.value <> "" Then
      Dim execName, execAddr As String
      
      Dim execData As Dictionary
      Set execData = New Dictionary
      Dim execObjs As Dictionary
      Set execObjs = New Dictionary
      
      execName = cell.Offset(0, 1).value
      execAddr = cell.value
      execData.Add Key:="address", Item:=execAddr
      execData.Add Key:="objs", Item:=execObjs
      'Debug.Print (execData.Item("address"))
      result.Add Key:=execName, Item:=execData
      currentExec = execName
    ElseIf cell.Interior.Color = 255 Then
    Else
      Dim tempExecs As New Dictionary
      Dim tempObjs As New Dictionary
      Set tempExecs = result.Item(currentExec)
      Set tempObjs = tempExecs.Item("objs")

      Dim obj As PrimitiveRow
      Set obj = PrimitiveRowFactory.CreatePrimitiveRow(dataWS, cell.value)
      obj.setSuperv (cell.Offset(0, -1).value)
      tempObjs.Add Key:=obj.name, Item:=obj
      tempExecs.Remove Key:="objs"
      tempExecs.Add Key:="objs", Item:=tempObjs
      result.Remove Key:=currentExec
      result.Add Key:=currentExec, Item:=tempExecs
    End If
  Next

  Set selectRedactedObjsByColor = result
End Function

Public Sub printRedacted(redacted As Dictionary)
  For Each k In redacted.Keys
    Dim itemData As Dictionary
    Set itemData = redacted.Item(Key:=k)
    Debug.Print (itemData.Item("address") & " Отв. исп: " & k)
    For Each itm In itemData.Item("objs").Items
      Debug.Print (itm.superv & " " & itm.address & " " & itm.name)
    Next
  Next
End Sub

Public Sub printNewObjs(newObjs As Dictionary)
  For Each ename In newObjs.Keys
    Debug.Print ("ename " & ename)
    Dim edata As Dictionary
    Set edata = newObjs.Item(ename)
  
    Debug.Print ("Ответственный исполнитель: " & edata.Item("address") & " " & ename)
  
    For Each eobj In edata.Item("objs").Items
      Debug.Print ("--- " & eobj.address & " " & eobj.name)
    Next
  Next
End Sub

Public Function printNewObjsWS(newObjs As Dictionary, ws As Worksheet, i As Integer, startColumn As String, startRow As String) As Variant
  ws.Range(startColumn & startRow).value = "Добавлено"
  For Each ename In newObjs.Keys
    r = startRow + i
    Dim edata As Dictionary
    Set edata = newObjs.Item(ename)
    
    ws.Range(startColumn & r).value = edata.Item("address")
    ws.Range(startColumn & r).Interior.Color = 65535
    ws.Range(startColumn & r).Offset(0, 1).value = ename
    ws.Range(startColumn & r).Offset(0, 1).Interior.Color = 65535
    i = i + 1
    For Each eobj In edata.Item("objs").Items
      r = startRow + i
      ws.Range(startColumn & r).value = eobj.address
      ws.Range(startColumn & r).Offset(0, 1).value = eobj.name
      i = i + 1
    Next
  Next
  printNewObjsWS = i
End Function

Public Function printDeletedObjsWS(newObjs As Dictionary, ws As Worksheet, i As Integer, startColumn As String, startRow As String) As Variant
  ws.Range(startColumn & startRow).value = "Удалено"
  For Each ename In newObjs.Keys
    r = startRow + i
    Dim edata As Dictionary
    Set edata = newObjs.Item(ename)
    
    ws.Range(startColumn & r).value = edata.Item("address")
    ws.Range(startColumn & r).Interior.Color = 65535
    ws.Range(startColumn & r).Offset(0, 1).value = ename
    ws.Range(startColumn & r).Offset(0, 1).Interior.Color = 65535
    i = i + 1
    For Each eobj In edata.Item("objs").Items
      r = startRow + i
      ws.Range(startColumn & r).value = eobj.address
      ws.Range(startColumn & r).Offset(0, 1).value = eobj.name
      i = i + 1
    Next
  Next
  printDeletedObjsWS = i
End Function
