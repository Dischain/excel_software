Attribute VB_Name = "doctStructTable"
Public Function buildDocStruct(outRows As String, executiveColor As Long, supervColor As String, outWS As Worksheet) As Dictionary
  Dim docStructTable As New Dictionary
  
  For Each row In outWS.Range(outRows)
    acc = acc & " " & row.address
    
    If row.Interior.Color = executiveColor Then
      executiveAddr = row.address
      executiveName = Trim(eraseTrailingPeriod(eraseSPs(eraseEOLs(s:=row.Value))))
      executiveLastAddr = getLastRowByColors((executiveAddr), ws:=outWS, sep1:=(executiveColor), sep2:=(supervColor))
  
      supervAddr = getParentByColor((supervColor), (executiveAddr), outWS)
      supervName = Trim(eraseTrailingPeriod(eraseSPs(eraseEOLs(s:=outWS.Range(supervAddr).Value))))
      If Not docStructTable.Exists(supervName) Then
        Dim supervDict As Dictionary
        Set supervDict = New Dictionary
        Dim execsDict As Dictionary
        Set execsDict = New Dictionary
        supervDict.Add Key:="address", Item:=supervAddr
        supervDict.Add Key:="execs", Item:=execsDict
        docStructTable.Add Key:=supervName, Item:=supervDict
      End If

      Dim supervs, executives, execsObj As Dictionary
      Set supervs = New Dictionary
      Set executives = New Dictionary
      Set execsObj = New Dictionary
      
      Set supervs = docStructTable.Item(Key:=supervName)
      Set executives = supervs.Item(Key:="execs")
      execsObj.Add Key:="address", Item:=executiveAddr
      execsObj.Add Key:="lastRowAddr", Item:=executiveLastAddr
     
      executives.Add Key:=executiveName, Item:=execsObj ' <-----
      supervs.Remove ("execs")
      supervs.Add Key:="execs", Item:=executives
      docStructTable.Remove (supervName)
      docStructTable.Add Key:=supervName, Item:=supervs
    End If
  Next
  Set buildDocStruct = docStructTable
End Function

Public Sub printDocStruct(docStructTable As Dictionary)
  For Each sprv In docStructTable.Keys
    Dim sprvData As Dictionary
    Set sprvData = docStructTable.Item(sprv)
  
    Debug.Print ("Дирекция: " & sprvData.Item("address") & " " & sprv)
    For Each exc In sprvData.Item("execs").Keys
      Dim excData As Dictionary
      Set excData = sprvData.Item("execs").Item(exc)
  
      Debug.Print ("Отв. исполнитель: " & " " & excData.Item("address") & " " & excData.Item("lastRowAddr") & " " & exc)
    Next
  Next
End Sub

