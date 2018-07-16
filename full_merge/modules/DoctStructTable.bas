Attribute VB_Name = "doctStructTable"
Public Function buildDocStruct(outRows As String, executiveColor As Long, supervColor As String, outWS As Worksheet) As Dictionary
  Dim docstructTable As New Dictionary
  
  For Each row In outWS.Range(outRows)
    If row.Interior.Color = executiveColor Then
      executiveAddr = row.address
      executiveName = Trim(eraseTrailingPeriod(eraseSPs(eraseEOLs(s:=row.value))))
      executiveLastAddr = getLastRowByColors((executiveAddr), ws:=outWS, sep1:=(executiveColor), sep2:=(supervColor))
  
      supervAddr = getParentByColor((supervColor), (executiveAddr), outWS)
      supervName = mapSupervs(Trim(eraseTrailingPeriod(eraseSPs(eraseEOLs(s:=outWS.Range(supervAddr).value)))))
      If Not docstructTable.Exists(supervName) Then
        Dim supervDict As Dictionary
        Set supervDict = New Dictionary
        Dim execsDict As Dictionary
        Set execsDict = New Dictionary
        supervDict.Add Key:="address", Item:=supervAddr
        supervDict.Add Key:="execs", Item:=execsDict
        docstructTable.Add Key:=supervName, Item:=supervDict
      End If

      Dim supervs, executives, execsObj As Dictionary
      Set supervs = New Dictionary
      Set executives = New Dictionary
      Set execsObj = New Dictionary
      
      Set supervs = docstructTable.Item(Key:=supervName)
      Set executives = supervs.Item(Key:="execs")
      execsObj.Add Key:="address", Item:=executiveAddr
      execsObj.Add Key:="lastRowAddr", Item:=executiveLastAddr
     
      executives.Add Key:=executiveName, Item:=execsObj ' <-----
      supervs.Remove ("execs")
      supervs.Add Key:="execs", Item:=executives
      docstructTable.Remove (supervName)
      docstructTable.Add Key:=supervName, Item:=supervs
    End If
  Next
  Set buildDocStruct = docstructTable
End Function

Public Function mapSupervs(s As String) As String
  If s = "Дирекция по строительству объектов инженерной инфраструктуры" Then
    mapSupervs = "ИНЖ"
  ElseIf s = "Дирекция по строительству объектов социальной сферы" Then
    mapSupervs = "СОЦ"
  ElseIf s = "Дирекция по строительству объектов транспортной инфраструктуры" Then
    mapSupervs = "ТР"
  Else
    Debug.Print ("Unsupported supervisor name: " & s)
    mapSupervs = ""
  End If
End Function

' Не сработает, если у разных дирекций есть одинаковые отв. исполнители.
Public Function buildDoctStructWithoutSupervs(dict As Dictionary) As Dictionary
  Dim result As Dictionary
  Set result = New Dictionary
  For Each superv In dict.Items
    For Each exec In superv.Item("execs").Keys ' имя каждого отв. исполнителя по дирекции
      Dim execData As Dictionary
      Set execData = superv.Item("execs").Item(exec) ' берем данные данного исполнителя
      Debug.Print (exec)
      result.Add Key:=exec, Item:=execData
    Next
  Next
End Function

Public Sub printDocStruct(docstructTable As Dictionary)
  For Each sprv In docstructTable.Keys
    Dim sprvData As Dictionary
    Set sprvData = docstructTable.Item(sprv)
  
    Debug.Print ("Дирекция: " & sprvData.Item("address") & " " & sprv)
    For Each exc In sprvData.Item("execs").Keys
      Dim excData As Dictionary
      Set excData = sprvData.Item("execs").Item(exc)
  
      Debug.Print ("Отв. исполнитель: " & " " & excData.Item("address") & " " & excData.Item("lastRowAddr") & " " & exc)
    Next
  Next
End Sub



