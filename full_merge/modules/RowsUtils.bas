Attribute VB_Name = "RowsUtils"
' Возвращает номер вставленной строки
' r - номе строки, после которой будет выполнена вставка
' value - значение, которое будет вставленно в valueColl
' valueColl - колонка, в которую будет вставлено value
' collsRange - диапозон колонок, который будет скопирован (должен быть полный
' диапазон колонок документа). Пример: "A:Z"
Public Sub appendRow(r As Long, value As String, valueColl As String, collsRange As String, ws As Worksheet)
  ws.Range(valueColl & r).Offset(1).EntireRow.Insert shift:=xlDown
  splittedAddr = Split(collsRange, ":")
  first = splittedAddr(0)
  last = splittedAddr(1)
  copyRowWithFormulas first:=first & r, last:=last & r, ws:=ws
  ws.Range(valueColl & r + 1).value = value
  Debug.Print ("row appended")
End Sub

Sub copyRowWithFormulas(first As String, last As String, ws As Worksheet)
  Dim strPattern As String: strPattern = ".*[a-z]+($)?[0-9]+.*"
  Dim regEx As New RegExp
  
  regEx.Pattern = strPattern
  regEx.IgnoreCase = True
  regEx.Global = True
  
  With ws.Range(first, last)
  .Offset(1).Insert
  .Copy
  .Offset(1).PasteSpecial xlPasteFormulas
  .Offset(1).PasteSpecial xlPasteFormats
  
  Application.CutCopyMode = False
  End With
  
  ' Удалить все ячейки, не содержащие ссылок на другие ячейки
  For Each c In ws.Range(first, last).Offset(1, 0)
    
    If Not regEx.Test(c.Formula) Then
      c.value = ""
    End If
  Next
End Sub

Public Sub addNewObjs(docstructTable As Dictionary, _
               redacted As Dictionary, _
               insertionCellAddr As String, _
               appendRange As String, _
               outWS As Worksheet)

  Count = 0
  For Each inExecName In redacted.Keys
    Dim curInExec As Dictionary
    Set curInExec = redacted.Item(inExecName)

    For Each inObjName In curInExec.Item("objs").Keys
      Dim curInExecObjs As Dictionary
      Set curInExecObjs = curInExec.Item("objs")
      Dim curInSupervName As String
      Dim curRow As PrimitiveRow

      Set curRow = curInExecObjs.Item(Key:=inObjName)
      curInSupervName = curRow.superv
      
      If docstructTable.Exists(curInSupervName) Then
        Dim outSuperv As Dictionary
        Set outSuperv = docstructTable(curInSupervName)
        Dim outSupervExecs As Dictionary
        Set outSupervExecs = outSuperv.Item("execs")
        
        Debug.Print (inExecName)
        Debug.Print ("outSupervExecs.Exists(inExecName) " & outSupervExecs.Exists(inExecName))
        If outSupervExecs.Exists(inExecName) Then
          Dim outExec As Dictionary
          Set outExec = outSupervExecs.Item(inExecName)
          rowAddr = outWS.Range(outExec.Item("lastRowAddr")).Offset(Count, 0).row
          Debug.Print (rowAddr)
          appendRow r:=(rowAddr), _
                    value:=(inObjName), _
                    valueColl:=insertionCellAddr, _
                    collsRange:=appendRange, ws:=outWS
          Count = Count + 1
        End If
      End If
    Next
  Next
End Sub

