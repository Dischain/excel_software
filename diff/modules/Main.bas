Attribute VB_Name = "Main"
Public Sub main()
  ' Чистим список несовпадений по строкам из результата
  ' предыдущего запуска программы (20 строк)
  Dim activeWS As Worksheet
  Set activeWS = ActiveWorkbook.ActiveSheet
  
  For Each c In activeWS.Range("B17:C307")
    c.Value = ""
  Next

  ' Источник данных, подлежащий форматированию
  inputFilePath = activeWS.Range("C3").Value
  Dim inSrc As Workbook
  Set inSrc = Workbooks.Open(inputFilePath, True, True)
  inSheet = activeWS.Range("C4").Value
  Dim inWS As Worksheet
  Set inWS = inSrc.Worksheets(inSheet)
  
  ' Файл, в который будет осуществлена заливка из inputFile
  outputFilePath = activeWS.Range("E3").Value
  Dim outSrc As Workbook
  Set outSrc = Workbooks.Open(outputFilePath, True, True)
  outSheet = activeWS.Range("E4").Value
  Dim outWS As Worksheet
  Set outWS = outSrc.Worksheets(outSheet)
  
  ' Строки, подлежащие объединению
  Dim inRows, outRows As String
  inRows = activeWS.Range("C5").Value
  outRows = activeWS.Range("E5").Value

  '------------------------------------------------'
  Dim inRowMap As New Dictionary
  Dim outRowMap As New Dictionary
  Set inRowMap = createRowMap((inRows), ws:=inWS)
  Set outRowMap = createRowMap((outRows), ws:=outWS)
  
  Dim diff As New Dictionary
  Set diff = diffRows(inRowMap, outRowMap)
  Dim newRows As New Dictionary
  Dim deletedRows As New Dictionary
  Set newRows = diff.Item("new")
  Set deletedRows = diff.Item("deleted")
  
  executiveColor = activeWS.Range("B8").Interior.Color
  supervColor = activeWS.Range("B7").Interior.Color
  '------------------------------------------------'
  
  i = 1
  Dim newRowsTable As New Dictionary
  Set newRowsTable = newObjsDict(newRows, (executiveColor), inWS)
  Dim deletedRowsTable As New Dictionary
  Set deletedRowsTable = newObjsDict(deletedRows, (executiveColor), outWS)
  'Debug.Print ("Новые объекты: ")
  'printNewObjs newObjs:=newRowsTable
  'Debug.Print ("Удаленные объекты: ")
  'printNewObjs newObjs:=deletedRowsTable
  i = printNewObjsWS(newObjs:=newRowsTable, ws:=activeWS, i:=(i), startColumn:="B", startRow:="16")
  i = printDeletedObjsWS(newObjs:=deletedRowsTable, ws:=activeWS, i:=(1), startColumn:="B", startRow:=i + 17)
  '------------------------------------------------'
  Dim docStructTable As New Dictionary
  Set docStructTable = buildDocStruct(outRows, (executiveColor), (supervColor), outWS)
  'Debug.Print ("Структура документа: ")
  'printDocStruct docStructTable:=docStructTable
  '------------------------------------------------'
  MsgBox ("Готово!")
End Sub
