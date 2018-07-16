Attribute VB_Name = "Main"
' Step 1. Определение diff, границы добавленных и удаленных объектов.
Public Sub diff()
  Dim activeWS As Worksheet
  Set activeWS = ActiveWorkbook.ActiveSheet
  
  ' Чистим список несовпадений по строкам из результата
  ' предыдущего запуска программы (100 строк)
  For Each c In activeWS.Range("A17:C107")
    c.value = ""
    c.Interior.Color = 16777215
  Next
  
  ' ----------------------------------------------------'
  ' Выборка исходных данных программы
  ' ----------------------------------------------------'
  
  ' Источник данных, подлежащий форматированию
  inputFilePath = activeWS.Range("C3").value
  Dim inSrc As Workbook
  Set inSrc = Workbooks.Open(inputFilePath, True, True)
  inSheet = activeWS.Range("C4").value
  Dim inWS As Worksheet
  Set inWS = inSrc.Worksheets(inSheet)
  
  ' Файл, в который будет осуществлена заливка из inputFile
  outputFilePath = activeWS.Range("E3").value
  Dim outSrc As Workbook
  Set outSrc = Workbooks.Open(outputFilePath, True, True)
  outSheet = activeWS.Range("E4").value
  Dim outWS As Worksheet
  Set outWS = outSrc.Worksheets(outSheet)

  ' Строки, подлежащие объединению
  Dim inRows, outRows As String
  inRows = activeWS.Range("C7").value
  outRows = activeWS.Range("E7").value
    
  '------------------------------------------------'
  '                   InHTBuild                    '
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
  
  executiveColor = activeWS.Range("G4").Interior.Color '14277081
  supervColor = activeWS.Range("G3").Interior.Color '39423
  reportStartingRow = 17
  
  Dim addedObjsStartCell, addedObjsEndCell, deletedObjsStartCell, deletedObjsEndCell As Range
  Set addedObjsStartCell = activeWS.Range("P10")
  Set addedObjsEndCell = activeWS.Range("P11")
  Set deletedObjsStartCell = activeWS.Range("P12")
  Set deletedObjsEndCell = activeWS.Range("P13")
  '------------------------------------------------'
  
  i = 1
  addedObjsStart = reportStartingRow + 1 ' <--
  addedObjsStartCell.value = addedObjsStart
  
  Dim newRowsTable As New Dictionary
  Set newRowsTable = newObjsDict(newRows, (executiveColor), inWS)
  Dim deletedRowsTable As New Dictionary
  Set deletedRowsTable = newObjsDict(deletedRows, (executiveColor), outWS)
  i = printNewObjsWS(newObjs:=newRowsTable, ws:=activeWS, i:=(i), startColumn:="B", startRow:=(reportStartingRow))
  addedObjsEnd = addedObjsStart + i - 2 ' <--
  addedObjsEndCell.value = addedObjsEnd
  
  deletedObjsStart = addedObjsEnd + 3 ' <--
  deletedObjsStartCell.value = deletedObjsStart
  
  i = printDeletedObjsWS(newObjs:=deletedRowsTable, ws:=activeWS, i:=(1), startColumn:="B", startRow:=i + reportStartingRow + 1)
  deletedObjsEnd = deletedObjsStart + i - 2 ' <--
  deletedObjsEndCell.value = deletedObjsEnd

  '------------------------------------------------'
  Dim docstructTable As New Dictionary
  Set docstructTable = buildDocStruct(outRows, (executiveColor), (supervColor), outWS)
  printDocStruct docstructTable:=docstructTable

  MsgBox ("Сравнение выполнено")
End Sub

' Step 2. Добавление новых строк согласно списку отредактированных объектов.
Public Sub merge()
  Dim activeWS As Worksheet
  Set activeWS = ActiveWorkbook.ActiveSheet
  
  ' ----------------------------------------------------'
  ' Выборка исходных данных программы                   '
  ' ----------------------------------------------------'
  
  executiveColor = activeWS.Range("G4").Interior.Color '14277081
  supervColor = activeWS.Range("G3").Interior.Color '39423
  reportStartingRow = 17
  
  Dim addedObjsStartCell, addedObjsEndCell, deletedObjsStartCell, deletedObjsEndCell As Range
  Set addedObjsStartCell = activeWS.Range("P10")
  Set addedObjsEndCell = activeWS.Range("P11")
  Set deletedObjsStartCell = activeWS.Range("P12")
  Set deletedObjsEndCell = activeWS.Range("P13")
  
  ' Источник данных, подлежащий форматированию
  inputFilePath = activeWS.Range("C3").value
  Dim inSrc As Workbook
  Set inSrc = Workbooks.Open(inputFilePath, True, True)
  inSheet = activeWS.Range("C4").value
  Dim inWS As Worksheet
  Set inWS = inSrc.Worksheets(inSheet)
  
  ' Файл, в который будет осуществлена заливка из inputFile
  outputFilePath = activeWS.Range("E3").value
  Dim outSrc As Workbook
  Set outSrc = Workbooks.Open(outputFilePath, True, True)
  outSheet = activeWS.Range("E4").value
  Dim outWS As Worksheet
  Set outWS = outSrc.Worksheets(outSheet)

  ' Строки, подлежащие объединению
  Dim inRows, outRows As String
  inRows = activeWS.Range("C7").value
  outRows = activeWS.Range("E7").value
  
  ' Дополнительные признаки
  inSignsStr = activeWS.Range("C8").value
  outSignsStr = activeWS.Range("E8").value
  
  ' Одинаковые по смыслу колонки из источника, заливаемого файла
  ' и кол. строк под ними.
  inFields1 = activeWS.Range("C5").value
  outFields1 = activeWS.Range("E5").value
  subFields1 = activeWS.Range("C6").value
    
  Dim inRowMap As New Dictionary
  Dim outRowMap As New Dictionary
  Set inRowMap = createRowMap((inRows), ws:=inWS)
  Set outRowMap = createRowMap((outRows), ws:=outWS)
  
  Dim inFieldMap As New Dictionary
  Dim outFieldMap As New Dictionary
  Set inFieldMap = createFieldMap((inFields1), complLevel:=(subFields1), ws:=inWS)
  Set outFieldMap = createFieldMap((outFields1), complLevel:=(subFields1), ws:=outWS)
  
  Dim redactedDict As New Dictionary
  Set redactedDict = selectRedactedObjsByColor(addedObjsStartCell.value, addedObjsEndCell.value, activeWS, inWS)
  'printRedacted redacted:=redactedDict
  
  '------------------------------------------------'
  Dim docstructTable As New Dictionary
  Set docstructTable = buildDocStruct(outRows, (executiveColor), (supervColor), outWS)

  'printDocStruct docstructTable:=docstructTable
  
  addNewObjs docstructTable:=docstructTable, _
              redacted:=redactedDict, _
              insertionCellAddr:="C", _
              appendRange:="A:CR", _
              outWS:=outWS
              
  ' использовать для определения одинаковых исполнителей
  ' и последующего их блокирования
  'buildDoctStructWithoutSupervs dict:=docStructTable <-- не должно раб. при дубл. executors
  
  
  '------------------------------------------------'
  Dim inRM As New Dictionary
  Dim outRM As New Dictionary
  If inSignsStr <> "" And outSignsStr <> "" Then
    inSigns = Split(inSignsStr, " ")
    outSigns = Split(outSignsStr, " ")
    Set inRM = createRowMap((inRows), ws:=inWS, signs:=inSigns)
    Set outRM = createRowMap((outRows), ws:=outWS, signs:=outSigns)
    Set unmatched = mergeRowsWithSigns(inRM, outRM, inFieldMap, outFieldMap)
  Else
    Set inRM = createRowMap((inRows), ws:=inWS)
    Set outRM = createRowMap((outRows), ws:=outWS)
    Set unmatched = mergeRows(inRM, outRM, inFieldMap, outFieldMap)
  End If
  
  MsgBox ("Слияние выполнено")
End Sub
