Attribute VB_Name = "Main"
Public Sub main()
  ' ������ ������ ������������ �� ������� �� ����������
  ' ����������� ������� ��������� (20 �����)
  Dim activeWS As Worksheet
  Set activeWS = ActiveWorkbook.ActiveSheet
  
  For Each c In activeWS.Range("B17:C307")
    c.Value = ""
  Next

  ' �������� ������, ���������� ��������������
  inputFilePath = activeWS.Range("C3").Value
  Dim inSrc As Workbook
  Set inSrc = Workbooks.Open(inputFilePath, True, True)
  inSheet = activeWS.Range("C4").Value
  Dim inWS As Worksheet
  Set inWS = inSrc.Worksheets(inSheet)
  
  ' ����, � ������� ����� ������������ ������� �� inputFile
  outputFilePath = activeWS.Range("E3").Value
  Dim outSrc As Workbook
  Set outSrc = Workbooks.Open(outputFilePath, True, True)
  outSheet = activeWS.Range("E4").Value
  Dim outWS As Worksheet
  Set outWS = outSrc.Worksheets(outSheet)
  
  ' ������, ���������� �����������
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
  'Debug.Print ("����� �������: ")
  'printNewObjs newObjs:=newRowsTable
  'Debug.Print ("��������� �������: ")
  'printNewObjs newObjs:=deletedRowsTable
  i = printNewObjsWS(newObjs:=newRowsTable, ws:=activeWS, i:=(i), startColumn:="B", startRow:="16")
  i = printDeletedObjsWS(newObjs:=deletedRowsTable, ws:=activeWS, i:=(1), startColumn:="B", startRow:=i + 17)
  '------------------------------------------------'
  Dim docStructTable As New Dictionary
  Set docStructTable = buildDocStruct(outRows, (executiveColor), (supervColor), outWS)
  'Debug.Print ("��������� ���������: ")
  'printDocStruct docStructTable:=docStructTable
  '------------------------------------------------'
  MsgBox ("������!")
End Sub
