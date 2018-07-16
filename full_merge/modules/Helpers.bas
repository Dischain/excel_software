Attribute VB_Name = "Helpers"
' Выполняет операцию равенства по ключам для двух хэш-таблиц
Public Function dictEquals(dict1 As Dictionary, dict2 As Dictionary) As Boolean
  For Each k In dict1.Keys
    If dict1.Item(k) <> dict2.Item(k) Then
      dictEquals = False
      Exit Function
    End If
  Next
  dictEquals = True
End Function

Public Function containsEscapeWords(name As String) As Boolean
  For Each w In escapeWords
    word = LCase(w)
    Dim str As String
    str = LCase(name)
    If startsWith((word), str) Then
      containsEscapeWords = True
      Exit Function
    End If
  Next
  containsEscapeWords = False
End Function

Public Function startsWith(s As String, seed As String) As Boolean
  If InStr(1, seed, s) = 1 Then
    startsWith = True
  Else
    startsWith = False
  End If
End Function

Public Function arrayToString(arr As Variant) As String
  res = ""
  
  For Each itm In arr
    Debug.Print (itm)
    res = res & itm & Chr(13)
  Next
  arrayToString = res
End Function

Public Function stringToArray(str As String) As Variant
  arr = Split(str, ",")
  Dim res() As String
  For i = 0 To UBound(arr)
    ReDim Preserve res(i + 1)
    s = Trim(arr(i))
    res(i) = s
  Next
  stringToArray = res
End Function

Public Function eraseEOLs(s As String) As String
  tempStr = ""
  For c = 1 To Len(s)
    If Mid(s, c, 1) = vbCr Or Mid(s, c, 1) = vbLf Then
      tempStr = tempStr + ""
    Else
      tempStr = tempStr + Mid(s, c, 1)
    End If
  Next
  eraseEOLs = tempStr
End Function

Public Function eraseSPs(s As String) As String
  tempStr = ""
  For c = 2 To Len(s)
    Dim prev As Integer
    prev = c - 1
    
    If c > 2 And Mid(s, c, 1) = " " And Mid(s, prev, 1) = " " Then
      tempStr = tempStr + ""
    Else
      tempStr = tempStr + Mid(s, c, 1)
    End If
  Next
  eraseSPs = Mid(s, 1, 1) & tempStr
End Function

Public Function eraseTrailingPeriod(s As String) As String
  tempStr = ""
  l = Len(s)
  For c = 1 To l
    If c = l And Mid(s, c, 1) = "." Then
      tempStr = tempStr + ""
    Else
      tempStr = tempStr + Mid(s, c, 1)
    End If
  Next
  eraseTrailingPeriod = tempStr
End Function

Public Function concat(arr1 As Variant, arr2 As Variant) As Variant
  arr1Length = UBound(arr1)
  arr2Length = UBound(arr2)
  For i = 0 To arr2Length
    ReDim Preserve arr1(arr1Length + i + 1)
    Set arr1(arr1Length + i) = arr2(i)
  Next i
  concat = arr1
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
  IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Public Function isRowAtRange(row As Long, area As String) As Boolean
  arr = Split(area, ":")
  If Range(arr(0)).row <= row And Range(arr(1)).row >= row Then
    isRowAtRange = True
  Else
    isRowAtRange = False
  End If
End Function

Public Function containsDiapasone(addr As String) As Boolean
  Dim strPattern As String: strPattern = "^=.*\([a-z]+[0-9]+:[a-z]+[0-9]+\).*$"
  Dim regEx As New RegExp
  
  regEx.Pattern = strPattern
  regEx.IgnoreCase = True
  regEx.Global = True
  
  containsDiapasone = regEx.Test(addr)
End Function

Public Sub updateVertFormula(r As Range, i As Integer)
  Dim digitPattern As String: digitPattern = ".*([0-9]+).*$"
  Dim regExDigit As New RegExp
  
  regExDigit.Pattern = digitPattern
  regExDigit.IgnoreCase = True
  regExDigit.Global = True
  
  Dim columnPattern As String: columnPattern = ".*([a-z]+).*"
  Dim regExColumn As New RegExp
  
  regExColumn.Pattern = columnPattern
  regExColumn.IgnoreCase = True
  regExColumn.Global = True
  
  For Each c In r
    If containsDiapasone(c.Formula) Then
      temp = Split(c.Formula, ":")
      regExDigit.Execute (temp(1))
      'row = CInt(matched(0))
      regExColumn.Execute (temp(1))
      
      'Debug.Print (Column & "" & row)
    End If
  Next
End Sub

Public Sub fff()
  Dim activeWS As Worksheet
  Set activeWS = ActiveWorkbook.ActiveSheet
  updateVertFormula r:=activeWS.Range("G17"), i:=2
End Sub
