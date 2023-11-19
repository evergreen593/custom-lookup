Attribute VB_Name = "Custom_lookup"
Dim separator0 As String

Function superlookup(lookup_value, array1 As Range, return1 As Range)
'Return the first non-empty result of the lookup
Dim result As Variant
Dim count As Integer
Dim x As Variant
Dim sheet As Variant

If array1.Rows.count = 1 And array1.Columns.count > 1 Then GoTo horizontal:

'handling the user selecting entire column for lookup
Set sheet = array1.Worksheet

With sheet
If array1.Rows.count = array1.EntireColumn.Rows.count Then
Set array1 = .Range(array1.Range("A1").Address & ":" & .Cells(array1.EntireColumn.Rows.count, array1.Column).End(xlUp).Address)
End If

If return1.Rows.count = return1.EntireColumn.Rows.count Then
Set return1 = .Range(return1.Range("A1").Address & ":" & .Cells(array1.EntireColumn.Rows.count, array1.Column).End(xlUp).Offset(0, return1.Column - array1.Column).Address)
End If
End With

'Check if lookup_value does exist in lookup_array
If Application.WorksheetFunction.CountIf(array1, lookup_value) = 0 Then
superlookup = CVErr(xlErrNA)
Exit Function
End If
'start index&matching
On Error GoTo errorhandler:
For count = 1 To array1.Rows.count
result = Application.WorksheetFunction.index(return1, Application.Match(lookup_value, array1, 0))
If result = "" Then
x = Application.Match(lookup_value, array1, 0)
Set array1 = array1.Offset(x, 0).Resize(array1.Rows.count - x, array1.Columns.count)
Set return1 = return1.Offset(x, 0).Resize(return1.Rows.count - x, return1.Columns.count)
Else


'For not accumulating
superlookup = result
Exit Function
End If
Next count


horizontal:
Set sheet = array1.Worksheet
With sheet
If array1.Columns.count = array1.EntireRow.Columns.count Then
Set array1 = .Range(array1.Range("A1").Address & ":" & .Cells(array1.Row, array1.EntireRow.Columns.count).End(xlToLeft).Address)
End If

If return1.Columns.count = return1.EntireRow.Columns.count Then
Set return1 = .Range(return1.Range("A1").Address & ":" & .Cells(array1.Row, array1.EntireRow.Columns.count).End(xlToLeft).Offset(return1.Row - array1.Row, 0).Address)
End If
End With

'Check if lookup_value does exist in lookup_array
If Application.WorksheetFunction.CountIf(array1, lookup_value) = 0 Then
superlookup = CVErr(xlErrNA)
Exit Function
End If
'start index&matching
On Error GoTo errorhandler:
For count = 1 To array1.Columns.count
result = Application.WorksheetFunction.index(return1, Application.Match(lookup_value, array1, 0))
If result = "" Then
x = Application.Match(lookup_value, array1, 0)
Set array1 = array1.Offset(0, x).Resize(array1.Rows.count, array1.Columns.count - x)
Set return1 = return1.Offset(0, x).Resize(return1.Rows.count, return1.Columns.count - x)
Else


'For not accumulating
superlookup = result
Exit Function
End If
Next count


errorhandler:
result = ""
Resume Next
End Function
Function hyperlookup(lookup_value, array1 As Range, return1 As Range, Optional index As Integer = -1, Optional separator As String = "/")
'Return all non-empty value that matches the lookup
Dim result As Variant
Dim result2 As Variant
Dim count As Integer
Dim x As Variant
Dim sheet As Variant
Dim storage() As Variant
Dim i As Integer
ReDim storage(1 To 1)
i = 1

If array1.Rows.count = 1 And array1.Columns.count > 1 Then GoTo horizontal:

'handling the user selecting entire column for lookup
Set sheet = array1.Worksheet

With sheet
If array1.Rows.count = array1.EntireColumn.Rows.count Then
Set array1 = .Range(array1.Range("A1").Address & ":" & .Cells(array1.EntireColumn.Rows.count, array1.Column).End(xlUp).Address)
End If

If return1.Rows.count = return1.EntireColumn.Rows.count Then
Set return1 = .Range(return1.Range("A1").Address & ":" & .Cells(array1.EntireColumn.Rows.count, array1.Column).End(xlUp).Offset(0, return1.Column - array1.Column).Address)
End If
End With

'Check if lookup_value does exist in lookup_array
If Application.WorksheetFunction.CountIf(array1, lookup_value) = 0 Then
hyperlookup = CVErr(xlErrNA)
Exit Function
End If

'start index&matching
On Error GoTo errorhandler:
For count = 1 To array1.Rows.count
result = Application.WorksheetFunction.index(return1, Application.Match(lookup_value, array1, 0))
    'For indexing results
    ReDim Preserve storage(1 To i)
    storage(i) = result
        ' If index is 0 then regard as need to show all indexes as separator
        If index = 0 Then
            separator = " " & i + 1 & "."
        End If
    i = i + 1
If result = "" Then
result = "(empty)"
End If
On Error GoTo 0
'For accumulating results
result2 = result2 & result & separator
separator0 = separator
On Error GoTo errorhandler2:
x = Application.Match(lookup_value, array1, 0)

If array1.Rows.count - x > 0 Then
    Set array1 = array1.Offset(x, 0).Resize(array1.Rows.count - x, array1.Columns.count)
    Set return1 = return1.Offset(x, 0).Resize(return1.Rows.count - x, return1.Columns.count)
Else
    result2 = Left(result2, Len(result2) - Len(separator))
    hyperlookup = result2
    'if index =-1, which is the default value, then disable this feature and show all results with original separator. Will not enter below If statement
    'if index = 0, will show all results, indexed
    'if index is positive, will show that indexed value
    If Not (index = -1 Or index = 0) Then
        hyperlookup = storage(index)
    ElseIf index = 0 Then
        hyperlookup = "1." & result2
    End If
    Exit Function
End If

'End If
Next count
' vertical hyperlookup (except error handler) ends here. Below codes are for horizontal hyperlookup.


horizontal:
Set sheet = array1.Worksheet
With sheet
If array1.Columns.count = array1.EntireRow.Columns.count Then
Set array1 = .Range(array1.Range("A1").Address & ":" & .Cells(array1.Row, array1.EntireRow.Columns.count).End(xlToLeft).Address)
End If

If return1.Columns.count = return1.EntireRow.Columns.count Then
Set return1 = .Range(return1.Range("A1").Address & ":" & .Cells(array1.Row, array1.EntireRow.Columns.count).End(xlToLeft).Offset(return1.Row - array1.Row, 0).Address)
End If
End With

'start index&matching
On Error GoTo errorhandler:
For count = 1 To array1.Columns.count
result = Application.WorksheetFunction.index(return1, Application.Match(lookup_value, array1, 0))
    'For indexing results
    ReDim Preserve storage(1 To i)
    storage(i) = result
        ' If index is 0 then regard as need to show all indexes as separator
        If index = 0 Then
            separator = " " & i + 1 & "."
        End If
    i = i + 1
If result = "" Then
result = "(empty)"
End If
On Error GoTo 0
'For accumulating results
result2 = result2 & result & separator
separator0 = separator
On Error GoTo errorhandler2:
x = Application.Match(lookup_value, array1, 0)

If array1.Columns.count - x > 0 Then
    Set array1 = array1.Offset(0, x).Resize(array1.Rows.count, array1.Columns.count - x)
    Set return1 = return1.Offset(0, x).Resize(return1.Rows.count, return1.Columns.count - x)
Else
    result2 = Left(result2, Len(result2) - Len(separator))
    hyperlookup = result2
    'if index =-1, which is the default value, then disable this feature and show all results with original separator. Will not enter below If statement
    'if index = 0, will show all results, indexed
    'if index is positive, will show that indexed value
    If Not (index = -1 Or index = 0) Then
        hyperlookup = storage(index)
    ElseIf index = 0 Then
        hyperlookup = "1." & result2
    End If
    Exit Function
End If

'End If
Next count

errorhandler:
result = ""
'MsgBox (Err.Description)
Resume Next

errorhandler2:
result2 = Left(result2, Len(result2) - Len(separator))
hyperlookup = result2
    'if index =-1, which is the default value, then disable this feature and show all results with original separator. Will not enter below If statement
    'if index = 0, will show all results, indexed
    'if index is positive, will show that indexed value
    If Not (index = -1 Or index = 0) Then
        hyperlookup = storage(index)
    ElseIf index = 0 Then
        hyperlookup = "1." & result2
    End If
Exit Function

End Function


Sub splitting_hyperlookup()
Dim v As Variant
Dim target As Range
Dim source As Range
Dim i As Integer
Set source = Application.InputBox("select the source cell", Type:=8)
Set target = Application.InputBox("Select cell to contain the first data.", Type:=8)
Dim c As Long, count As Integer
Dim newline As String
Dim separator As String
separator = InputBox("Please specify the separator used:", , separator0)
newline = MsgBox("Do you want to add a new line for each result?", vbYesNo)
Range(source.Address & ":" & source.End(xlDown).Address).Select
c = Application.WorksheetFunction.CountA(Selection)

For count = 1 To c

v = Split(source.Text, separator)

If UBound(v) = 0 Then
target.Value2 = source.Value2
GoTo errorhandler1:
End If
'MsgBox (Rows("""" & target.Row & ":" & target.Row + UBound(v) & """"))
If newline = vbYes Then
Rows(target.Offset(1, 0).Row & ":" & target.Row + UBound(v)).Insert
End If

For i = 0 To UBound(v)
target.Offset(i, 0).Activate

target.Offset(i, 0) = v(i)
'MsgBox (target.Offset(i, 0))
Next i
'GoTo normal:
errorhandler1:

If UBound(v) = 0 Then
    Set target = target.Offset(1, 0)
    Set source = source.Offset(1, 0)
Else
    Set target = target.End(xlDown).Offset(1, 0)
    Set source = source.End(xlDown)
End If
'MsgBox (source.Address & "," & target.Address)

'normal:
Next count

End Sub

