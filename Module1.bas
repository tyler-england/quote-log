Attribute VB_Name = "Module1"
Sub PopulateLog()

On Error GoTo errhandler

Dim mateerFolder As String, subFolder As String, subSubFolder As String, yearFolders() As String, i As Integer
Dim curPath As String, quotePath As String, quoteFolder As String, rowNum As Integer, badNumber As Boolean
Dim quoteWB As Workbook, custName As String, modelNO As String, machDesc As String, quoteFile As String
Dim rowStart As Integer, rowStop As Integer, headRng As Range, sLayout As String

Sheet1.Activate 'clear current content
If Range("A2").Value > 0 Then
    If Range("A3").Value > 0 Then
        Range("A2").Select
        Selection.End(xlDown).Select
        rowStart = ActiveCell.Row + 1
    Else
        rowStart = 2
    End If
    
    Rows("2:" & rowStart).ClearContents
    Range("A1").Select
    rowStart = 0
    
End If

mateerFolder = "\\PSACLW02\Home\Applications\Applications-Share\Quotes\Mateer\"

subFolder = Dir(mateerFolder, vbDirectory)

Application.ScreenUpdating = False

Do While subFolder Like ".*" 'get to first subfolder
    subFolder = Dir()
Loop

i = 0

Do While subFolder > "" 'do for each folder in quote folder
    ReDim Preserve yearFolders(i)
    yearFolders(i) = subFolder
    i = i + 1
    subFolder = Dir()
Loop

For Each yearFolder In yearFolders 'for each year folder
    
    ThisWorkbook.Activate
    Sheet1.Activate
    'set rownum
    If Range("A2").Value > 0 Then
        Range("A1").Select
        Selection.End(xlDown).Select
        rowNum = ActiveCell.Row + 1
    Else
        rowNum = 2
    End If

    curPath = mateerFolder & yearFolder & "\"
    quoteFolder = Dir(curPath, vbDirectory)
    Do While quoteFolder Like ".*" 'get to first quote folder
        quoteFolder = Dir()
    Loop
    
    Do While quoteFolder > ""
        Range("A" & rowNum).Value = Left(quoteFolder, 12)
        rowNum = rowNum + 1
        quoteFolder = Dir()
    Loop
    
Next yearFolder

Range("A:A").Sort Key1:=Range("A1"), Order1:=xlDescending
'get deets for all quotes in A (delete if weird format)

rowNum = 2
Do While Range("A" & rowNum).Value > 0

    For i = 1 To Len(Range("A" & rowNum).Value)
        If Not (IsNumeric(Mid(Range("A" & rowNum).Value, i, 1))) And _
            Mid(Range("A" & rowNum).Value, i, 1) <> "-" Then 'text in quote number
                badNumber = True
        End If
    Next i
    
    If badNumber Then 'delete row
        badNumber = False
        Rows(rowNum).Delete
    Else 'get info
        curPath = mateerFolder & "20" & Left(Range("A" & rowNum), 2) & " Quotes\"
        quotePath = curPath & Range("A" & rowNum).Value & "*"
        quoteFolder = Dir(quotePath, vbDirectory)
        quotePath = curPath & quoteFolder & "\"
        
        Range("A" & rowNum).Formula = "=HYPERLINK(" & """" & quotePath & """" & "," & _
                                        """" & Range("A" & rowNum).Value & """" & ")"
                                        
        If quoteFolder > "" Then 'it exists where expected
            'find customer name
            custName = Right(quoteFolder, Len(quoteFolder) - 16)
            If InStr(custName, "-") > 0 Then 'name & model exist
                modelNO = Right(custName, Len(custName) - InStr(custName, "-"))
                If Not (IsNumeric(Left(modelNO, 1))) Then 'hyphen in customer name
                    modelNO = ""
                Else 'no hyphen in customer name
                    custName = Left(custName, InStr(custName, "-") - 1)
                    If Len(modelNO) > 5 Then
                        If UCase(Right(modelNO, 6)) = "FILLER" Or UCase(Right(modelNO, 6)) = "ROTARY" Then
                            modelNO = Left(modelNO, Len(modelNO) - 6)
                        End If
                    End If
                End If
            Else 'only customer name exists
                modelNO = ""
            End If
            
            Range("B" & rowNum).Value = custName
            Range("C" & rowNum).Value = modelNO
            
            'find file
            quoteFile = FindLatestWkbk(quotePath)
            If quoteFile <> "" Then 'xls file returned
                
                'open file
                Set quoteWB = Workbooks.Open(quotePath & quoteFile, UpdateLinks:=0)
                
                If Not quoteWB Is Nothing Then
                
                    'base machine info
                    rowStart = 0
                    quoteWB.Activate
                    On Error Resume Next
                        rowStart = Application.WorksheetFunction.Match("Base*Machine*", Range("A:A"), 0)
                        If rowStart = 0 Then
                            rowStart = Application.WorksheetFunction.Match("Base*Machine*", Range("B:B"), 0)
                        End If
                    On Error GoTo 0 'errhandler
                    
                    If rowStart = 0 Then 'aftermarket/weird format
                        rowStart = 1
                    End If
                    
                    If rowStart = 1 Then 'aftermarket/weird format
                        If UCase(Range("A1").Value) = "LINE ITEM" Then ' definitely aftermarket quote
                            ThisWorkbook.Sheets(1).Range("D" & rowNum).Value = "Aftermarket/Budgetary"
                            ThisWorkbook.Sheets(1).Range("E" & rowNum).Value = Range("A4").Value
                            If Range("A5").Value > 0 Then
                                Range("A4").Select
                                Selection.End(xlDown).Select
                                If ActiveCell.Row < 100 Then
                                    For i = 5 To ActiveCell.Row
                                        ThisWorkbook.Sheets(1).Range("E" & rowNum).Value = ThisWorkbook.Sheets(1).Range("E" & _
                                                                                    rowNum).Value & vbCrLf & Range("A" & i).Value
                                    Next i
                                End If
                            End If
                        Else 'weird unknown format
                            ThisWorkbook.Sheets(1).Range("D" & rowNum).Value = "Aftermarket/Budgetary (?)"
                        End If
                    Else 'new machine quote
                        ThisWorkbook.Sheets(1).Range("D" & rowNum).Value = "New machine"
                        Set headRng = Range("A" & rowStart + 1 & ":B" & rowStart + 3)
                        For Each cell In headRng
                            If UCase(cell.Value) <> "DESCRIPTION" And cell.Value > 0 And Not (cell.Value Like "*Price*") Then
                                If ThisWorkbook.Sheets(1).Range("E" & rowNum).Value = 0 Then
                                    ThisWorkbook.Sheets(1).Range("E" & rowNum).Value = cell.Value
                                Else
                                    ThisWorkbook.Sheets(1).Range("E" & rowNum).Value = ThisWorkbook.Sheets(1).Range("E" & _
                                                                                    rowNum).Value & vbCrLf & cell.Value
                                End If
                                Exit For
                            End If
                        Next
                        
                        'options info
                        quoteWB.Activate
                        rowStart = 0
                        On Error Resume Next
                            rowStart = Application.WorksheetFunction.Match("*Options*", Range("A:A"), 0)
                            If rowStart = 0 Then
                                rowStart = Application.WorksheetFunction.Match("*Options*", Range("B:B"), 0)
                            End If
                        On Error GoTo 0 'errhandler
                        
                        If rowStart > 0 Then 'options are listed
                            rowStart = rowStart + 1
                            If UCase(Range("A" & rowStart).Value) = "DESCRIPTION" Or _
                                UCase(Range("B" & rowStart).Value) = "DESCRIPTION" Then
                                rowStart = rowStart + 1
                            End If
                            Set headRng = Range("A" & rowStart + 1 & ":B" & rowStart + 1) 'to check for multiple options
                            rowStop = rowStart 'default (only 1 option)
                            For Each cell In headRng
                                If cell.Value > 0 And Not (IsNumeric(cell.Value)) And Not (UCase(cell.Value) = "TBD") Then 'redefine rowstop
                                    cell.Select
                                    Selection.Offset(-1, 0).Select
                                    Selection.End(xlDown).Select
                                    rowStop = ActiveCell.Row 'last row under options
                                End If
                            Next cell
                            
                            Set headRng = Range("A" & rowStart & ":B" & rowStop) 'new range (all the options)
                            
                            For Each cell In headRng
                                If cell.Value > 0 And Not (IsNumeric(cell.Value)) And Not (UCase(cell.Value) = "TBD") Then 'add to log
                                    ThisWorkbook.Sheets(1).Range("E" & rowNum).Value = ThisWorkbook.Sheets(1).Range("E" & _
                                                                                rowNum).Value & vbCrLf & cell.Value
                                End If
                            Next cell
                        End If
                        
                    End If
                    
                    quoteWB.Close savechanges:=False
                    
                End If
            Else 'no xls file returned
            
                ThisWorkbook.Sheets(1).Range("D" & rowNum).Value = "Aftermarket/Budgetary (?)"
            
            End If
            
            sLayout = Dir$(quotePath & "*.dwg")
            If sLayout <> "" Then
                ThisWorkbook.Sheets(1).Range("F" & rowNum).Value = "YES"
            Else
                ThisWorkbook.Sheets(1).Range("F" & rowNum).Value = ""
            End If
            
        End If
        
        rowNum = rowNum + 1
        
    End If

Loop

Columns("B:B").Cells.HorizontalAlignment = xlHAlignLeft
Columns("E:E").Cells.HorizontalAlignment = xlHAlignLeft
Columns("A:E").AutoFit
Rows("1:1").Cells.HorizontalAlignment = xlHAlignCenter
Range("A1").Select
Selection.Offset(1, 0).Select
Application.ScreenUpdating = True
Exit Sub
errhandler:
MsgBox "Error in PopulateLog sub"
End Sub

Function FindLatestWkbk(quotePath As String) As String

Dim latestWB As String, testWB As String
Dim latestDate As Date, testDate As Date

latestWB = ""

On Error GoTo errhandler

latestWB = Dir(quotePath & "*.xls*")

If latestWB = "" Then
    Exit Function
End If

latestDate = FileDateTime(quotePath & latestWB)

testWB = Dir()

Do While testWB > ""

    testDate = FileDateTime(quotePath & testWB)
    If testDate > latestDate And IsNumeric(Left(testWB, 1)) Then
        latestWB = testWB
        latestDate = testDate
    End If
    testWB = Dir()
Loop

If IsNumeric(Left(latestWB, 1)) Then
    FindLatestWkbk = latestWB
End If

Exit Function
errhandler:
MsgBox "Error in FindLatestWkbk function"
End Function


