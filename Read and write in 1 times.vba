Sub WriteCsvFile()

'=================================================
' READ HEADER
'=================================================
' Header at first row
' Active cell A1 - Cell start read header
Range("A1").Select

' Get header
Header1 = ActiveCell.Offset(0, 0).Value
Header2 = ActiveCell.Offset(0, 1).Value
Header3 = ActiveCell.Offset(0, 2).Value
Header4 = ActiveCell.Offset(0, 3).Value
Header5 = ActiveCell.Offset(0, 4).Value
Header6 = ActiveCell.Offset(0, 5).Value
Header7 = ActiveCell.Offset(0, 6).Value
Header8 = ActiveCell.Offset(0, 7).Value
Header9 = ActiveCell.Offset(0, 8).Value
Header10 = ActiveCell.Offset(0, 9).Value
Header11 = ActiveCell.Offset(0, 10).Value

' Define result return
Dim headerContent As String

' Append header
headerContent = Header1 _
    & "," _
    & Header2 _
    & "," _
    & Header3 _
    & "," _
    & Header4 _
    & "," _
    & Header5 _
    & "," _
    & Header6 _
    & "," _
    & Header7 _
    & "," _
    & Header8 _
    & "," _
    & Header9 _
    & "," _
    & Header10 _
    & "," _
    & Header11


'=================================================
' READ DATA
'=================================================
' CSV data from line 2st
' Active cell A2 - Cell start read data
Range("A2").Select

' Define row offset index start
Dim rowOff As Integer
rowOff = 0

' Define result return
Dim dataContent As String

' Read data - Using trigger "END-END-END-END" to stop the reading process
While (ActiveCell.Offset(rowOff, 0).Value <> "END-END-END-END")
    ' Read 11 column
    Data1 = ActiveCell.Offset(rowOff, 0).Value
    Data2 = ActiveCell.Offset(rowOff, 1).Value
    Data3 = ActiveCell.Offset(rowOff, 2).Value
    Data4 = ActiveCell.Offset(rowOff, 3).Value
    Data5 = ActiveCell.Offset(rowOff, 4).Value
    Data6 = ActiveCell.Offset(rowOff, 5).Value
    Data7 = ActiveCell.Offset(rowOff, 6).Value
    Data8 = ActiveCell.Offset(rowOff, 7).Value
    Data9 = ActiveCell.Offset(rowOff, 8).Value
    Data10 = ActiveCell.Offset(rowOff, 9).Value
    Data11 = ActiveCell.Offset(rowOff, 10).Value
    
    ' Append result to write in 1 line
    dataContent = dataContent _
        & vbNewLine _
        & Data1 _
        & "," _
        & Data2 _
        & "," _
        & Data3 _
        & "," _
        & Data4 _
        & "," _
        & Data5 _
        & "," _
        & Data6 _
        & "," _
        & Data7 _
        & "," _
        & Data8 _
        & "," _
        & Data9 _
        & "," _
        & Data10 _
        & "," _
        & Data11
    
    
    ' Next row
    rowOff = rowOff + 1

Wend


'=================================================
' CREATE FILE
' Write in 1 time
'=================================================

' Create file with path: path of workbook + sheet name + .csv
pFolder = Application.ActiveWorkbook.Path
' Get sheet name to using as file name
pFile = Application.ActiveWorkbook.ActiveSheet.Name


Dim ADODBStream As ADODB.stream
    Set ADODBStream = CreateObject("ADODB.Stream")
        ADODBStream.Charset = "shift_jis"
        ADODBStream.Open
        ADODBStream.WriteText headerContent & dataContent 'Write all content in 1 time
        ADODBStream.SaveToFile pFolder & "\" & pFile & ".csv", adSaveCreateOverWrite
        ADODBStream.Close
    Set ADODBStream = Nothing
    

' Finished
MsgBox ("Created file: " & pFolder & "\" & pFile & ".csv")

End Sub
