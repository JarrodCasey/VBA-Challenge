Attribute VB_Name = "Module1"
Sub VBAofWallSt()

'===== DEFINE VARIABLES APPLICABLE TO ALL WORKSHEET =====
Dim YrStart As Double
Dim YrEnd As Double
Dim YrlChg As Double

'===== START LOOP FOR WORKSHEET=====
' For each Worksheet in this workbook
For Each ws In Worksheets

    '----- CREATE COLUMN HEADERS -----
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change ($)"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    '----- CREATE VARIABLE "LASTROW" -----
    'This sets the number of rows with records into the variable "LastRow"
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
    '----- FORMAT COLUMN "K" AS A % -----
    ws.Range("K:K").NumberFormat = "0.00%"
    
    '----- FORMAT COLUMN "J" AS "0.00" -----
    ws.Range("J:J").NumberFormat = "0.00"
    
    '----- SET CONDITIONAL FORMATTING FOR COLUMN "J" -----
    'Delete previous conditional formats (in case one is previously created in the range)
    ws.Range("J2:J" & LastRow).FormatConditions.Delete
    
    'Add Condition 1: if cell value is <blank> then no colour
    ws.Range("J2:J" & LastRow).FormatConditions.Add Type:=xlExpression, Formula1:= _
                               "=LEN(TRIM(J2))=0"
    ws.Range("J2:J" & LastRow).FormatConditions(1).Interior.Pattern = xlNone
    
    'Add Condition 2: if cell value is less than 0.00 then colour cell Red
    ws.Range("J2:J" & LastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                               Formula1:="=0.00"
    ws.Range("J2:J" & LastRow).FormatConditions(2).Interior.ColorIndex = 3
    
    'Add Condition 3: if cell value is greater than 0.00 then colour cell green
    ws.Range("J2:J" & LastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
                               Formula1:="=0.00"
    ws.Range("J2:J" & LastRow).FormatConditions(3).Interior.ColorIndex = 4
    
    '----- CREATE VARIABLE "TICKERCOUNTER" -----
    'Variable will be used to count the number of time a Ticker appears in a row before it changes to a different ticker
    Dim TickerCounter As Integer
    'Set TickerCounter starting value to "0"
    TickerCounter = 0
    
    '----- CREATE VARIABLE "SUMMROW" -----
    'Variable will be used to manage where the data will go in the summary table
    Dim SummRow As Integer
    'Set SummRow starting value to "2"; This is because the data will start in row 2 (Due to Column Headers)
    SummRow = 2
    
    '~~~~~ BONUS: CREATE TABLE FOR BONUS TABLE ~~~~~
    'Create the table
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    '----- FORMAT Cells Q2 and Q3 as A % -----
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    '===== START LOOP FOR ROWS PER WORKSHEET =====
    'Loop through rows in the column
    For i = 2 To LastRow
        
        '===== START IF STATEMENT TO LOOK FOR A CHANGE IN TICKER CODE IN ROW A OF WS =====
        'If the Ticker Code change is detected
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        '===== IF TICKER CODE CHANGE IS DETECTED THNE COMPLETE SUMMARY TABLE FOR THAT TICKER =====
        '-----Put the ticker code in the next available space in Column I-----
        ws.Cells(SummRow, 9).Value = ws.Cells(i, 1).Value
        
        '----- CALCULATE THE YEARLY CHANGE OF THE TICKER -----
        'Identify the Ticker's Opening Price for the year
        YrStart = ws.Cells(i - TickerCounter, 3).Value
        'Identify the Ticker's Closing Price for the year
        YrEnd = ws.Cells(i, 6).Value
        'Calculate the yearly change
        YrlChg = YrStart - YrEnd
        '----- PUT YEARLY CHANGE OF THE TICKER IN THE NEXT AVAILABLE SPEACE IN COLUMN J -----
        ws.Cells(SummRow, 10).Value = YrlChg
        
        '----- CALCULATE THE % YRLY CHANGE FROM OPENING PRICE AND POPULATE IN THE NEXT AVAILABLE SPACE IN COLUMN K -----
        ws.Cells(SummRow, 11).Value = (YrlChg / YrStart)
        
        '----- CALCULATE THE TOTAL STOCK VOLUME AND POPULATE IN THE NEXT AVAILABLE SPACE IN COLUMN L -----
        ws.Cells(SummRow, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(i - TickerCounter, 7), ws.Cells(i, 7)))
        
        '----- PREPARE KEY VARIABLES FOR THE NEXT SET OF TICKER -----
        'Reset TickerCounter to "0"
        TickerCounter = 0
        'Add "1" to the SummRow value
        SummRow = SummRow + 1
        
        
        '===== IF TICKER CODE CHANGE IS *NOT* DETECTED =====
        Else
        'Add "1" to the TickerCounter
            TickerCounter = TickerCounter + 1
            
        '===== END IF STATEMENT =====
        End If
        
    '===== END LOOP FOR ROWS PER WORKSHEET =====
    Next i
    
'~~~~~ BONUS: POPULATE BONUS TABLE ~~~~~
'----- Identify Greatest % Increase -----
'Create Variables for Greatest % Increase
Dim GrtPerIncTic As String
Dim GrtPerIncVal As Double

'Set Variable GrtPerIncVal to "0"
GrtPerIncVal = 0

'----- CREATE VARIABLE "LASTROWBON" -----
'This sets the number of rows with records into the variable "LastRowBon" based on the table created in the main loop
LastRowBon = ws.Cells(Rows.Count, "J").End(xlUp).Row

'----- START LOOP TO IDENTIFY GREATES % INCREASE -----
'Start loop
For j = 2 To LastRowBon

    '----- START IF STATEMENT TO IDENTIFY GREATEST % -----
    'If the current cell's value is greater than the next cell's value
    If ws.Cells(j, 11).Value > GrtPerIncVal Then
    
    GrtPerIncTic = ws.Cells(j, 9).Value
    GrtPerIncVal = ws.Cells(j, 11).Value
    
    '----- END IF STATEMENT -----
    End If
    
    '----- END LOOP TO IDENTIFY GREATES % INCREASE -----
    Next j
    
'----- POPULATE BONUS TABLE FOR GREATEST % INCREASE -----
ws.Range("P2").Value = GrtPerIncTic
ws.Range("Q2").Value = GrtPerIncVal

'----- Identify Greatest % Decrease -----
'Create Variables for Greatest % Decrease
Dim GrtPerDecTic As String
Dim GrtPerDecVal As Double

'Set Variable GrtPerDecVal to "0"
GrtPerDecVal = 0

'----- START LOOP TO IDENTIFY GREATES % DECREASE -----
'Start loop
For k = 2 To LastRowBon

    '----- START IF STATEMENT TO IDENTIFY GREATEST % DECREASE-----
    'If the current cell's value is greater than the next cell's value
    If ws.Cells(k, 11).Value < GrtPerDecVal Then
    
    GrtPerDecTic = ws.Cells(k, 9).Value
    GrtPerDecVal = ws.Cells(k, 11).Value
    
    '----- END IF STATEMENT -----
    End If
    
    '----- END LOOP TO IDENTIFY GREATES % DECREASE -----
    Next k
    
'----- POPULATE BONUS TABLE FOR GREATEST % DECREASE -----
ws.Range("P3").Value = GrtPerDecTic
ws.Range("Q3").Value = GrtPerDecVal

'----- Identify Greatest Total Volume -----
'Create Variables for Greatest Total Volume
Dim GrtTotVolTic As String
Dim GrtTotVolVal As Double

'Set Variable GrtPerDecVal to "0"
GrtTotVolVal = 0

'----- START LOOP TO IDENTIFY GREATES TOTAL VOLUME -----
'Start loop
For l = 2 To LastRowBon

    '----- START IF STATEMENT TO IDENTIFY GREATEST TOTAL VOLUME -----
    'If the current cell's value is greater than the next cell's value
    If ws.Cells(l, 12).Value > GrtTotVolVal Then
    
    GrtTotVolTic = ws.Cells(l, 9).Value
    GrtTotVolVal = ws.Cells(l, 12).Value
    
    '----- END IF STATEMENT -----
    End If
    
    '----- END LOOP TO IDENTIFY GREATES TOTAL VOLUME -----
    Next l
    
'----- POPULATE BONUS TABLE FOR GREATEST TOTAL VOLUME -----
ws.Range("P4").Value = GrtTotVolTic
ws.Range("Q4").Value = GrtTotVolVal
    
'===== END LOOP FOR WORKSHEET =====
Next ws

End Sub
