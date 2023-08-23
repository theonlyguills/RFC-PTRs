Attribute VB_Name = "RFCModule"
Sub Rockliffe_Report(ByRef control As Office.IRibbonControl)
 Dim ws1 As Worksheet
 Dim ws2 As Worksheet
 Dim i As Integer
 Dim split_array As Variant
 Dim split_array2 As Variant
 Dim r As Range
 Dim rCell As Range
 Dim col As Integer
 Dim col2 As Integer
 Dim col3 As Integer
 Dim col4 As Integer
 Dim col5 As Integer
 Dim col6 As Integer
 Dim col7 As Integer
 Dim col8, col9, col10, col11, col12, col13, col14, col15, col16 As Integer
 Dim fly_total As Single
 Dim temp1 As Single
Dim temp2 As Single
Dim z As Integer
Dim tempstring As String
Dim header_array As Variant
Dim last_row As Integer
Dim name_header As String
Set ws1 = ActiveSheet

Set r = ws1.Range(Cells(1, 1), Cells(1, ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column))
'sets the Row 1 of the original worksheet to be range "r"
last_row = (ws1.Cells(Rows.Count, "A").End(xlUp).Row) - 1
    With ActiveWorkbook
        Set ws2 = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws2.Name = "Report"
    End With
'creates a new sheet and names it Report

ws2.Cells(1, "A") = "Date"
ws2.Cells(1, "B") = "Student"
ws2.Cells(1, "C") = "Instructor"
ws2.Cells(1, "D") = "Aircraft Type"
ws2.Cells(1, "E") = "Tailnumber"
ws2.Cells(1, "F") = "Dual Day Total"
ws2.Cells(1, "G") = "Solo Day Total"
ws2.Cells(1, "H") = "Instrument (Hood)"
ws2.Cells(1, "I") = "Flight Sim (Instrument)"
ws2.Cells(1, "J") = "Dual Day XC"
ws2.Cells(1, "K") = "Solo Day XC"
ws2.Cells(1, "L") = "From Airport"
ws2.Cells(1, "M") = "Via Airports"
ws2.Cells(1, "N") = "To Airport"
ws2.Cells(1, "O") = "Lesson"
ws2.Cells(1, "P") = "2a - Documents and Airworthiness"
ws2.Cells(1, "Q") = "2b - Aeroplane Performance"
ws2.Cells(1, "R") = "2c - Wt. and Balance, Loading"
ws2.Cells(1, "S") = "2d - Pre-flight Inspection"
ws2.Cells(1, "T") = "2e - Engine Start/Run-Up/Check List"
ws2.Cells(1, "U") = "2f - Operation of A/C Systems"
ws2.Cells(1, "V") = "3 - Ancillary Controls"
ws2.Cells(1, "W") = "4 - Taxiing"
ws2.Cells(1, "X") = "5 - Attitudes & Movements"
ws2.Cells(1, "Y") = "6 - Straight & Level Flight"
ws2.Cells(1, "Z") = "7 - Climbing"
ws2.Cells(1, "AA") = "8 - Descending"
ws2.Cells(1, "AB") = "9 - Turning"
ws2.Cells(1, "AC") = "9s - Steep Turn"
ws2.Cells(1, "AD") = "10 - Range & Endurance"
ws2.Cells(1, "AE") = "11 - Slow Flight"
ws2.Cells(1, "AF") = "12a - Power-Off Stall"
ws2.Cells(1, "AG") = "12b - Power-On Stall"
ws2.Cells(1, "AH") = "13 - Spin"
ws2.Cells(1, "AI") = "14 - Spiral"
ws2.Cells(1, "AJ") = "15 - Slipping"
ws2.Cells(1, "AK") = "16a - Normal Takeoff"
ws2.Cells(1, "AL") = "16b - Crosswind"
ws2.Cells(1, "AM") = "16b - Obstacle"
ws2.Cells(1, "AN") = "16b - Short/Minimum Run"
ws2.Cells(1, "AO") = "16b - Soft/Rough"
ws2.Cells(1, "AP") = "17 - Circuit"
ws2.Cells(1, "AQ") = "18a - 180 Power Off"
ws2.Cells(1, "AR") = "18a - Normal Landing"
ws2.Cells(1, "AS") = "18b - Crosswind"
ws2.Cells(1, "AT") = "18b - Obstacle"
ws2.Cells(1, "AU") = "18b - Short Field"
ws2.Cells(1, "AV") = "18b - Soft/Rough"
ws2.Cells(1, "AW") = "18c - Overshoot"
ws2.Cells(1, "AX") = "19 - First Solo"
ws2.Cells(1, "AY") = "20 - Illusions"
ws2.Cells(1, "AZ") = "21a - Precautionary - On Aerodrome"
ws2.Cells(1, "BA") = "21b - Precautionary - Off Aerodrome"
ws2.Cells(1, "BB") = "22a - Forced - (Control / Approach)"
ws2.Cells(1, "BC") = "22b - Forced - (Cockpit Management)"
ws2.Cells(1, "BD") = "23 - Navigation"
ws2.Cells(1, "BE") = "23a - Pre-Flight Planning Procedures"
ws2.Cells(1, "BF") = "23b - Departure Procedure"
ws2.Cells(1, "BG") = "23c - En-Route Procedure"
ws2.Cells(1, "BH") = "23d - Diversion to an Alternate"
ws2.Cells(1, "BI") = "24a - Full Panel"
ws2.Cells(1, "BJ") = "24b - Limited Panel"
ws2.Cells(1, "BK") = "24c - Unusual Attitude"
ws2.Cells(1, "BL") = "24d - Radio Navigation"
ws2.Cells(1, "BM") = "29 - Emergencies"
ws2.Cells(1, "BN") = "30 - Radio"
ws2.Cells(1, "BO") = ""
ws2.Cells(1, "BP") = "Date"
ws2.Cells(1, "BQ") = "Comments"
'prints the new headers in the report

For i = 2 To last_row
On Error Resume Next
split_array = Split(ws1.Cells(i, "A").Value, " ", 2)
ws2.Cells(i, "A").Value = split_array(0)
ws2.Cells(i, "BP").Value = split_array(0)
Next
'prints the date without the time stamp


For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Student" Then
    col = rCell.Column
End If
Next
name_header = ws1.Cells(2, col)
'finds the Student column in the original document

For i = 2 To last_row
On Error Resume Next
split_array = Split(ws1.Cells(i, col).Value, " ", 2)
ws2.Cells(i, "B").Value = split_array(1)
Next
'prints last name only of student

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Instructor" Then
    col = rCell.Column
End If
Next
'finds which cell has instructor name

For i = 2 To last_row
On Error Resume Next
split_array = Split(ws1.Cells(i, col).Value, " ", 2)
ws2.Cells(i, "C").Value = split_array(1)
Next
'prints last name only of student

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Aircraft Type" Then
    col = rCell.Column
End If
Next
'finds which cell has Aircraft type

For i = 2 To last_row
On Error Resume Next
    If InStr(ws1.Cells(i, col).Value, "172") > 0 Then
        ws2.Cells(i, "D").Value = "C172"
    ElseIf InStr(ws1.Cells(i, col).Value, "150") > 0 Then
        ws2.Cells(i, "D").Value = "C150"
    ElseIf InStr(ws1.Cells(i, col).Value, "edbird") > 0 Then
        ws2.Cells(i, "D").Value = "RB"
    ElseIf InStr(ws1.Cells(i, col).Value, "iamond") > 0 Then
        ws2.Cells(i, "D").Value = "DA20"
    Else
        ws2.Cells(i, "D").Value = ws1.Cells(i, col).Value
    End If
Next
'prints abbreviated aircraft type; if the type does not match one listed above, it will be printed as is

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Tailnumber" Then
    col = rCell.Column
End If
Next
'finds which cell has Tailnumber (or Registration Type)

For i = 2 To last_row
On Error Resume Next
    If InStr(ws1.Cells(i, col).Value, "4073") > 0 Then
        ws2.Cells(i, "E").Value = "4073"
    ElseIf InStr(ws1.Cells(i, col).Value, "GABC") > 0 Then
        ws2.Cells(i, "E").Interior.ColorIndex = 8
        ws2.Cells(i, "E").Value = ws1.Cells(i, col).Value
    Else
        If InStr(ws1.Cells(i, col).Value, "-") > 0 Then
        split_array = Split(ws1.Cells(i, col).Value, "-", 2)
        ws2.Cells(i, "E").Value = split_array(1)
        Else
        ws2.Cells(i, "E").Value = ws1.Cells(i, col).Value
        End If
    End If
Next
'prints abbreviated Tailnumbers - if no match found, prints original as is

ActiveWorkbook.PrecisionAsDisplayed = True
'avoids the space-time continuum problem.  Makes sure decimals are counted as shown

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Dual Day Local" Then
    col = rCell.Column
ElseIf rCell.Value = "Dual Day XC" Then
    col2 = rCell.Column
End If
Next
'finds which cell has Dual Day Local and Dual Day XC

For i = 2 To last_row
On Error Resume Next
temp1 = ws1.Cells(i, col).Value
temp2 = ws1.Cells(i, col2).Value
fly_total = temp1 + temp2
ws2.Cells(i, "F").Value = fly_total
ws2.Cells(i, "F").NumberFormat = "0.0"
fly_total = 0
temp1 = 0
temp2 = 0
Next
'adds together dual day local and dual day xc then prints them with only one decimal place

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Solo Day Local" Then
    col = rCell.Column
ElseIf rCell.Value = "Solo Day XC" Then
    col2 = rCell.Column
End If
Next
'finds which cell has Solo Day Local and Solo Day XC


For i = 2 To last_row
On Error Resume Next
temp1 = ws1.Cells(i, col).Value
temp2 = ws1.Cells(i, col2).Value
fly_total = temp1 + temp2
ws2.Cells(i, "G").Value = fly_total
ws2.Cells(i, "G").NumberFormat = "0.0"
fly_total = 0
temp1 = 0
temp2 = 0
Next
'adds together Solo Day Local and Solo Day XC

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Instrument (Hood)" Then
    col = rCell.Column
ElseIf rCell.Value = "Flight Sim (Instrument)" Then
    col2 = rCell.Column
ElseIf rCell.Value = "Dual Day XC" Then
    col3 = rCell.Column
ElseIf rCell.Value = "Solo Day XC" Then
    col4 = rCell.Column
ElseIf rCell.Value = "From Airport" Then
    col5 = rCell.Column
ElseIf rCell.Value = "Via Airports" Then
    col6 = rCell.Column
ElseIf rCell.Value = "To Airport" Then
    col7 = rCell.Column
End If
Next
'identifies 7 columns we will print as is

For i = 2 To last_row
On Error Resume Next
ws2.Cells(i, "H").NumberFormat = "0.0"
ws2.Cells(i, "H").Value = ws1.Cells(i, col).Value
ws2.Cells(i, "I").NumberFormat = "0.0"
ws2.Cells(i, "I").Value = ws1.Cells(i, col2).Value
ws2.Cells(i, "J").NumberFormat = "0.0"
ws2.Cells(i, "J").Value = ws1.Cells(i, col3).Value
ws2.Cells(i, "K").NumberFormat = "0.0"
ws2.Cells(i, "K").Value = ws1.Cells(i, col4).Value
ws2.Cells(i, "L").Value = ws1.Cells(i, col5).Value
ws2.Cells(i, "M").Value = ws1.Cells(i, col6).Value
ws2.Cells(i, "N").Value = ws1.Cells(i, col7).Value
Next
'prints these 7 columns with no changes except formatting the numbers to one decimal place

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Lesson" Then
    col = rCell.Column
End If
Next
'finds which cell has Lesson

For i = 2 To last_row
On Error Resume Next
    If InStr(ws1.Cells(i, col).Value, " ") > 0 Then
    split_array = Split(ws1.Cells(i, col).Value, "-", 2)
        If IsNumeric(split_array(0)) Then
            split_array2 = Split(ws1.Cells(i, col).Value, " ", 2)
            ws2.Cells(i, "O").NumberFormat = "@"
            ws2.Cells(i, "O").Value = split_array2(0)
        Else
            ws2.Cells(i, "O").NumberFormat = "@"
            ws2.Cells(i, "O").Value = ws1.Cells(i, col).Value
        End If
    Else
    ws2.Cells(i, "O").NumberFormat = "@"
    ws2.Cells(i, "O").Value = ws1.Cells(i, col).Value
    End If
Next
'prints lesson information with numbers only (if they exist) and formatted as text

col = 0
col1 = 0
col2 = 0
col3 = 0
col4 = 0
col5 = 0
col6 = 0
col7 = 0
col8 = 0
col9 = 0
col10 = 0
col11 = 0
col12 = 0
col13 = 0
col14 = 0
col15 = 0
col16 = 0

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "2a - Documents and Airworthiness" Then
    col = rCell.Column
ElseIf rCell.Value = "2b - Aeroplane Performance" Then
    col2 = rCell.Column
ElseIf rCell.Value = "2c - Wt. and Balance, Loading" Then
    col3 = rCell.Column
ElseIf rCell.Value = "2d - Pre-flight Inspection" Then
    col4 = rCell.Column
ElseIf rCell.Value = "2e - Engine Start/Run-Up/Check List" Then
    col5 = rCell.Column
ElseIf rCell.Value = "2f - Operation of A/C Systems" Then
    col6 = rCell.Column
ElseIf rCell.Value = "3 - Ancillary Controls" Then
    col7 = rCell.Column
ElseIf rCell.Value = "4 - Taxiing" Then
   col8 = rCell.Column
ElseIf rCell.Value = "5 - Attitudes & Movements" Then
   col9 = rCell.Column
ElseIf rCell.Value = "6 - Straight & Level Flight" Then
   col10 = rCell.Column
ElseIf rCell.Value = "7 - Climbing" Then
   col11 = rCell.Column
ElseIf rCell.Value = "8 - Descending" Then
   col12 = rCell.Column
ElseIf rCell.Value = "9 - Steep Turn" Then
   col13 = rCell.Column
ElseIf rCell.Value = "9s - Steep Turn" Then
   col14 = rCell.Column
ElseIf rCell.Value = "10 - Range & Endurance" Then
   col15 = rCell.Column
ElseIf rCell.Value = "11 - Slow Flight" Then
    col16 = rCell.Column
End If
Next
'identifies first 15 lesson columns

For i = 2 To last_row
On Error Resume Next
    If col > 0 Then
        ws2.Cells(i, "P").Value = ws1.Cells(i, col).Value
    End If
    If col2 > 0 Then
        ws2.Cells(i, "Q").Value = ws1.Cells(i, col2).Value
    End If
    If col3 > 0 Then
        ws2.Cells(i, "R").Value = ws1.Cells(i, col3).Value
    End If
    If col4 > 0 Then
        ws2.Cells(i, "S").Value = ws1.Cells(i, col4).Value
    End If
    If col5 > 0 Then
        ws2.Cells(i, "T").Value = ws1.Cells(i, col5).Value
    End If
    If col6 > 0 Then
        ws2.Cells(i, "U").Value = ws1.Cells(i, col6).Value
    End If
    If col7 > 0 Then
        ws2.Cells(i, "V").Value = ws1.Cells(i, col7).Value
    End If
    If col8 > 0 Then
        ws2.Cells(i, "W").Value = ws1.Cells(i, col8).Value
    End If
    If col9 > 0 Then
        ws2.Cells(i, "X").Value = ws1.Cells(i, col9).Value
    End If
    If col10 > 0 Then
        ws2.Cells(i, "Y").Value = ws1.Cells(i, col10).Value
    End If
    If col11 > 0 Then
        ws2.Cells(i, "Z").Value = ws1.Cells(i, col11).Value
    End If
    If col12 > 0 Then
        ws2.Cells(i, "AA").Value = ws1.Cells(i, col12).Value
    End If
    If col13 > 0 Then
        ws2.Cells(i, "AB").Value = ws1.Cells(i, col13).Value
    End If
    If col14 > 0 Then
        ws2.Cells(i, "AC").Value = ws1.Cells(i, col14).Value
    End If
    If col15 > 0 Then
        ws2.Cells(i, "AD").Value = ws1.Cells(i, col15).Value
    End If
    If col16 > 0 Then
        ws2.Cells(i, "AE").Value = ws1.Cells(i, col16).Value
    End If
Next
'populates first 15 lesson columns if they exist in the original

col = 0
col1 = 0
col2 = 0
col3 = 0
col4 = 0
col5 = 0
col6 = 0
col7 = 0
col8 = 0
col9 = 0
col10 = 0
col11 = 0
col12 = 0
col13 = 0
col14 = 0
col15 = 0

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "12a - Power-Off Stall" Then
    col = rCell.Column
ElseIf rCell.Value = "12b - Power-On Stall" Then
    col2 = rCell.Column
ElseIf rCell.Value = "13 - Spin" Then
    col3 = rCell.Column
ElseIf rCell.Value = "14 - Spiral" Then
    col4 = rCell.Column
ElseIf rCell.Value = "15 - Slipping" Then
    col5 = rCell.Column
ElseIf rCell.Value = "16a - Normal Takeoff" Then
    col6 = rCell.Column
ElseIf rCell.Value = "16b - Crosswind" Then
    col7 = rCell.Column
ElseIf rCell.Value = "16b - Obstacle" Then
   col8 = rCell.Column
ElseIf rCell.Value = "16b - Short/Minimum Run" Then
   col9 = rCell.Column
ElseIf rCell.Value = "16b - Soft/Rough" Then
   col10 = rCell.Column
ElseIf rCell.Value = "17 - Circuit" Then
   col11 = rCell.Column
ElseIf rCell.Value = "18a - 180 Power Off" Then
   col12 = rCell.Column
ElseIf rCell.Value = "18a - Normal Landing" Then
   col13 = rCell.Column
ElseIf rCell.Value = "18b - Crosswind" Then
   col14 = rCell.Column
ElseIf rCell.Value = "18b - Obstacle" Then
   col15 = rCell.Column
End If
Next
'identifies second 15 lesson columns

For i = 2 To last_row
On Error Resume Next
    If col > 0 Then
        ws2.Cells(i, "AF").Value = ws1.Cells(i, col).Value
    End If
    If col2 > 0 Then
        ws2.Cells(i, "AG").Value = ws1.Cells(i, col2).Value
    End If
    If col3 > 0 Then
        ws2.Cells(i, "AH").Value = ws1.Cells(i, col3).Value
    End If
    If col4 > 0 Then
        ws2.Cells(i, "AI").Value = ws1.Cells(i, col4).Value
    End If
    If col5 > 0 Then
        ws2.Cells(i, "AJ").Value = ws1.Cells(i, col5).Value
    End If
    If col6 > 0 Then
        ws2.Cells(i, "AK").Value = ws1.Cells(i, col6).Value
    End If
    If col7 > 0 Then
        ws2.Cells(i, "AL").Value = ws1.Cells(i, col7).Value
    End If
    If col8 > 0 Then
        ws2.Cells(i, "AM").Value = ws1.Cells(i, col8).Value
    End If
    If col9 > 0 Then
        ws2.Cells(i, "AN").Value = ws1.Cells(i, col9).Value
    End If
    If col10 > 0 Then
        ws2.Cells(i, "AO").Value = ws1.Cells(i, col10).Value
    End If
    If col11 > 0 Then
        ws2.Cells(i, "AP").Value = ws1.Cells(i, col11).Value
    End If
    If col12 > 0 Then
        ws2.Cells(i, "AQ").Value = ws1.Cells(i, col12).Value
    End If
    If col13 > 0 Then
        ws2.Cells(i, "AR").Value = ws1.Cells(i, col13).Value
    End If
    If col14 > 0 Then
        ws2.Cells(i, "AS").Value = ws1.Cells(i, col14).Value
    End If
    If col15 > 0 Then
        ws2.Cells(i, "AT").Value = ws1.Cells(i, col15).Value
    End If
Next
'populates second 15 lesson columns if they exist in the original

col = 0
col1 = 0
col2 = 0
col3 = 0
col4 = 0
col5 = 0
col6 = 0
col7 = 0
col8 = 0
col9 = 0
col10 = 0
col11 = 0
col12 = 0
col13 = 0
col14 = 0
col15 = 0

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "18b - Short Field" Then
    col = rCell.Column
ElseIf rCell.Value = "18b - Soft/Rough" Then
    col2 = rCell.Column
ElseIf rCell.Value = "18c - Overshoot" Then
    col3 = rCell.Column
ElseIf rCell.Value = "19 - First Solo" Then
    col4 = rCell.Column
ElseIf rCell.Value = "20 - Illusions" Then
    col5 = rCell.Column
ElseIf rCell.Value = "21a - Precautionary - On Aerodrome" Then
    col6 = rCell.Column
ElseIf rCell.Value = "21b - Precautionary - Off Aerodrome" Then
    col7 = rCell.Column
ElseIf rCell.Value = "22a - Forced - (Control / Approach)" Then
   col8 = rCell.Column
ElseIf rCell.Value = "22b - Forced - (Cockpit Management)" Then
   col9 = rCell.Column
ElseIf rCell.Value = "23 - Navigation" Then
   col10 = rCell.Column
ElseIf rCell.Value = "23a - Pre-Flight Planning Procedures" Then
   col11 = rCell.Column
ElseIf rCell.Value = "23b - Departure Procedure" Then
   col12 = rCell.Column
ElseIf rCell.Value = "23c - En Route Procedure" Then
   col13 = rCell.Column
ElseIf rCell.Value = "23d - Diversion to an Alternate" Then
   col14 = rCell.Column
ElseIf rCell.Value = "24a - Full Panel" Then
   col15 = rCell.Column
End If
Next
'identifies third 15 lesson columns


For i = 2 To last_row
On Error Resume Next
    If col > 0 Then
        ws2.Cells(i, "AU").Value = ws1.Cells(i, col).Value
    End If
    If col2 > 0 Then
        ws2.Cells(i, "AV").Value = ws1.Cells(i, col2).Value
    End If
    If col3 > 0 Then
        ws2.Cells(i, "AW").Value = ws1.Cells(i, col3).Value
    End If
    If col4 > 0 Then
        ws2.Cells(i, "AX").Value = ws1.Cells(i, col4).Value
    End If
    If col5 > 0 Then
        ws2.Cells(i, "AY").Value = ws1.Cells(i, col5).Value
    End If
    If col6 > 0 Then
        ws2.Cells(i, "AZ").Value = ws1.Cells(i, col6).Value
    End If
    If col7 > 0 Then
        ws2.Cells(i, "BA").Value = ws1.Cells(i, col7).Value
    End If
    If col8 > 0 Then
        ws2.Cells(i, "BB").Value = ws1.Cells(i, col8).Value
    End If
    If col9 > 0 Then
        ws2.Cells(i, "BC").Value = ws1.Cells(i, col9).Value
    End If
    If col10 > 0 Then
        ws2.Cells(i, "BD").Value = ws1.Cells(i, col10).Value
    End If
    If col11 > 0 Then
        ws2.Cells(i, "BE").Value = ws1.Cells(i, col11).Value
    End If
    If col12 > 0 Then
        ws2.Cells(i, "BF").Value = ws1.Cells(i, col12).Value
    End If
    If col13 > 0 Then
        ws2.Cells(i, "BG").Value = ws1.Cells(i, col13).Value
    End If
    If col14 > 0 Then
        ws2.Cells(i, "BH").Value = ws1.Cells(i, col14).Value
    End If
    If col15 > 0 Then
        ws2.Cells(i, "BI").Value = ws1.Cells(i, col15).Value
    End If
Next
'populates third 15 lesson columns if they exist in the original



col = 0
col1 = 0
col2 = 0
col3 = 0
col4 = 0
col5 = 0
col6 = 0

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "24b - Limited Panel" Then
    col = rCell.Column
ElseIf rCell.Value = "24c - Unusual Attitude" Then
    col2 = rCell.Column
ElseIf rCell.Value = "24d - Radio Navigation" Then
    col3 = rCell.Column
ElseIf rCell.Value = "29 - Emergencies" Then
    col4 = rCell.Column
ElseIf rCell.Value = "30 - Radio" Then
    col5 = rCell.Column
ElseIf rCell.Value = "Instructor Remarks" Then
    col6 = rCell.Column
End If
Next
'identifies remaining 5 lesson columns and instructor remarks

For i = 2 To last_row
On Error Resume Next
    If col > 0 Then
        ws2.Cells(i, "BJ").Value = ws1.Cells(i, col).Value
    End If
    If col2 > 0 Then
        ws2.Cells(i, "BK").Value = ws1.Cells(i, col2).Value
    End If
    If col3 > 0 Then
        ws2.Cells(i, "BL").Value = ws1.Cells(i, col3).Value
    End If
    If col4 > 0 Then
        ws2.Cells(i, "BM").Value = ws1.Cells(i, col4).Value
    End If
    If col5 > 0 Then
        ws2.Cells(i, "BN").Value = ws1.Cells(i, col5).Value
    End If
    If col6 > 0 Then
        ws2.Cells(i, "BQ").Value = ws1.Cells(i, col6).Value
    End If
Next
'populates remaining 5 lesson columns

ws2.PageSetup.Orientation = xlLandscape
ws2.PageSetup.PaperSize = xlPaperTabloid
'sets page to landscape tabloid size

last_row = (ws2.Cells(Rows.Count, "A").End(xlUp).Row)
col = 69

ws2.Columns("BP").PageBreak = xlPageBreakManual
ws2.Columns("BR").PageBreak = xlPageBreakManual
'adds the two page breaks

With ws2.PageSetup
 .LeftMargin = Application.InchesToPoints(0.25)
 .RightMargin = Application.InchesToPoints(0.25)
 .TopMargin = Application.InchesToPoints(0.75)
 .BottomMargin = Application.InchesToPoints(0.75)
 .HeaderMargin = Application.InchesToPoints(0.3)
 .FooterMargin = Application.InchesToPoints(0.3)
 .PrintArea = ws2.Range(Cells(1, 1), Cells(last_row, col))
 .CenterHeader = "&B&10" & name_header
 .LeftHeader = "&B&10" & "Pilot Training Record (Private Pilot Licence)"
 .RightHeader = "&B&10" & "Rockcliffe Flying Club (0195)"
 .LeftFooter = "&11" & "Page " & "&P" & " of " & "&N"
 .RightFooter = "&11" & "Printed on " & "&D"
 .PrintTitleRows = "$1:$1"
End With
'setting print area, margins, headers, and footers

With ws2.Range(Cells(1, 1), Cells(last_row, 1))
 .Cells.ColumnWidth = 8
End With
'column A

With ws2.Range(Cells(1, 2), Cells(last_row, 3))
 .Cells.ColumnWidth = 6
End With
'columns B-C

With ws2.Range(Cells(1, 67), Cells(last_row, 68))
 .Cells.ColumnWidth = 8
End With
'columns BO-BP

With ws2.Range(Cells(1, 4), Cells(last_row, 5))
    .Cells.ColumnWidth = 4.3
End With
'columns D-E

With ws2.Range(Cells(1, 6), Cells(last_row, 11))
    .Cells.ColumnWidth = 3
End With
'columns F-K

With ws2.Range(Cells(1, 12), Cells(last_row, 15))
    .Cells.ColumnWidth = 4
End With
'columns L-O

With ws2.Range(Cells(1, 16), Cells(last_row, 66))
    .Cells.ColumnWidth = 2
End With
'columns P-BN

With ws2.Range(Cells(1, 69), Cells(last_row, 69))
    .Cells.ColumnWidth = 180
End With
'last column (remarks)

With ws2.Range(Cells(1, 1), Cells(last_row, 69))
    .HorizontalAlignment = xlCenter
    .Borders.LineStyle = xlContinuous
    .Font.Size = 8
    .Rows.AutoFit
 .WrapText = True
End With
'sets header at angle

With ws2.Range(Cells(1, 1), Cells(1, 69))
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = True
    .Orientation = 45
    .WrapText = False
    .LineStyle = xlContinuous
    .Weight = xlThin
    .ColorIndex = 3
    .Range(.Cells(last_row + 1, "D"), .Cells(last_row + 1, "E")).Merge
    .Range(.Cells(last_row + 2, "D"), .Cells(last_row + 2, "E")).Merge
    .Range(.Cells(last_row + 3, "D"), .Cells(last_row + 3, "E")).Merge
End With
'wrap text and trying to add border to everything but that's not working

With ws2.Range(Cells(last_row + 1, 4), Cells(last_row + 3, "K"))
    .Borders.LineStyle = xlContinuous
End With
'adds thin lines to totals sections

With ws2.Range(Cells(1, 4), Cells(last_row + 3, 7))
.Borders(xlEdgeRight).LineStyle = xlContinuous
.Borders(xlEdgeRight).Weight = xlThick
.Borders(xlEdgeLeft).LineStyle = xlContinuous
.Borders(xlEdgeLeft).Weight = xlThick
End With
'adds thick borders

With ws2.Range(Cells(1, 10), Cells(last_row + 3, 14))
.Borders(xlEdgeRight).LineStyle = xlContinuous
.Borders(xlEdgeRight).Weight = xlThick
.Borders(xlEdgeLeft).LineStyle = xlContinuous
.Borders(xlEdgeLeft).Weight = xlThick
End With
'adds thick borders

ws2.Cells(last_row + 1, "D").Value = "Subtotals:"
ws2.Cells(last_row + 2, "D").Value = "Forwarded:"
ws2.Cells(last_row + 3, "D").Value = "Totals:"
ws2.Cells(last_row + 3, "A").Value = "Grand Total:"
ws2.Range("F" & last_row + 2, "K" & last_row + 2).Value = 0
ws2.Cells(last_row + 1, "M").Value = "All times"
ws2.Cells(last_row + 2, "M").Value = "Certified"
ws2.Cells(last_row + 3, "M").Value = "Correct:"
ws2.Cells(last_row + 3, "Q").Value = "CFI or Delegate"
ws2.Cells(last_row + 2, "Q").Value = "______________________________________________"
ws2.Cells(last_row + 3, "AD").Value = "Date"
ws2.Cells(last_row + 2, "AD").Value = "_______________________________"
ws2.Cells(last_row + 3, "AN").Value = "Student"
ws2.Cells(last_row + 2, "AN").Value = "______________________________________________"
ws2.Cells(last_row + 3, "BB").Value = "Date"
ws2.Cells(last_row + 2, "BB").Value = "_______________________________"

ws2.Cells(last_row + 1, "F").Formula = "=sum(F2:F" & last_row & ")"
ws2.Cells(last_row + 3, "F").Formula = "=sum(F" & (last_row + 1) & ":F" & (last_row + 2) & ")"

ws2.Cells(last_row + 1, "G").Formula = "=sum(G2:G" & last_row & ")"
ws2.Cells(last_row + 3, "G").Formula = "=sum(G" & (last_row + 1) & ":G" & (last_row + 2) & ")"

ws2.Cells(last_row + 1, "H").Formula = "=sum(H2:H" & last_row & ")"
ws2.Cells(last_row + 3, "H").Formula = "=sum(H" & (last_row + 1) & ":H" & (last_row + 2) & ")"

ws2.Cells(last_row + 1, "I").Formula = "=sum(I2:I" & last_row & ")"
ws2.Cells(last_row + 3, "I").Formula = "=sum(I" & (last_row + 1) & ":I" & (last_row + 2) & ")"

ws2.Cells(last_row + 1, "J").Formula = "=sum(J2:J" & last_row & ")"
ws2.Cells(last_row + 3, "J").Formula = "=sum(J" & (last_row + 1) & ":J" & (last_row + 2) & ")"

ws2.Cells(last_row + 1, "K").Formula = "=sum(K2:K" & last_row & ")"
ws2.Cells(last_row + 3, "K").Formula = "=sum(K" & (last_row + 1) & ":K" & (last_row + 2) & ")"

ws2.Cells(last_row + 3, "C").Formula = "=sum(F" & (last_row + 3) & ":G" & (last_row + 3) & ")"

With ws2.Range(Cells(last_row + 1, 1), Cells(last_row + 3, 69))
 .Font.Size = 8
 .Rows.AutoFit
 .WrapText = False
End With

Worksheets("Sheet1").Range("A1:A5").Font.Bold = True
ws2.Range(Cells(last_row + 3, 1), Cells(last_row + 3, 69)).Font.Bold = True
ws2.Range(Cells(last_row + 1, "M"), Cells(last_row + 2, "M")).Font.Bold = True
ws2.Range(Cells(1, 1), Cells(1, 69)).Font.Bold = True
 
ws2.Range("BQ:BQ").Cut Range("BR:BR")
ws2.Range("O:O").Cut Range("BQ:BQ")
ws2.Range("O:O").Delete

End Sub


Sub Rockliffe_Report_CPL(ByRef control As Office.IRibbonControl)
 Dim ws1 As Worksheet
 Dim ws2 As Worksheet
 Dim i As Integer
 Dim split_array As Variant
 Dim split_array2 As Variant
 Dim r As Range
 Dim rCell As Range
 Dim col As Integer
 Dim col2 As Integer
 Dim col3 As Integer
 Dim col4 As Integer
 Dim col5 As Integer
 Dim col6 As Integer
 Dim col7 As Integer
 Dim col8, col9, col10, col11, col12, col13, col14, col15, col16 As Integer
 Dim fly_total As Single
 Dim temp1 As Single
Dim temp2 As Single
Dim z As Integer
Dim tempstring As String
Dim header_array As Variant
Dim last_row As Integer
Dim name_header As String
Set ws1 = ActiveSheet

Set r = ws1.Range(Cells(1, 1), Cells(1, ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column))
'sets the Row 1 of the original worksheet to be range "r"
last_row = (ws1.Cells(Rows.Count, "A").End(xlUp).Row) - 1
    With ActiveWorkbook
        Set ws2 = .Sheets.Add(After:=.Sheets(.Sheets.Count))
        ws2.Name = "Report"
    End With
'creates a new sheet and names it Report

ws2.Cells(1, "A") = "Date"
ws2.Cells(1, "B") = "Student"
ws2.Cells(1, "C") = "Instructor"
ws2.Cells(1, "D") = "Aircraft Type"
ws2.Cells(1, "E") = "Tailnumber"
ws2.Cells(1, "F") = "Dual Day Total"
ws2.Cells(1, "G") = "Dual Night Total"
ws2.Cells(1, "H") = "Solo Day Total"
ws2.Cells(1, "I") = "Solo Night Total"
ws2.Cells(1, "J") = "Instrument (Hood)"
ws2.Cells(1, "K") = "Flight Sim (Instrument)"
ws2.Cells(1, "L") = "Dual Day XC"
ws2.Cells(1, "M") = "Dual Night XC"
ws2.Cells(1, "N") = "Solo Day XC"
ws2.Cells(1, "O") = "Solo Night XC"
ws2.Cells(1, "P") = "From Airport"
ws2.Cells(1, "Q") = "Via Airports"
ws2.Cells(1, "R") = "To Airport"
ws2.Cells(1, "S") = "Lesson"
ws2.Cells(1, "T") = "9s - Steep Turn"
ws2.Cells(1, "U") = "10 - Range & Endurance"
ws2.Cells(1, "V") = "11 - Slow Flight"
ws2.Cells(1, "W") = "12a - Power-Off Stall"
ws2.Cells(1, "X") = "12b - Power-On Stall"
ws2.Cells(1, "Y") = "13 - Spin"
ws2.Cells(1, "Z") = "14 - Spiral"
ws2.Cells(1, "AA") = "15 - Slipping"
ws2.Cells(1, "AB") = "16a - Normal Takeoff"
ws2.Cells(1, "AC") = "16b - Crosswind"
ws2.Cells(1, "AD") = "16b - Obstacle"
ws2.Cells(1, "AE") = "16b - Short/Minimum Run"
ws2.Cells(1, "AF") = "16b - Soft/Rough"
ws2.Cells(1, "AG") = "17 - Circuit"
ws2.Cells(1, "AH") = "18a - 180 Power Off"
ws2.Cells(1, "AI") = "18a - Normal Landing"
ws2.Cells(1, "AJ") = "18b - Crosswind"
ws2.Cells(1, "AK") = "18b - Obstacle"
ws2.Cells(1, "AL") = "18b - Short Field"
ws2.Cells(1, "AM") = "18b - Soft/Rough"
ws2.Cells(1, "AN") = "18c - Overshoot"
ws2.Cells(1, "AO") = "19 - First Solo"
ws2.Cells(1, "AP") = "20 - Illusions"
ws2.Cells(1, "AQ") = "21a - Precautionary - On Aerodrome"
ws2.Cells(1, "AR") = "21b - Precautionary - Off Aerodrome"
ws2.Cells(1, "AS") = "22a - Forced - (Control / Approach)"
ws2.Cells(1, "AT") = "22b - Forced - (Cockpit Management)"
ws2.Cells(1, "AU") = "23 - Navigation"
ws2.Cells(1, "AV") = "23a - Pre-Flight Planning Procedures"
ws2.Cells(1, "AW") = "23b - Departure Procedure"
ws2.Cells(1, "AX") = "23c - En Route Procedure"
ws2.Cells(1, "AY") = "23d - Diversion to an Alternate"
ws2.Cells(1, "AZ") = "24a - Full Panel"
ws2.Cells(1, "BA") = "24b - Limited Panel"
ws2.Cells(1, "BB") = "24c - Unusual Attitude"
ws2.Cells(1, "BC") = "24d - Radio Navigation"
ws2.Cells(1, "BD") = "29 - Emergencies"
ws2.Cells(1, "BE") = "30 - Radio"
ws2.Cells(1, "BF") = ""
ws2.Cells(1, "BG") = "Date"
ws2.Cells(1, "BH") = "Comments"
'prints the new headers in the report

For i = 2 To last_row
On Error Resume Next
split_array = Split(ws1.Cells(i, "A").Value, " ", 2)
ws2.Cells(i, "A").Value = split_array(0)
ws2.Cells(i, "BG").Value = split_array(0)
Next
'prints the date without the time stamp


For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Student" Then
    col = rCell.Column
End If
Next
name_header = ws1.Cells(2, col)
'finds the Student column in the original document

For i = 2 To last_row
On Error Resume Next
split_array = Split(ws1.Cells(i, col).Value, " ", 2)
ws2.Cells(i, "B").Value = split_array(1)
Next
'prints last name only of student

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Instructor" Then
    col = rCell.Column
End If
Next
'finds which cell has instructor name

For i = 2 To last_row
On Error Resume Next
split_array = Split(ws1.Cells(i, col).Value, " ", 2)
ws2.Cells(i, "C").Value = split_array(1)
Next
'prints last name only of student

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Aircraft Type" Then
    col = rCell.Column
End If
Next
'finds which cell has Aircraft type

For i = 2 To last_row
On Error Resume Next
    If InStr(ws1.Cells(i, col).Value, "172") > 0 Then
        ws2.Cells(i, "D").Value = "C172"
    ElseIf InStr(ws1.Cells(i, col).Value, "150") > 0 Then
        ws2.Cells(i, "D").Value = "C150"
    ElseIf InStr(ws1.Cells(i, col).Value, "edbird") > 0 Then
        ws2.Cells(i, "D").Value = "RB"
    ElseIf InStr(ws1.Cells(i, col).Value, "iamond") > 0 Then
        ws2.Cells(i, "D").Value = "DA20"
    Else
        ws2.Cells(i, "D").Value = ws1.Cells(i, col).Value
    End If
Next
'prints abbreviated aircraft type; if the type does not match one listed above, it will be printed as is

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Tailnumber" Then
    col = rCell.Column
End If
Next
'finds which cell has Tailnumber (or Registration Type)

For i = 2 To last_row
On Error Resume Next
    If InStr(ws1.Cells(i, col).Value, "4073") > 0 Then
        ws2.Cells(i, "E").Value = "4073"
    ElseIf InStr(ws1.Cells(i, col).Value, "GABC") > 0 Then
        ws2.Cells(i, "E").Interior.ColorIndex = 8
        ws2.Cells(i, "E").Value = ws1.Cells(i, col).Value
    Else
        If InStr(ws1.Cells(i, col).Value, "-") > 0 Then
        split_array = Split(ws1.Cells(i, col).Value, "-", 2)
        ws2.Cells(i, "E").Value = split_array(1)
        Else
        ws2.Cells(i, "E").Value = ws1.Cells(i, col).Value
        End If
    End If
Next
'prints abbreviated Tailnumbers - if no match found, prints original as is

ActiveWorkbook.PrecisionAsDisplayed = True
'avoids the space-time continuum problem.  Makes sure decimals are counted as shown

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Dual Day Local" Then
    col = rCell.Column
ElseIf rCell.Value = "Dual Day XC" Then
    col2 = rCell.Column
End If
Next
'finds which cell has Dual Day Local and Dual Day XC

For i = 2 To last_row
On Error Resume Next
temp1 = ws1.Cells(i, col).Value
temp2 = ws1.Cells(i, col2).Value
fly_total = temp1 + temp2
ws2.Cells(i, "F").Value = fly_total
ws2.Cells(i, "F").NumberFormat = "0.0"
fly_total = 0
temp1 = 0
temp2 = 0
Next
'adds together dual day local and dual day xc then prints them with only one decimal place

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Solo Day Local" Then
    col = rCell.Column
ElseIf rCell.Value = "Solo Day XC" Then
    col2 = rCell.Column
End If
Next
'finds which cell has Solo Day Local and Solo Day XC

For i = 2 To last_row
On Error Resume Next
temp1 = ws1.Cells(i, col).Value
temp2 = ws1.Cells(i, col2).Value
fly_total = temp1 + temp2
ws2.Cells(i, "H").Value = fly_total
ws2.Cells(i, "H").NumberFormat = "0.0"
fly_total = 0
temp1 = 0
temp2 = 0
Next
'adds together Solo Day Local and Solo Day XC

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Solo Night Local" Then
    col = rCell.Column
ElseIf rCell.Value = "Solo Night XC" Then
    col2 = rCell.Column
End If
Next
'finds which cell has Solo Night Local and Solo Night XC

For i = 2 To last_row
On Error Resume Next
temp1 = ws1.Cells(i, col).Value
temp2 = ws1.Cells(i, col2).Value
fly_total = temp1 + temp2
ws2.Cells(i, "I").Value = fly_total
ws2.Cells(i, "I").NumberFormat = "0.0"
fly_total = 0
temp1 = 0
temp2 = 0
Next
'adds together Solo Night Local and Solo Night XC

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Dual Night Local" Then
    col = rCell.Column
ElseIf rCell.Value = "Dual Night XC" Then
    col2 = rCell.Column
End If
Next
'finds which cell has Dual Night Local and Dual XC

For i = 2 To last_row
On Error Resume Next
temp1 = ws1.Cells(i, col).Value
temp2 = ws1.Cells(i, col2).Value
fly_total = temp1 + temp2
ws2.Cells(i, "G").Value = fly_total
ws2.Cells(i, "G").NumberFormat = "0.0"
fly_total = 0
temp1 = 0
temp2 = 0
Next
'adds together Dual Night Local and Dual Night XC

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Instrument (Hood)" Then
    col = rCell.Column
ElseIf rCell.Value = "Flight Sim (Instrument)" Then
    col2 = rCell.Column
ElseIf rCell.Value = "Dual Day XC" Then
    col3 = rCell.Column
ElseIf rCell.Value = "Dual Night XC" Then
    col4 = rCell.Column
ElseIf rCell.Value = "Solo Day XC" Then
    col5 = rCell.Column
ElseIf rCell.Value = "Solo Night XC" Then
    col6 = rCell.Column
ElseIf rCell.Value = "From Airport" Then
    col7 = rCell.Column
ElseIf rCell.Value = "Via Airports" Then
    col8 = rCell.Column
ElseIf rCell.Value = "To Airport" Then
    col9 = rCell.Column
End If
Next
'identifies 9 columns we will print as is

For i = 2 To last_row
On Error Resume Next
ws2.Cells(i, "J").NumberFormat = "0.0"
ws2.Cells(i, "J").Value = ws1.Cells(i, col).Value
ws2.Cells(i, "K").NumberFormat = "0.0"
ws2.Cells(i, "K").Value = ws1.Cells(i, col2).Value
ws2.Cells(i, "L").NumberFormat = "0.0"
ws2.Cells(i, "L").Value = ws1.Cells(i, col3).Value
ws2.Cells(i, "M").NumberFormat = "0.0"
ws2.Cells(i, "M").Value = ws1.Cells(i, col4).Value
ws2.Cells(i, "N").NumberFormat = "0.0"
ws2.Cells(i, "N").Value = ws1.Cells(i, col5).Value
ws2.Cells(i, "O").NumberFormat = "0.0"
ws2.Cells(i, "O").Value = ws1.Cells(i, col6).Value
ws2.Cells(i, "P").Value = ws1.Cells(i, col7).Value
ws2.Cells(i, "Q").Value = ws1.Cells(i, col8).Value
ws2.Cells(i, "R").Value = ws1.Cells(i, col9).Value
Next
'prints these 9 columns with no changes except formatting the numbers to one decimal place

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "Lesson" Then
    col = rCell.Column
End If
Next
'finds which cell has Lesson

For i = 2 To last_row
On Error Resume Next
    If InStr(ws1.Cells(i, col).Value, " ") > 0 Then
    split_array = Split(ws1.Cells(i, col).Value, "-", 2)
        If IsNumeric(split_array(0)) Then
            split_array2 = Split(ws1.Cells(i, col).Value, " ", 2)
            ws2.Cells(i, "S").NumberFormat = "@"
            ws2.Cells(i, "S").Value = split_array2(0)
        Else
            ws2.Cells(i, "S").NumberFormat = "@"
            ws2.Cells(i, "S").Value = ws1.Cells(i, col).Value
        End If
    Else
    ws2.Cells(i, "S").NumberFormat = "@"
    ws2.Cells(i, "S").Value = ws1.Cells(i, col).Value
    End If
Next
'prints lesson information with numbers only (if they exist) and formatted as text

col = 0
col1 = 0
col2 = 0
col3 = 0
col4 = 0
col5 = 0
col6 = 0
col7 = 0
col8 = 0
col9 = 0
col10 = 0
col11 = 0
col12 = 0
col13 = 0
col14 = 0
col15 = 0
col16 = 0

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "9s - Steep Turn" Then
   col14 = rCell.Column
ElseIf rCell.Value = "10 - Range & Endurance" Then
   col15 = rCell.Column
ElseIf rCell.Value = "11 - Slow Flight" Then
    col16 = rCell.Column
End If
Next
'identifies first 3 lesson columns

For i = 2 To last_row
On Error Resume Next
    If col14 > 0 Then
        ws2.Cells(i, "T").Value = ws1.Cells(i, col14).Value
    End If
    If col15 > 0 Then
        ws2.Cells(i, "U").Value = ws1.Cells(i, col15).Value
    End If
    If col16 > 0 Then
        ws2.Cells(i, "V").Value = ws1.Cells(i, col16).Value
    End If
Next
'populates first 3 lesson columns if they exist in the original

col = 0
col1 = 0
col2 = 0
col3 = 0
col4 = 0
col5 = 0
col6 = 0
col7 = 0
col8 = 0
col9 = 0
col10 = 0
col11 = 0
col12 = 0
col13 = 0
col14 = 0
col15 = 0

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "12a - Power-Off Stall" Then
    col = rCell.Column
ElseIf rCell.Value = "12b - Power-On Stall" Then
    col2 = rCell.Column
ElseIf rCell.Value = "13 - Spin" Then
    col3 = rCell.Column
ElseIf rCell.Value = "14 - Spiral" Then
    col4 = rCell.Column
ElseIf rCell.Value = "15 - Slipping" Then
    col5 = rCell.Column
ElseIf rCell.Value = "16a - Normal Takeoff" Then
    col6 = rCell.Column
ElseIf rCell.Value = "16b - Crosswind" Then
    col7 = rCell.Column
ElseIf rCell.Value = "16b - Obstacle" Then
   col8 = rCell.Column
ElseIf rCell.Value = "16b - Short/Minimum Run" Then
   col9 = rCell.Column
ElseIf rCell.Value = "16b - Soft/Rough" Then
   col10 = rCell.Column
ElseIf rCell.Value = "17 - Circuit" Then
   col11 = rCell.Column
ElseIf rCell.Value = "18a - 180 Power Off" Then
   col12 = rCell.Column
ElseIf rCell.Value = "18a - Normal Landing" Then
   col13 = rCell.Column
ElseIf rCell.Value = "18b - Crosswind" Then
   col14 = rCell.Column
ElseIf rCell.Value = "18b - Obstacle" Then
   col15 = rCell.Column
End If
Next
'identifies second 15 lesson columns

For i = 2 To last_row
On Error Resume Next
    If col > 0 Then
        ws2.Cells(i, "W").Value = ws1.Cells(i, col).Value
    End If
    If col2 > 0 Then
        ws2.Cells(i, "X").Value = ws1.Cells(i, col2).Value
    End If
    If col3 > 0 Then
        ws2.Cells(i, "Y").Value = ws1.Cells(i, col3).Value
    End If
    If col4 > 0 Then
        ws2.Cells(i, "Z").Value = ws1.Cells(i, col4).Value
    End If
    If col5 > 0 Then
        ws2.Cells(i, "AA").Value = ws1.Cells(i, col5).Value
    End If
    If col6 > 0 Then
        ws2.Cells(i, "AB").Value = ws1.Cells(i, col6).Value
    End If
    If col7 > 0 Then
        ws2.Cells(i, "AC").Value = ws1.Cells(i, col7).Value
    End If
    If col8 > 0 Then
        ws2.Cells(i, "AD").Value = ws1.Cells(i, col8).Value
    End If
    If col9 > 0 Then
        ws2.Cells(i, "AE").Value = ws1.Cells(i, col9).Value
    End If
    If col10 > 0 Then
        ws2.Cells(i, "AF").Value = ws1.Cells(i, col10).Value
    End If
    If col11 > 0 Then
        ws2.Cells(i, "AG").Value = ws1.Cells(i, col11).Value
    End If
    If col12 > 0 Then
        ws2.Cells(i, "AH").Value = ws1.Cells(i, col12).Value
    End If
    If col13 > 0 Then
        ws2.Cells(i, "AI").Value = ws1.Cells(i, col13).Value
    End If
    If col14 > 0 Then
        ws2.Cells(i, "AJ").Value = ws1.Cells(i, col14).Value
    End If
    If col15 > 0 Then
        ws2.Cells(i, "AK").Value = ws1.Cells(i, col15).Value
    End If
Next
'populates second 15 lesson columns if they exist in the original

col = 0
col1 = 0
col2 = 0
col3 = 0
col4 = 0
col5 = 0
col6 = 0
col7 = 0
col8 = 0
col9 = 0
col10 = 0
col11 = 0
col12 = 0
col13 = 0
col14 = 0
col15 = 0

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "18b - Short Field" Then
    col = rCell.Column
ElseIf rCell.Value = "18b - Soft/Rough" Then
    col2 = rCell.Column
ElseIf rCell.Value = "18c - Overshoot" Then
    col3 = rCell.Column
ElseIf rCell.Value = "19 - First Solo" Then
    col4 = rCell.Column
ElseIf rCell.Value = "20 - Illusions" Then
    col5 = rCell.Column
ElseIf rCell.Value = "21a - Precautionary - On Aerodrome" Then
    col6 = rCell.Column
ElseIf rCell.Value = "21b - Precautionary - Off Aerodrome" Then
    col7 = rCell.Column
ElseIf rCell.Value = "22a - Forced - (Control / Approach)" Then
   col8 = rCell.Column
ElseIf rCell.Value = "22b - Forced - (Cockpit Management)" Then
   col9 = rCell.Column
ElseIf rCell.Value = "23 - Navigation" Then
   col10 = rCell.Column
ElseIf rCell.Value = "23a - Pre-Flight Planning Procedures" Then
   col11 = rCell.Column
ElseIf rCell.Value = "23b - Departure Procedure" Then
   col12 = rCell.Column
ElseIf rCell.Value = "23c - En Route Procedure" Then
   col13 = rCell.Column
ElseIf rCell.Value = "23d - Diversion to an Alternate" Then
   col14 = rCell.Column
ElseIf rCell.Value = "24a - Full Panel" Then
   col15 = rCell.Column
End If
Next
'identifies third 15 lesson columns


For i = 2 To last_row
On Error Resume Next
    If col > 0 Then
        ws2.Cells(i, "AL").Value = ws1.Cells(i, col).Value
    End If
    If col2 > 0 Then
        ws2.Cells(i, "AM").Value = ws1.Cells(i, col2).Value
    End If
    If col3 > 0 Then
        ws2.Cells(i, "AN").Value = ws1.Cells(i, col3).Value
    End If
    If col4 > 0 Then
        ws2.Cells(i, "AO").Value = ws1.Cells(i, col4).Value
    End If
    If col5 > 0 Then
        ws2.Cells(i, "AP").Value = ws1.Cells(i, col5).Value
    End If
    If col6 > 0 Then
        ws2.Cells(i, "AQ").Value = ws1.Cells(i, col6).Value
    End If
    If col7 > 0 Then
        ws2.Cells(i, "AR").Value = ws1.Cells(i, col7).Value
    End If
    If col8 > 0 Then
        ws2.Cells(i, "AS").Value = ws1.Cells(i, col8).Value
    End If
    If col9 > 0 Then
        ws2.Cells(i, "AT").Value = ws1.Cells(i, col9).Value
    End If
    If col10 > 0 Then
        ws2.Cells(i, "AU").Value = ws1.Cells(i, col10).Value
    End If
    If col11 > 0 Then
        ws2.Cells(i, "AV").Value = ws1.Cells(i, col11).Value
    End If
    If col12 > 0 Then
        ws2.Cells(i, "AW").Value = ws1.Cells(i, col12).Value
    End If
    If col13 > 0 Then
        ws2.Cells(i, "AX").Value = ws1.Cells(i, col13).Value
    End If
    If col14 > 0 Then
        ws2.Cells(i, "AY").Value = ws1.Cells(i, col14).Value
    End If
    If col15 > 0 Then
        ws2.Cells(i, "AZ").Value = ws1.Cells(i, col15).Value
    End If
Next
'populates third 15 lesson columns if they exist in the original



col = 0
col1 = 0
col2 = 0
col3 = 0
col4 = 0
col5 = 0
col6 = 0

For Each rCell In r.Cells
On Error Resume Next
If rCell.Value = "24b - Limited Panel" Then
    col = rCell.Column
ElseIf rCell.Value = "24c - Unusual Attitude" Then
    col2 = rCell.Column
ElseIf rCell.Value = "24d - Radio Navigation" Then
    col3 = rCell.Column
ElseIf rCell.Value = "29 - Emergencies" Then
    col4 = rCell.Column
ElseIf rCell.Value = "30 - Radio" Then
    col5 = rCell.Column
ElseIf rCell.Value = "Instructor Remarks" Then
    col6 = rCell.Column
End If
Next
'identifies remaining 5 lesson columns and instructor remarks

For i = 2 To last_row
On Error Resume Next
    If col > 0 Then
        ws2.Cells(i, "BA").Value = ws1.Cells(i, col).Value
    End If
    If col2 > 0 Then
        ws2.Cells(i, "BB").Value = ws1.Cells(i, col2).Value
    End If
    If col3 > 0 Then
        ws2.Cells(i, "BC").Value = ws1.Cells(i, col3).Value
    End If
    If col4 > 0 Then
        ws2.Cells(i, "BD").Value = ws1.Cells(i, col4).Value
    End If
    If col5 > 0 Then
        ws2.Cells(i, "BE").Value = ws1.Cells(i, col5).Value
    End If
    If col6 > 0 Then
        ws2.Cells(i, "BH").Value = ws1.Cells(i, col6).Value
    End If
Next
'populates remaining 5 lesson columns

ws2.PageSetup.Orientation = xlLandscape
ws2.PageSetup.PaperSize = xlPaperTabloid
'sets page to landscape tabloid size

last_row = (ws2.Cells(Rows.Count, "A").End(xlUp).Row)
col = 60

ws2.Columns("BG").PageBreak = xlPageBreakManual
ws2.Columns("BI").PageBreak = xlPageBreakManual
'adds the two page breaks

With ws2.PageSetup
 .LeftMargin = Application.InchesToPoints(0.25)
 .RightMargin = Application.InchesToPoints(0.25)
 .TopMargin = Application.InchesToPoints(0.75)
 .BottomMargin = Application.InchesToPoints(0.75)
 .HeaderMargin = Application.InchesToPoints(0.3)
 .FooterMargin = Application.InchesToPoints(0.3)
 .PrintArea = ws2.Range(Cells(1, 1), Cells(last_row, col))
 .CenterHeader = "&B&10" & name_header
 .LeftHeader = "&B&10" & "Pilot Training Record (Commercial Pilot Licence)"
 .RightHeader = "&B&10" & "Rockcliffe Flying Club (0195)"
 .LeftFooter = "&11" & "Page " & "&P" & " of " & "&N"
 .RightFooter = "&11" & "Printed on " & "&D"
 .PrintTitleRows = "$1:$1"
End With
'setting print area, margins, headers, and footers

With ws2.Range(Cells(1, 1), Cells(last_row, 1))
 .Cells.ColumnWidth = 8
End With
'column A

With ws2.Range(Cells(1, 2), Cells(last_row, 3))
 .Cells.ColumnWidth = 6.8
End With
'columns B-C

With ws2.Range(Cells(1, 58), Cells(last_row, 59))
 .Cells.ColumnWidth = 8
End With
'columns BF-BG

With ws2.Range(Cells(1, 4), Cells(last_row, 5))
    .Cells.ColumnWidth = 4.5
End With
'columns D-E

With ws2.Range(Cells(1, 6), Cells(last_row, 15))
    .Cells.ColumnWidth = 3.67
End With
'columns F-O

With ws2.Range(Cells(1, 16), Cells(last_row, 19))
    .Cells.ColumnWidth = 4
End With
'columns P-S

With ws2.Range(Cells(1, 20), Cells(last_row, 57))
    .Cells.ColumnWidth = 2
End With
'columns T-BF

With ws2.Range(Cells(1, 60), Cells(last_row, 60))
    .Cells.ColumnWidth = 160
End With
'last column (remarks)

With ws2.Range(Cells(1, 1), Cells(last_row, 60))
    .HorizontalAlignment = xlCenter
    .Borders.LineStyle = xlContinuous
    .Font.Size = 8
    .Rows.AutoFit
 .WrapText = True
End With
'sets header at angle

With ws2.Range(Cells(1, 1), Cells(1, 60))
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = True
    .Orientation = 45
    .WrapText = False
    .LineStyle = xlContinuous
    .Weight = xlThin
    .Font.Bold = True
    .ColorIndex = 3
    .Range(.Cells(last_row + 1, "D"), .Cells(last_row + 1, "E")).Merge
    .Range(.Cells(last_row + 2, "D"), .Cells(last_row + 2, "E")).Merge
    .Range(.Cells(last_row + 3, "D"), .Cells(last_row + 3, "E")).Merge
End With
'wrap text and trying to add border to everything but that's not working

With ws2.Range(Cells(last_row + 1, 4), Cells(last_row + 3, "O"))
    .Borders.LineStyle = xlContinuous
End With
'adds thin lines to totals sections

With ws2.Range(Cells(1, 6), Cells(last_row + 5, 9))
.Borders(xlEdgeRight).LineStyle = xlContinuous
.Borders(xlEdgeRight).Weight = xlThick
.Borders(xlEdgeLeft).LineStyle = xlContinuous
.Borders(xlEdgeLeft).Weight = xlThick
End With
'adds thick borders

With ws2.Range(Cells(1, 12), Cells(last_row, 18))
.Borders(xlEdgeRight).LineStyle = xlContinuous
.Borders(xlEdgeRight).Weight = xlThick
.Borders(xlEdgeLeft).LineStyle = xlContinuous
.Borders(xlEdgeLeft).Weight = xlThick
End With
'adds thick borders

With ws2.Range(Cells(last_row + 1, 12), Cells(last_row + 5, 15))
.Borders(xlEdgeRight).LineStyle = xlContinuous
.Borders(xlEdgeRight).Weight = xlThin
.Borders(xlEdgeLeft).LineStyle = xlContinuous
.Borders(xlEdgeLeft).Weight = xlThick
End With

With ws2.Range(Cells(last_row + 4, 8), Cells(last_row + 5, 13))
.Borders(xlEdgeRight).LineStyle = xlContinuous
.Borders(xlEdgeRight).Weight = xlThin
.Borders(xlEdgeLeft).LineStyle = xlContinuous
.Borders(xlEdgeLeft).Weight = xlThin
End With
'adds thin borders in small section of totals section

With ws2.Range(Cells(last_row + 5, 6), Cells(last_row + 5, 15))
.Borders(xlEdgeBottom).LineStyle = xlContinuous
.Borders(xlEdgeBottom).Weight = xlThin
.Borders(xlEdgeTop).LineStyle = xlContinuous
.Borders(xlEdgeTop).Weight = xlThin
End With
'adds thin borders in small section of totals section

ws2.Cells(last_row + 1, "D").Value = "Subtotals:"
ws2.Cells(last_row + 2, "D").Value = "Forwarded:"
ws2.Cells(last_row + 3, "D").Value = "Totals:"
ws2.Cells(last_row + 1, "A").Value = "Grand Total:"

ws2.Cells(last_row + 4, "F").Value = "Dual Total:"
ws2.Cells(last_row + 4, "H").Value = "Solo Total:"
ws2.Cells(last_row + 4, "J").Value = "Inst Total:"
ws2.Cells(last_row + 4, "L").Value = "Dual XC:"
ws2.Cells(last_row + 4, "N").Value = "Solo XC:"

ws2.Range("F" & last_row + 2, "O" & last_row + 2).Value = 0
ws2.Cells(last_row + 1, "Q").Value = "All times"
ws2.Cells(last_row + 2, "Q").Value = "Certified"
ws2.Cells(last_row + 3, "Q").Value = "Correct:"
ws2.Cells(last_row + 3, "S").Value = "CFI or Delegate"
ws2.Cells(last_row + 2, "S").Value = "________________________________________________________"
ws2.Cells(last_row + 3, "AD").Value = "Date"
ws2.Cells(last_row + 2, "AD").Value = "________________________________________________________"
ws2.Cells(last_row + 3, "AL").Value = "Student"
ws2.Cells(last_row + 2, "AL").Value = "________________________________________________________"

ws2.Cells(last_row + 1, "F").Formula = "=sum(F2:F" & last_row & ")"
ws2.Cells(last_row + 3, "F").Formula = "=sum(F" & (last_row + 1) & ":F" & (last_row + 2) & ")"
'formula for dual day total

ws2.Cells(last_row + 1, "G").Formula = "=sum(G2:G" & last_row & ")"
ws2.Cells(last_row + 3, "G").Formula = "=sum(G" & (last_row + 1) & ":G" & (last_row + 2) & ")"
'formula for dual night total

ws2.Cells(last_row + 5, "G").Formula = "=sum(F" & (last_row + 3) & ":G" & (last_row + 3) & ")"
'dual total

ws2.Cells(last_row + 1, "H").Formula = "=sum(H2:H" & last_row & ")"
ws2.Cells(last_row + 3, "H").Formula = "=sum(H" & (last_row + 1) & ":H" & (last_row + 2) & ")"

ws2.Cells(last_row + 1, "I").Formula = "=sum(I2:I" & last_row & ")"
ws2.Cells(last_row + 3, "I").Formula = "=sum(I" & (last_row + 1) & ":I" & (last_row + 2) & ")"

ws2.Cells(last_row + 5, "I").Formula = "=sum(H" & (last_row + 3) & ":I" & (last_row + 3) & ")"
'solo total

ws2.Cells(last_row + 1, "J").Formula = "=sum(J2:J" & last_row & ")"
ws2.Cells(last_row + 3, "J").Formula = "=sum(J" & (last_row + 1) & ":J" & (last_row + 2) & ")"

ws2.Cells(last_row + 1, "K").Formula = "=sum(K2:K" & last_row & ")"
ws2.Cells(last_row + 3, "K").Formula = "=sum(K" & (last_row + 1) & ":K" & (last_row + 2) & ")"

ws2.Cells(last_row + 5, "K").Formula = "=sum(J" & (last_row + 3) & ":K" & (last_row + 3) & ")"
'Inst total

ws2.Cells(last_row + 1, "L").Formula = "=sum(L2:L" & last_row & ")"
ws2.Cells(last_row + 3, "L").Formula = "=sum(L" & (last_row + 1) & ":L" & (last_row + 2) & ")"

ws2.Cells(last_row + 1, "M").Formula = "=sum(M2:M" & last_row & ")"
ws2.Cells(last_row + 3, "M").Formula = "=sum(M" & (last_row + 1) & ":M" & (last_row + 2) & ")"

ws2.Cells(last_row + 5, "M").Formula = "=sum(L" & (last_row + 3) & ":M" & (last_row + 3) & ")"
'Dual XC total

ws2.Cells(last_row + 1, "N").Formula = "=sum(N2:N" & last_row & ")"
ws2.Cells(last_row + 3, "N").Formula = "=sum(N" & (last_row + 1) & ":N" & (last_row + 2) & ")"

ws2.Cells(last_row + 1, "O").Formula = "=sum(O2:O" & last_row & ")"
ws2.Cells(last_row + 3, "O").Formula = "=sum(O" & (last_row + 1) & ":O" & (last_row + 2) & ")"

ws2.Cells(last_row + 5, "O").Formula = "=sum(N" & (last_row + 3) & ":O" & (last_row + 3) & ")"
'Solo XC total


ws2.Cells(last_row + 1, "B").Formula = "=sum(F" & (last_row + 3) & ":I" & (last_row + 3) & ")"

With ws2.Range(Cells(last_row + 1, 1), Cells(last_row + 5, 60))
 .Font.Size = 8
 .Rows.AutoFit
 .WrapText = False
End With

ws2.Range(Cells(last_row + 1, 1), Cells(last_row + 1, 2)).Font.Bold = True
ws2.Range(Cells(last_row + 1, "F"), Cells(last_row + 1, "O")).Font.Bold = True
ws2.Range(Cells(last_row + 3, "F"), Cells(last_row + 5, "O")).Font.Bold = True
ws2.Range(Cells(last_row + 1, "Q"), Cells(last_row + 3, "AL")).Font.Bold = True

End Sub

