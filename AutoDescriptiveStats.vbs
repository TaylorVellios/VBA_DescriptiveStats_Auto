Sub Statistics()
    Dim dataset As String
    Dim rangeEnd As Integer
    Dim datasheet As String
    Dim ColumnNum As Integer
    
'-------------Confirming Active Sheet Is the One We Want----------------
    confirm = MsgBox("Please Make Sure the Sheet You Want to Analyze is Your Active Sheet", vbOKCancel)
    If confirm = vbCancel Then
        Exit Sub        '<-----------------Simple Conditional to Kill this macro if the user is not on the right active page.
    Else
        datasheet = ActiveSheet.Name
    End If
    
'-------------Pull Column Name to Process----------------
    dataset = InputBox("Enter The ExcelColumn with Numeric Data to Analyze:" & vbCrLf & "Ex:  A")
    dataset = UCase(dataset)
    ColumnNum = (Asc(UCase(dataset)) - 64) '<--------------------------------------------------------------This gives us a numerical value for the column if needed
    
'-------------Find the End of the DataSet----------------
    rangeEnd = Cells(Rows.Count, dataset).End(xlUp).Row
    
'-------------Creating The Output Sheet----------------
    Dim NameSheet As String
    Dim Title As String
    Dim Default As String
    Dim today As String
    
    today = Format(Date, "Medium Date")
    
    Default = "Analysis " & today
    
'-------------Get New Sheet Name From User----------------------
    NameSheet = InputBox("Please Enter the Name of Your New Sheet:", Default, Default)

'-------------Creates New Sheet at the End of Workbook, Named---------------
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = NameSheet
    
'-------------Create Headers For New Sheet...STAY ORGANIZED!-----
    Title = InputBox("Please Enter a Title for Your Analysis:", Default, Default)

    Range("B2").Value = "Fixed Data"
    Range("G2").Value = "Formula"
    
    Range("B10,G10").Value = "Quartiles(INC)"
    Range("B15,G15").Value = "Quartiles(EXC)"
    
    Range("B20,G20").Value = "Upper Outliers"
    Range("B21,G21").Value = "Min/Max INC"
    Range("B24,G24").Value = "Min/Max EXC"
    
    Range("E20,J20").Value = "Count Of Outliers"
'-------------Create Title Cells for Data-----------------------
    Range("A1").Value = Title
    
    Range("C3,H3").Value = "Mean"
    Range("C4,H4").Value = "Median"
    Range("C5,H5").Value = "Max Value"
    Range("C6,H6").Value = "Min Value"
    Range("C7,H7").Value = "Standard Deviation P"
    Range("C8,H8").Value = "Standard Deviation S"
    
    Range("C11,C16,H11,H16").Value = "Upper Quartile"
    Range("C12,C17,H12,H17").Value = "Lower Quartile"
    Range("C13,C18,H13,H18").Value = "IQR"

    Range("C21,H21").Value = "IQR Rule"
    Range("C22,H22").Value = "Std. Dev. Mean Rule"
    
    Range("C24,H24").Value = "IQR Rule"

'-------------Formulaic Data Output----------------
    formrange = "(" & datasheet & "!" & dataset & ":" & dataset
    
    Range("i3").Value = "=AVERAGE" & formrange & ")"
    Range("i4").Value = "=MEDIAN" & formrange & ")"
    Range("i5").Value = "=MAX" & formrange & ")"
    Range("i6").Value = "=MIN" & formrange & ")"
    Range("i7").Value = "=STDEV.P" & formrange & ")"
    Range("i8").Value = "=STDEV.S" & formrange & ")"
    
    Range("i11").Value = "=QUARTILE.INC" & formrange & ",3)"
    Range("i12").Value = "=QUARTILE.INC" & formrange & ",1)"
    Range("i13").Value = "=(I11 - I12)"
    
    Range("i16").Value = "=QUARTILE.EXC" & formrange & ",3)"
    Range("i17").Value = "=QUARTILE.EXC" & formrange & ",1)"
    Range("i18").Value = "=(I16 - I17)"
    
'-------------Formulaic Data Output----------------
    Range("D3").Value = Range("I3").Value
    Range("D4").Value = Range("I4").Value
    Range("D5").Value = Range("I5").Value
    Range("D6").Value = Range("I6").Value
    Range("D7").Value = Range("I7").Value
    Range("D8").Value = Range("I8").Value
    
    Range("D11").Value = Range("I11").Value
    Range("D12").Value = Range("I12").Value
    Range("D13").Value = Range("I13").Value
    
    Range("D16").Value = Range("I16").Value
    Range("D17").Value = Range("I17").Value
    Range("D18").Value = Range("I18").Value

'------------Outliers---------------------
    Range("i21").Value = "=(1.5 * I13) + I11"
    Range("i22").Value = "=(3 * I7) + I3"

    Range("i24").Value = "=(1.5 * I18) + I16"
    
    
    Range("J21").Value = "=COUNTIFS" & formrange & ","">""& I21)"
    Range("J22").Value = "=COUNTIFS" & formrange & ","">""& I22)"

    Range("J24").Value = "=COUNTIFS" & formrange & ","">""& I24)"
    
    '-----Fixed
    Range("D21").Value = Range("I21").Value
    Range("D22").Value = Range("I22").Value
    
    Range("D24").Value = Range("I24").Value
    
    Range("E21").Value = Range("J21").Value
    Range("E22").Value = Range("J22").Value

    Range("E24").Value = Range("J24").Value

'-------------Format Header Cell and Table-----------

    Call FormatNew
    
End Sub
Sub FormatNew()
'-------------Worksheet Title-----------
    Range("A1").Font.Size = 18
    Range("A1").Font.Bold = True

'-------------Section Headers-----------
    Range("B2:B25,G2:G25,C3:C24,H3:H24").Font.Bold = True

    Range("B2:B25").Font.Size = 16
    Range("G2:G25").Font.Size = 16
'-------------Borders-----------
    Range("B2:D2,G2:I2,B10:D10,G10:I10,B15:D15,G15:I15,B20:D20,G20:I20,E20,J20").BorderAround ColorIndex:=1
    
    Range("C3:D8,C11:D13,C16:D18,C21:E24,H3:I8,H11:I13,H16:I18,H21:J24").Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Range("C3:D8,C11:D13,C16:D18,C21:E24,H3:I8,H11:I13,H16:I18,H21:J24").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("C3:D8,C11:D13,C16:D18,C21:E24,H3:I8,H11:I13,H16:I18,H21:J24").BorderAround ColorIndex:=1
    
'-------------Data Beautification - Fonts and Colors--------
    Range("B2:D2,G2:I2,B10:D10,G10:I10,B15:D15,G15:I15,B20:D20,G20:I20,E20,J20").Interior.ColorIndex = 24
    
    Range("B21:b24,G21:G24").Font.Size = 10
    
    Range("E20,J20").Font.Bold = True
    
    Range("C3:D8,C11:D13,C16:D18,C21:E24,H3:I8,H11:I13,H16:I18,H21:J24").Interior.Color = RGB(242, 233, 245)
        
    Columns.AutoFit
    
    Range("C23:J23").Clear
    
End Sub

