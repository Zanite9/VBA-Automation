Attribute VB_Name = "OpsMetricsAutomation"
' Written by Sunveer Dhillon '
' Copyright 2025 '

Option Explicit
Sub RunOpsMetricsAutomation()
FilterWarrantyData
ImportParts
AddCategoryColumn
FormatNotes
SummaryStats
ExportToTxt
End Sub
' Copy WARRANTY rows from Raw Data into Warranty Sheet - rev2 functional '
Sub FilterWarrantyData()
    Dim wsRaw As Worksheet, wsWarranty As Worksheet
    Set wsRaw = ThisWorkbook.Sheets("Raw Data")
 
    Application.DisplayAlerts = False
    On Error Resume Next: Sheets("Warranty").Delete: On Error GoTo 0
    Application.DisplayAlerts = True
 
    Set wsWarranty = ThisWorkbook.Sheets.Add(After:=wsRaw)
    wsWarranty.Name = "Warranty"
 
    wsRaw.Rows(1).AutoFilter Field:=1, Criteria1:="WARRANTY"
    wsRaw.Range("A1").CurrentRegion.SpecialCells(xlCellTypeVisible).Copy Destination:=wsWarranty.Range("A1")
    wsRaw.AutoFilterMode = False
End Sub
' Import Parts sheet from shared drive '
Sub ImportParts()
    Dim folderPath As String
    folderPath = "S:\_BIS\Empower\Part Data Dump\CT\"
    
    Dim fileArray() As String
    Dim fileName As String
    Dim fileCount As Integer
    Dim i As Integer
    
    ' Collect all csv files in folder '
    fileName = Dir(folderPath & "*.csv")
    Do While fileName <> ""
        fileCount = fileCount + 1
        ReDim Preserve fileArray(1 To fileCount)
        fileArray(fileCount) = fileName
        fileName = Dir
    Loop
    
    ' Check for at least 3 files '
    If fileCount < 2 Then
        MsgBox "Fewer than 3 .csv files in the folder.", vbExclamation
        Exit Sub
    End If
    
    ' Full path of the 3rd file '
    Dim targetFile As String
    targetFile = folderPath & fileArray(3)
    
    ' Delete existing "Parts" sheet if it exists '
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("Parts").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    ' Add new sheet after Warranty '
    Dim wsParts As Worksheet
    Set wsParts = ThisWorkbook.Sheets.Add(After:=Sheets("Warranty"))
    wsParts.Name = "Parts"
    
    ' Import CSV into Parts sheet '
    With wsParts.QueryTables.Add(Connection:="TEXT;" & targetFile, Destination:=wsParts.Range("A1"))
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .TextFilePlatform = xlWindows
        .Refresh BackgroundQuery:=False
    End With
    
    MsgBox "Imported file: " & fileArray(3), vbInformation, "Parts Sheet Created"
End Sub
' Add Category Column To Sheet - rev2 functional '
Sub AddCategoryColumn()
    Dim wsWarranty As Worksheet, wsParts As Worksheet
    Set wsWarranty = Sheets("Warranty")
    Set wsParts = Sheets("Parts")
        
    Dim partDict As Object:
    Set partDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Long, lastRow As Long
    Dim partNum As String

    'Assume parts sheet: Col A = Part Num, Col D = Category'
    Dim partsLastRow As Long
    partsLastRow = wsParts.Cells(wsParts.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To partsLastRow
        partNum = Trim(wsParts.Cells(i, "A").Value)
        If Len(partNum) > 0 Then
            partDict(partNum) = wsParts.Cells(i, "D").Value
        End If
    Next i
    
    ' Insert new column before "Return Code" '
    wsWarranty.Columns("A:A").Insert Shift:=xlToRight
    wsWarranty.Cells(1, "A").Value = "Category"
    
    ' Copy formatting from shifted column to column A '
    wsWarranty.Columns("B:B").Copy
    wsWarranty.Columns("A:A").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ' Populate Category column using part num from column F (after insertion) '
    lastRow = wsWarranty.Cells(wsWarranty.Rows.Count, "B").End(xlUp).Row
    For i = 2 To lastRow
        partNum = Trim(wsWarranty.Cells(i, "F").Value) ' F is part num after insertion '
        If partDict.exists(partNum) Then
            wsWarranty.Cells(i, "A").Value = partDict(partNum)
        Else
            wsWarranty.Cells(i, "A").Value = "UNKNOWN"
        End If
        wsWarranty.Cells(i, "A").HorizontalAlignment = xlCenter
    Next i
End Sub
 ' Format the note section in a better way for presenting - rev2 functional '
Sub FormatNotes()
    Dim ws As Worksheet
    Set ws = Sheets("Warranty")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Determine where to place new column '
    Dim noteColIndex As Long
    noteColIndex = ws.Cells(1, Columns.Count).End(xlToLeft).Column + 1
    
    ' Add header for new column '
    ws.Cells(1, noteColIndex).Value = "Formatted Notes"
    
    ' Copy formatting from existing note column '
    
    ws.Columns("K:K").Copy
    ws.Columns(noteColIndex).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    Dim i As Long
    For i = 2 To lastRow
        Dim rmaNum As String: rmaNum = Trim(ws.Cells(i, "C").Value) ' RMA # col C '
        Dim qty As String: qty = Trim(ws.Cells(i, "G").Value) ' Qty Disposition col G'
        Dim cust As String: cust = Trim(ws.Cells(i, "E").Value) ' customer col E '
        Dim part As String: part = Trim(ws.Cells(i, "F").Value) ' part num col F '
        Dim desc As String: desc = Trim(ws.Cells(i, "J").Value) ' Note col J '
        
        Dim formattedNote As String
        formattedNote = rmaNum & " QTY:" & qty & vbCrLf & cust & ", " & part & vbCrLf & desc
        
        ws.Cells(i, noteColIndex).Value = formattedNote
        ws.Cells(i, noteColIndex).HorizontalAlignment = xlCenter
        
        ' Copy formatting from shifted column to column A '
        ws.Columns("J:J").Copy
        ws.Columns("K:K").PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        
    Next i
End Sub
 ' Summarize category totals for PCB and Electro-Mechanical - rev2 functional  '
Sub SummaryStats()
    Dim ws As Worksheet
    Set ws = Sheets("Warranty")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim pcbSum As Long, emSum As Long
    pcbSum = 0: emSum = 0
    
    Dim i As Long
    For i = 2 To lastRow
        Select Case Trim(ws.Cells(i, "A").Value) ' Col A is category '
            Case "ASSEMBLY-PCB"
                pcbSum = pcbSum + CLng(ws.Cells(i, "G").Value) ' qty disposition in cell G '
            Case "ASSEMBLY-ELECTRO MECHANICAL"
                emSum = emSum + CLng(ws.Cells(i, "G").Value)
        End Select
    Next i
    
    ' Write headers and totals into new columns '
    ws.Cells(1, lastCol + 1).Value = "PCB Total Qty"
    ws.Cells(1, lastCol + 2).Value = "Electro-Mech Total Qty"
    
    ws.Cells(2, lastCol + 1).Value = pcbSum
    ws.Cells(2, lastCol + 2).Value = emSum
    
    ' Copy formatting from qty disposition col G '
    ws.Columns("G:G").Copy
    ws.Columns(lastCol + 1).PasteSpecial Paste:=xlPasteFormats
    ws.Columns(lastCol + 2).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False
    
    ws.Cells(1, lastCol + 1).HorizontalAlignment = xlCenter
    ws.Cells(1, lastCol + 2).HorizontalAlignment = xlCenter
End Sub
' Export Formatted Notes to .txt File - rev 6 functional'
Sub ExportToTxt()
    Dim ws As Worksheet
    Set ws = Sheets("Warranty")
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Get column indicess dynamically '
    Dim colCategory As Long, colNote As Long, colQty As Long, colDate As Long
    colCategory = 1 ' col A '
    colNote = 0: colQty = 0: colDate = 0
    
    Dim i As Long
    For i = 1 To ws.Cells(1, Columns.Count).End(xlToLeft).Column
        If Trim(ws.Cells(1, i).Value) = "Formatted Notes" Then colNote = i
        If Trim(ws.Cells(1, i).Value) = "RMA Date" Then colDate = i
    Next i
    
    If colNote = 0 Or colDate = 0 Then
        MsgBox "Missing one of the required columns: Formatted Notes or RMA Date.", vbExclamation
        Exit Sub
    End If
    
    ' Build arrays for filtered notes '
    Dim pcbNotes As Collection, emNotes As Collection
    Set pcbNotes = New Collection
    Set emNotes = New Collection
    
    ' Read PCB/EM total from Warranty sheet L:2 and M:2 '
    Dim pcbQty As Long, emQty As Long
    pcbQty = CLng(ws.Range("L2").Value)
    emQty = CLng(ws.Range("M2").Value)
    
    ' Initialize min/max dates '
    Dim minDate As Date, maxDate As Date
    minDate = #12/31/9999#: maxDate = #1/1/1900#
    
    ' Loop through rows to build notes and stats '
    For i = 2 To lastRow
        Dim cat As String: cat = Trim(ws.Cells(i, colCategory).Value)
        Dim noteText As String: noteText = Trim(ws.Cells(i, colNote).Value)
        Dim rmaDate As Variant: rmaDate = ws.Cells(i, colDate).Value
        
        ' Track min/max dates '
        If IsDate(rmaDate) Then
            If CDate(rmaDate) < minDate Then minDate = CDate(rmaDate)
            If CDate(rmaDate) > maxDate Then maxDate = CDate(rmaDate)
        End If
        
        ' Collect formatted notes '
        If Len(noteText) > 0 Then
            Select Case cat
                Case "ASSEMBLY-PCB": pcbNotes.Add noteText
                Case "ASSEMBLY-ELECTRO MECHANICAL": emNotes.Add noteText
            End Select
        End If
    Next i
    
    ' Format date range: MM-DD-YYYY to MM-DD-YYYY '
    Dim dateLabel As String
    If minDate <= maxDate Then
        dateLabel = Format(minDate, "mm-dd-yyyy") & "_to_" & Format(maxDate, "mm-dd-yyyy")
    Else
        dateLabel = "no_dates_found"
    End If
    
    ' Ensure workbook is saved before proceeding '
    If ThisWorkbook.Path = "" Then
        MsgBox "Please save the workbook first before exporting notes.", vbExclamation
        Exit Sub
    End If
    
    ' Build output paths on user's Downloads folder '
    Dim dlPath As String
    dlPath = Environ("USERPROFILE") & "\Downloads\"
    
    Dim pcbFile As String, emFile As String
    pcbFile = dlPath & "RMA_PCB_(QTY=" & pcbQty & ")_" & dateLabel & ".txt"
    emFile = dlPath & "RMA_ElectroMech_(QTY=" & emQty & ")_" & dateLabel & ".txt"
 
    ' Write files using FileSystemObject '
    Dim fso As Object, outFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
 
    ' Write PCB file '
    Set outFile = fso.CreateTextFile(pcbFile, True, True)
    Dim item As Variant
    For Each item In pcbNotes
        outFile.WriteLine item
        outFile.WriteLine ' blank line separation '
    Next item
    outFile.Close
 
    ' Write Electro-Mech file '
    Set outFile = fso.CreateTextFile(emFile, True, True)
    For Each item In emNotes
        outFile.WriteLine item
        outFile.WriteLine ' blank line separation '
    Next item
    outFile.Close
 
    MsgBox "Notes exported to Downloads:" & vbCrLf & vbCrLf & pcbFile & vbCrLf & emFile, vbInformation, "Export Complete"
End Sub

