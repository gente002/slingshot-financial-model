Attribute VB_Name = "KernelCsvIO"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelCsvIO.bas
' Purpose: Writes computation output to CSV files using atomic write pattern.
' =============================================================================


' =============================================================================
' WriteCSV
' Writes a CSV file with headers and data rows.
' Uses atomic write pattern (PT-002): write to .tmp, verify, rename.
' =============================================================================
Public Function WriteCSV(ByRef outputs() As Variant, _
                         totalRows As Long, _
                         csvPath As String) As Boolean
    On Error GoTo ErrHandler

    WriteCSV = False

    Dim totalCols As Long
    totalCols = KernelConfig.GetColumnCount()

    ' Ensure directory exists
    Dim dirPath As String
    dirPath = Left(csvPath, InStrRev(csvPath, "\") - 1)
    If Dir(dirPath, vbDirectory) = "" Then
        MkDir dirPath
    End If

    ' Step 1: Write to temp file
    Dim tmpPath As String
    tmpPath = csvPath & ".tmp"

    Dim fileNum As Integer
    fileNum = FreeFile
    Open tmpPath For Output As #fileNum

    ' Write header row
    Dim headerLine As String
    headerLine = BuildHeaderLine(totalCols)
    Print #fileNum, headerLine

    ' Write data rows
    Dim rowIdx As Long
    Dim rowsWritten As Long
    rowsWritten = 0

    For rowIdx = 1 To totalRows
        Dim dataLine As String
        dataLine = BuildDataLine(outputs, rowIdx, totalCols)
        Print #fileNum, dataLine
        rowsWritten = rowsWritten + 1
    Next rowIdx

    Close #fileNum

    ' Step 2: Verify row count
    If rowsWritten <> totalRows Then
        KernelConfig.LogError SEV_ERROR, "KernelCsvIO", "E-400", _
                              "CSV row count mismatch: expected " & totalRows & _
                              ", wrote " & rowsWritten, _
                              "MANUAL BYPASS: The .tmp file at " & tmpPath & " has partial data. Copy Detail tab data to CSV manually."
        ' Keep .tmp for debugging
        WriteCSV = False
        Exit Function
    End If

    ' Step 3: Verify by re-reading and counting lines
    Dim verifyCount As Long
    verifyCount = VerifyTmpRowCount(tmpPath)
    ' verifyCount includes header row, so data rows = verifyCount - 1
    If (verifyCount - 1) <> totalRows Then
        KernelConfig.LogError SEV_ERROR, "KernelCsvIO", "E-401", _
                              "CSV verification failed: expected " & totalRows & _
                              " data rows, found " & (verifyCount - 1), _
                              "MANUAL BYPASS: The .tmp file at " & tmpPath & " failed verification. Copy Detail tab data to CSV manually."
        WriteCSV = False
        Exit Function
    End If

    ' Step 4: Delete old file if exists, rename .tmp to final
    If Dir(csvPath) <> "" Then
        Kill csvPath
    End If
    Name tmpPath As csvPath

    KernelConfig.LogError SEV_INFO, "KernelCsvIO", "I-400", _
                          "CSV written successfully", csvPath & " (" & totalRows & " rows)"

    WriteCSV = True
    Exit Function

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelCsvIO", "E-499", _
                          "Error writing CSV: " & Err.Description, _
                          "MANUAL BYPASS: Copy the Detail tab data to a CSV file manually. Use headers from row 1 of Detail. Save to: " & csvPath
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    WriteCSV = False
End Function


' =============================================================================
' BuildHeaderLine
' Builds the CSV header line from ColumnRegistry names.
' Columns are ordered by CsvIndex (0-indexed).
' =============================================================================
Private Function BuildHeaderLine(totalCols As Long) As String
    ' Build array ordered by CSV column position
    Dim headers() As String
    ReDim headers(0 To totalCols - 1)

    Dim regIdx As Long
    For regIdx = 1 To totalCols
        Dim colName As String
        colName = KernelConfig.GetColName(regIdx)
        Dim csvCol As Long
        csvCol = KernelConfig.CsvIndex(colName)
        If csvCol >= 0 And csvCol < totalCols Then
            headers(csvCol) = colName
        End If
    Next regIdx

    ' Quote each header value
    Dim cidx As Long
    For cidx = 0 To totalCols - 1
        headers(cidx) = CsvQuote(headers(cidx))
    Next cidx

    BuildHeaderLine = Join(headers, ",")
End Function


' =============================================================================
' BuildDataLine
' Builds a CSV data line for the given row from the outputs array.
' Columns are ordered by CsvIndex mapping.
' =============================================================================
Private Function BuildDataLine(ByRef outputs() As Variant, _
                               rowIdx As Long, totalCols As Long) As String
    ' Map Detail columns to CSV columns
    Dim csvValues() As String
    ReDim csvValues(0 To totalCols - 1)

    Dim regIdx As Long
    For regIdx = 1 To totalCols
        Dim colName As String
        colName = KernelConfig.GetColName(regIdx)

        Dim csvCol As Long
        csvCol = KernelConfig.CsvIndex(colName)

        Dim detCol As Long
        detCol = KernelConfig.GetDetailCol(regIdx)

        If csvCol >= 0 And csvCol < totalCols And detCol >= 1 Then
            Dim cellVal As Variant
            cellVal = outputs(rowIdx, detCol)

            If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then
                ' Numeric: round to 6 decimal places, no quoting
                csvValues(csvCol) = Format(CDbl(cellVal), "0.000000")
            Else
                ' Text: quote if needed
                csvValues(csvCol) = CsvQuote(CStr(cellVal))
            End If
        End If
    Next regIdx

    BuildDataLine = Join(csvValues, ",")
End Function


' =============================================================================
' CsvQuote
' Quotes a string value if it contains commas, quotes, or newlines.
' =============================================================================
Private Function CsvQuote(val As String) As String
    If InStr(1, val, ",") > 0 Or InStr(1, val, """") > 0 Or _
       InStr(1, val, vbCr) > 0 Or InStr(1, val, vbLf) > 0 Then
        ' Escape double quotes by doubling them
        CsvQuote = """" & Replace(val, """", """""") & """"
    Else
        CsvQuote = val
    End If
End Function


' =============================================================================
' VerifyTmpRowCount
' Re-reads the temp file and counts total lines.
' =============================================================================
Private Function VerifyTmpRowCount(tmpPath As String) As Long
    Dim fileNum As Integer
    fileNum = FreeFile

    Dim lineCount As Long
    lineCount = 0

    Dim lineText As String

    Open tmpPath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineCount = lineCount + 1
    Loop
    Close #fileNum

    VerifyTmpRowCount = lineCount
End Function


' =============================================================================
' LoadCSVToDetail
' Reads a CSV and populates the Detail tab.
' Stub for Phase 1 - implement in Phase 2.
' =============================================================================
Public Function LoadCSVToDetail(csvPath As String) As Boolean
    KernelConfig.LogError SEV_INFO, "KernelCsvIO", "I-410", _
                          "LoadCSVToDetail is a Phase 2 stub", csvPath
    LoadCSVToDetail = False
End Function
