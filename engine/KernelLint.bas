Attribute VB_Name = "KernelLint"
Option Explicit

' Copyright (c) 2026 Ethan Genteman. All rights reserved.
' Proprietary and confidential. Unauthorized use prohibited.
' See LICENSE.txt for terms.

' =============================================================================
' KernelLint.bas
' Purpose: Anti-pattern scanner. Reads .bas files from disk and checks them
'          against the anti-pattern library. The RDK's immune system.
' Phase 4: Observability + Hardening
' =============================================================================

' Violation record structure (parallel arrays)
Private m_vFiles() As String
Private m_vLines() As Long
Private m_vCheckIDs() As String
Private m_vSeverities() As String
Private m_vDescriptions() As String
Private m_vCount As Long

' =============================================================================
' RunLint
' Scans all .bas files in engine/ directory (or targetPath if specified).
' Results written to TestResults sheet under a LINT RESULTS section.
' =============================================================================
Public Sub RunLint(Optional targetPath As String = "")
    On Error GoTo ErrHandler

    Dim enginePath As String
    If Len(targetPath) > 0 Then
        enginePath = targetPath
    Else
        enginePath = ThisWorkbook.Path & "\..\engine"
    End If

    ' Initialize violation arrays
    ResetViolations

    ' List all .bas files using FileSystemObject (PT-019)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(enginePath) Then
        KernelConfig.LogError SEV_ERROR, "KernelLint", "E-700", _
            "Engine directory not found: " & enginePath, _
            "MANUAL BYPASS: Review .bas files manually against data/anti_patterns.csv. " & _
            "Check for non-ASCII characters, magic numbers, and missing contract functions."
        MsgBox "Engine directory not found: " & enginePath, vbExclamation, "RDK -- Lint"
        Exit Sub
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(enginePath)
    Dim fileObj As Object
    Dim fileCount As Long
    fileCount = 0

    For Each fileObj In folder.Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "bas" Then
            fileCount = fileCount + 1
            LintSingleFile fileObj.Path, fileObj.Name, False
        End If
    Next fileObj

    Set folder = Nothing
    Set fso = Nothing

    ' Also scan live VBA project modules (catches edits made in VBA editor)
    Dim vbaCount As Long
    vbaCount = LintVBAModules(False)

    ' Write results to TestResults sheet
    WriteLintResults

    ' Summary
    Dim errCount As Long
    Dim warnCount As Long
    CountBySeverity errCount, warnCount

    ' Unhide TestResults tab when violations found
    If m_vCount > 0 Then
        On Error Resume Next
        Dim wsResults As Worksheet
        Set wsResults = ThisWorkbook.Sheets(TAB_TEST_RESULTS)
        If Not wsResults Is Nothing Then
            wsResults.Visible = xlSheetVisible
            wsResults.Activate
        End If
        On Error GoTo ErrHandler
    End If

    Dim summary As String
    Dim sourceDesc As String
    If vbaCount > 0 Then
        sourceDesc = fileCount & " disk files + " & vbaCount & " VBA modules"
    Else
        sourceDesc = fileCount & " files"
    End If
    If m_vCount = 0 Then
        summary = "Lint CLEAN: " & sourceDesc & " scanned, 0 violations."
    Else
        summary = "Lint complete: " & sourceDesc & " scanned, " & _
                  m_vCount & " violations (" & errCount & " errors, " & warnCount & " warnings)"
    End If

    KernelConfig.LogError SEV_INFO, "KernelLint", "I-700", summary, ""
    MsgBox summary, vbInformation, "RDK -- Lint"
    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelLint", "E-799", _
        "Lint failed: " & Err.Description, _
        "MANUAL BYPASS: Review .bas files manually against data/anti_patterns.csv. " & _
        "Check for non-ASCII characters, magic numbers, and missing contract functions."
    MsgBox "Lint failed: " & Err.Description, vbCritical, "RDK -- Lint"
End Sub


' =============================================================================
' RunLintOnFile
' Scans a single .bas file. Returns violation count.
' =============================================================================
Public Function RunLintOnFile(filePath As String) As Long
    On Error GoTo ErrHandler

    ResetViolations

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fileName As String
    fileName = fso.GetFileName(filePath)
    Set fso = Nothing

    LintSingleFile filePath, fileName, False
    RunLintOnFile = m_vCount
    Exit Function

ErrHandler:
    RunLintOnFile = -1
End Function


' =============================================================================
' RunLintQuick
' Runs only ERROR-severity checks (faster). Skips WARN-level checks.
' =============================================================================
Public Sub RunLintQuick()
    On Error GoTo ErrHandler

    Dim enginePath As String
    enginePath = ThisWorkbook.Path & "\..\engine"

    ResetViolations

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(enginePath) Then
        Set fso = Nothing
        Exit Sub
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(enginePath)
    Dim fileObj As Object
    Dim fileCount As Long
    fileCount = 0

    For Each fileObj In folder.Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "bas" Then
            fileCount = fileCount + 1
            LintSingleFile fileObj.Path, fileObj.Name, True
        End If
    Next fileObj

    Set folder = Nothing
    Set fso = Nothing

    Dim errCount As Long
    Dim warnCount As Long
    CountBySeverity errCount, warnCount

    Dim summary As String
    If m_vCount = 0 Then
        summary = "Lint Quick CLEAN: " & fileCount & " files, 0 errors."
    Else
        summary = "Lint Quick: " & fileCount & " files, " & errCount & " errors found."
    End If

    KernelConfig.LogError SEV_INFO, "KernelLint", "I-701", summary, ""
    MsgBox summary, vbInformation, "RDK -- Lint Quick"
    Exit Sub

ErrHandler:
    KernelConfig.LogError SEV_ERROR, "KernelLint", "E-798", _
        "Lint Quick failed: " & Err.Description, _
        "MANUAL BYPASS: Review .bas files manually against data/anti_patterns.csv."
    MsgBox "Lint Quick failed: " & Err.Description, vbCritical, "RDK -- Lint Quick"
End Sub


' =============================================================================
' RunLintQuickSilent
' Same as RunLintQuick but no MsgBox or ErrorLog. For internal use by
' DiagnosticDump so it does not pop dialogs during error handling.
' =============================================================================
Public Sub RunLintQuickSilent()
    On Error Resume Next

    Dim enginePath As String
    enginePath = ThisWorkbook.Path & "\..\engine"

    ResetViolations

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(enginePath) Then
        Set fso = Nothing
        Exit Sub
    End If

    Dim folder As Object
    Set folder = fso.GetFolder(enginePath)
    Dim fileObj As Object

    For Each fileObj In folder.Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "bas" Then
            LintSingleFile fileObj.Path, fileObj.Name, True
        End If
    Next fileObj

    Set folder = Nothing
    Set fso = Nothing
    On Error GoTo 0
End Sub


' =============================================================================
' GetQuickViolationCount
' Returns the violation count from the last RunLintQuick (for DiagnosticDump).
' =============================================================================
Public Function GetQuickViolationCount() As Long
    GetQuickViolationCount = m_vCount
End Function


' =============================================================================
' GetQuickViolationSummary
' Returns a summary string of the last quick lint run (for DiagnosticDump).
' =============================================================================
Public Function GetQuickViolationSummary() As String
    If m_vCount = 0 Then
        GetQuickViolationSummary = "0 errors found."
        Exit Function
    End If
    Dim summary As String
    summary = m_vCount & " error(s):" & vbCrLf
    Dim i As Long
    For i = 1 To m_vCount
        summary = summary & "  " & m_vFiles(i) & " L" & m_vLines(i) & _
                  " [" & m_vCheckIDs(i) & "] " & m_vDescriptions(i) & vbCrLf
    Next i
    GetQuickViolationSummary = summary
End Function


' =============================================================================
' LintSingleFile
' Runs all lint checks on a single file.
' If errorsOnly = True, only runs ERROR-severity checks.
' =============================================================================
Private Sub LintSingleFile(filePath As String, fileName As String, errorsOnly As Boolean)
    On Error GoTo FileErr

    ' Read entire file as binary via byte array (to detect CRLF vs LF and non-ASCII)
    Dim fileNum As Integer
    fileNum = FreeFile
    Dim fileContent As String
    Dim fileSize As Long

    Open filePath For Binary Access Read As #fileNum
    fileSize = LOF(fileNum)
    If fileSize = 0 Then
        Close #fileNum
        Exit Sub
    End If
    Dim rawBytes() As Byte
    ReDim rawBytes(0 To fileSize - 1)
    Get #fileNum, , rawBytes
    Close #fileNum
    ' Convert byte array to string preserving each byte as a character
    fileContent = StrConv(rawBytes, vbUnicode)

    Dim isDomainModule As Boolean
    isDomainModule = ((Left(LCase(fileName), 6) = "domain") Or _
                      (Left(LCase(fileName), 12) = "sampledomain")) And _
                     (InStr(1, LCase(fileName), "tests") = 0)

    ' LINT-12: CRLF check (ERROR) -- scan raw bytes for reliability
    CheckCRLF fileName, rawBytes, fileSize

    ' LINT-11: Module size (WARN/ERROR)
    CheckModuleSize fileName, fileSize, errorsOnly

    ' LINT-01: Non-ASCII (ERROR) -- scan raw bytes for reliability
    CheckNonASCII fileName, rawBytes, fileSize

    ' LINT-10: Sub/Function balance (ERROR)
    Dim lines() As String
    lines = Split(Replace(fileContent, vbCrLf, vbLf), vbLf)
    CheckSubFuncBalance fileName, lines

    ' LINT-04: ReDim Preserve in loop (ERROR)
    CheckReDimInLoop fileName, lines

    ' LINT-08: Missing contract functions -- Domain only (ERROR)
    If isDomainModule Then
        CheckMissingContract fileName, lines
    End If

    ' WARN-level checks (skip in errorsOnly mode)
    If Not errorsOnly Then
        ' LINT-02: Formula injection in .Value (WARN)
        CheckFormulaInject fileName, lines

        ' LINT-03: Magic numbers in array indexing (WARN)
        CheckMagicNumbers fileName, lines

        ' LINT-05: Sized Dim after executable code (WARN)
        CheckSizedDimAfterCode fileName, lines

        ' LINT-06: 2-char vars i/m/f + uppercase (WARN)
        CheckTwoCharVars fileName, lines

        ' LINT-07: Cumulative storage in Domain code (WARN)
        If isDomainModule Then
            CheckCumulativeDomain fileName, lines
        End If

        ' LINT-09: = string without NumberFormat (WARN)
        CheckEqStringNoFormat fileName, lines
    End If

    Exit Sub

FileErr:
    AddViolation fileName, 0, "LINT-00", LINT_SEV_ERROR, _
        "Could not read file: " & Err.Description
End Sub


' =============================================================================
' LINT-01: Non-ASCII characters (AP-06)
' =============================================================================
Private Sub CheckNonASCII(fileName As String, rawBytes() As Byte, fileSize As Long)
    Dim i As Long
    Dim lineNum As Long
    lineNum = 1
    For i = 0 To fileSize - 1
        If rawBytes(i) > 127 Then
            AddViolation fileName, lineNum, LINT_NON_ASCII, LINT_SEV_ERROR, _
                "Non-ASCII byte (code " & rawBytes(i) & ") at position " & (i + 1)
            ' Only report first occurrence per file to avoid flood
            Exit Sub
        End If
        If rawBytes(i) = 10 Then lineNum = lineNum + 1
    Next i
End Sub


' =============================================================================
' LINT-02: Formula injection strings in .Value (AP-07)
' =============================================================================
Private Sub CheckFormulaInject(fileName As String, lines() As String)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = LTrim(lines(i))
        ' Skip comments
        If Left(ln, 1) = "'" Then GoTo NextFI
        ' Look for .Value = "=... or .Value = "+... etc. without preceding NumberFormat
        Dim fiValPos As Long
        fiValPos = InStr(1, ln, ".Value", vbTextCompare)
        If fiValPos > 0 Then
            If Not IsInsideString(ln, fiValPos) Then
                If MatchesFormulaInjectPattern(ln) Then
                    ' Check preceding 3 lines for NumberFormat = "@"
                    If Not HasNumberFormatGuard(lines, i) Then
                        AddViolation fileName, i + 1, LINT_FORMULA_INJECT, LINT_SEV_WARN, _
                            "AP-07: .Value write with formula-injection string without NumberFormat guard"
                    End If
                End If
            End If
        End If
NextFI:
    Next i
End Sub

Private Function MatchesFormulaInjectPattern(ln As String) As Boolean
    MatchesFormulaInjectPattern = False
    Dim pos As Long
    pos = InStr(1, ln, ".Value", vbTextCompare)
    If pos = 0 Then Exit Function
    ' Find the = after .Value
    Dim afterValue As String
    afterValue = Mid(ln, pos + 6)
    Dim eqPos As Long
    eqPos = InStr(1, afterValue, "=")
    If eqPos = 0 Then Exit Function
    ' Check what follows the =
    Dim afterEq As String
    afterEq = LTrim(Mid(afterValue, eqPos + 1))
    If Len(afterEq) < 2 Then Exit Function
    If Left(afterEq, 1) = """" Then
        Dim secondChar As String
        secondChar = Mid(afterEq, 2, 1)
        If secondChar = "=" Or secondChar = "+" Or secondChar = "-" Or secondChar = "@" Then
            MatchesFormulaInjectPattern = True
        End If
    End If
End Function

Private Function HasNumberFormatGuard(lines() As String, currentLine As Long) As Boolean
    HasNumberFormatGuard = False
    Dim startLine As Long
    startLine = currentLine - 3
    If startLine < 0 Then startLine = 0
    Dim i As Long
    For i = startLine To currentLine - 1
        If i >= 0 And i <= UBound(lines) Then
            If InStr(1, lines(i), "NumberFormat", vbTextCompare) > 0 Then
                If InStr(1, lines(i), """@""", vbTextCompare) > 0 Then
                    HasNumberFormatGuard = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Private Function IsInsideString(ln As String, pos As Long) As Boolean
    ' Count double-quote characters before the given position.
    ' If odd, the position is inside a string literal.
    Dim quoteCount As Long
    quoteCount = 0
    Dim k As Long
    For k = 1 To pos - 1
        If Mid(ln, k, 1) = """" Then quoteCount = quoteCount + 1
    Next k
    IsInsideString = ((quoteCount Mod 2) = 1)
End Function


' =============================================================================
' LINT-03: Magic numbers in array indexing (AP-08)
' =============================================================================
Private Sub CheckMagicNumbers(fileName As String, lines() As String)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = LTrim(lines(i))
        ' Skip comments, Dim, ReDim, Const lines
        If Left(ln, 1) = "'" Then GoTo NextMN
        If Left(ln, 3) = "Dim" Then GoTo NextMN
        If Left(ln, 5) = "ReDim" Then GoTo NextMN
        If Left(ln, 5) = "Const" Then GoTo NextMN
        If Left(ln, 9) = "Attribute" Then GoTo NextMN
        ' Skip lines that reference ColIndex (they are legitimate)
        If InStr(1, ln, "ColIndex", vbTextCompare) > 0 Then GoTo NextMN
        ' Check 1: outputs(... , <number>) pattern
        If MatchesMagicNumberPattern(ln) Then
            AddViolation fileName, i + 1, LINT_MAGIC_NUMBER, LINT_SEV_WARN, _
                "AP-08: Possible magic number in array indexing"
            GoTo NextMN
        End If
        ' Check 2: col<Name> = <number> (column variable assigned bare numeric)
        If MatchesColAssignmentMagic(ln) Then
            AddViolation fileName, i + 1, LINT_MAGIC_NUMBER, LINT_SEV_WARN, _
                "AP-08: Column variable assigned numeric literal instead of ColIndex()"
        End If
NextMN:
    Next i
End Sub

Private Function MatchesMagicNumberPattern(ln As String) As Boolean
    MatchesMagicNumberPattern = False
    ' Look for outputs( or output( followed by , <digit>)
    Dim lnLower As String
    lnLower = LCase(ln)
    Dim pos As Long
    pos = 1
    Do
        Dim foundPos As Long
        foundPos = InStr(pos, lnLower, "outputs(")
        If foundPos = 0 Then foundPos = InStr(pos, lnLower, "output(")
        If foundPos = 0 Then Exit Function
        pos = foundPos
        ' Find the comma
        Dim commaPos As Long
        commaPos = InStr(pos, ln, ",")
        If commaPos = 0 Then Exit Function
        ' Check if what follows the comma (trimmed) is a digit then )
        Dim afterComma As String
        afterComma = LTrim(Mid(ln, commaPos + 1))
        If Len(afterComma) > 0 Then
            If IsNumeric(Left(afterComma, 1)) Then
                ' Check it ends with )
                Dim j As Long
                For j = 1 To Len(afterComma)
                    Dim c As String
                    c = Mid(afterComma, j, 1)
                    If c = ")" Then
                        MatchesMagicNumberPattern = True
                        Exit Function
                    ElseIf Not IsNumeric(c) And c <> " " Then
                        Exit For
                    End If
                Next j
            End If
        End If
        pos = commaPos + 1
    Loop
End Function

Private Function MatchesColAssignmentMagic(ln As String) As Boolean
    MatchesColAssignmentMagic = False
    ' Detect patterns like: col<Name> = <number>
    ' This catches column variables assigned bare numeric literals
    Dim lnLower As String
    lnLower = LCase(LTrim(ln))
    ' Must start with "col" followed by an alpha character (not "collection" etc.)
    If Left(lnLower, 3) <> "col" Then Exit Function
    If Len(lnLower) < 4 Then Exit Function
    Dim fourthChar As String
    fourthChar = Mid(lnLower, 4, 1)
    If fourthChar < "a" Or fourthChar > "z" Then Exit Function
    ' Find the = sign (assignment)
    Dim eqPos As Long
    eqPos = InStr(1, ln, "=")
    If eqPos = 0 Then Exit Function
    ' Make sure it is not == or <= or >= (not applicable in VBA but be safe)
    If eqPos > 1 Then
        Dim before As String
        before = Mid(ln, eqPos - 1, 1)
        If before = "<" Or before = ">" Or before = "!" Then Exit Function
    End If
    ' Get the right side of the assignment
    Dim rhs As String
    rhs = Trim(Mid(ln, eqPos + 1))
    If Len(rhs) = 0 Then Exit Function
    ' Check if right side is a bare numeric literal (possibly negative)
    Dim startChar As String
    startChar = Left(rhs, 1)
    If startChar = "-" Then rhs = LTrim(Mid(rhs, 2))
    If Len(rhs) = 0 Then Exit Function
    ' Must start with a digit
    If Not IsNumeric(Left(rhs, 1)) Then Exit Function
    ' All remaining chars must be digits (allow trailing comment)
    Dim k As Long
    For k = 1 To Len(rhs)
        Dim ch As String
        ch = Mid(rhs, k, 1)
        If ch = " " Or ch = "'" Then Exit For
        If Not IsNumeric(ch) Then Exit Function
    Next k
    MatchesColAssignmentMagic = True
End Function


' =============================================================================
' LINT-04: ReDim Preserve in loop (AP-18)
' =============================================================================
Private Sub CheckReDimInLoop(fileName As String, lines() As String)
    Dim nestDepth As Long
    nestDepth = 0
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = LTrim(lines(i))
        ' Skip comments
        If Left(ln, 1) = "'" Then GoTo NextRL
        Dim lnUpper As String
        lnUpper = UCase(ln)
        ' Track nesting: For, Do, While increase; Next, Loop, Wend decrease
        If Left(lnUpper, 4) = "FOR " Or Left(lnUpper, 3) = "DO " Or _
           Left(lnUpper, 3) = "DO" & vbCr Or lnUpper = "DO" Or _
           Left(lnUpper, 6) = "WHILE " Then
            nestDepth = nestDepth + 1
        End If
        If Left(lnUpper, 5) = "NEXT " Or lnUpper = "NEXT" Or _
           Left(lnUpper, 4) = "LOOP" Or Left(lnUpper, 4) = "WEND" Then
            nestDepth = nestDepth - 1
            If nestDepth < 0 Then nestDepth = 0
        End If
        ' Check for ReDim Preserve while inside a loop
        ' Use Left() to match actual statements only (not string references)
        If nestDepth > 0 Then
            If Left(lnUpper, 15) = "REDIM PRESERVE " Then
                ' Skip if guarded by a bounds check (amortized growth pattern)
                If Not IsGuardedGrowth(lines, i) Then
                    AddViolation fileName, i + 1, LINT_REDIM_IN_LOOP, LINT_SEV_ERROR, _
                        "AP-18: ReDim Preserve inside loop (depth=" & nestDepth & ")"
                End If
            End If
        End If
NextRL:
    Next i
End Sub

Private Function IsGuardedGrowth(lines() As String, currentLine As Long) As Boolean
    ' Check if the ReDim Preserve is guarded by a bounds/size check on
    ' preceding lines (amortized growth pattern, not catastrophic O(n^2))
    IsGuardedGrowth = False
    Dim startLine As Long
    startLine = currentLine - 3
    If startLine < 0 Then startLine = 0
    Dim j As Long
    For j = startLine To currentLine - 1
        If j >= 0 And j <= UBound(lines) Then
            Dim guardLine As String
            guardLine = UCase(LTrim(lines(j)))
            If Left(guardLine, 3) = "IF " Then
                If InStr(1, guardLine, "UBOUND") > 0 Or _
                   InStr(1, guardLine, "COUNT") > 0 Then
                    IsGuardedGrowth = True
                    Exit Function
                End If
            End If
        End If
    Next j
End Function


' =============================================================================
' LINT-05: Sized Dim after executable code (AP-34)
' =============================================================================
Private Sub CheckSizedDimAfterCode(fileName As String, lines() As String)
    Dim inProcedure As Boolean
    inProcedure = False
    Dim hasExecutableCode As Boolean
    hasExecutableCode = False
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = LTrim(lines(i))
        Dim lnUpper As String
        lnUpper = UCase(ln)
        ' Skip blank lines and comments
        If Len(Trim(ln)) = 0 Then GoTo NextSD
        If Left(ln, 1) = "'" Then GoTo NextSD
        ' Detect procedure start
        If Left(lnUpper, 11) = "PUBLIC SUB " Or Left(lnUpper, 12) = "PRIVATE SUB " Or _
           Left(lnUpper, 16) = "PUBLIC FUNCTION " Or Left(lnUpper, 17) = "PRIVATE FUNCTION " Or _
           Left(lnUpper, 16) = "PUBLIC PROPERTY " Or Left(lnUpper, 17) = "PRIVATE PROPERTY " Or _
           Left(lnUpper, 4) = "SUB " Or Left(lnUpper, 9) = "FUNCTION " Then
            inProcedure = True
            hasExecutableCode = False
            GoTo NextSD
        End If
        ' Detect procedure end
        If Left(lnUpper, 7) = "END SUB" Or Left(lnUpper, 12) = "END FUNCTION" Or _
           Left(lnUpper, 12) = "END PROPERTY" Then
            inProcedure = False
            hasExecutableCode = False
            GoTo NextSD
        End If
        If Not inProcedure Then GoTo NextSD
        ' Non-executable lines: Dim, Const, Option, Attribute, labels, #If/#Else/#End If
        If Left(lnUpper, 4) = "DIM " Or Left(lnUpper, 6) = "CONST " Or _
           Left(lnUpper, 7) = "OPTION " Or Left(lnUpper, 10) = "ATTRIBUTE " Or _
           Left(lnUpper, 7) = "STATIC " Or Left(lnUpper, 3) = "#IF" Or _
           Left(lnUpper, 5) = "#ELSE" Or Left(lnUpper, 7) = "#END IF" Then
            ' Check if it is a sized Dim after executable code
            If hasExecutableCode And Left(lnUpper, 4) = "DIM " Then
                ' Check for sized array: Dim x(N) or Dim x(N To M)
                If HasSizedArrayDecl(ln) Then
                    AddViolation fileName, i + 1, LINT_SIZED_DIM_AFTER_CODE, LINT_SEV_WARN, _
                        "AP-34: Sized array Dim after executable code"
                End If
            End If
            GoTo NextSD
        End If
        ' Check for label (word followed by colon at start of line)
        If Right(Trim(ln), 1) = ":" And InStr(1, ln, " ") = 0 Then GoTo NextSD
        ' Everything else is executable code
        hasExecutableCode = True
NextSD:
    Next i
End Sub

Private Function HasSizedArrayDecl(ln As String) As Boolean
    HasSizedArrayDecl = False
    ' Only look at the Dim portion (before any colon separator)
    Dim dimPart As String
    Dim colonPos As Long
    colonPos = InStr(1, ln, ":")
    If colonPos > 0 Then
        dimPart = Left(ln, colonPos - 1)
    Else
        dimPart = ln
    End If
    ' Look for pattern like Dim varName(digits
    Dim parenPos As Long
    parenPos = InStr(1, dimPart, "(")
    If parenPos = 0 Then Exit Function
    ' Check if there is a digit after the paren
    Dim afterParen As String
    afterParen = LTrim(Mid(dimPart, parenPos + 1))
    If Len(afterParen) > 0 Then
        If IsNumeric(Left(afterParen, 1)) Then
            HasSizedArrayDecl = True
        End If
    End If
End Function


' =============================================================================
' LINT-06: 2-char vars i/m/f + uppercase (AP-35)
' =============================================================================
Private Sub CheckTwoCharVars(fileName As String, lines() As String)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = LTrim(lines(i))
        If Left(ln, 1) = "'" Then GoTo NextTV
        Dim lnUpper As String
        lnUpper = UCase(ln)
        ' Look for Dim [imf][A-Z] but skip inline Dim+assign (Dim x As T: x = expr)
        If Left(lnUpper, 4) = "DIM " Then
            ' Skip colon-separated Dim+assignment combos (compact local temps)
            If InStr(1, ln, ":") > 0 Then GoTo NextTV
            Dim afterDim As String
            afterDim = LTrim(Mid(ln, 5))
            If Len(afterDim) >= 2 Then
                Dim firstCh As String
                firstCh = Left(afterDim, 1)
                Dim secondCh As String
                secondCh = Mid(afterDim, 2, 1)
                If (firstCh = "i" Or firstCh = "m" Or firstCh = "f") Then
                    If secondCh >= "A" And secondCh <= "Z" Then
                        ' Check third char is space or As (i.e. it is a 2-char var)
                        If Len(afterDim) = 2 Or Mid(afterDim, 3, 1) = " " Then
                            AddViolation fileName, i + 1, LINT_TWO_CHAR_VAR, LINT_SEV_WARN, _
                                "AP-35: 2-char variable '" & Left(afterDim, 2) & "' matches i/m/f + uppercase pattern"
                        End If
                    End If
                End If
            End If
        End If
NextTV:
    Next i
End Sub


' =============================================================================
' LINT-07: Cumulative storage in Domain code (AP-42)
' =============================================================================
Private Sub CheckCumulativeDomain(fileName As String, lines() As String)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = LTrim(lines(i))
        If Left(ln, 1) = "'" Then GoTo NextCD
        Dim lnLower As String
        lnLower = LCase(ln)
        ' Look for running sum patterns: cumulative, running total, += pattern
        If InStr(1, lnLower, "cumulative") > 0 Or _
           InStr(1, lnLower, "runningtotal") > 0 Or _
           InStr(1, lnLower, "running_total") > 0 Then
            AddViolation fileName, i + 1, LINT_CUMULATIVE_DOMAIN, LINT_SEV_WARN, _
                "AP-42: Possible cumulative storage in domain code"
        End If
NextCD:
    Next i
End Sub


' =============================================================================
' LINT-08: Missing contract functions in Domain*.bas (AP-43)
' =============================================================================
Private Sub CheckMissingContract(fileName As String, lines() As String)
    Dim hasInit As Boolean
    Dim hasValidate As Boolean
    Dim hasReset As Boolean
    Dim hasExecute As Boolean
    hasInit = False
    hasValidate = False
    hasReset = False
    hasExecute = False

    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = UCase(LTrim(lines(i)))
        If InStr(1, ln, "PUBLIC SUB INITIALIZE") > 0 Then hasInit = True
        If InStr(1, ln, "PUBLIC FUNCTION VALIDATE") > 0 Then hasValidate = True
        If InStr(1, ln, "PUBLIC SUB RESET") > 0 Then hasReset = True
        If InStr(1, ln, "PUBLIC SUB EXECUTE") > 0 Then hasExecute = True
    Next i

    If Not hasInit Then
        AddViolation fileName, 0, LINT_MISSING_CONTRACT, LINT_SEV_ERROR, _
            "AP-43: Missing Public Sub Initialize()"
    End If
    If Not hasValidate Then
        AddViolation fileName, 0, LINT_MISSING_CONTRACT, LINT_SEV_ERROR, _
            "AP-43: Missing Public Function Validate() As Boolean"
    End If
    If Not hasReset Then
        AddViolation fileName, 0, LINT_MISSING_CONTRACT, LINT_SEV_ERROR, _
            "AP-43: Missing Public Sub Reset()"
    End If
    If Not hasExecute Then
        AddViolation fileName, 0, LINT_MISSING_CONTRACT, LINT_SEV_ERROR, _
            "AP-43: Missing Public Sub Execute()"
    End If
End Sub


' =============================================================================
' LINT-09: = string without NumberFormat (AP-50)
' =============================================================================
Private Sub CheckEqStringNoFormat(fileName As String, lines() As String)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = LTrim(lines(i))
        If Left(ln, 1) = "'" Then GoTo NextEQ
        ' Look for .Value = "= (only when .Value is actual code, not inside a string)
        Dim dotValPos As Long
        dotValPos = InStr(1, ln, ".Value", vbTextCompare)
        If dotValPos > 0 Then
            ' Skip if .Value is inside a string (odd number of quotes before it)
            If Not IsInsideString(ln, dotValPos) Then
                If InStr(1, ln, """=""", vbTextCompare) > 0 Then
                    If Not HasNumberFormatGuard(lines, i) Then
                        AddViolation fileName, i + 1, LINT_EQ_NO_FORMAT, LINT_SEV_WARN, _
                            "AP-50: .Value write with ""="" without preceding NumberFormat = ""@"""
                    End If
                End If
            End If
        End If
NextEQ:
    Next i
End Sub


' =============================================================================
' LINT-10: Sub/Function balance
' =============================================================================
Private Sub CheckSubFuncBalance(fileName As String, lines() As String)
    Dim subCount As Long
    Dim endSubCount As Long
    Dim funcCount As Long
    Dim endFuncCount As Long
    subCount = 0
    endSubCount = 0
    funcCount = 0
    endFuncCount = 0

    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = UCase(LTrim(lines(i)))
        ' Skip comments
        If Left(ln, 1) = "'" Then GoTo NextSF
        ' Count Sub declarations (Public Sub, Private Sub, Sub)
        If Left(ln, 11) = "PUBLIC SUB " Or Left(ln, 12) = "PRIVATE SUB " Or _
           Left(ln, 4) = "SUB " Then
            subCount = subCount + 1
        End If
        If Left(ln, 7) = "END SUB" Then endSubCount = endSubCount + 1
        ' Count Function declarations
        If Left(ln, 16) = "PUBLIC FUNCTION " Or Left(ln, 17) = "PRIVATE FUNCTION " Or _
           Left(ln, 9) = "FUNCTION " Then
            funcCount = funcCount + 1
        End If
        If Left(ln, 12) = "END FUNCTION" Then endFuncCount = endFuncCount + 1
NextSF:
    Next i

    If subCount <> endSubCount Then
        AddViolation fileName, 0, LINT_SUB_FUNC_BALANCE, LINT_SEV_ERROR, _
            "Sub/End Sub mismatch: " & subCount & " Sub vs " & endSubCount & " End Sub"
    End If
    If funcCount <> endFuncCount Then
        AddViolation fileName, 0, LINT_SUB_FUNC_BALANCE, LINT_SEV_ERROR, _
            "Function/End Function mismatch: " & funcCount & " Function vs " & endFuncCount & " End Function"
    End If
End Sub


' =============================================================================
' LINT-11: Module size check
' =============================================================================
Private Sub CheckModuleSize(fileName As String, fileSize As Long, errorsOnly As Boolean)
    If fileSize > MODULE_SIZE_ERROR Then
        AddViolation fileName, 0, LINT_MODULE_SIZE, LINT_SEV_ERROR, _
            "Module size " & fileSize & " bytes exceeds 64KB hard limit"
    ElseIf fileSize > MODULE_SIZE_WARN And Not errorsOnly Then
        AddViolation fileName, 0, LINT_MODULE_SIZE, LINT_SEV_WARN, _
            "Module size " & fileSize & " bytes exceeds 50KB WARN threshold"
    End If
End Sub


' =============================================================================
' LINT-12: CRLF line endings check
' =============================================================================
Private Sub CheckCRLF(fileName As String, rawBytes() As Byte, fileSize As Long)
    ' Check for at least one CRLF. If file has LF but no CRLF, flag it.
    Dim hasCRLF As Boolean
    Dim hasLF As Boolean
    hasCRLF = False
    hasLF = False
    Dim i As Long
    For i = 0 To fileSize - 1
        If rawBytes(i) = 10 Then
            hasLF = True
            If i > 0 Then
                If rawBytes(i - 1) = 13 Then
                    hasCRLF = True
                    Exit For
                End If
            End If
        End If
    Next i
    If hasLF And Not hasCRLF Then
        AddViolation fileName, 0, LINT_CRLF, LINT_SEV_ERROR, _
            "File uses LF line endings instead of CRLF (VBA import will fail)"
    End If
End Sub


' =============================================================================
' LintVBAModules
' Scans live VBA project modules for code-content checks.
' Requires "Trust access to the VBA project object model" in Excel settings.
' Returns the number of VBA modules scanned (0 if access denied).
' =============================================================================
Private Function LintVBAModules(errorsOnly As Boolean) As Long
    On Error Resume Next

    LintVBAModules = 0

    Dim vbProj As Object
    Set vbProj = ThisWorkbook.VBProject
    If vbProj Is Nothing Then Exit Function
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    Dim vbComp As Object
    Dim moduleCount As Long
    moduleCount = 0

    For Each vbComp In vbProj.VBComponents
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        ' Type 1 = standard module (vbext_ct_StdModule)
        If vbComp.Type = 1 Then
            Dim codeModule As Object
            Set codeModule = vbComp.CodeModule
            If Not codeModule Is Nothing Then
                If codeModule.CountOfLines > 0 Then
                    moduleCount = moduleCount + 1
                    Dim codeText As String
                    codeText = codeModule.Lines(1, codeModule.CountOfLines)
                    LintVBAModule "[VBA] " & vbComp.Name & ".bas", codeText, errorsOnly
                End If
            End If
        End If
    Next vbComp

    LintVBAModules = moduleCount
    On Error GoTo 0
End Function


' =============================================================================
' LintVBAModule
' Runs code-content checks on a single VBA module's source code.
' Skips LINT-11 (file size) and LINT-12 (CRLF) which are disk-only checks.
' =============================================================================
Private Sub LintVBAModule(fileName As String, codeText As String, errorsOnly As Boolean)
    On Error Resume Next

    ' Determine if this is a domain module (strip [VBA] prefix for detection)
    Dim bareFileName As String
    bareFileName = fileName
    If Left(bareFileName, 6) = "[VBA] " Then bareFileName = Mid(bareFileName, 7)
    Dim isDomainModule As Boolean
    isDomainModule = ((Left(LCase(bareFileName), 6) = "domain") Or _
                      (Left(LCase(bareFileName), 12) = "sampledomain")) And _
                     (InStr(1, LCase(bareFileName), "tests") = 0)

    ' LINT-01: Non-ASCII via AscW (works on VBA Unicode strings)
    CheckNonASCIIString fileName, codeText

    ' Split into lines for other checks
    Dim lines() As String
    lines = Split(Replace(codeText, vbCrLf, vbLf), vbLf)

    ' LINT-10: Sub/Function balance (ERROR)
    CheckSubFuncBalance fileName, lines

    ' LINT-04: ReDim Preserve in loop (ERROR)
    CheckReDimInLoop fileName, lines

    ' LINT-08: Missing contract functions -- Domain only (ERROR)
    If isDomainModule Then
        CheckMissingContract fileName, lines
    End If

    ' WARN-level checks (skip in errorsOnly mode)
    If Not errorsOnly Then
        ' LINT-02: Formula injection in .Value (WARN)
        CheckFormulaInject fileName, lines

        ' LINT-03: Magic numbers in array indexing (WARN)
        CheckMagicNumbers fileName, lines

        ' LINT-05: Sized Dim after executable code (WARN)
        CheckSizedDimAfterCode fileName, lines

        ' LINT-06: 2-char vars i/m/f + uppercase (WARN)
        CheckTwoCharVars fileName, lines

        ' LINT-07: Cumulative storage in Domain code (WARN)
        If isDomainModule Then
            CheckCumulativeDomain fileName, lines
        End If

        ' LINT-09: = string without NumberFormat (WARN)
        CheckEqStringNoFormat fileName, lines
    End If

    On Error GoTo 0
End Sub


' =============================================================================
' CheckNonASCIIString
' String-based non-ASCII check for VBA module code (no raw bytes available).
' =============================================================================
Private Sub CheckNonASCIIString(fileName As String, codeText As String)
    Dim i As Long
    Dim lineNum As Long
    lineNum = 1
    For i = 1 To Len(codeText)
        Dim ch As Long
        ch = AscW(Mid(codeText, i, 1))
        If ch > 127 Or ch < 0 Then
            AddViolation fileName, lineNum, LINT_NON_ASCII, LINT_SEV_ERROR, _
                "Non-ASCII character (code " & ch & ") at position " & i
            ' Only report first occurrence per module to avoid flood
            Exit Sub
        End If
        If ch = 10 Then lineNum = lineNum + 1
    Next i
End Sub


' =============================================================================
' Violation array management
' =============================================================================
Private Sub ResetViolations()
    m_vCount = 0
    ReDim m_vFiles(1 To 100)
    ReDim m_vLines(1 To 100)
    ReDim m_vCheckIDs(1 To 100)
    ReDim m_vSeverities(1 To 100)
    ReDim m_vDescriptions(1 To 100)
End Sub

Private Sub AddViolation(fileName As String, lineNum As Long, checkID As String, _
                          severity As String, description As String)
    m_vCount = m_vCount + 1
    ' Grow arrays if needed (double the size)
    If m_vCount > UBound(m_vFiles) Then
        Dim newSize As Long
        newSize = UBound(m_vFiles) * 2
        ReDim Preserve m_vFiles(1 To newSize)
        ReDim Preserve m_vLines(1 To newSize)
        ReDim Preserve m_vCheckIDs(1 To newSize)
        ReDim Preserve m_vSeverities(1 To newSize)
        ReDim Preserve m_vDescriptions(1 To newSize)
    End If
    m_vFiles(m_vCount) = fileName
    m_vLines(m_vCount) = lineNum
    m_vCheckIDs(m_vCount) = checkID
    m_vSeverities(m_vCount) = severity
    m_vDescriptions(m_vCount) = description
End Sub

Private Sub CountBySeverity(ByRef errCount As Long, ByRef warnCount As Long)
    errCount = 0
    warnCount = 0
    Dim i As Long
    For i = 1 To m_vCount
        If m_vSeverities(i) = LINT_SEV_ERROR Then
            errCount = errCount + 1
        Else
            warnCount = warnCount + 1
        End If
    Next i
End Sub


' =============================================================================
' WriteLintResults
' Writes lint results to the TestResults sheet above existing content (newest on top).
' Inserts rows at row 3 (below title + run-info), pushing test results down.
' =============================================================================
Private Sub WriteLintResults()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(TAB_TEST_RESULTS)
    If ws Is Nothing Then Exit Sub
    On Error GoTo 0

    ' Calculate how many rows we need for the lint section
    Dim dataRows As Long
    If m_vCount = 0 Then
        dataRows = 1  ' "No violations found" row
    Else
        dataRows = m_vCount
    End If
    Dim rowsNeeded As Long
    rowsNeeded = 2 + dataRows + 1  ' section header + col header + data + spacer

    ' Insert blank rows at row 3 (below merged title at row 1 and run-info at row 2)
    Dim insertAt As Long
    insertAt = 3
    ws.Rows(insertAt & ":" & (insertAt + rowsNeeded - 1)).Insert Shift:=xlDown

    ' Write section header
    Dim startRow As Long
    startRow = insertAt
    ws.Cells(startRow, 1).NumberFormat = "@"
    ws.Cells(startRow, 1).Value = "=== LINT RESULTS === " & Format(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(startRow, 1).Font.Bold = True
    startRow = startRow + 1

    ' Write column headers
    ws.Cells(startRow, 1).Value = "File"
    ws.Cells(startRow, 2).Value = "Line"
    ws.Cells(startRow, 3).Value = "CheckID"
    ws.Cells(startRow, 4).Value = "Severity"
    ws.Cells(startRow, 5).Value = "Description"
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 5))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    startRow = startRow + 1

    If m_vCount = 0 Then
        ws.Cells(startRow, 1).Value = "No violations found."
        Exit Sub
    End If

    ' Write violations using array batch write (PT-001)
    Dim resultArr() As Variant
    ReDim resultArr(1 To m_vCount, 1 To 5)
    Dim i As Long
    For i = 1 To m_vCount
        resultArr(i, 1) = m_vFiles(i)
        resultArr(i, 2) = m_vLines(i)
        resultArr(i, 3) = m_vCheckIDs(i)
        resultArr(i, 4) = m_vSeverities(i)
        resultArr(i, 5) = m_vDescriptions(i)
    Next i

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + m_vCount - 1, 5)).Value = resultArr

    ' Color-code severity column
    For i = 1 To m_vCount
        Dim dataRow As Long
        dataRow = startRow + i - 1
        If m_vSeverities(i) = LINT_SEV_ERROR Then
            ws.Cells(dataRow, 4).Interior.Color = RGB(255, 199, 206)
            ws.Cells(dataRow, 4).Font.Color = RGB(156, 0, 6)
            ws.Cells(dataRow, 4).Font.Bold = True
        Else
            ws.Cells(dataRow, 4).Interior.Color = RGB(255, 235, 156)
            ws.Cells(dataRow, 4).Font.Color = RGB(156, 101, 0)
        End If
    Next i
End Sub
