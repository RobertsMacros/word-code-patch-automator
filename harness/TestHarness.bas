Attribute VB_Name = "TestHarness"
'===========================================================================
' TestHarness — MVP VBA Test Runner
'
' Runs all registered tests, writes JSON results to the path specified
' by the TEST_RESULT_PATH environment variable (set by the controller).
'
' Test functions must:
'   - Be Public Subs in modules whose names start with "Test_"
'   - Follow the naming pattern: Sub Test_<category>_<name>()
'   - Call Assert* helpers to record pass/fail
'
' Three test layers:
'   1. Fast logic tests    — pure VBA logic, no document access
'   2. Regression tests    — test against known-bug fixtures
'   3. Integration tests   — full Word document operations
'===========================================================================
Option Explicit

' --- Internal state ---
Private p_Results As Collection  ' Collection of test result dictionaries
Private p_CurrentTest As String
Private p_CurrentPassed As Boolean
Private p_CurrentMessage As String
Private p_TestStartTime As Double
Private p_RunStartTime As Double

'===========================================================================
' Entry point — called by the controller
'===========================================================================
Public Sub RunAllTests()
    Dim resultPath As String
    resultPath = Environ("TEST_RESULT_PATH")

    ' Fallback: read from config file next to result location
    If resultPath = "" Then
        resultPath = GetConfigResultPath()
    End If

    If resultPath = "" Then
        MsgBox "TEST_RESULT_PATH not set. Cannot write results.", vbCritical
        Exit Sub
    End If

    ' Initialize
    Set p_Results = New Collection
    p_RunStartTime = Timer

    ' Discover and run test modules
    Dim comp As Object  ' VBComponent
    Dim proj As Object  ' VBProject
    Set proj = ThisDocument.VBProject

    Dim testCount As Long
    testCount = 0

    Dim i As Long
    For i = 1 To proj.VBComponents.Count
        Set comp = proj.VBComponents(i)
        ' Run modules whose names start with "Test_"
        If Left$(comp.Name, 5) = "Test_" Then
            RunTestModule comp
        End If
    Next i

    ' Also run any tests registered in this module (smoke tests)
    RunSmokeTests

    ' Calculate elapsed
    Dim totalElapsed As Double
    totalElapsed = Timer - p_RunStartTime

    ' Write results
    WriteResultsJSON resultPath, totalElapsed
End Sub

'===========================================================================
' Run all Public Subs in a test module
'===========================================================================
Private Sub RunTestModule(comp As Object)
    ' Parse the code module for Public Sub Test_* declarations
    Dim codeModule As Object
    Set codeModule = comp.CodeModule

    Dim lineNum As Long
    Dim lineText As String
    Dim subName As String

    For lineNum = 1 To codeModule.CountOfLines
        lineText = Trim$(codeModule.Lines(lineNum, 1))

        ' Look for "Public Sub Test_..." or "Sub Test_..."
        subName = ExtractTestSubName(lineText)
        If subName <> "" Then
            RunSingleTest comp.Name & "." & subName, subName
        End If
    Next lineNum
End Sub

'===========================================================================
' Extract sub name from a line like "Public Sub Test_Foo()"
'===========================================================================
Private Function ExtractTestSubName(lineText As String) As String
    Dim cleaned As String
    cleaned = Trim$(lineText)

    ' Remove "Public " prefix if present
    If Left$(LCase$(cleaned), 7) = "public " Then
        cleaned = Trim$(Mid$(cleaned, 8))
    End If

    ' Must start with "Sub Test_"
    If Left$(LCase$(cleaned), 9) <> "sub test_" Then
        ExtractTestSubName = ""
        Exit Function
    End If

    ' Extract the sub name (everything up to the opening paren or end of line)
    Dim subPart As String
    subPart = Mid$(cleaned, 5) ' skip "Sub "

    Dim parenPos As Long
    parenPos = InStr(subPart, "(")
    If parenPos > 0 Then
        subPart = Left$(subPart, parenPos - 1)
    End If

    ExtractTestSubName = Trim$(subPart)
End Function

'===========================================================================
' Run a single test sub by name
'===========================================================================
Private Sub RunSingleTest(displayName As String, subName As String)
    p_CurrentTest = displayName
    p_CurrentPassed = True  ' Assume pass unless Assert fails
    p_CurrentMessage = ""
    p_TestStartTime = Timer

    Dim elapsed As Double
    Dim timedOut As Boolean
    timedOut = False

    On Error GoTo TestError
    Application.Run subName
    GoTo TestDone

TestError:
    p_CurrentPassed = False
    p_CurrentMessage = "Runtime error " & Err.Number & ": " & Err.Description
    Resume TestDone

TestDone:
    On Error GoTo 0
    elapsed = Timer - p_TestStartTime

    ' Record result
    Dim result As String
    result = "{"
    result = result & """name"": " & JsonStr(p_CurrentTest) & ", "
    result = result & """passed"": " & IIf(p_CurrentPassed, "true", "false") & ", "
    result = result & """message"": " & JsonStr(p_CurrentMessage) & ", "
    result = result & """duration_ms"": " & Format$(elapsed * 1000, "0.0") & ", "
    result = result & """timed_out"": " & IIf(timedOut, "true", "false")
    result = result & "}"

    p_Results.Add result
End Sub

'===========================================================================
' Built-in smoke tests
'===========================================================================
Private Sub RunSmokeTests()
    ' Smoke test: VBA environment is functional
    p_CurrentTest = "Smoke_VBAEnvironment"
    p_CurrentPassed = True
    p_CurrentMessage = ""
    p_TestStartTime = Timer

    ' Simple sanity check
    Dim x As Long
    x = 1 + 1
    If x <> 2 Then
        p_CurrentPassed = False
        p_CurrentMessage = "Basic arithmetic failed"
    End If

    Dim elapsed As Double
    elapsed = Timer - p_TestStartTime

    Dim result As String
    result = "{"
    result = result & """name"": ""Smoke_VBAEnvironment"", "
    result = result & """passed"": " & IIf(p_CurrentPassed, "true", "false") & ", "
    result = result & """message"": " & JsonStr(p_CurrentMessage) & ", "
    result = result & """duration_ms"": " & Format$(elapsed * 1000, "0.0") & ", "
    result = result & """timed_out"": false"
    result = result & "}"
    p_Results.Add result

    ' Smoke test: Document is accessible
    p_CurrentTest = "Smoke_DocumentAccess"
    p_CurrentPassed = True
    p_CurrentMessage = ""
    p_TestStartTime = Timer

    On Error Resume Next
    Dim docName As String
    docName = ThisDocument.Name
    If Err.Number <> 0 Then
        p_CurrentPassed = False
        p_CurrentMessage = "Cannot access ThisDocument: " & Err.Description
    End If
    On Error GoTo 0

    elapsed = Timer - p_TestStartTime

    result = "{"
    result = result & """name"": ""Smoke_DocumentAccess"", "
    result = result & """passed"": " & IIf(p_CurrentPassed, "true", "false") & ", "
    result = result & """message"": " & JsonStr(p_CurrentMessage) & ", "
    result = result & """duration_ms"": " & Format$(elapsed * 1000, "0.0") & ", "
    result = result & """timed_out"": false"
    result = result & "}"
    p_Results.Add result
End Sub

'===========================================================================
' Assert helpers — call these from test subs
'===========================================================================
Public Sub AssertTrue(condition As Boolean, Optional msg As String = "")
    If Not condition Then
        p_CurrentPassed = False
        If msg <> "" Then
            p_CurrentMessage = msg
        Else
            p_CurrentMessage = "AssertTrue failed"
        End If
    End If
End Sub

Public Sub AssertFalse(condition As Boolean, Optional msg As String = "")
    AssertTrue Not condition, IIf(msg <> "", msg, "AssertFalse failed")
End Sub

Public Sub AssertEqual(expected As Variant, actual As Variant, Optional msg As String = "")
    If expected <> actual Then
        p_CurrentPassed = False
        If msg <> "" Then
            p_CurrentMessage = msg
        Else
            p_CurrentMessage = "Expected [" & CStr(expected) & "] but got [" & CStr(actual) & "]"
        End If
    End If
End Sub

Public Sub AssertNotEqual(notExpected As Variant, actual As Variant, Optional msg As String = "")
    If notExpected = actual Then
        p_CurrentPassed = False
        If msg <> "" Then
            p_CurrentMessage = msg
        Else
            p_CurrentMessage = "Expected value to differ from [" & CStr(notExpected) & "]"
        End If
    End If
End Sub

Public Sub AssertContains(haystack As String, needle As String, Optional msg As String = "")
    If InStr(1, haystack, needle, vbTextCompare) = 0 Then
        p_CurrentPassed = False
        If msg <> "" Then
            p_CurrentMessage = msg
        Else
            p_CurrentMessage = "String does not contain [" & needle & "]"
        End If
    End If
End Sub

Public Sub Fail(msg As String)
    p_CurrentPassed = False
    p_CurrentMessage = msg
End Sub

'===========================================================================
' JSON output
'===========================================================================
Private Sub WriteResultsJSON(filePath As String, totalElapsed As Double)
    Dim fNum As Integer
    fNum = FreeFile

    Open filePath For Output As #fNum

    Print #fNum, "{"
    Print #fNum, "  ""harness_version"": ""mvp-1.0"","
    Print #fNum, "  ""timestamp"": """ & Format$(Now, "yyyy-mm-dd hh:nn:ss") & ""","
    Print #fNum, "  ""total_elapsed_seconds"": " & Format$(totalElapsed, "0.000") & ","
    Print #fNum, "  ""test_count"": " & p_Results.Count & ","

    ' Count pass/fail
    Dim passCount As Long, failCount As Long
    Dim r As Variant
    passCount = 0: failCount = 0
    Dim idx As Long
    For idx = 1 To p_Results.Count
        If InStr(1, p_Results(idx), """passed"": true") > 0 Then
            passCount = passCount + 1
        Else
            failCount = failCount + 1
        End If
    Next idx

    Print #fNum, "  ""passed"": " & passCount & ","
    Print #fNum, "  ""failed"": " & failCount & ","
    Print #fNum, "  ""tests"": ["

    For idx = 1 To p_Results.Count
        If idx < p_Results.Count Then
            Print #fNum, "    " & p_Results(idx) & ","
        Else
            Print #fNum, "    " & p_Results(idx)
        End If
    Next idx

    Print #fNum, "  ]"
    Print #fNum, "}"

    Close #fNum
End Sub

'===========================================================================
' Helpers
'===========================================================================
Private Function JsonStr(s As String) As String
    ' Escape a string for JSON
    Dim escaped As String
    escaped = Replace(s, "\", "\\")
    escaped = Replace(escaped, """", "\""")
    escaped = Replace(escaped, vbCr, "\r")
    escaped = Replace(escaped, vbLf, "\n")
    escaped = Replace(escaped, vbTab, "\t")
    JsonStr = """" & escaped & """"
End Function

Private Function GetConfigResultPath() As String
    ' Try to read result path from a config file next to the document
    Dim configPath As String
    configPath = ThisDocument.Path & "\_harness_config.json"

    On Error Resume Next
    Dim fNum As Integer
    fNum = FreeFile
    Open configPath For Input As #fNum
    If Err.Number <> 0 Then
        GetConfigResultPath = ""
        Exit Function
    End If

    Dim content As String
    Dim lineText As String
    Do Until EOF(fNum)
        Line Input #fNum, lineText
        content = content & lineText
    Loop
    Close #fNum
    On Error GoTo 0

    ' Simple JSON parse for "result_file": "..."
    Dim startPos As Long, endPos As Long
    startPos = InStr(1, content, """result_file""")
    If startPos = 0 Then
        GetConfigResultPath = ""
        Exit Function
    End If

    startPos = InStr(startPos, content, ":") + 1
    ' Find the opening quote
    startPos = InStr(startPos, content, """") + 1
    endPos = InStr(startPos, content, """")

    If startPos > 1 And endPos > startPos Then
        GetConfigResultPath = Mid$(content, startPos, endPos - startPos)
    Else
        GetConfigResultPath = ""
    End If
End Function
