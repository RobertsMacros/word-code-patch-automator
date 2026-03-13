Attribute VB_Name = "TestHarness"
'===========================================================================
' TestHarness — MVP VBA Test Runner with Fixture Support
'
' Called by the controller. Runs all Test_* modules, including:
'   - Smoke tests (no fixture needed)
'   - Logic tests (no fixture needed)
'   - Fixture-based tests (iterated over each fixture document)
'
' Config is read from:
'   1. Environ("TEST_RESULT_PATH") — where to write results
'   2. Environ("TEST_HARNESS_CONFIG") — path to JSON with fixture list
'   Fallback: %TEMP%\_harness_config.json
'
' Test naming conventions:
'   Sub Test_<Category>_<Name>()       — runs once (no fixture)
'   Sub TestFixture_<Category>_<Name>(doc As Document)
'       — runs once per fixture document
'===========================================================================
Option Explicit

' --- Internal state ---
Private p_Results As Collection
Private p_FixtureResults As Collection
Private p_CurrentTest As String
Private p_CurrentPassed As Boolean
Private p_CurrentMessage As String
Private p_TestStartTime As Double
Private p_RunStartTime As Double
Private p_CurrentFixture As String

'===========================================================================
' Entry point
'===========================================================================
Public Sub RunAllTests()
    Dim resultPath As String
    Dim configPath As String
    Dim fixtures() As String
    Dim fixtureCount As Long

    ' Resolve result path
    resultPath = Environ("TEST_RESULT_PATH")
    If resultPath = "" Then
        resultPath = ReadConfigValue("result_file")
    End If
    If resultPath = "" Then
        MsgBox "TEST_RESULT_PATH not set and no config found.", vbCritical
        Exit Sub
    End If

    ' Resolve fixture list
    fixtures = ReadFixtureList()
    fixtureCount = UBound(fixtures) - LBound(fixtures) + 1
    If fixtureCount = 1 And fixtures(0) = "" Then fixtureCount = 0

    ' Initialize
    Set p_Results = New Collection
    Set p_FixtureResults = New Collection
    p_RunStartTime = Timer
    p_CurrentFixture = ""

    ' --- Phase 1: Run non-fixture tests (smoke, logic) ---
    RunSmokeTests
    RunNonFixtureTests

    ' --- Phase 2: Run fixture-based tests ---
    If fixtureCount > 0 Then
        Dim i As Long
        For i = LBound(fixtures) To UBound(fixtures)
            If fixtures(i) <> "" Then
                RunFixtureTests fixtures(i)
            End If
        Next i
    End If

    ' --- Write results ---
    Dim totalElapsed As Double
    totalElapsed = Timer - p_RunStartTime
    WriteResultsJSON resultPath, totalElapsed, fixtures
End Sub

'===========================================================================
' Run all non-fixture Test_* subs across all Test_* modules
'===========================================================================
Private Sub RunNonFixtureTests()
    Dim proj As Object
    Set proj = ThisDocument.VBProject

    Dim i As Long
    For i = 1 To proj.VBComponents.Count
        Dim comp As Object
        Set comp = proj.VBComponents(i)
        If Left$(comp.Name, 5) = "Test_" Then
            RunNonFixtureTestsInModule comp
        End If
    Next i
End Sub

Private Sub RunNonFixtureTestsInModule(comp As Object)
    Dim codeModule As Object
    Set codeModule = comp.CodeModule

    Dim lineNum As Long
    For lineNum = 1 To codeModule.CountOfLines
        Dim lineText As String
        lineText = Trim$(codeModule.Lines(lineNum, 1))

        Dim subName As String
        subName = ExtractSubName(lineText, "Test_")

        ' Skip fixture-based subs (TestFixture_*)
        If subName <> "" And Left$(subName, 12) <> "TestFixture_" Then
            RunSingleTest comp.Name & "." & subName, subName, ""
        End If
    Next lineNum
End Sub

'===========================================================================
' Run fixture-based tests for one fixture document
'===========================================================================
Private Sub RunFixtureTests(fixturePath As String)
    Dim fixtureDoc As Document
    Dim fixtureName As String
    fixtureName = Mid$(fixturePath, InStrRev(fixturePath, "\") + 1)

    Dim fixtureStart As Double
    fixtureStart = Timer

    Dim fixtureError As String
    fixtureError = ""

    ' Open the fixture document (read-only to keep it immutable)
    On Error Resume Next
    Set fixtureDoc = Application.Documents.Open( _
        FileName:=fixturePath, _
        ReadOnly:=True, _
        AddToRecentFiles:=False)
    If Err.Number <> 0 Then
        fixtureError = "Cannot open fixture: " & Err.Description
        On Error GoTo 0

        ' Record a single failure for this fixture
        RecordResult "Fixture_Open." & fixtureName, False, fixtureError, 0, False, fixtureName
        RecordFixtureResult fixtureName, 0, 0, 1, Timer - fixtureStart, fixtureError
        Exit Sub
    End If
    On Error GoTo 0

    p_CurrentFixture = fixtureName

    ' Find and run all TestFixture_* subs, passing the fixture doc
    Dim proj As Object
    Set proj = ThisDocument.VBProject
    Dim passCount As Long, failCount As Long
    passCount = 0: failCount = 0

    Dim i As Long
    For i = 1 To proj.VBComponents.Count
        Dim comp As Object
        Set comp = proj.VBComponents(i)
        If Left$(comp.Name, 5) = "Test_" Then
            RunFixtureTestsInModule comp, fixtureDoc, fixtureName, passCount, failCount
        End If
    Next i

    ' Close fixture document without saving
    On Error Resume Next
    fixtureDoc.Close SaveChanges:=False
    On Error GoTo 0

    p_CurrentFixture = ""

    Dim fixtureElapsed As Double
    fixtureElapsed = Timer - fixtureStart
    RecordFixtureResult fixtureName, passCount, failCount, passCount + failCount, fixtureElapsed, ""
End Sub

Private Sub RunFixtureTestsInModule(comp As Object, fixtureDoc As Document, fixtureName As String, ByRef passCount As Long, ByRef failCount As Long)
    Dim codeModule As Object
    Set codeModule = comp.CodeModule

    Dim lineNum As Long
    For lineNum = 1 To codeModule.CountOfLines
        Dim lineText As String
        lineText = Trim$(codeModule.Lines(lineNum, 1))

        Dim subName As String
        subName = ExtractSubName(lineText, "TestFixture_")
        If subName <> "" Then
            ' Run the fixture test, passing the doc as argument
            Dim beforeCount As Long
            beforeCount = p_Results.Count

            RunSingleFixtureTest comp.Name & "." & subName, subName, fixtureDoc, fixtureName

            ' Check if it passed (look at the last added result)
            If p_Results.Count > beforeCount Then
                Dim lastResult As String
                lastResult = p_Results(p_Results.Count)
                If InStr(1, lastResult, """passed"": true") > 0 Then
                    passCount = passCount + 1
                Else
                    failCount = failCount + 1
                End If
            End If
        End If
    Next lineNum
End Sub

'===========================================================================
' Run a single non-fixture test
'===========================================================================
Private Sub RunSingleTest(displayName As String, subName As String, fixtureName As String)
    p_CurrentTest = displayName
    p_CurrentPassed = True
    p_CurrentMessage = ""
    p_TestStartTime = Timer

    On Error GoTo TestErr
    Application.Run subName
    GoTo TestEnd

TestErr:
    p_CurrentPassed = False
    p_CurrentMessage = "Runtime error " & Err.Number & ": " & Err.Description
    Resume TestEnd

TestEnd:
    On Error GoTo 0
    Dim elapsed As Double
    elapsed = Timer - p_TestStartTime
    RecordResult p_CurrentTest, p_CurrentPassed, p_CurrentMessage, elapsed * 1000, False, fixtureName
End Sub

'===========================================================================
' Run a single fixture test (passes Document argument)
'===========================================================================
Private Sub RunSingleFixtureTest(displayName As String, subName As String, fixtureDoc As Document, fixtureName As String)
    p_CurrentTest = displayName
    p_CurrentPassed = True
    p_CurrentMessage = ""
    p_TestStartTime = Timer

    On Error GoTo FTestErr
    Application.Run subName, fixtureDoc
    GoTo FTestEnd

FTestErr:
    p_CurrentPassed = False
    p_CurrentMessage = "Runtime error " & Err.Number & ": " & Err.Description
    Resume FTestEnd

FTestEnd:
    On Error GoTo 0
    Dim elapsed As Double
    elapsed = Timer - p_TestStartTime
    RecordResult p_CurrentTest, p_CurrentPassed, p_CurrentMessage, elapsed * 1000, False, fixtureName
End Sub

'===========================================================================
' Built-in smoke tests
'===========================================================================
Private Sub RunSmokeTests()
    ' Smoke: VBA environment
    p_CurrentTest = "Smoke_VBAEnvironment"
    p_CurrentPassed = True
    p_CurrentMessage = ""
    p_TestStartTime = Timer
    Dim x As Long: x = 1 + 1
    If x <> 2 Then
        p_CurrentPassed = False
        p_CurrentMessage = "Basic arithmetic failed"
    End If
    RecordResult "Smoke_VBAEnvironment", p_CurrentPassed, p_CurrentMessage, (Timer - p_TestStartTime) * 1000, False, ""

    ' Smoke: host document accessible
    p_CurrentTest = "Smoke_HostDocument"
    p_CurrentPassed = True
    p_CurrentMessage = ""
    p_TestStartTime = Timer
    On Error Resume Next
    Dim n As String: n = ThisDocument.Name
    If Err.Number <> 0 Then
        p_CurrentPassed = False
        p_CurrentMessage = "Cannot access ThisDocument: " & Err.Description
    End If
    On Error GoTo 0
    RecordResult "Smoke_HostDocument", p_CurrentPassed, p_CurrentMessage, (Timer - p_TestStartTime) * 1000, False, ""
End Sub

'===========================================================================
' Assert helpers
'===========================================================================
Public Sub AssertTrue(condition As Boolean, Optional msg As String = "")
    If Not condition Then
        p_CurrentPassed = False
        p_CurrentMessage = IIf(msg <> "", msg, "AssertTrue failed")
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
' Result recording
'===========================================================================
Private Sub RecordResult(testName As String, passed As Boolean, message As String, durationMs As Double, timedOut As Boolean, fixtureName As String)
    Dim r As String
    r = "{"
    r = r & """name"": " & JsonStr(testName) & ", "
    r = r & """passed"": " & IIf(passed, "true", "false") & ", "
    r = r & """message"": " & JsonStr(message) & ", "
    r = r & """duration_ms"": " & Format$(durationMs, "0.0") & ", "
    r = r & """timed_out"": " & IIf(timedOut, "true", "false") & ", "
    r = r & """fixture"": " & JsonStr(fixtureName)
    r = r & "}"
    p_Results.Add r
End Sub

Private Sub RecordFixtureResult(fixtureName As String, passed As Long, failed As Long, total As Long, elapsedSec As Double, errorMsg As String)
    Dim r As String
    r = "{"
    r = r & """fixture"": " & JsonStr(fixtureName) & ", "
    r = r & """passed"": " & passed & ", "
    r = r & """failed"": " & failed & ", "
    r = r & """total"": " & total & ", "
    r = r & """elapsed_seconds"": " & Format$(elapsedSec, "0.000") & ", "
    r = r & """error"": " & JsonStr(errorMsg)
    r = r & "}"
    p_FixtureResults.Add r
End Sub

'===========================================================================
' Extract sub name matching a prefix pattern
'===========================================================================
Private Function ExtractSubName(lineText As String, prefix As String) As String
    Dim cleaned As String
    cleaned = Trim$(lineText)

    ' Remove "Public " if present
    If Left$(LCase$(cleaned), 7) = "public " Then
        cleaned = Trim$(Mid$(cleaned, 8))
    End If

    ' Must start with "Sub <prefix>"
    Dim target As String
    target = "sub " & LCase$(prefix)
    If Left$(LCase$(cleaned), Len(target)) <> target Then
        ExtractSubName = ""
        Exit Function
    End If

    ' Extract sub name after "Sub "
    Dim subPart As String
    subPart = Mid$(cleaned, 5) ' skip "Sub "

    Dim parenPos As Long
    parenPos = InStr(subPart, "(")
    If parenPos > 0 Then
        subPart = Left$(subPart, parenPos - 1)
    End If

    ExtractSubName = Trim$(subPart)
End Function

'===========================================================================
' JSON output
'===========================================================================
Private Sub WriteResultsJSON(filePath As String, totalElapsed As Double, fixtures() As String)
    Dim fNum As Integer
    fNum = FreeFile

    Open filePath For Output As #fNum

    Print #fNum, "{"
    Print #fNum, "  ""harness_version"": ""mvp-2.0"","
    Print #fNum, "  ""timestamp"": """ & Format$(Now, "yyyy-mm-dd hh:nn:ss") & ""","
    Print #fNum, "  ""total_elapsed_seconds"": " & Format$(totalElapsed, "0.000") & ","

    ' Count pass/fail
    Dim passCount As Long, failCount As Long
    passCount = 0: failCount = 0
    Dim idx As Long
    For idx = 1 To p_Results.Count
        If InStr(1, p_Results(idx), """passed"": true") > 0 Then
            passCount = passCount + 1
        Else
            failCount = failCount + 1
        End If
    Next idx

    Print #fNum, "  ""test_count"": " & p_Results.Count & ","
    Print #fNum, "  ""passed"": " & passCount & ","
    Print #fNum, "  ""failed"": " & failCount & ","

    ' Fixture results
    Print #fNum, "  ""fixtures"": ["
    For idx = 1 To p_FixtureResults.Count
        If idx < p_FixtureResults.Count Then
            Print #fNum, "    " & p_FixtureResults(idx) & ","
        Else
            Print #fNum, "    " & p_FixtureResults(idx)
        End If
    Next idx
    Print #fNum, "  ],"

    ' Test results
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
' Config reading helpers
'===========================================================================
Private Function ReadConfigValue(key As String) As String
    ' Read a value from the harness config JSON
    Dim content As String
    content = ReadConfigFile()
    If content = "" Then
        ReadConfigValue = ""
        Exit Function
    End If

    Dim searchKey As String
    searchKey = """" & key & """"

    Dim startPos As Long
    startPos = InStr(1, content, searchKey)
    If startPos = 0 Then
        ReadConfigValue = ""
        Exit Function
    End If

    startPos = InStr(startPos, content, ":") + 1
    startPos = InStr(startPos, content, """") + 1
    Dim endPos As Long
    endPos = InStr(startPos, content, """")

    If startPos > 1 And endPos > startPos Then
        ReadConfigValue = Mid$(content, startPos, endPos - startPos)
    Else
        ReadConfigValue = ""
    End If
End Function

Private Function ReadFixtureList() As String()
    ' Read fixture paths from harness config
    Dim content As String
    content = ReadConfigFile()

    Dim emptyArr(0) As String
    emptyArr(0) = ""

    If content = "" Then
        ReadFixtureList = emptyArr
        Exit Function
    End If

    ' Find "fixtures": [...]
    Dim startPos As Long
    startPos = InStr(1, content, """fixtures""")
    If startPos = 0 Then
        ReadFixtureList = emptyArr
        Exit Function
    End If

    startPos = InStr(startPos, content, "[")
    If startPos = 0 Then
        ReadFixtureList = emptyArr
        Exit Function
    End If

    Dim endPos As Long
    endPos = InStr(startPos, content, "]")
    If endPos = 0 Then
        ReadFixtureList = emptyArr
        Exit Function
    End If

    ' Extract the array content between [ and ]
    Dim arrContent As String
    arrContent = Mid$(content, startPos + 1, endPos - startPos - 1)

    ' Split by comma, extract quoted strings
    Dim parts() As String
    parts = Split(arrContent, ",")

    Dim count As Long
    count = 0
    Dim i As Long

    ' First pass: count non-empty entries
    For i = LBound(parts) To UBound(parts)
        Dim cleaned As String
        cleaned = Trim$(parts(i))
        cleaned = Replace(cleaned, """", "")
        cleaned = Replace(cleaned, vbLf, "")
        cleaned = Replace(cleaned, vbCr, "")
        cleaned = Trim$(cleaned)
        If cleaned <> "" Then count = count + 1
    Next i

    If count = 0 Then
        ReadFixtureList = emptyArr
        Exit Function
    End If

    ReDim result(0 To count - 1) As String
    Dim j As Long: j = 0
    For i = LBound(parts) To UBound(parts)
        cleaned = Trim$(parts(i))
        cleaned = Replace(cleaned, """", "")
        cleaned = Replace(cleaned, vbLf, "")
        cleaned = Replace(cleaned, vbCr, "")
        cleaned = Trim$(cleaned)
        If cleaned <> "" Then
            result(j) = cleaned
            j = j + 1
        End If
    Next i

    ReadFixtureList = result
End Function

Private Function ReadConfigFile() As String
    ' Try env var path first, then %TEMP% fallback
    Dim configPath As String
    configPath = Environ("TEST_HARNESS_CONFIG")

    If configPath = "" Then
        Dim tempDir As String
        tempDir = Environ("TEMP")
        If tempDir = "" Then tempDir = Environ("TMP")
        If tempDir = "" Then
            ReadConfigFile = ""
            Exit Function
        End If
        configPath = tempDir & "\_harness_config.json"
    End If

    On Error Resume Next
    Dim fNum As Integer
    fNum = FreeFile
    Open configPath For Input As #fNum
    If Err.Number <> 0 Then
        ReadConfigFile = ""
        Exit Function
    End If

    Dim content As String
    Dim lineText As String
    Do Until EOF(fNum)
        Line Input #fNum, lineText
        content = content & lineText & vbLf
    Loop
    Close #fNum
    On Error GoTo 0

    ReadConfigFile = content
End Function

'===========================================================================
' JSON string escaping
'===========================================================================
Private Function JsonStr(s As String) As String
    Dim escaped As String
    escaped = Replace(s, "\", "\\")
    escaped = Replace(escaped, """", "\""")
    escaped = Replace(escaped, vbCr, "\r")
    escaped = Replace(escaped, vbLf, "\n")
    escaped = Replace(escaped, vbTab, "\t")
    JsonStr = """" & escaped & """"
End Function
