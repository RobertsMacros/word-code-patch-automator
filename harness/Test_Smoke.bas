Attribute VB_Name = "Test_Smoke"
'===========================================================================
' Test_Smoke — smoke tests for the harness and environment
'
' Non-fixture tests: run once, no document argument.
'===========================================================================
Option Explicit

Public Sub Test_Smoke_Arithmetic()
    AssertEqual 4, 2 + 2, "2+2 should equal 4"
    AssertEqual 6, 2 * 3, "2*3 should equal 6"
End Sub

Public Sub Test_Smoke_StringOps()
    AssertEqual "hello world", LCase$("Hello World"), "LCase should work"
    AssertEqual 5, Len("hello"), "Len should return correct length"
    AssertContains "hello world", "world", "Should contain 'world'"
End Sub

Public Sub Test_Smoke_HostDocumentName()
    Dim docName As String
    docName = ThisDocument.Name
    AssertTrue Len(docName) > 0, "Host document name should not be empty"
End Sub

'===========================================================================
' Example fixture test (runs once per fixture document)
'===========================================================================
Public Sub TestFixture_Smoke_FixtureOpens(doc As Document)
    ' Verify the fixture document opened successfully and is readable
    AssertTrue Len(doc.Name) > 0, "Fixture document name should not be empty"
    AssertTrue doc.Paragraphs.Count >= 1, "Fixture should have at least one paragraph"
End Sub
