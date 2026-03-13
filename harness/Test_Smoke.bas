Attribute VB_Name = "Test_Smoke"
'===========================================================================
' Test_Smoke — Basic smoke tests to verify the harness is working
'
' These tests validate that the test infrastructure itself is functional
' before running any project-specific tests.
'===========================================================================
Option Explicit

Public Sub Test_Smoke_Arithmetic()
    ' Verify basic VBA arithmetic works
    AssertEqual 4, 2 + 2, "2+2 should equal 4"
    AssertEqual 6, 2 * 3, "2*3 should equal 6"
End Sub

Public Sub Test_Smoke_StringOps()
    ' Verify basic string operations
    AssertEqual "hello world", LCase$("Hello World"), "LCase should work"
    AssertEqual 5, Len("hello"), "Len should return correct length"
    AssertContains "hello world", "world", "Should contain 'world'"
End Sub

Public Sub Test_Smoke_DocumentName()
    ' Verify we can access the active document
    Dim docName As String
    docName = ThisDocument.Name
    AssertTrue Len(docName) > 0, "Document name should not be empty"
End Sub
