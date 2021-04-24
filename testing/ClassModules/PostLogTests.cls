VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PostLogTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iTestCase

Dim Foo As Variant

Sub iTestCase_Setup()

End Sub

Sub iTestCase_TearDown()

End Sub

' Function: iTestCase_GetAllTests
' A list of every test in this test case
Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array()
End Function

' Function: iTestCase_GetObject
' Returns this object
Function iTestCase_GetObject()
    Set iTestCase_GetObject = New PostLogChecks
End Function

' Function: ErrorLoggedTest
' Ensure that when a runtime error is generated before an assertion, the log records the error, not the lack of assertion.
Sub CreateErrorTest()
    ' Generate an error. We test that the failure is reported properly in the TearDown()
    A = 5 / 0
End Sub


Sub ExpectExceptionSetupTest()
    ExpectError
    ' Generate an error
    A = 5 / 0
End Sub

Sub TwoFailingAssertSetupTest()
    AssertTrue False, "First Failure"
    AssertTrue False, "Second Failure"
End Sub

Sub NoAssertionMadeSetupTest()
    A = 5 / 1
End Sub

Sub NoErrorOccuredSetupTest()
    ' Tell the test environment that we expect an error but do not generate one.
    ' We test that the lack of an expected error creates a failed test in TearDown().
    ExpectError
End Sub
