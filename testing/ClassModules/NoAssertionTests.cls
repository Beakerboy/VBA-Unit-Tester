VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NoAssertionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iTestCase

Dim TestCase As iTestCase

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
    Set TestCase = CreateTestCase()
End Sub

Sub iTestCase_Setup()
    Dim SpecialTestCase As iTestCase
    Set SpecialTestCase = New PostLogTests
    SpecialTestCase.RunTest "NoAssertionMadeSetupTest"
End Sub

Sub iTestCase_TearDown()
End Sub

' Function: iTestCase_GetObject
' Returns this object
Function iTestCase_GetObject()
    Set iTestCase_GetObject = New NoAssertionTests
End Function

' Function: NoAssertionMadeTest
' Ensure that when a test does not make any assertions, it fails.
Sub NoAssertionMadeTest()
    Status = TestReporter.GetLastTestStatus
    PassFail = Status(2)
    TestReporter.RemoveLastTest
    AssertEquals PassFail, "Failing", "Test should be failing because no assertions are made. Status is " & PassFail
End Sub

' Sub: iTestCase_RunTest
' Run a specific test.
Public Sub iTestCase_RunTest(Test As String, Optional clsObj = Nothing)
    TestCase.RunTest Test, Me
End Sub

' Sub: iTestCase_RunAllTests
' Run all tests.
Public Sub iTestCase_RunAllTests(Optional ByVal clsObj = Nothing)
    TestCase.RunAllTests Me
End Sub
