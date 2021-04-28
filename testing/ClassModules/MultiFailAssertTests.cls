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
    SpecialTestCase.RunTest "TwoFailingAssertSetupTest"
End Sub

Sub iTestCase_TearDown()

End Sub

' Function: iTestCase_GetObject
' Returns this object
Function iTestCase_GetObject()
    Set iTestCase_GetObject = New MultiFailAssertTests
End Function

' Function: MultipleAssertsTest
' When a function has multiple assertions that fail, only report the first.
Sub MultipleAssertsTest()
    Status = TestReporter.GetLastTestStatus
    Message = Status(3)
    TestReporter.RemoveLastTest
    AssertEquals Message, "First Failure", "The logger should latch onto the first failure. Message seen: " & Message
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
