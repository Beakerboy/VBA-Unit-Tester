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
    SpecialTestCase.RunTest "NoErrorOccuredSetupTest"
End Sub

Sub iTestCase_TearDown()

End Sub

' Function: iTestCase_GetObject
' Returns this object
Function iTestCase_GetObject()
    Set iTestCase_GetObject = New NoErrorOccurredTests
End Function

Function NoErrorOcurredTest()
    TestReporter.OutputFailure = bOutputFailure
    Status = TestReporter.GetLastTestStatus
    PassFail = Status(2)
    TestReporter.RemoveLastTest
    AssertEquals PassFail, "Failing", "Test should be failing because no error was generated. Status is " & PassFail
End Function

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
