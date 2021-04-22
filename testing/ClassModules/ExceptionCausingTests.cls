Implements iTestCase

Dim Foo As Variant

Sub iTestCase_Setup()

End Sub

Sub iTestCase_TearDown()
    TestReporter.OutputFailure = bOutputFailure
    Status = TestReporter.GetCurrentTestStatus
    PassFail = Status(2)
    TestReporter.ForceTestToPassing
    AssertEquals PassFail, "Failing", "Test should be failing because unexpected error was generated. Status is " & PassFail
End Sub

' Function: iTestCase_GetAllTests
' A list of every test in this test case
Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("ErrorLoggedTest")
End Function

' Function: iTestCase_GetObject
' Returns this object
Function iTestCase_GetObject()
    Set iTestCase_GetObject = New ExceptionCausingTests
End Function

' Function: ErrorLoggedTest
' Ensure that when a runtime error is generated, the logger catches it and loggs the test as failed.
Sub ErrorLoggedTest()
    bOutputFailure = TestReporter.OutputFailure
    TestReporter.OutputFailure = False
    ' call function that generates an error, but do not do any real-time error output
    SubThatGenratesAnError
End Sub

Sub SubThatGenratesAnError()
    Dim A As Integer
    A = 5 / 0
End Sub
