Implements iTestCase

Dim Foo As Variant

Sub iTestCase_Setup()

End Sub

Sub iTestCase_TearDown()
End Sub

' Function: iTestCase_GetAllTests
' A list of every test in this test case
Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("ErrorCaughtTest")
End Function

' Function: iTestCase_GetObject
' Returns this object
Function iTestCase_GetObject()
    Set iTestCase_GetObject = New LoggerErrorHandlingTests
End Function


Sub SubThatGenratesAnError()
    Dim A As Integer
    A = 5 / 0
End Sub

'Function: ErrorCaughtTest
' Ensure that when an expected runtime error is generated, we can get its information, and the test is still passing
Sub ErrorCaughtTest()

    ExpectException
    ' call function that generates an error
    SubThatGenratesAnError
    ' Ensure the error is caught in a segragated way and test is still flagged as passing
    ' This stuff does not run due to error catching
    Status = TestReporter.GetCurrentTestStatus
    AssertEquals Status(2), "Passing", "Test should be passing despite error because it was expected"
End Sub

' Function: FailsWhenErrorMissingTest
' If a test expects an exception, ensure that the test fails when one does not occur
' This one might be hard to test
Sub FailsWhenErrorMissingTest()
    ' Call the test that expects an exception, but never generates one
    RunTest New LoggerErrorHandlingTests, "FailsEhenErrorMissingExampleTest"
    
    ' Check the logger to ensure that that method is flagged as failed.
    ' How do we do this without the failure actually generating outputs?
    ' We could force all outputs to the screen to not occur.
    ' We could grab a copy of the entire run/error log before running the test, and then restore it after this one runs.
End Sub

Sub FirstFailureLatchedTest()
    ' Generate a failure followed by an error
    ' Ensure the logger kept the first

    ' Generate a failure followed by an error
    ' Ensure the logger kept the first
End Sub

Function FailsEhenErrorMissingExampleTest()
    ExpectException
End Function
