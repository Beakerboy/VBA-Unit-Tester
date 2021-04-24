Implements iTestCase

Dim Foo As Variant

Sub iTestCase_Setup()
    Dim TesterRunner As TestRunner
    Set TestRunner = CreateTestRunnerNoInit()
    
    Set TestRunner.TestConfig = New NoOutputConfig
    TestRunner.RunTest New PostLogTests, "ExpectExceptionSetupTest"
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

'Function: ErrorCaughtTest
' Ensure that when an expected runtime error is generated it is not logged, and the test is still passing.
Sub ErrorCaughtTest()
    Status = TestReporter.GetLastTestStatus
    PassFail = Status(2)
    TestReporter.RemoveLastTest
    AssertEquals PassFail, "Passing", "Test should be passing because an expected error was generated. Status is " & PassFail
End Sub
