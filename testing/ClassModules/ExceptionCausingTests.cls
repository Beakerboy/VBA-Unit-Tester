Implements iTestCase

Dim Foo As Variant

Sub iTestCase_Setup()
    Dim TesterRunner As TestRunner
    Set TestRunner = CreateTestRunnerNoInit()
    
    Set TestRunner.TestConfig = New NoOutputConfig
    TestRunner.RunTest New PostLogTests, "CreateErrorTest"
End Sub

Sub iTestCase_TearDown()

End Sub

' Function: iTestCase_GetAllTests
' A list of every test in this test case
Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("ErrorLoggedTest")
End Function

Function ErrorLoggedTest()
    Status = TestReporter.GetLastTestStatus
    PassFail = Status(2)
    TestReporter.RemoveLastTest
    AssertEquals PassFail, "Failing", "Test should be failing because unexpected error was generated. Status is " & PassFail
    Message = Status(3)
    AssertNotEquals Message, "Test made no assertions.", "The message should not be for a missing assertion."
End Function

' Function: iTestCase_GetObject
' Returns this object
Function iTestCase_GetObject()
    Set iTestCase_GetObject = New ExceptionCausingTests
End Function
