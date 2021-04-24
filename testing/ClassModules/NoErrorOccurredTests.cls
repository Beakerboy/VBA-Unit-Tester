Implements iTestCase

Dim Foo As Variant

Sub iTestCase_Setup()
    Dim TesterRunner As TestRunner
    Set TestRunner = CreateTestRunnerNoInit()
    
    Set TestRunner.TestConfig = New NoOutputConfig
    TestRunner.RunTest New PostLogTests, "NoErrorOccuredSetupTest"
End Sub

Sub iTestCase_TearDown()

End Sub

' Function: iTestCase_GetAllTests
' A list of every test in this test case
Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("NoErrorOcurredTest")
End Function

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
