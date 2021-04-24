Implements iTestCase

Dim Foo As Variant

Sub iTestCase_Setup()
    Dim TesterRunner As TestRunner
    Set TestRunner = CreateTestRunnerNoInit()
    
    Set TestRunner.TestConfig = New NoOutputConfig
    TestRunner.RunTest New PostLogTests, "NoAssertionMadeSetupTest"
End Sub

Sub iTestCase_TearDown()
End Sub

' Function: iTestCase_GetAllTests
' A list of every test in this test case
Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("NoAssertionMadeTest")
End Function

' Function: iTestCase_GetObject
' Returns this object
Function iTestCase_GetObject()
    Set iTestCase_GetObject = New NoAssertionTests
End Function

' Function: NoAssertionMadeTest
' Ensure that when a test does not make any assertions, it fails.
Sub NoAssertionMadeTest()
    TestReporter.OutputFailure = bOutputFailure
    Status = TestReporter.GetLastTestStatus
    PassFail = Status(2)
    TestReporter.RemoveLastTest
    AssertEquals PassFail, "Failing", "Test should be failing because no assertions are made. Status is " & PassFail
End Sub
