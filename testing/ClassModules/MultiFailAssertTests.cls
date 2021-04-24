Implements iTestCase

Dim Foo As Variant

Sub iTestCase_Setup()
    Dim TesterRunner As TestRunner
    Set TestRunner = CreateTestRunnerNoInit()
    
    Set TestRunner.TestConfig = New NoOutputConfig
    TestRunner.RunTest New PostLogTests, "TwoFailingAssertSetupTest"
End Sub

Sub iTestCase_TearDown()

End Sub

' Function: iTestCase_GetAllTests
' A list of every test in this test case
Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("MultipleAssertsTest")
End Function

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
