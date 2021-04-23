Implements iTestCase

Dim Foo As Variant

Sub iTestCase_Setup()

End Sub

Sub iTestCase_TearDown()
    TestReporter.OutputFailure = bOutputFailure
    Status = TestReporter.GetCurrentTestStatus
    Message = Status(3)
    TestReporter.ForceTestToPassing
    AssertEquals Message, "First Failure", "The logger should latch onto the first failure. Message seen: " & Message
End Sub

' Function: iTestCase_GetAllTests
' A list of every test in this test case
Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("ErrorLoggedTest")
End Function

' Function: iTestCase_GetObject
' Returns this object
Function iTestCase_GetObject()
    Set iTestCase_GetObject = New MultiFailAssertTests
End Function

' Function: ErrorLoggedTest
' Ensure that when a runtime error is generated, the logger catches it and loggs the test as failed.
Sub ErrorLoggedTest()
    AssertTrue False, "First Failure"
    AssertTrue False, "Second Failure"
End Sub
