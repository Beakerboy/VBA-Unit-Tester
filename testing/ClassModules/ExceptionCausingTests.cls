VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExceptionCausingTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
    SpecialTestCase.RunTest "CreateErrorTest"
End Sub

Sub iTestCase_TearDown()

End Sub

Function ErrorLoggedTest()
    Status = TestReporter.GetLastTestStatus
    PassFail = Status(2)
    TestReporter.RemoveLastTest
    AssertEquals PassFail, "Failing", "Test should be failing because unexpected error was generated. Status is " & PassFail
    Message = Status(3)
    AssertNotEquals Message, "Test made no assertions.", "The message should not be for a missing assertion."
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
