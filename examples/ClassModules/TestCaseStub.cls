VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestCaseStub"
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

' Sub: iTestCase_Setup
' Add code here that will be run before every test.
' For example, create objects, populate data.
Sub iTestCase_Setup()

End Sub

' Sub: iTestCase_Setup
' Add code here that will be run after every test.
' For example, remove a worksheet, close a connection.
Sub iTestCase_TearDown()
End Sub

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

' Sub: MyFirstTest
' All test subs should end with "Test" and include at least one assertion.
Sub MyFirstTest()
    AssertTrue(2 = 1 + 1, "AssertTrue that one plus one equals two")
    AssertEquals(1 + 1, 2,  "AssertEquals that one plus one equals two")
End Sub
