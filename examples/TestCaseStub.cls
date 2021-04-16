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

Dim Foo As Variant

Sub iTestCase_Setup()

End Sub

Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("MyFirstTest")
End Function

Function iTestCase_GetObject()
    ' Ensure this is a copy of "self"
    Set iTestCase_GetObject = New TestCaseStub
End Function

Sub MyFirstTest()
    AssertTrue(2 = 1 + 1, "AssertTrue that one plus one equals two")
    AssertEquals(1 + 1, 2,  "AssertEquals that one plus one equals two")
End Sub
