VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FalseAssertionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iTestCase

Sub iTestCase_Setup()
   
End Sub

Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("FailingAssertionTests")
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New FalseAssertionTests
End Function

Sub FailingAssertionTests()
    Dim Foo As Variant
    ' Test that assertions evaluate to false appropriately
    Foo = AssertFalse(AssertTrue(False, "", False), "Testing boolean False literal")
    Foo = AssertFalse(AssertFalse(True, "", False), "Expecting False when AssertFalse is given True")
    Foo = AssertFalse(AssertEquals(1, 2, "", False), "Expecting False when AssertEquals is given unequal integers")
    Foo = AssertFalse(AssertEquals("one", "One", "", False), "Expecting False when AssertEquals is given unequal strings")
    Foo = AssertFalse(AssertNotEquals(1, 1, "", False), "Expecting False when AssertNotEquals is given equal integers")
    Foo = AssertFalse(AssertNotEquals("one", "one", "", False), "Expecting False when AssertNotEquals is given uequal strings")
    Foo = AssertFalse(AssertNotEquals(1.15, 1.15, "", False), "Expecting False when AssertNotEquals is given uequal floats")
End Sub
