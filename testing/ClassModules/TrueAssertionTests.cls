VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrueAssertionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iTestCase

Sub iTestCase_Setup()
   
End Sub

Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("AssertTrueTest", "AssertFalseTest", "AssertIntEqualsTest", "AssertStringEqualsTest", "AssertFloatEqualsTest", "AssertIntNotEqualsTest", "AssertStringNotEqualsTest")
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New TrueAssertionTests
End Function

Sub AssertTrueTest()
    Dim Foo As Variant
    Foo = AssertTrue(AssertTrue(True, "", False), "Expecting True when AssertTrue is given True")
End Sub

Sub AssertFalseTest()
    Dim Foo As Variant
    Foo = AssertTrue(AssertFalse(False, "", False), "Expecting True when AssertFalse is given False")
End Sub

Sub AssertIntEqualsTest()
    Dim Foo As Variant
    Foo = AssertTrue(AssertEquals(1, 1, "", False), "Expecting True when AssertEquals is given equal integers")
End Sub

Sub AssertStringEqualsTest()
    Dim Foo As Variant
    Foo = AssertTrue(AssertEquals("one", "one", "", False), "Expecting True when AssertEquals is given equal strings")
End Sub

Sub AssertFloatEqualsTest()
    Dim Foo As Variant
    Foo = AssertTrue(AssertEquals(1.15, 1.15, "", False), "Expecting True when AssertEquals is given equal floats")
End Sub

Sub AssertIntNotEqualsTest()
    Dim Foo As Variant
    Foo = AssertTrue(AssertNotEquals(1, 2, "", False), "Expecting True when AssertNotEquals is given unequal integers")
End Sub

Sub AssertStringNotEqualsTest()
    Dim Foo As Variant
    Foo = AssertTrue(AssertNotEquals("one", "One", "", False), "Expecting True when AssertNotEquals is given unequal strings")
End Sub
