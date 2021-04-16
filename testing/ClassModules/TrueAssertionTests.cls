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
    ' Test that assertions evaluate to True appropriately
    AssertTrue True, "Expecting True when AssertTrue is given True"
End Sub

Sub AssertFalseTest()
    AssertFalse False, "Expecting True when AssertFalse is given False"
End Sub

Sub AssertIntEqualsTest()
    AssertEquals 1, 1, "Expecting True when AssertEquals is given equal integers"
End Sub

Sub AssertStringEqualsTest()
    AssertEquals "one", "one", "Expecting True when AssertEquals is given equal strings"
End Sub

Sub AssertFloatEqualsTest()
    AssertEquals 1.15, 1.15, "Expecting True when AssertEquals is given equal floats"
End Sub

Sub AssertIntNotEqualsTest()
    AssertNotEquals 1, 2, "Expecting True when AssertNotEquals is given unequal integers"
End Sub

Sub AssertStringNotEqualsTest()
    AssertNotEquals "one", "One", "Expecting True when AssertNotEquals is given unequal strings"
End Sub

