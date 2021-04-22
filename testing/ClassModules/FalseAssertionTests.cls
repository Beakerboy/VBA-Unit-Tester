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

Sub iTestCase_TearDown()
   
End Sub

Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("FailingAssertionTests", "AssertEqualsProviderTest", "AssertNotEqualsProviderTest", "AssertSameProviderTest")
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New FalseAssertionTests
End Function

Sub FailingAssertionTests()
    Dim Foo As Variant
    ' Test that assertions evaluate to false appropriately
    Foo = AssertFalse(AssertTrue(False, "", False), "Testing boolean False literal")
    Foo = AssertFalse(AssertFalse(True, "", False), "Expecting False when AssertFalse is given True")
End Sub

' Function: AssertEqualsProviderTest(
' Test that AssertEquals() returns false appropriately
Sub AssertEqualsProviderTest(Inputs, Expected, Message As String)
    Foo = AssertFalse(AssertEquals(Inputs, Expected, "", False), Message)
End Sub

' Function: AssertEqualsProviderTest_Data
' Test data for AssertEqualsProviderTest()
Function AssertEqualsProviderTest_Data()
    AssertEqualsProviderTest_Data = Array( _
        Array(1, 2, "Expecting False when AssertEquals is given unequal integers"), _
        Array("one", "One", "Expecting False when AssertEquals is given unequal strings"), _
        Array(1.15, 1.16, "Expecting False when AssertEquals is given unequal floats") _
    )
End Function

' Function: AssertNotEqualsProviderTest(
' Test that AssertNotEquals() returns false appropriately
Sub AssertNotEqualsProviderTest(Inputs, Expected, Message As String)
    Foo = AssertFalse(AssertNotEquals(Inputs, Expected, "", False), Message)
End Sub

' Function: AssertNotEqualsProviderTest_Data
' Test data for AssertNotEqualsProviderTest()
Function AssertNotEqualsProviderTest_Data()
    AssertNotEqualsProviderTest_Data = Array( _
        Array(1, 1, "Expecting False when AssertNotEquals is given equal integers"), _
        Array("one", "one", "Expecting False when AssertNotEquals is given uequal strings"), _
        Array(1.15, 1.15, "Expecting False when AssertNotEquals is given uequal floats") _
    )
End Function

' Function: AssertEqualsProviderTest(
' Test that AssertEquals() returns false appropriately
Sub AssertSameProviderTest(Inputs, Expected, Message As String)
    Foo = AssertFalse(AssertSame(Inputs, Expected, "", False), Message)
End Sub

' Function: AssertEqualsProviderTest_Data
' Test data for AssertEqualsProviderTest()
Function AssertSameProviderTest_Data()
    AssertSameProviderTest_Data = Array( _
        Array(1, 1#, "Expecting False when AssertSame is given a float and an integer"), _
        Array(1, 2, "Expecting False when AssertEquals is given unequal integers"), _
        Array("one", "One", "Expecting False when AssertEquals is given unequal strings"), _
        Array(1.15, 1.16, "Expecting False when AssertEquals is given unequal floats") _
    )
End Function
