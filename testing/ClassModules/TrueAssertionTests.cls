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

Dim Foo As Variant

Dim TestCase As iTestCase

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
    Set TestCase = CreateTestCase()
End Sub

Sub iTestCase_Setup()

End Sub

Sub iTestCase_TearDown()

End Sub

' Function: iTestCase_GetObject
' Returns this object
Function iTestCase_GetObject()
    Set iTestCase_GetObject = New TrueAssertionTests
End Function

' Function: AssertTrueTest
' Test that assertTrue() returns true appropriately
Sub AssertTrueTest()
    Foo = AssertTrue(AssertTrue(True, "", False), "Expecting True when AssertTrue is given True")
End Sub

' Function: AssertFalseTest
' Test that assertFalse() returns true appropriately
Sub AssertFalseTest()
    Foo = AssertTrue(AssertFalse(False, "", False), "Expecting True when AssertFalse is given False")
End Sub

' Function: AssertEqualsProviderTest(
' Test that AssertEquals() returns true appropriately
Sub AssertEqualsProviderTest(Inputs, Expected, Message As String)
    Foo = AssertTrue(AssertEquals(Inputs, Expected, "", False), Message)
End Sub

' Function: AssertNotEqualsProviderTest(
' Test that AssertNotEquals() returns true appropriately
Sub AssertNotEqualsProviderTest(Inputs, Expected, Message As String)
    Foo = AssertTrue(AssertNotEquals(Inputs, Expected, "", False), Message)
End Sub

' Function: AssertEqualsProviderTest_Data
' Test data for AssertEqualsProviderTest()
Function AssertEqualsProviderTest_Data()
    AssertEqualsProviderTest_Data = Array( _
        Array(1, 1, "Expecting True when AssertEquals is given equal integers"), _
        Array(1, 1#, "Expecting True when AssertEquals is given an integer and a float"), _
        Array("one", "one", "Expecting True when AssertEquals is given equal strings"), _
        Array(1.15, 1.15, "Expecting True when AssertEquals is given equal floats") _
    )
End Function

' Function: AssertNotEqualsProviderTest_Data
' Test data for AssertNotEqualsProviderTest()
Function AssertNotEqualsProviderTest_Data()
    AssertNotEqualsProviderTest_Data = Array( _
        Array(1, 2, "Expecting True when AssertNotEquals is given unequal integers"), _
        Array("one", "One", "Expecting True when AssertNotEquals is given unequal strings"), _
        Array(1.15, 1.16, "Expecting True when AssertNotEquals is given unequal floats") _
    )
End Function

' Function: AssertEqualsProviderTest(
' Test that AssertEquals() returns true appropriately
Sub AssertSameProviderTest(Inputs, Expected, Message As String)
    Foo = AssertTrue(AssertSame(Inputs, Expected, "", False), Message)
End Sub

' Function: AssertEqualsProviderTest_Data
' Test data for AssertEqualsProviderTest()
Function AssertSameProviderTest_Data()
    AssertSameProviderTest_Data = Array( _
        Array(1, 1, "Expecting True when AssertEquals is given equal integers"), _
        Array("one", "one", "Expecting True when AssertEquals is given equal strings"), _
        Array(1.15, 1.15, "Expecting True when AssertEquals is given equal floats") _
    )
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
