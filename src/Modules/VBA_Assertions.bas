Attribute VB_Name = "VBA_Assertions"
Dim OutputType As Variant

' Sub: AssertTrue
' Assert that the provided parameter is true
'
' Parameters:
'   MyTest - The parameter under test
Function AssertTrue(MyTest, Optional MessageString As String = "")
    TestReporter.LogAssertion
    AssertTrue = True
    If Not MyTest Then
        If MessageString = "" Then
            MessageString = "Expected: TRUE" & vbNewLine & "Provided: " & MyTest
        End If
        TestReporter.LogFailure MessageString
        AssertTrue = False
    Else
        'TestReporter.LogSuccess
    End If
End Function

' Sub: AssertFalse
' Assert that the provided parameter is false
'
' Parameters:
'   MyTest - The parameter under test
Function AssertFalse(MyTest, Optional MessageString As String = "")
    If MessageString = "" Then
        MessageString = "Expected: FALSE" & vbNewLine & "Provided: " & MyTest
    End If
    AssertFalse = AssertTrue(Not MyTest, MessageString)
End Function

' Sub: AssertEquals
' Assert that two variables have the same value
'
' Parameters:
'   MyTest        - The parameter under test
'   ExpectedValue - The expected value of MyTest
'
Function AssertEquals(MyTest, ExpectedValue, Optional MessageString As String = "")
    If MessageString = "" Then
        MessageString = "Expected: " & ExpectedValue & vbNewLine & "Provided: " & MyTest
    End If
    AssertEquals = AssertTrue(MyTest = ExpectedValue, MessageString)
End Function

' Sub: AssertNotEquals
' Assert that two variables do not have the same value
'
' Parameters:
'   MyTest          - The parameter under test
'   UnexpectedValue - The unexpected value of MyTest
'
Function AssertNotEquals(MyTest, UnexpectedValue, Optional MessageString As String = "")
    If MessageString = "" Then
        MessageString = "Expected any value other than: " & MyTest
    End If
    AssertNotEquals = AssertTrue(MyTest <> UnexpectedValue, MessageString)
End Function

' Function: AssertObjectStringEquals
' Assert that an objcts toString() function returns a specific value.
'
' Parameters:
'   MyTest        - The object under test
'   ExpectedValue - The expected value of MyTest.toString
'
Function AssertObjectStringEquals(MyObject, ExpectedValue, Optional MessageString As String = "")
    ObjectString = MyObject.toString
    If MessageString = "" Then
        ObjectString = MyObject.toString
        MessageString = "Expected: " & ExpectedValue & vbNewLine & "Provided: " & ObjectString
    End If
    AssertObjectStringEquals = AssertTrue(ObjectString = ExpectedValue, MessageString)
End Function
