Attribute VB_Name = "VBA_Assertions"
Dim OutputType As Variant

' Sub: AssertTrue
' Assert that the provided parameter is true
'
' Parameters:
'   MyTest - The parameter under test
Function AssertTrue(MyTest, Optional MessageString As String = "", Optional Logging = True)
    If Logging Then TestReporter.LogAssertion
    AssertTrue = True
    If Not MyTest Then
        If MessageString = "" Then
            MessageString = "Expected: TRUE" & vbNewLine & "Provided: " & MyTest
        End If
        If Logging Then TestReporter.LogFailure MessageString
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
Function AssertFalse(MyTest, Optional MessageString As String = "", Optional Logging = True)
    If MessageString = "" Then
        MessageString = "Expected: FALSE" & vbNewLine & "Provided: " & MyTest
    End If
    AssertFalse = AssertTrue(Not MyTest, MessageString, Logging)
End Function

' Sub: AssertEquals
' Assert that two variables have the same value
'
' Parameters:
'   MyTest        - The parameter under test
'   ExpectedValue - The expected value of MyTest
'
Function AssertEquals(MyTest, ExpectedValue, Optional MessageString As String = "", Optional Logging = True)
    If MessageString = "" Then
        MessageString = "Expected: " & ExpectedValue & vbNewLine & "Provided: " & MyTest
    End If
    AssertEquals = AssertTrue(MyTest = ExpectedValue, MessageString, Logging)
End Function

' Sub: AssertSame
' Assert that two variables have the same value and Type
'
' Parameters:
'   MyTest        - The parameter under test
'   ExpectedValue - The expected value of MyTest
'
Function AssertSame(MyTest, ExpectedValue, Optional MessageString As String = "", Optional Logging = True)
    If MessageString = "" Then
        MessageString = "Expected: " & ExpectedValue & vbNewLine & "Provided: " & MyTest
    End If
    bBool = MyTest = ExpectedValue And TypeName(MyTest) = TypeName(ExpectedValue)
    AssertSame = AssertTrue(bBool, MessageString, Logging)
End Function

' Sub: AssertArrayEquals
' Assert that two arrays have the same dimensions and values
'
' Parameters:
'   MyTest        - The parameter under test
'   ExpectedValue - The expected value of MyTest
'
Function AssertArrayEquals(MyTest, ExpectedValue, Optional MessageString As String = "", Optional Logging = True)
    AssertArrayEquals = True
    ' Check the number of dimensions of MyTest, verify the same in Expected
    ' Iterate each dimension. If it's an atomic value, measure equality. If it's an array, continue iterating.
    
    If MessageString = "" Then
        MessageString = "Expected: " & ExpectedValue & vbNewLine & "Provided: " & MyTest
    End If
    AssertEquals = AssertTrue(MyTest = ExpectedValue, MessageString, Logging)
End Function

Private Function getArrayRank(aVar)

End Function

' Sub: AssertNotEquals
' Assert that two variables do not have the same value
'
' Parameters:
'   MyTest          - The parameter under test
'   UnexpectedValue - The unexpected value of MyTest
'
Function AssertNotEquals(MyTest, UnexpectedValue, Optional MessageString As String = "", Optional Logging = True)
    If MessageString = "" Then
        MessageString = "Expected any value other than: " & MyTest
    End If
    AssertNotEquals = AssertTrue(MyTest <> UnexpectedValue, MessageString, Logging)
End Function

' Function: AssertObjectStringEquals
' Assert that an objcts toString() function returns a specific value.
'
' Parameters:
'   MyTest        - The object under test
'   ExpectedValue - The expected value of MyTest.toString
'
Function AssertObjectStringEquals(MyObject, ExpectedValue, Optional MessageString As String = "", Optional Logging = True)
    ObjectString = MyObject.toString
    If MessageString = "" Then
        ObjectString = MyObject.toString
        MessageString = "Expected: " & ExpectedValue & vbNewLine & "Provided: " & ObjectString
    End If
    AssertObjectStringEquals = AssertTrue(ObjectString = ExpectedValue, MessageString, Logging)
End Function

Public Sub ExpectError(Optional Code As Integer = 0, Optional Message As String = "")
    TestReporter.ExpectException = True
End Sub
