' Sub: AssertTrue
' Assert that the provided parameter is true
'
' Parameters:
'   MyTest - The parameter under test
Sub AssertTrue(MyTest, Optional MessageString As String = "")
    If Not MyTest Then
        If MessageString = "" Then
            MsgBox "Expected: TRUE" & vbNewLine & "Provided: " & MyTest
        Else
            MsgBox MessageString
        End If
    End If
End Sub

' Sub: AssertEquals
' Assert that two variables have the same value
'
' Parameters:
'   MyTest        - The parameter under test
'   ExpectedValue - The expected value of MyTest
'
Sub AssertEquals(MyTest, ExpectedValue, Optional MessageString As String = "")
    If MessageString = "" Then
        MessageString = "Expected: " & ExpectedValue & vbNewLine & "Provided: " & MyTest
    End If
    AssertTrue MyTest = ExpectedValue, MessageString
End Sub

' Sub: AssertNotEquals
' Assert that two variables do not have the same value
'
' Parameters:
'   MyTest          - The parameter under test
'   UnexpectedValue - The unexpected value of MyTest
'
Sub AssertNotEquals(MyTest, UnexpectedValue, Optional MessageString As String = "")
    If MessageString = "" Then
        MessageString = "Expected any value other than: " & MyTest
    End If
    AssertTrue MyTest <> ExpectedValue, MessageString
End Sub

' Sub: AssertObjectStringEquals
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
    AssertTrue ObjectString = ExpectedValue, MessageString
End Function
