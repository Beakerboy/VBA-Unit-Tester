Dim OutputType As Variant
' Function: RunAllModuleTests
' Run all the defined tests for a given module
'
' Parameters:
'   Module_name - The name of the module
'
' Returns:
'   True if all tests pass
Function RunAllModuleTests(Module_name, Optional Output = "MsgBox")
  OutputType = Output
  RunAllModuleTests = Application.Run(Module_name & "_RunAllTests")
End Function

' Sub: AssertTrue
' Assert that the provided parameter is true
'
' Parameters:
'   MyTest - The parameter under test
Function AssertTrue(MyTest, Optional MessageString As String = "", Optional Output = 1)
    Old_Output = OutputType
    If Output <> 1 Then
        OutputType = Output
    End If
    AssertTrue = True
    If Not MyTest Then
        If MessageString = "" Then
            MessageString = "Expected: TRUE" & vbNewLine & "Provided: " & MyTest
        End If
        If OutputType = "MsgBox" Then
            MsgBox MessageString
        End If
        AssertTrue = False
    End If
    OutputType = Old_Output
End Function

' Sub: AssertFalse
' Assert that the provided parameter is false
'
' Parameters:
'   MyTest - The parameter under test
Function AssertFalse(MyTest, Optional MessageString As String = "", Optional Output = 1)
    If MessageString = "" Then
        MessageString = "Expected: FALSE" & vbNewLine & "Provided: " & MyTest
    End If
    AssertFalse = AssertTrue(Not MyTest, MessageString, Output)
End Function

' Sub: AssertEquals
' Assert that two variables have the same value
'
' Parameters:
'   MyTest        - The parameter under test
'   ExpectedValue - The expected value of MyTest
'
Function AssertEquals(MyTest, ExpectedValue, Optional MessageString As String = "", Optional Output = 1)
    If MessageString = "" Then
        MessageString = "Expected: " & ExpectedValue & vbNewLine & "Provided: " & MyTest
    End If
    AssertEquals = AssertTrue(MyTest = ExpectedValue, MessageString, Output)
End Function

' Sub: AssertNotEquals
' Assert that two variables do not have the same value
'
' Parameters:
'   MyTest          - The parameter under test
'   UnexpectedValue - The unexpected value of MyTest
'
Function AssertNotEquals(MyTest, UnexpectedValue, Optional MessageString As String = "", Optional Output = 1)
    If MessageString = "" Then
        MessageString = "Expected any value other than: " & MyTest
    End If
    AssertNotEquals = AssertTrue(MyTest <> UnexpectedValue, MessageString, Output)
End Function

' Function: AssertObjectStringEquals
' Assert that an objcts toString() function returns a specific value.
'
' Parameters:
'   MyTest        - The object under test
'   ExpectedValue - The expected value of MyTest.toString
'
Function AssertObjectStringEquals(MyObject, ExpectedValue, Optional MessageString As String = "", Optional Output = 1)
    ObjectString = MyObject.toString
    If MessageString = "" Then
        ObjectString = MyObject.toString
        MessageString = "Expected: " & ExpectedValue & vbNewLine & "Provided: " & ObjectString
    End If
    AssertObjectStringEquals = AssertTrue(ObjectString = ExpectedValue, MessageString, Output)
End Function
