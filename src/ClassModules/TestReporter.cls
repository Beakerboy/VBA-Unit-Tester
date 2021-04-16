VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Count of passing Test Cases
Dim Passing As Integer

'Count of failing Test Cases
Dim FailingCase As Integer

' Count of Failing Tests
Dim FailingTests As Integer

' Count of Assertions
Dim Assertions As Integer

Dim colLog As Collection

Dim sMessage As String

Dim sTestCase As String
Dim sTest As String

Dim PassFail As String

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
    ResetCounters
    PassFail = "Passing"
    sMessage = ""
    Set colLog = New Collection
End Sub

Public Sub ResultsMsgBox()
    Total = Passing + FailingCase
    MsgBox "Results" & vbNewLine & "Total Run: " & Total & vbNewLine & "Passing: " & Passing & vbNewLine & "Failing: " & Failing & vbNewLine & "Assertions: " & Assertions
End Sub

Public Sub LogAssertion()
    Assertions = Assertions + 1
End Sub

Public Sub LogFailure(Optional Message = "")
    If FailingTests = 0 Then
        FailingCase = FailingCase + 1
    End If
    FailingTests = FailingTests + 1
    sMessage = Message
    PassFail = "Failing"
End Sub

Public Sub ResetCounters()
    Passing = 0
    Failing = 0
    FailingTests = 0
    Assertions = 0
    Set colLog = New Collection
End Sub

Public Sub LogTest(TestCase, TestName)
    If sTest <> "" Then
        ' Write everything to the collection and clear the counts from the previous test
        colLog.Add Array(sTestCase, sTest, PassFail, sMessage)
    End If
    If TestCase <> sTestCase And FailingTests = 0 Then
        Passing = Passing + 1
    End If
    sTestCase = TestCase
    sTest = TestName
    sMessage = ""
    PassFail = "Passing"
End Sub

Public Function PrintLog()
    If sTest <> "" Then
        ' Write everything to the collection because we are done
        colLog.Add Array(sTestCase, sTest, PassFail, sMessage)
    End If
    sTestCase = ""
    sTest = ""
    sMessage = ""
    PassFail = "Passing"
    
    If FailingTests = 0 Then
        Passing = Passing + 1
    End If
    msg = ""
    For Each Test In colLog
        msg = msg & Test(0) & " " & Test(1) & " " & Test(2) & " " & Test(3) & vbNewLine
    Next Test
    MsgBox msg
End Function

Public Function PrintFailureLog()
    If sTest <> "" Then
        ' Write everything to the collection because we are done
        colLog.Add Array(sTestCase, sTest, PassFail, sMessage)
    End If
    sTestCase = ""
    sTest = ""
    sMessage = ""
    PassFail = "Passing"
    If FailingTests = 0 Then
        Passing = Passing + 1
    End If
    msg = ""
    For Each Test In colLog
        If Test(2) = "Failing" Then
            msg = msg & Test(0) & " " & Test(1) & " " & Test(2) & " " & Test(3) & vbNewLine
        End If
    Next Test
    MsgBox msg
End Function
