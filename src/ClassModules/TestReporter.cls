VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Assertions made in the current test
Dim Assertions As Integer

Dim colLog As Collection

Dim sMessage As String

Dim sTestCase As String

Dim sTest As String

Dim bOutputFailures As Boolean

Dim PassFail As String

Dim bExpectException As Boolean

Public Property Let OutputFailure(OutputFailures As Boolean)
    bOutputFailures = OutputFailures
End Property

Public Property Get OutputFailure() As Boolean
    OutputFailure = bOutputFailures
End Property

Public Property Let ExpectException(ExpectExceptions As Boolean)
    bExpectException = ExpectExceptions
    Assertions = Assertions + 1
End Property

Public Property Get ExpectException() As Boolean
    ExpectException = bExpectException
End Property

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
    ResetCounters
End Sub

Public Function GetCurrentTestStatus()
    GetCurrentTestStatus = Array(sTestCase, sTest, PassFail, sMessage, Assertions)
End Function

Public Function GetLastTestStatus()
    LastTestNumber = colLog.Count
    GetLastTestStatus = colLog.Item(LastTestNumber)
End Function

Public Function FinalizeTest()
    If sTest <> "" Then
        If Assertions = 0 And PassFail = "Passing" Then
            LogFailure "Test made no assertions."
        End If
        If ExpectException And PassFail = "Passing" Then
            LogFailure "Expected Exception but none occurred"
            ExpectException = False
        End If
        ' Write everything to the collection and clear the counts from the previous test
        colLog.Add Array(sTestCase, sTest, PassFail, sMessage, Assertions)
    End If
End Function

Public Sub RemoveLastTest()
    LastTestNumber = colLog.Count
    colLog.Remove LastTestNumber
End Sub

Public Sub ResultsMsgBox()
Dim TotalFailingTests As Integer
Dim TotalAssertions As Integer
TotalFailingTests = 0
TotalAssertions = 0
    For Each test In colLog
        If test(2) = "Failing" Then
            TotalFailingTests = TotalFailingTests + 1
        End If
        TotalAssertions = TotalAssertions + test(4)
    Next test
    MsgBox "Results" & vbNewLine & "Cases Run: " & Total & vbNewLine & "Failing: " & FailingCase & vbNewLine & vbNewLine & "Total Tests: " & colLog.Count & vbNewLine & "Failing: " & TotalFailingTests & vbNewLine & vbNewLine & "Assertions: " & TotalAssertions
End Sub

Public Sub LogAssertion()
    TotalAssertions = TotalAssertions + 1
    Assertions = Assertions + 1
End Sub

Public Sub LogFailure(Optional Message = "")
    If PassFail = "Passing" Then
        sMessage = Message
        PassFail = "Failing"
        If bOutputFailures Then
            MsgBox sTestCase & vbNewLine & sTest & vbNewLine & sMessage
        End If
    End If
End Sub

Public Sub LogError(Optional Message = "")
    If Not bExpectException Then
        LogFailure Message
    Else
        LogAssertion
        bExpectException = False
    End If
End Sub

Public Sub ResetCounters()
    Assertions = 0
    Set colLog = New Collection
    bExpectException = False
    PassFail = "Passing"
    sMessage = ""
End Sub

Public Sub LogTest(TestCase, TestName)
    sTestCase = TestCase
    sTest = TestName
    sMessage = ""
    PassFail = "Passing"
    bExpectException = False
    Assertions = 0
End Sub

Public Function PrintLog()
    sTestCase = ""
    sTest = ""
    sMessage = ""
    PassFail = "Passing"
    bExpectException = False
    msg = ""
    For Each test In colLog
        msg = msg & test(0) & " " & test(1) & " " & test(2) & " " & test(3) & vbNewLine
    Next test
    MsgBox msg
End Function

Public Function PrintFailureLog()
    msg = ""
    For Each test In colLog
        If test(2) = "Failing" Then
            msg = msg & test(0) & " " & test(1) & " " & test(2) & " " & test(3) & vbNewLine
        End If
    Next test
    If msg = "" Then
        msg = "No Failures"
    End If
    MsgBox msg
End Function
