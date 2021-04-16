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

' Count of Failing Tests for the current TestCase
Dim FailingTests As Integer

' Count of the Total Failing Tests
Dim TotalFailingTests As Integer

' Count of Assertions
Dim Assertions As Integer

Dim colLog As Collection

Dim sMessage As String

Dim sTestCase As String
Dim sTest As String

Dim bOutputFailures As Boolean
Dim PassFail As String

Public Property Let OutputFailure(OutputFailures As Boolean)
    bOutputFailures = OutputFailures
End Property

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
    ResetCounters
    PassFail = "Passing"
    FailingTests = 0
    TotalFailingTests = 0
    sMessage = ""
    Set colLog = New Collection
End Sub

Public Sub ResultsMsgBox()
    FinalizeLog
    Total = Passing + FailingCase
    MsgBox "Results" & vbNewLine & "Cases Run: " & Total & vbNewLine & "Failing: " & FailingCase & vbNewLine & vbNewLine & "Total Tests: " & colLog.Count & vbNewLine & "Failing: " & TotalFailingTests & vbNewLine & vbNewLine & "Assertions: " & Assertions
End Sub

Public Sub LogAssertion()
    Assertions = Assertions + 1
End Sub

Public Sub LogFailure(Optional Message = "")
    If FailingTests = 0 Then
        FailingCase = FailingCase + 1
    End If
    FailingTests = FailingTests + 1
    TotalFailingTests = TotalFailingTests + 1
    sMessage = Message
    PassFail = "Failing"
    If bOutputFailures Then
        MsgBox sTestCase & vbNewLine & sTest & vbNewLine & sMessage
    End If
End Sub

Public Sub ResetCounters()
    Passing = 0
    FailingCase = 0
    FailingTests = 0
    TotalFailingTests = 0
    Assertions = 0
    Set colLog = New Collection
End Sub

Public Sub LogTest(TestCase, TestName)
    If sTest <> "" Then
        ' Write everything to the collection and clear the counts from the previous test
        colLog.Add Array(sTestCase, sTest, PassFail, sMessage)
    End If
    If TestCase <> sTestCase Then
        If FailingTests = 0 Then
            Passing = Passing + 1
        Else
            FailingTests = 0
        End If
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

Private Sub FinalizeLog()
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
End Sub

Public Function PrintFailureLog()
    FinalizeLog
    msg = ""
    For Each Test In colLog
        If Test(2) = "Failing" Then
            msg = msg & Test(0) & " " & Test(1) & " " & Test(2) & " " & Test(3) & vbNewLine
        End If
    Next Test
    If msg = "" Then
        msg = "No Failures"
    End If
    MsgBox msg
End Function
