VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim Passing As Integer
Dim Failing As Integer

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
    Passing = 0
    Failing = 0
End Sub

' Function: TestAll
' Given a TestCase object, run all of its tests
Public Function TestAll(ByVal clsObj As iTestCase)
    Result = True
    ' get the list of all test methods
    List = clsObj.GetAllTests()
    Dim parObj As Object
    Set parObj = clsObj.GetObject()
    For Each Method In List
        TestResult = RunTest(parObj, Method)
        If TestResult Then
            Passing = Passing + 1
        Else
            Failing = Failing + 1
        End If
        Result = Result And TestResult
    Next Method
    TestAll = Result
End Function

' Function: TestAllClasses
' Given an array of TestCase objects, run all of their tests
Public Function TestAllClasses(clsArr)
    Result = True
    For Each clsObj In clsArr
        Result = Result And TestAll(clsObj)
    Next clsObj
    TestAllClasses = Result
End Function

' Function: RunTest
' Run a test
Public Function RunTest(clsObj, Test)
    CallByName clsObj, "iTestCase_Setup", VbMethod
    RunTest = CallByName(clsObj, Test, VbMethod)
End Function

Public Sub ResultsMsgBox()
    Total = Passing + Failing
    MsgBox "Results" & vbNewLine & "Total Run: " & Total & vbNewLine & "Passing: " & Passing & vbNewLine & "Failing: " & Failing
End Sub
