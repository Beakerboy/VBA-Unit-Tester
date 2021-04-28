VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim oTestConfig As iTestableProject

Property Set TestConfig(oConfig As iTestableProject)
    Set oTestConfig = oConfig
    TestReporter.OutputFailure = oConfig.GetOutputFailures
End Property


' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
End Sub

Public Sub Initialize()
    TestReporter.ResetCounters
End Sub

' Function: TestAllCases
' Given an array of TestCase objects, run all of their tests
Public Sub TestAllCases()
    clsArr = oTestConfig.getTestCases
    Dim clsObj As Variant
    Dim TestCase As iTestCase
    For Each clsObj In clsArr
        Set TestCase = clsObj
        TestCase.RunAllTests
    Next clsObj
End Sub
