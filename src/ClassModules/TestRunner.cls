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
   TestReporter.ResetCounters
End Sub

' Function: TestCase
' Given a TestCase object, run all of its tests
Public Sub TestCase(ByVal clsObj As iTestCase)
    ' get the list of all test methods
    List = clsObj.GetAllTests()
    Dim parObj As Object
    Dim Method As Variant
    Set parObj = clsObj.GetObject()

    For Each Method In List
        ' Run the test
        RunTest parObj, Method
    Next Method
End Sub

' Function: TestAllCases
' Given an array of TestCase objects, run all of their tests
Public Sub TestAllCases()
    clsArr = oTestConfig.getTestCases
    Dim clsObj As Variant

    For Each clsObj In clsArr
        TestCase clsObj
    Next clsObj
End Sub

' Function: RunTest
' Run a test
Public Sub RunTest(clsObj, Test)
    CallByName clsObj, "iTestCase_Setup", VbMethod
    ' Tell the logger what test we are running
    TestReporter.LogTest TypeName(clsObj), Test
    Call CallByName(clsObj, Test, VbMethod)
    CallByName clsObj, "iTestCase_TearDown", VbMethod
End Sub
