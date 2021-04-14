VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Function: TestAll
' Given a TestCase object, run all of its tests
Public Function TestAll(clsObj)
    Result = True
    ' get the list of all test methods
    List = clsObj.GetAllTests()
    For Each Method In List
        Result = Result And RunTest(clsObj, Method)
    Next Method
    TestAll = Result
End Function

' Function: RunTest
' Run a test
Function RunTest(clsObj, Test)
    CallByName clsObj, "iTestCase_Setup", VbMethod
    RunTest = CallByName(clsObj, Test, VbMethod)
End Function
