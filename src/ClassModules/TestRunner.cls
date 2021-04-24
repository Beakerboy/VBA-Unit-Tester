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

' Function: TestCase
' Given a TestCase object, run all of its tests
Public Sub TestCase(ByVal clsObj As iTestCase)
    ' get the list of all test methods
    List = clsObj.GetAllTests()
    Dim parObj As Object
    Dim Item As Variant
    Set parObj = clsObj.GetObject()

    For Each Item In List
        Dim Method As String
        Method = Item
        ' Run the test
        RunTest parObj, Method
    Next Item
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
Public Sub RunTest(clsObj, Test As String)
    CallByName clsObj, "iTestCase_Setup", VbMethod

    'If the test cas a data provider, fetch the data array and run the test multiple times
    If EndsWith(Test, "ProviderTest") Then
        Dim i As Integer
        i = 1
        
        DataList = CallByName(clsObj, Test & "_Data", VbMethod)
        Dim Message As String
        Dim Inputs As Variant
        Dim Expected As Variant
        For Each Data In DataList
            ' Tell the logger what test we are running and which data set
            TestReporter.LogTest TypeName(clsObj), Test & " (" & i & ")"
            i = i + 1
            Inputs = Data(0)
            Message = Data(2)
            Expected = Data(1)

            ' Catch any runtime errors produced within the test
            On Error Resume Next
            CallByName clsObj, Test, VbMethod, Inputs, Expected, Message
            If Err.Number <> 0 Then
                ' If a runtime error occured, log it and continue
                TestReporter.LogError Err.Number & " - " & Err.Description
                Err.Clear
            End If
            TestReporter.FinalizeTest
            On Error GoTo 0
        Next Data
    Else
        ' Tell the logger what test we are running
        TestReporter.LogTest TypeName(clsObj), Test
        ' Catch any runtime errors produced within the test
        On Error Resume Next
        CallByName clsObj, Test, VbMethod
        If Err.Number <> 0 Then
            ' If a runtime error occured, log it and continue
            TestReporter.LogError Err.Number & " - " & Err.Description
            Err.Clear
        End If
        TestReporter.FinalizeTest
        On Error GoTo 0

    End If
    CallByName clsObj, "iTestCase_TearDown", VbMethod
End Sub

Private Function EndsWith(str As String, ending As String) As Boolean
    Dim endingLen As Integer
    endingLen = Len(ending)
    EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function
