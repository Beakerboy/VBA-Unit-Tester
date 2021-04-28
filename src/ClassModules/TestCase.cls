Implements iTestCase

Sub iTestCase_Setup()
   
End Sub

Sub iTestCase_TearDown()
   
End Sub

Sub iTestCase_RunTest(test As String, Optional clsObj = Nothing)
    CallByName clsObj, "iTestCase_Setup", VbMethod

    'If the test cas a data provider, fetch the data array and run the test multiple times
    If EndsWith(test, "ProviderTest") Then
        Dim i As Integer
        i = 1
        
        DataList = CallByName(clsObj, test & "_Data", VbMethod)
        Dim Message As String
        Dim Inputs As Variant
        Dim Expected As Variant
        For Each Data In DataList
            ' Tell the logger what test we are running and which data set
            TestReporter.LogTest TypeName(clsObj), test & " (" & i & ")"
            i = i + 1
            Inputs = Data(0)
            Message = Data(2)
            Expected = Data(1)

            ' Catch any runtime errors produced within the test
            On Error Resume Next
            CallByName clsObj, test, VbMethod, Inputs, Expected, Message
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
        TestReporter.LogTest TypeName(clsObj), test
        ' Catch any runtime errors produced within the test
        On Error Resume Next
        CallByName clsObj, test, VbMethod
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

' Function: RunAllTests
' Given a TestCase object, run all of its tests
Public Sub iTestCase_RunAllTests(Optional ByVal clsObj = Nothing)
    ' get the list of all test methods
    Dim FullList As Variant
    
    FullList = GetObjectFunctions(TheObject:=clsObj, FuncType:=VbMethod)
    Dim List As Variant
    MaxIndex = UBound(FullList, 1)
    ReDim List(MaxIndex)
    Dim i As Integer
    For i = 0 To MaxIndex
        List(i) = FullList(i, 0)
    Next i
    
    Dim Item As Variant

    For Each Item In List
        Dim Method As String
        Method = Item
        If EndsWith(Method, "Test") And Method <> "iTestCase_RunTest" Then
            ' Run the test
            iTestCase_RunTest Method, clsObj
        End If
    Next Item
End Sub
