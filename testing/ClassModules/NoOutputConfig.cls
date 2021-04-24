VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NoOutputConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = FalseImplements iTestableProject

Dim bOutputFailures As Boolean

Dim TestRunner As TestRunner

' Constructor: Class_Initialize
' Configure the test runner
Private Sub Class_Initialize()
    ' Print each failure as it occurs
    bOutputFailures = False
End Sub

Public Property Get iTestableProject_GetOutputFailures()
    iTestableProject_GetOutputFailures = bOutputFailures
End Property

Public Function iTestableProject_GetTestCases()
    iTestableProject_GetTestCases = Array()
End Function

' Function: Run
' Perform all the actions.
Public Function iTestableProject_Run()

End Function
