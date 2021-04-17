VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iTestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Interface: iTestCase
'

' Function: Setup
' Set up the environment for every test in the test case
Public Sub Setup()

End Sub

' Function: TearDown
' Restore the test environment after the test has been run.
Public Sub TearDown()

End Sub

' Function: GetAllTests
' Get a list of all functions in the class that perfom an assertion
'
' Returns:
'   An array of function names as strings.
Public Function GetAllTests()

End Function

Public Function GetObject()

End Function
