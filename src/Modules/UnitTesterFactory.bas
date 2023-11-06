Attribute VB_Name = "UnitTesterFactory"
Function CreateTestRunner()
    Dim Runner As New TestRunner
    Runner.Initialize
    Set CreateTestRunner = Runner
End Function

Function CreateTestCase()
    Dim TestCase As New TestCase
    Set CreateTestCase = TestCase
End Function

Public Function EndsWith(str As String, ending As String) As Boolean
    Dim endingLen As Integer
    endingLen = Len(ending)
    EndsWith = (Right(Trim(UCase(str)), endingLen) = UCase(ending))
End Function
