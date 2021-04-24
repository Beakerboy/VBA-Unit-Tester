Attribute VB_Name = "TestRunnerFactory"
Function CreateTestRunner()
    Dim Runner As New TestRunner
    Runner.Initialize
    Set CreateTestRunner = Runner
End Function

Function CreateTestRunnerNoInit()
    Dim Runner As New TestRunner
    Set CreateTestRunnerNoInit = Runner
End Function
