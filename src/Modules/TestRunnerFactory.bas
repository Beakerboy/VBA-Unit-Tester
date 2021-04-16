Attribute VB_Name = "TestRunnerFactory"
Function CreateTestRunner()
    Dim Runner As New TestRunner
    Set CreateTestRunner = Runner
End Function
