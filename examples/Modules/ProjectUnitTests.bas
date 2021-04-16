Public Function RunTests()
    
    Dim TestConfig As iTestableProject
    Dim UnitTestConfig As New TestConfigStub
    Set TestConfig = UnitTestConfig
    
    TestConfig.Run
End Function
