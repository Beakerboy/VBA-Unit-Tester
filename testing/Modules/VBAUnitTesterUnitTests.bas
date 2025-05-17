Attribute VB_Name = "VBAUnitTesterUnitTests"
Public Function RunTests()
    
    Dim TestConfig As iTestableProject
    Dim UnitTestertestConfig As New UnitTestertestConfig
    Set TestConfig = UnitTestertestConfig
    
    TestConfig.Run
End Function
