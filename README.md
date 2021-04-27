VBA Unit Tester
=====================

### Test your VBA code with VBA

Easily create Unit tests, execute them, and see reports, all within VBA.

Features
--------
 * [Setting up the software](#Setup)
 * [Using the library](#usage)
 * [Assertions](#assertions)
 
Setup
-----

Open Microsoft Visual Basic For Applications and import each cls and bas from the src directory into a new project. Name the project VBAUnitTester and save it as an xlam file. Enable the addin.

Setting up a module to be tested
-----
This project is used to test itself. Refer to the testing directory to see a working example.
In your VBA project, you will need to create one class module that implements the iTestableProject interface and at least one class that implements the iTestCase interface. Each TestCase should contain at least one test function, each of which should contain at least one assertion. Information on the format of each class is provided in their section of this guide.
Within Microsoft Visual Basic For Applications, select Tools>References and ensure that VBAUnitTester is selected.
Create a module with a publicly visible function that calls the TestRunner with the desired options and run this macro.

Projects that use this code
-----
 * [VBA-Math-Objects](https://github.com/Beakerboy/VBA-Math-Objects)
 * [VBA-SQL-Library](https://github.com/Beakerboy/VBA-SQL-Library)

 Usage
-----
## Classes

### TestRunner
A class that manages running the specified tests:

Tests can be executed at three different levels:
 * RunTest (iTestCase, Method) - Executes one test in one TestCase
 * TestCase(iTestCase) - Executes all tests in a supplied TestCase
 * TestAllCases - Requests a list of all TestCase classes and executes all tests in each

### TestReporter
A static (singleton) class that records which tests have run, and the results.

## Interfaces
These interfaces must be implemented by users who wish to use this project to test their code.

### iTestCase
A Test Case is a collection of tests, each of which share a common set up. The examples directory has a stub version of this class which can be used as a starting point. All tests should end with the string "Test"

 * SetUp() - This function is run before every test in this test case.
 * TearDown() - This function is run after every test in the test case. 
       - TODO: Match arrays. How to handle objects? introspection on public properties?
 * RunAllTests() - Run all tests in the test case. For ease of use, call the parent `TestCase.RunAllTests Me`
 * RunTest() - Run a specific test within this test case. For ease of use, call the parent `TestCase.RunTest Test, Me`

There are two types of tests that can be written. A basic test runs without any input parameteres. If you wish to run a test multiple times with different parameters, indicate this by ensuring the function name ends with "ProviderTest", and create another function with the same name and append "\_Data" to it.
For example:
```vba
Sub MathProviderTest(Inputs, Expected, Message As String)
    Call AssertEquals(Inputs(0) + Inputs(1), Expected, Message)
End Sub

Function MathProviderTest_Data()
    MathProviderTest_Data = Array( _
        Array(Array(1, 1), 2, "1+1 Works"), _
        Array(Array(2, 2), 4, "2+2 Works") _
    )
```

### iTestableProject
A TestableProject is the set of instructions on how to execute testing and reporting. The examples directory has a stub file than be quickly edited as a starting point.


Assertions
-----
This library contains several assertion functions. These differ from the native Debug.Assert in that they will run any time, not only while in debug mode. Next, they do not halt execution. Instead a record is made of the number of passes and failures, and reports can be generated.
 * AssertTrue() - Verify that the provided parameter is a boolean True value.
 * AssertFalse() - Verify that the provided parameter is a boolean False value.
 * AssertEquals() - Verify that parameters equal one another. They may have a different type. 
       - TODO: Match arrays. How to handle objects? introspection on public properties?
 * AssertNotEquals()
 * AssertObjectStringEquals() - Attempts to call a method named `.ToString()` and compares the results to the provided string.
 * AssertSame() - Assert that the parameters match value and type. ToDo: check that Objects are the same instances.
 * ExpectError() - If you would like to test that a given set of inputs will trigger an error, place a call to `EpectError()` immediately preceeding the call in the test. No code after the error-producing call will execute, so make that the last call of the test function.

