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
A Test Case is a collection of tests, each of which share a common set up. The examples directory has a stub version of this class which can be used as a starting point.

 * SetUp() - This function is run before every test in this test case.
 * TearDown() - This function is run after every test in the test case. 
       - TODO: Match arrays. How to handle objects? introspection on public properties?
 * GetTests()

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
 * AssertObjectStringEquals() - Attempts to call a method named .ToString() and compares the results to the provided string.

To Be Implemented
 * AssertSame() - Assert that the parameters match value and type. Objects are the same instances.
 * 
