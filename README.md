VBA Unit Tester
=====================

### Test your VBA code with VBA

Easily create Unit tests and execute them within VBA

Features
--------
 * [Standard Formatting](#format)
 * [Test Execution](#execution)
 * [Assertions](#assertions)
 
Setup
-----

Open Microsoft Visual Basic For Applications and import each cls and bas into a new project. Name the project VBAUnitTester and save it as an xlam file. Enable the addin.

Setting up a module to be tested
-----
In your VBA project, you will need to create one class module that implements the iTestableProject interface and at least one class that implements the iTestCase interface. Each TestCase must contain at least one test function, each of which should contain at least one assertion. Information on the format of each class is provided in their section of this guide.
Within Microsoft Visual Basic For Applications, select Tools>References and ensure that VBAUnitTester is selected.
Create a module with a function that calls the TestRunner with the desired options and run this macro.

 Usage
-----
## Classes

### TestRunner
A class that manages running the specified tests:
```vb
Dim MyTestConfig As iTestableProject
Set MyTestConfig = New {projectClass}

MyTestConfig.Run()

```
### TestReporter
A static (singleton) class that records which tests have run, and the results.

## Interfaces

### iTestCase
A Test Case is a collection of tests, each of which share a common set up.

### iTestableProject
A TestableProject is the set of instructions on how to execute testing and reporting. 


Assertions
-----
This library contains several assertion functions. These differ from the native Debug.Assert in that they will run any time, not only while in debug mode. Next, they do not halt execution. Instead a record is made of the number of passes and failures, and reports can be generated.
AssertTrue()
AssertFalse()
AssertEquals()
AssertNotEquals()
AssertObjectStringEquals() - Attempts to call a method named .ToString() and compares the results to the provided string.
