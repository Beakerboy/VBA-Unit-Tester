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

Open Microsoft Visual Basic For Applications and import each cls and bas into a new project. Name the project VBAUnitTester and save it as an xlam file. Enable the addin. Within Microsoft Visual Basic For Applications, select Tools>References and ensure that VBAUnitTester is selected.

Setting up a module to be tested
-----
In you VBA project, you will need to create one class module that implements the iTestableProject interface and at least one class that implements the iTestCase interface. Each TestCase must contain at least one test function, each of which should contain at least one assertion.

Assertions
-----
This library contains several assertion functions which will open a messagebox and display a message if there is a failure.
AssertTrue()
AssertFalse()
AssertEquals()
AssertNotEquals()
AssertObjectStringEquals()
