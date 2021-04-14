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

Running Tests
-----
If the module to test has been correctly configured, type "=RunAllModuleTests(_module_name_)" in cell A1 to run all the tests for a particular module.

Setting up a module to be tested
-----
Create a module and create a function called {_module_name_}\_RunAllTests(). Place all the test code, or calls to every test within this function.

Assertions
-----
This library contains several assertion functions which will open a messagebox and display a message if there is a failure.
AssertTrue()
AssertFalse()
AssertEquals()
AssertNotEquals()
AssertObjectStringEquals()
