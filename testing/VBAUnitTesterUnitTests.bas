Function VBAUnitTester_RunAllTests()
    VBAUnitTester_RunAllTests = True

    ' Test that assertions evaluate to True appropriately
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertTrue(True, "Expecting True when AssertTrue is given True")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertFalse(False, "Expecting True when AssertFalse is given False")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertEquals(1, 1, "Expecting True when AssertEquals is given equal integers")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertEquals("one", "one", "Expecting True when AssertEquals is given equal strings")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertEquals(1.15, 1.15, "Expecting True when AssertEquals is given equal floats")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertNotEquals(1, 2, "Expecting True when AssertNotEquals is given unequal integers")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertNotEquals("one", "One", "Expecting True when AssertNotEquals is given unequal strings")

    ' Test that assertions evaluate to false appropriately
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertFalse(AssertTrue(False, "Testing boolean False literal", "None"))
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertFalse(AssertFalse(True, "", "None"), "Expecting False when AssertFalse is given True")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertFalse(AssertEquals(1, 2, "", "None"), "Expecting False when AssertEquals is given unequal integers")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertFalse(AssertEquals("one", "One", "", "None"), "Expecting False when AssertEquals is given unequal strings")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertFalse(AssertNotEquals(1, 1, "", "None"), "Expecting False when AssertNotEquals is given equal integers")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertFalse(AssertNotEquals("one", "one", "", "None"), "Expecting False when AssertNotEquals is given uequal strings")
    VBAUnitTester_RunAllTests = VBAUnitTester_RunAllTests And AssertFalse(AssertNotEquals(1.15, 1.15, "", "None"), "Expecting False when AssertNotEquals is given uequal floats")
End Function
