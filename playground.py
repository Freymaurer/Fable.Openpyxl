from __future__ import annotations
from abc import abstractmethod
import openpyxl as openpyxl_1
from openpyxl import Workbook as Workbook_1
from typing import (Any, Protocol)
from fable_modules.fable_library.array_ import equals_with
from fable_modules.fable_library.list import of_array
from fable_modules.fable_library.option import some
from fable_modules.fable_library.seq import (to_list, delay, append, singleton)
from fable_modules.fable_library.types import Array
from fable_modules.fable_library.util import (IEnumerable_1, IEnumerable, equals, get_enumerator)
from fable_modules.fable_pyxpecto.pyxpecto import (TranspilerHelper_op_BangBang, Assert_AreEqual, Helper_expectError)
from fable_modules.fable_pyxpecto.pyxpecto import (Test_testList, Test_testCase, Expect_pass, Expect_hasLength, Model_TestCase, Pyxpecto_runTests)

class Cell(Protocol):
    @property
    @abstractmethod
    def value(self) -> Any:
        ...


class Worksheet(Protocol):
    @property
    @abstractmethod
    def title(self) -> str:
        ...

    @title.setter
    @abstractmethod
    def title(self, __arg0: str) -> None:
        ...


class Workbook(IEnumerable_1, IEnumerable[Any]):
    @abstractmethod
    def Item(self, __arg0: str) -> Worksheet:
        ...

    @abstractmethod
    def copy_worksheet(self, __arg0: Worksheet) -> Worksheet:
        ...

    @abstractmethod
    def create_sheet(self, __arg0: str, position: int | None) -> Worksheet:
        ...

    @property
    @abstractmethod
    def active(self) -> Worksheet:
        ...

    @property
    @abstractmethod
    def sheetnames(self) -> Array[str]:
        ...


class WorkbookStatic(Protocol):
    @abstractmethod
    def create(self) -> Workbook:
        ...


class OpenPyXL(Protocol):
    @abstractmethod
    def load_workbook(self, __arg0: str) -> Workbook:
        ...


openpyxl: OpenPyXL = openpyxl_1

def _arrow3(__unit: None=None) -> IEnumerable_1[Model_TestCase]:
    def body(__unit: None=None) -> None:
        wb_obj: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
        sheet_obj: Worksheet = wb_obj.active
        cell_obj: Cell = sheet_obj.cell(1, 1)
        actual: Any = cell_obj.value
        expected: Any = "A1"
        if equals(TranspilerHelper_op_BangBang(actual), TranspilerHelper_op_BangBang(expected)):
            Assert_AreEqual(actual, expected, "")

        else: 
            Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


    def _arrow2(__unit: None=None) -> IEnumerable_1[Model_TestCase]:
        def body_1(__unit: None=None) -> None:
            wb: Workbook = Workbook_1(None)
            Expect_pass()

        def body_2(__unit: None=None) -> None:
            wb_1: Workbook = Workbook_1(None)
            ws: Worksheet = wb_1.create_sheet("New Sheet")
            Expect_pass()

        def body_3(__unit: None=None) -> None:
            wb_2: Workbook = Workbook_1(None)
            ws_1: Worksheet = wb_2.create_sheet("New Sheet", 2)
            Expect_pass()

        def body_4(__unit: None=None) -> None:
            wb_3: Workbook = Workbook_1(None)
            wb_3.create_sheet("New Sheet1")
            wb_3.create_sheet("New Sheet2")
            wb_3.create_sheet("New Sheet3")
            Expect_hasLength(wb_3.sheetnames, 4, "hasLength")
            actual_1: Array[str] = wb_3.sheetnames
            expected_1: Array[str] = ["Sheet", "New Sheet1", "New Sheet2", "New Sheet3"]
            def _arrow0(x: str, y: str) -> bool:
                return x == y

            if equals_with(_arrow0, TranspilerHelper_op_BangBang(actual_1), TranspilerHelper_op_BangBang(expected_1)):
                Assert_AreEqual(actual_1, expected_1, "equal")

            else: 
                Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_1)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_1)) + "") + ('[0m')) + " \n\b    Message: ") + "equal") + "")


        def body_5(__unit: None=None) -> None:
            wb_4: Workbook = Workbook_1(None)
            ws_2: Worksheet = wb_4["Sheet"]
            actual_2: str = ws_2.title
            if TranspilerHelper_op_BangBang(actual_2) == TranspilerHelper_op_BangBang("Sheet"):
                Assert_AreEqual(actual_2, "Sheet", "")

            else: 
                Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "Sheet") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_2) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


        def body_6(__unit: None=None) -> None:
            with get_enumerator(Workbook_1(None)) as enumerator:
                while enumerator.System_Collections_IEnumerator_MoveNext():
                    sheet: Worksheet = enumerator.System_Collections_Generic_IEnumerator_1_get_Current()
                    actual_3: str = sheet.title
                    if TranspilerHelper_op_BangBang(actual_3) == TranspilerHelper_op_BangBang("Sheet"):
                        Assert_AreEqual(actual_3, "Sheet", "only 1 sheet with the default init title")

                    else: 
                        Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "Sheet") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_3) + "") + ('[0m')) + " \n\b    Message: ") + "only 1 sheet with the default init title") + "")


        def body_7(__unit: None=None) -> None:
            wb_6: Workbook = Workbook_1(None)
            source: Worksheet = wb_6.active
            copy: Worksheet = wb_6.copy_worksheet(source)
            actual_4: str = source.title
            if TranspilerHelper_op_BangBang(actual_4) == TranspilerHelper_op_BangBang("Sheet"):
                Assert_AreEqual(actual_4, "Sheet", "source title")

            else: 
                Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "Sheet") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_4) + "") + ('[0m')) + " \n\b    Message: ") + "source title") + "")

            actual_5: str = copy.title
            if TranspilerHelper_op_BangBang(actual_5) == TranspilerHelper_op_BangBang("Sheet Copy"):
                Assert_AreEqual(actual_5, "Sheet Copy", "copy title")

            else: 
                Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "Sheet Copy") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_5) + "") + ('[0m')) + " \n\b    Message: ") + "copy title") + "")

            Expect_hasLength(wb_6.sheetnames, 2, "hasLenght")

        def _arrow1(__unit: None=None) -> IEnumerable_1[Model_TestCase]:
            def body_8(__unit: None=None) -> None:
                wb_7: Workbook = Workbook_1(None)
                ws_3: Worksheet = wb_7.create_sheet("New Sheet")
                actual_6: str = ws_3.title
                if TranspilerHelper_op_BangBang(actual_6) == TranspilerHelper_op_BangBang("New Sheet"):
                    Assert_AreEqual(actual_6, "New Sheet", "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "New Sheet") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_6) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_9(__unit: None=None) -> None:
                wb_8: Workbook = Workbook_1(None)
                ws_4: Worksheet = wb_8.create_sheet("New Sheet")
                ws_4.title = "New Title"
                actual_7: str = ws_4.title
                if TranspilerHelper_op_BangBang(actual_7) == TranspilerHelper_op_BangBang("New Title"):
                    Assert_AreEqual(actual_7, "New Title", "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "New Title") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_7) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_10(__unit: None=None) -> None:
                wb_obj_1: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                sheet_obj_1: Worksheet = wb_obj_1.active
                cell_obj_1: Cell = sheet_obj_1["A1"]
                actual_8: Any = cell_obj_1.value
                expected_10: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_8), TranspilerHelper_op_BangBang(expected_10)):
                    Assert_AreEqual(actual_8, expected_10, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_10)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_8)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_11(__unit: None=None) -> None:
                wb_9: Workbook = Workbook_1(None)
                ws_5: Worksheet = wb_9.active
                ws_5["A1"] = 42
                cell: Cell = ws_5["A1"]
                actual_9: Any = cell.value
                expected_11: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_9), TranspilerHelper_op_BangBang(expected_11)):
                    Assert_AreEqual(actual_9, expected_11, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_11)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_9)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_12(__unit: None=None) -> None:
                wb_10: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_6: Worksheet = wb_10.active
                cols: Array[Array[Cell]] = [list(inner_tuple) for inner_tuple in ws_6[(("A1", "C1"))[0] : (("A1", "C1"))[1]]]
                Expect_hasLength(cols, 1, "column lenght")
                Expect_hasLength(cols[0], 3, "rows lenght")
                actual_10: Any = cols[0][0].value
                expected_12: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_10), TranspilerHelper_op_BangBang(expected_12)):
                    Assert_AreEqual(actual_10, expected_12, "A1-value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_12)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_10)) + "") + ('[0m')) + " \n\b    Message: ") + "A1-value") + "")


            def body_13(__unit: None=None) -> None:
                wb_11: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_7: Worksheet = wb_11.active
                cols_1: Array[Cell] = list(ws_7["A"])
                Expect_hasLength(cols_1, 3, "column cell lenght")
                actual_11: Any = cols_1[0].value
                expected_13: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_11), TranspilerHelper_op_BangBang(expected_13)):
                    Assert_AreEqual(actual_11, expected_13, "A1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_13)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_11)) + "") + ('[0m')) + " \n\b    Message: ") + "A1 - value") + "")

                actual_12: Any = cols_1[1].value
                expected_14: Any = "A2"
                if equals(TranspilerHelper_op_BangBang(actual_12), TranspilerHelper_op_BangBang(expected_14)):
                    Assert_AreEqual(actual_12, expected_14, "A2 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_14)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_12)) + "") + ('[0m')) + " \n\b    Message: ") + "A2 - value") + "")

                actual_13: Any = cols_1[2].value
                expected_15: Any = "A3"
                if equals(TranspilerHelper_op_BangBang(actual_13), TranspilerHelper_op_BangBang(expected_15)):
                    Assert_AreEqual(actual_13, expected_15, "A3 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_15)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_13)) + "") + ('[0m')) + " \n\b    Message: ") + "A3 - value") + "")


            def body_14(__unit: None=None) -> None:
                wb_12: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_8: Worksheet = wb_12.active
                cols_2: Array[Cell] = list(ws_8[1])
                Expect_hasLength(cols_2, 3, "column cell lenght")
                actual_14: Any = cols_2[0].value
                expected_16: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_14), TranspilerHelper_op_BangBang(expected_16)):
                    Assert_AreEqual(actual_14, expected_16, "A1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_16)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_14)) + "") + ('[0m')) + " \n\b    Message: ") + "A1 - value") + "")

                actual_15: Any = cols_2[1].value
                expected_17: Any = "B1"
                if equals(TranspilerHelper_op_BangBang(actual_15), TranspilerHelper_op_BangBang(expected_17)):
                    Assert_AreEqual(actual_15, expected_17, "B1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_17)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_15)) + "") + ('[0m')) + " \n\b    Message: ") + "B1 - value") + "")

                actual_16: Any = cols_2[2].value
                expected_18: Any = "C1"
                if equals(TranspilerHelper_op_BangBang(actual_16), TranspilerHelper_op_BangBang(expected_18)):
                    Assert_AreEqual(actual_16, expected_18, "C1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_18)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_16)) + "") + ('[0m')) + " \n\b    Message: ") + "C1 - value") + "")


            def body_15(__unit: None=None) -> None:
                wb_13: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_9: Worksheet = wb_13.active
                cols_3: Array[Array[Cell]] = [list(inner_tuple) for inner_tuple in ws_9[('{start}:{end}'.format(start=(("A", "B"))[0],end=(("A", "B"))[1]))]]
                Expect_hasLength(cols_3, 2, "hasLength")
                for idx in range(0, (len(cols_3) - 1) + 1, 1):
                    Expect_hasLength(cols_3[idx], 3, "inner has Length")
                actual_17: Any = cols_3[0][0].value
                expected_19: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_17), TranspilerHelper_op_BangBang(expected_19)):
                    Assert_AreEqual(actual_17, expected_19, "A1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_19)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_17)) + "") + ('[0m')) + " \n\b    Message: ") + "A1 - value") + "")

                actual_18: Any = cols_3[1][2].value
                expected_20: Any = "B3"
                if equals(TranspilerHelper_op_BangBang(actual_18), TranspilerHelper_op_BangBang(expected_20)):
                    Assert_AreEqual(actual_18, expected_20, "B3 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_20)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_18)) + "") + ('[0m')) + " \n\b    Message: ") + "B3 - value") + "")


            def body_16(__unit: None=None) -> None:
                wb_14: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_10: Worksheet = wb_14.active
                rows: Array[Array[Cell]] = [list(inner_tuple) for inner_tuple in ws_10[((1, 2))[0] : ((1, 2))[1]]]
                Expect_hasLength(rows, 2, "hasLength")
                for idx_1 in range(0, (len(rows) - 1) + 1, 1):
                    Expect_hasLength(rows[idx_1], 3, "inner has Length")
                actual_19: Any = rows[0][0].value
                expected_21: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_19), TranspilerHelper_op_BangBang(expected_21)):
                    Assert_AreEqual(actual_19, expected_21, "A1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_21)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_19)) + "") + ('[0m')) + " \n\b    Message: ") + "A1 - value") + "")

                actual_20: Any = rows[1][2].value
                expected_22: Any = "C2"
                if equals(TranspilerHelper_op_BangBang(actual_20), TranspilerHelper_op_BangBang(expected_22)):
                    Assert_AreEqual(actual_20, expected_22, "C2 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_22)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_20)) + "") + ('[0m')) + " \n\b    Message: ") + "C2 - value") + "")


            def body_17(__unit: None=None) -> None:
                wb_15: Workbook = Workbook_1(None)
                ws_11: Worksheet = wb_15.active
                ws_11["A1"] = 42
                cell_1: Cell = ws_11.cell(1, 1)
                actual_21: Any = cell_1.value
                expected_23: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_21), TranspilerHelper_op_BangBang(expected_23)):
                    Assert_AreEqual(actual_21, expected_23, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_23)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_21)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_18(__unit: None=None) -> None:
                wb_16: Workbook = Workbook_1(None)
                ws_12: Worksheet = wb_16.active
                cell_2: Cell = ws_12.cell(1, 1, some(42))
                actual_22: Any = cell_2.value
                expected_24: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_22), TranspilerHelper_op_BangBang(expected_24)):
                    Assert_AreEqual(actual_22, expected_24, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_24)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_22)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            return singleton(Test_testList("Worksheet", of_array([Test_testCase("init title", body_8), Test_testCase("set title", body_9), Test_testList("Item", of_array([Test_testCase("get cell", body_10), Test_testCase("Create Cell", body_11), Test_testCase("get cell range", body_12), Test_testCase("get column", body_13), Test_testCase("get row", body_14), Test_testCase("get columns", body_15), Test_testCase("get rows", body_16)])), Test_testCase("get via _.cell", body_17), Test_testCase("set via _.cell", body_18)])))

        return append(singleton(Test_testList("Workbook", of_array([Test_testCase("Create new", body_1), Test_testCase("create new worksheet", body_2), Test_testCase("create new worksheet at position", body_3), Test_testCase("sheetnames", body_4), Test_testCase("Item", body_5), Test_testCase("for sheet in wb", body_6), Test_testCase("copy_worksheet", body_7)]))), delay(_arrow1))

    return append(singleton(Test_testCase("Minimal Read", body)), delay(_arrow2))


tests: Model_TestCase = Test_testList("main", to_list(delay(_arrow3)))

Pyxpecto_runTests([], tests)

