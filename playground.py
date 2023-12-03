from __future__ import annotations
from abc import abstractmethod
import openpyxl as openpyxl_1
from openpyxl import Workbook as Workbook_1
from openpyxl.worksheet.table import Table as Table_1
from typing import (Any, Protocol, Callable)
from fable_modules.fable_library.array_ import equals_with
from fable_modules.fable_library.date import now
from fable_modules.fable_library.list import of_array
from fable_modules.fable_library.option import some
from fable_modules.fable_library.reflection import (TypeInfo, union_type)
from fable_modules.fable_library.seq import (to_list, delay, append, singleton)
from fable_modules.fable_library.string_ import (to_console, printf)
from fable_modules.fable_library.types import (Array, Union)
from fable_modules.fable_library.util import (IEnumerable_1, IEnumerable, equals, get_enumerator, equal_arrays)
from fable_modules.fable_pyxpecto.pyxpecto import (TranspilerHelper_op_BangBang, Assert_AreEqual, Helper_expectError)
from fable_modules.fable_pyxpecto.pyxpecto import (Test_testList, Test_testCase, Expect_pass, Expect_hasLength, Model_TestCase, Pyxpecto_runTests)

def _expr0() -> TypeInfo:
    return union_type("Playground.CellType", [], CellType, lambda: [[], [], [], [], []])


class CellType(Union):
    def __init__(self, tag: int, *fields: Any) -> None:
        super().__init__()
        self.tag: int = tag or 0
        self.fields: Array[Any] = list(fields)

    @staticmethod
    def cases() -> list[str]:
        return ["Float", "Integer", "String", "Boolean", "DateTime"]


CellType_reflection = _expr0

def CellType_fromCellType_Z721C83C5(cell_type: str) -> CellType:
    if cell_type == "int":
        return CellType(1)

    elif cell_type == "float":
        return CellType(0)

    elif cell_type == "str":
        return CellType(2)

    elif cell_type == "bool":
        return CellType(3)

    elif cell_type == "datetime":
        return CellType(4)

    else: 
        raise Exception(("Unknown cell type of type: \'" + cell_type) + "\'")



class Cell(Protocol):
    @property
    @abstractmethod
    def cell_type(self) -> str:
        ...

    @property
    @abstractmethod
    def value(self) -> Any:
        ...

    @value.setter
    @abstractmethod
    def value(self, __arg0: Any) -> None:
        ...


class Table(Protocol):
    @property
    @abstractmethod
    def display_name(self) -> str:
        ...

    @display_name.setter
    @abstractmethod
    def display_name(self, __arg0: str) -> None:
        ...

    @property
    @abstractmethod
    def header_row_count(self) -> bool:
        ...

    @header_row_count.setter
    @abstractmethod
    def header_row_count(self, __arg0: bool) -> None:
        ...

    @property
    @abstractmethod
    def id(self) -> int:
        ...

    @id.setter
    @abstractmethod
    def id(self, __arg0: int) -> None:
        ...

    @property
    @abstractmethod
    def name(self) -> str:
        ...

    @name.setter
    @abstractmethod
    def name(self, __arg0: str) -> None:
        ...


class TableMap(Protocol):
    @abstractmethod
    def Item(self, __arg0: str) -> Table:
        ...

    @abstractmethod
    def delete(self, displayName: str) -> None:
        ...

    @abstractmethod
    def items(self) -> Array[tuple[str, str]]:
        ...

    @abstractmethod
    def values(self) -> Array[Table]:
        ...


class Worksheet(Protocol):
    @abstractmethod
    def add_table(self, __arg0: Table) -> None:
        ...

    @abstractmethod
    def append(self, __arg0: Array[Any]) -> None:
        ...

    @abstractmethod
    def delete_cols(self, start_index: int, count: int) -> None:
        ...

    @abstractmethod
    def delete_rows(self, start_index: int, count: int) -> None:
        ...

    @abstractmethod
    def delete_table(self, displayName: str) -> None:
        ...

    @property
    @abstractmethod
    def columns(self) -> Array[Array[Cell]]:
        ...

    @property
    @abstractmethod
    def rows(self) -> Array[Array[Cell]]:
        ...

    @property
    @abstractmethod
    def table_count(self) -> int:
        ...

    @property
    @abstractmethod
    def tables(self) -> TableMap:
        ...

    @property
    @abstractmethod
    def title(self) -> str:
        ...

    @title.setter
    @abstractmethod
    def title(self, __arg0: str) -> None:
        ...

    @property
    @abstractmethod
    def values(self) -> Array[Array[Any]]:
        ...

    @abstractmethod
    def insert_cols(self, __arg0: int) -> None:
        ...

    @abstractmethod
    def insert_rows(self, __arg0: int) -> None:
        ...

    @abstractmethod
    def iter_cols(self, min_row: int, max_col: int, max_row: int, action: Callable[[Array[Cell]], None]) -> None:
        ...

    @abstractmethod
    def iter_rows(self, min_row: int, max_col: int, max_row: int, action: Callable[[Array[Cell]], None]) -> None:
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
    def iso_dates(self) -> bool:
        ...

    @iso_dates.setter
    @abstractmethod
    def iso_dates(self, __arg0: bool) -> None:
        ...

    @property
    @abstractmethod
    def sheetnames(self) -> Array[str]:
        ...

    @property
    @abstractmethod
    def template(self) -> bool:
        ...

    @template.setter
    @abstractmethod
    def template(self, __arg0: bool) -> None:
        ...

    @abstractmethod
    def save(self, path: str) -> None:
        ...


class WorkbookStatic(Protocol):
    @abstractmethod
    def create(self) -> Workbook:
        ...


class TableStatic(Protocol):
    @abstractmethod
    def create(self, displayName: str, ref: str) -> Table:
        ...


class OpenPyXL(Protocol):
    @abstractmethod
    def Workbook(self) -> Workbook:
        ...

    @abstractmethod
    def load_workbook(self, __arg0: str) -> Workbook:
        ...


openpyxl: OpenPyXL = openpyxl_1

def _arrow22(__unit: None=None) -> IEnumerable_1[Model_TestCase]:
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


    def body_1(__unit: None=None) -> None:
        wb: Workbook = openpyxl.Workbook()
        ws: Worksheet = wb.active
        actual_1: str = ws.title
        if TranspilerHelper_op_BangBang(actual_1) == TranspilerHelper_op_BangBang("Sheet"):
            Assert_AreEqual(actual_1, "Sheet", "")

        else: 
            Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "Sheet") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_1) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


    def _arrow21(__unit: None=None) -> IEnumerable_1[Model_TestCase]:
        def body_2(__unit: None=None) -> None:
            wb_1: Workbook = Workbook_1(None)
            Expect_pass()

        def body_3(__unit: None=None) -> None:
            wb_2: Workbook = Workbook_1(None)
            ws_1: Worksheet = wb_2.create_sheet("New Sheet")
            Expect_pass()

        def body_4(__unit: None=None) -> None:
            wb_3: Workbook = Workbook_1(None)
            ws_2: Worksheet = wb_3.create_sheet("New Sheet", 2)
            Expect_pass()

        def body_5(__unit: None=None) -> None:
            wb_4: Workbook = Workbook_1(None)
            wb_4.create_sheet("New Sheet1")
            wb_4.create_sheet("New Sheet2")
            wb_4.create_sheet("New Sheet3")
            Expect_hasLength(wb_4.sheetnames, 4, "hasLength")
            actual_2: Array[str] = wb_4.sheetnames
            expected_2: Array[str] = ["Sheet", "New Sheet1", "New Sheet2", "New Sheet3"]
            def _arrow6(x: str, y: str) -> bool:
                return x == y

            if equals_with(_arrow6, TranspilerHelper_op_BangBang(actual_2), TranspilerHelper_op_BangBang(expected_2)):
                Assert_AreEqual(actual_2, expected_2, "equal")

            else: 
                Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_2)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_2)) + "") + ('[0m')) + " \n\b    Message: ") + "equal") + "")


        def body_6(__unit: None=None) -> None:
            wb_5: Workbook = Workbook_1(None)
            ws_3: Worksheet = wb_5["Sheet"]
            actual_3: str = ws_3.title
            if TranspilerHelper_op_BangBang(actual_3) == TranspilerHelper_op_BangBang("Sheet"):
                Assert_AreEqual(actual_3, "Sheet", "")

            else: 
                Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "Sheet") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_3) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


        def body_7(__unit: None=None) -> None:
            with get_enumerator(Workbook_1(None)) as enumerator:
                while enumerator.System_Collections_IEnumerator_MoveNext():
                    sheet: Worksheet = enumerator.System_Collections_Generic_IEnumerator_1_get_Current()
                    actual_4: str = sheet.title
                    if TranspilerHelper_op_BangBang(actual_4) == TranspilerHelper_op_BangBang("Sheet"):
                        Assert_AreEqual(actual_4, "Sheet", "only 1 sheet with the default init title")

                    else: 
                        Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "Sheet") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_4) + "") + ('[0m')) + " \n\b    Message: ") + "only 1 sheet with the default init title") + "")


        def body_8(__unit: None=None) -> None:
            wb_7: Workbook = Workbook_1(None)
            source: Worksheet = wb_7.active
            copy: Worksheet = wb_7.copy_worksheet(source)
            actual_5: str = source.title
            if TranspilerHelper_op_BangBang(actual_5) == TranspilerHelper_op_BangBang("Sheet"):
                Assert_AreEqual(actual_5, "Sheet", "source title")

            else: 
                Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "Sheet") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_5) + "") + ('[0m')) + " \n\b    Message: ") + "source title") + "")

            actual_6: str = copy.title
            if TranspilerHelper_op_BangBang(actual_6) == TranspilerHelper_op_BangBang("Sheet Copy"):
                Assert_AreEqual(actual_6, "Sheet Copy", "copy title")

            else: 
                Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "Sheet Copy") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_6) + "") + ('[0m')) + " \n\b    Message: ") + "copy title") + "")

            Expect_hasLength(wb_7.sheetnames, 2, "hasLenght")

        def _arrow20(__unit: None=None) -> IEnumerable_1[Model_TestCase]:
            def body_9(__unit: None=None) -> None:
                wb_8: Workbook = Workbook_1(None)
                ws_4: Worksheet = wb_8.create_sheet("New Sheet")
                actual_7: str = ws_4.title
                if TranspilerHelper_op_BangBang(actual_7) == TranspilerHelper_op_BangBang("New Sheet"):
                    Assert_AreEqual(actual_7, "New Sheet", "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "New Sheet") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_7) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_10(__unit: None=None) -> None:
                wb_9: Workbook = Workbook_1(None)
                ws_5: Worksheet = wb_9.create_sheet("New Sheet")
                ws_5.title = "New Title"
                actual_8: str = ws_5.title
                if TranspilerHelper_op_BangBang(actual_8) == TranspilerHelper_op_BangBang("New Title"):
                    Assert_AreEqual(actual_8, "New Title", "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "New Title") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_8) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_11(__unit: None=None) -> None:
                wb_obj_1: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                sheet_obj_1: Worksheet = wb_obj_1.active
                cell_obj_1: Cell = sheet_obj_1["A1"]
                actual_9: Any = cell_obj_1.value
                expected_11: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_9), TranspilerHelper_op_BangBang(expected_11)):
                    Assert_AreEqual(actual_9, expected_11, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_11)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_9)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_12(__unit: None=None) -> None:
                wb_10: Workbook = Workbook_1(None)
                ws_6: Worksheet = wb_10.active
                ws_6["A1"] = 42
                cell: Cell = ws_6["A1"]
                actual_10: Any = cell.value
                expected_12: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_10), TranspilerHelper_op_BangBang(expected_12)):
                    Assert_AreEqual(actual_10, expected_12, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_12)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_10)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_13(__unit: None=None) -> None:
                wb_11: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_7: Worksheet = wb_11.active
                cols: Array[Array[Cell]] = [list(inner_tuple) for inner_tuple in ws_7[(("A1", "C1"))[0] : (("A1", "C1"))[1]]]
                Expect_hasLength(cols, 1, "column lenght")
                Expect_hasLength(cols[0], 3, "rows lenght")
                actual_11: Any = cols[0][0].value
                expected_13: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_11), TranspilerHelper_op_BangBang(expected_13)):
                    Assert_AreEqual(actual_11, expected_13, "A1-value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_13)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_11)) + "") + ('[0m')) + " \n\b    Message: ") + "A1-value") + "")


            def body_14(__unit: None=None) -> None:
                wb_12: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_8: Worksheet = wb_12.active
                cols_1: Array[Cell] = list(ws_8["A"])
                Expect_hasLength(cols_1, 3, "column cell lenght")
                actual_12: Any = cols_1[0].value
                expected_14: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_12), TranspilerHelper_op_BangBang(expected_14)):
                    Assert_AreEqual(actual_12, expected_14, "A1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_14)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_12)) + "") + ('[0m')) + " \n\b    Message: ") + "A1 - value") + "")

                actual_13: Any = cols_1[1].value
                expected_15: Any = "A2"
                if equals(TranspilerHelper_op_BangBang(actual_13), TranspilerHelper_op_BangBang(expected_15)):
                    Assert_AreEqual(actual_13, expected_15, "A2 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_15)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_13)) + "") + ('[0m')) + " \n\b    Message: ") + "A2 - value") + "")

                actual_14: Any = cols_1[2].value
                expected_16: Any = "A3"
                if equals(TranspilerHelper_op_BangBang(actual_14), TranspilerHelper_op_BangBang(expected_16)):
                    Assert_AreEqual(actual_14, expected_16, "A3 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_16)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_14)) + "") + ('[0m')) + " \n\b    Message: ") + "A3 - value") + "")


            def body_15(__unit: None=None) -> None:
                wb_13: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_9: Worksheet = wb_13.active
                cols_2: Array[Cell] = list(ws_9[1])
                Expect_hasLength(cols_2, 3, "column cell lenght")
                actual_15: Any = cols_2[0].value
                expected_17: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_15), TranspilerHelper_op_BangBang(expected_17)):
                    Assert_AreEqual(actual_15, expected_17, "A1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_17)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_15)) + "") + ('[0m')) + " \n\b    Message: ") + "A1 - value") + "")

                actual_16: Any = cols_2[1].value
                expected_18: Any = "B1"
                if equals(TranspilerHelper_op_BangBang(actual_16), TranspilerHelper_op_BangBang(expected_18)):
                    Assert_AreEqual(actual_16, expected_18, "B1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_18)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_16)) + "") + ('[0m')) + " \n\b    Message: ") + "B1 - value") + "")

                actual_17: Any = cols_2[2].value
                expected_19: Any = "C1"
                if equals(TranspilerHelper_op_BangBang(actual_17), TranspilerHelper_op_BangBang(expected_19)):
                    Assert_AreEqual(actual_17, expected_19, "C1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_19)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_17)) + "") + ('[0m')) + " \n\b    Message: ") + "C1 - value") + "")


            def body_16(__unit: None=None) -> None:
                wb_14: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_10: Worksheet = wb_14.active
                cols_3: Array[Array[Cell]] = [list(inner_tuple) for inner_tuple in ws_10[('{start}:{end}'.format(start=(("A", "B"))[0],end=(("A", "B"))[1]))]]
                Expect_hasLength(cols_3, 2, "hasLength")
                for idx in range(0, (len(cols_3) - 1) + 1, 1):
                    Expect_hasLength(cols_3[idx], 3, "inner has Length")
                actual_18: Any = cols_3[0][0].value
                expected_20: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_18), TranspilerHelper_op_BangBang(expected_20)):
                    Assert_AreEqual(actual_18, expected_20, "A1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_20)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_18)) + "") + ('[0m')) + " \n\b    Message: ") + "A1 - value") + "")

                actual_19: Any = cols_3[1][2].value
                expected_21: Any = "B3"
                if equals(TranspilerHelper_op_BangBang(actual_19), TranspilerHelper_op_BangBang(expected_21)):
                    Assert_AreEqual(actual_19, expected_21, "B3 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_21)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_19)) + "") + ('[0m')) + " \n\b    Message: ") + "B3 - value") + "")


            def body_17(__unit: None=None) -> None:
                wb_15: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_11: Worksheet = wb_15.active
                rows: Array[Array[Cell]] = [list(inner_tuple) for inner_tuple in ws_11[((1, 2))[0] : ((1, 2))[1]]]
                Expect_hasLength(rows, 2, "hasLength")
                for idx_1 in range(0, (len(rows) - 1) + 1, 1):
                    Expect_hasLength(rows[idx_1], 3, "inner has Length")
                actual_20: Any = rows[0][0].value
                expected_22: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_20), TranspilerHelper_op_BangBang(expected_22)):
                    Assert_AreEqual(actual_20, expected_22, "A1 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_22)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_20)) + "") + ('[0m')) + " \n\b    Message: ") + "A1 - value") + "")

                actual_21: Any = rows[1][2].value
                expected_23: Any = "C2"
                if equals(TranspilerHelper_op_BangBang(actual_21), TranspilerHelper_op_BangBang(expected_23)):
                    Assert_AreEqual(actual_21, expected_23, "C2 - value")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_23)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_21)) + "") + ('[0m')) + " \n\b    Message: ") + "C2 - value") + "")


            def body_18(__unit: None=None) -> None:
                wb_16: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_12: Worksheet = wb_16.active
                def _arrow10(row_1: Array[Cell]) -> None:
                    for idx_2 in range(0, (len(row_1) - 1) + 1, 1):
                        cell_1: Cell = row_1[idx_2]
                        cell_1.value = 42

                for row in ws_12.iter_rows(min_row=1, max_col=3, max_row=2): _arrow10(row)
                actual_22: Any = ws_12["A1"].value
                expected_24: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_22), TranspilerHelper_op_BangBang(expected_24)):
                    Assert_AreEqual(actual_22, expected_24, "A1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_24)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_22)) + "") + ('[0m')) + " \n\b    Message: ") + "A1") + "")

                actual_23: Any = ws_12["A2"].value
                expected_25: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_23), TranspilerHelper_op_BangBang(expected_25)):
                    Assert_AreEqual(actual_23, expected_25, "A2")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_25)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_23)) + "") + ('[0m')) + " \n\b    Message: ") + "A2") + "")

                actual_24: Any = ws_12["B1"].value
                expected_26: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_24), TranspilerHelper_op_BangBang(expected_26)):
                    Assert_AreEqual(actual_24, expected_26, "B1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_26)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_24)) + "") + ('[0m')) + " \n\b    Message: ") + "B1") + "")

                actual_25: Any = ws_12["B2"].value
                expected_27: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_25), TranspilerHelper_op_BangBang(expected_27)):
                    Assert_AreEqual(actual_25, expected_27, "B2")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_27)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_25)) + "") + ('[0m')) + " \n\b    Message: ") + "B2") + "")

                actual_26: Any = ws_12["C1"].value
                expected_28: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_26), TranspilerHelper_op_BangBang(expected_28)):
                    Assert_AreEqual(actual_26, expected_28, "C1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_28)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_26)) + "") + ('[0m')) + " \n\b    Message: ") + "C1") + "")

                actual_27: Any = ws_12["C2"].value
                expected_29: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_27), TranspilerHelper_op_BangBang(expected_29)):
                    Assert_AreEqual(actual_27, expected_29, "C2")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_29)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_27)) + "") + ('[0m')) + " \n\b    Message: ") + "C2") + "")


            def body_19(__unit: None=None) -> None:
                wb_17: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_13: Worksheet = wb_17.active
                def _arrow11(col_1: Array[Cell]) -> None:
                    for idx_3 in range(0, (len(col_1) - 1) + 1, 1):
                        cell_2: Cell = col_1[idx_3]
                        cell_2.value = 42

                for col in ws_13.iter_cols(min_row=1, max_col=3, max_row=2): _arrow11(col)
                actual_28: Any = ws_13["A1"].value
                expected_30: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_28), TranspilerHelper_op_BangBang(expected_30)):
                    Assert_AreEqual(actual_28, expected_30, "A1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_30)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_28)) + "") + ('[0m')) + " \n\b    Message: ") + "A1") + "")

                actual_29: Any = ws_13["A2"].value
                expected_31: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_29), TranspilerHelper_op_BangBang(expected_31)):
                    Assert_AreEqual(actual_29, expected_31, "A2")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_31)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_29)) + "") + ('[0m')) + " \n\b    Message: ") + "A2") + "")

                actual_30: Any = ws_13["B1"].value
                expected_32: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_30), TranspilerHelper_op_BangBang(expected_32)):
                    Assert_AreEqual(actual_30, expected_32, "B1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_32)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_30)) + "") + ('[0m')) + " \n\b    Message: ") + "B1") + "")

                actual_31: Any = ws_13["B2"].value
                expected_33: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_31), TranspilerHelper_op_BangBang(expected_33)):
                    Assert_AreEqual(actual_31, expected_33, "B2")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_33)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_31)) + "") + ('[0m')) + " \n\b    Message: ") + "B2") + "")

                actual_32: Any = ws_13["C1"].value
                expected_34: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_32), TranspilerHelper_op_BangBang(expected_34)):
                    Assert_AreEqual(actual_32, expected_34, "C1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_34)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_32)) + "") + ('[0m')) + " \n\b    Message: ") + "C1") + "")

                actual_33: Any = ws_13["C2"].value
                expected_35: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_33), TranspilerHelper_op_BangBang(expected_35)):
                    Assert_AreEqual(actual_33, expected_35, "C2")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_35)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_33)) + "") + ('[0m')) + " \n\b    Message: ") + "C2") + "")


            def body_20(__unit: None=None) -> None:
                wb_18: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_14: Worksheet = wb_18.active
                rows_1: Array[Array[Cell]] = [list(inner_tuple) for inner_tuple in ws_14.rows]
                Expect_hasLength(rows_1, 3, "hasLenght")
                actual_34: Any = rows_1[0][0].value
                expected_36: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_34), TranspilerHelper_op_BangBang(expected_36)):
                    Assert_AreEqual(actual_34, expected_36, "A1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_36)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_34)) + "") + ('[0m')) + " \n\b    Message: ") + "A1") + "")


            def body_21(__unit: None=None) -> None:
                wb_19: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_15: Worksheet = wb_19.active
                rows_2: Array[Array[Cell]] = [list(inner_tuple) for inner_tuple in ws_15.columns]
                Expect_hasLength(rows_2, 3, "hasLenght")
                actual_35: Any = rows_2[0][0].value
                expected_37: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_35), TranspilerHelper_op_BangBang(expected_37)):
                    Assert_AreEqual(actual_35, expected_37, "A1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_37)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_35)) + "") + ('[0m')) + " \n\b    Message: ") + "A1") + "")


            def body_22(__unit: None=None) -> None:
                wb_20: Workbook = openpyxl.load_workbook("C:\\Users\\Kevin\\Desktop\\BookTest.xlsx")
                ws_16: Worksheet = wb_20.active
                rows_3: Array[Array[Any]] = [list(inner_tuple) for inner_tuple in ws_16.values]
                Expect_hasLength(rows_3, 3, "hasLenght")
                actual_36: Any = rows_3[0][0]
                expected_38: Any = "A1"
                if equals(TranspilerHelper_op_BangBang(actual_36), TranspilerHelper_op_BangBang(expected_38)):
                    Assert_AreEqual(actual_36, expected_38, "A1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_38)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_36)) + "") + ('[0m')) + " \n\b    Message: ") + "A1") + "")


            def body_23(__unit: None=None) -> None:
                wb_21: Workbook = Workbook_1(None)
                ws_17: Worksheet = wb_21.active
                treedata: Array[Array[Any]] = [["Type", "Leaf Color", "Height"], ["Maple", "Red", 549], ["Oak", "Green", 783], ["Pine", "Green", 1204]]
                for idx_4 in range(0, (len(treedata) - 1) + 1, 1):
                    row_2: Array[Any] = treedata[idx_4]
                    ws_17.append(row_2)
                Expect_hasLength([list(inner_tuple) for inner_tuple in ws_17.rows], 4, "row count")
                Expect_hasLength([list(inner_tuple) for inner_tuple in ws_17.columns], 3, "column count")
                actual_37: Any = ws_17["C4"].value
                expected_39: Any = 1204
                if equals(TranspilerHelper_op_BangBang(actual_37), TranspilerHelper_op_BangBang(expected_39)):
                    Assert_AreEqual(actual_37, expected_39, "value C4")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_39)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_37)) + "") + ('[0m')) + " \n\b    Message: ") + "value C4") + "")


            def body_24(__unit: None=None) -> None:
                wb_22: Workbook = Workbook_1(None)
                ws_18: Worksheet = wb_22.active
                t1: Table = Table_1(displayName="Table1", ref="A1:B2")
                t2: Table = Table_1(displayName="Table2", ref="C1:D2")
                ws_18.add_table(t1)
                ws_18.add_table(t2)
                del ws_18.tables["Table2"]
                tables: Array[Table] = list(ws_18.tables.values())
                Expect_hasLength(tables, 1, "lenght")
                actual_38: Table = tables[0]
                expected_40: Table = t1
                if equals(TranspilerHelper_op_BangBang(actual_38), TranspilerHelper_op_BangBang(expected_40)):
                    Assert_AreEqual(actual_38, expected_40, "1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_40)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_38)) + "") + ('[0m')) + " \n\b    Message: ") + "1") + "")


            def body_25(__unit: None=None) -> None:
                wb_23: Workbook = Workbook_1(None)
                ws_19: Worksheet = wb_23.active
                t1_1: Table = Table_1(displayName="Table1", ref="A1:B2")
                t2_1: Table = Table_1(displayName="Table2", ref="C1:D2")
                ws_19.add_table(t1_1)
                ws_19.add_table(t2_1)
                actual_40: int = (len(ws_19.tables)) or 0
                if TranspilerHelper_op_BangBang(actual_40) == TranspilerHelper_op_BangBang(2):
                    Assert_AreEqual(actual_40, 2, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(2)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_40)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_26(__unit: None=None) -> None:
                wb_24: Workbook = Workbook_1(None)
                ws_20: Worksheet = wb_24.active
                ws_20["A1"] = 42
                cell_3: Cell = ws_20.cell(1, 1)
                actual_41: Any = cell_3.value
                expected_42: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_41), TranspilerHelper_op_BangBang(expected_42)):
                    Assert_AreEqual(actual_41, expected_42, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_42)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_41)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_27(__unit: None=None) -> None:
                wb_25: Workbook = Workbook_1(None)
                ws_21: Worksheet = wb_25.active
                cell_4: Cell = ws_21.cell(1, 1, some(42))
                actual_42: Any = cell_4.value
                expected_43: Any = 42
                if equals(TranspilerHelper_op_BangBang(actual_42), TranspilerHelper_op_BangBang(expected_43)):
                    Assert_AreEqual(actual_42, expected_43, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_43)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_42)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_28(__unit: None=None) -> None:
                wb_26: Workbook = Workbook_1(None)
                ws_22: Worksheet = wb_26.active
                cell_5: Cell = ws_22.cell(1, 1, some(42))
                actual_44: CellType = CellType_fromCellType_Z721C83C5(type(cell_5.value).__name__)
                expected_45: CellType = CellType(1)
                if equals(TranspilerHelper_op_BangBang(actual_44), TranspilerHelper_op_BangBang(expected_45)):
                    Assert_AreEqual(actual_44, expected_45, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_45)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_44)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_29(__unit: None=None) -> None:
                wb_27: Workbook = Workbook_1(None)
                ws_23: Worksheet = wb_27.active
                cell_6: Cell = ws_23.cell(1, 1, some(42.0))
                actual_46: CellType = CellType_fromCellType_Z721C83C5(type(cell_6.value).__name__)
                expected_47: CellType = CellType(0)
                if equals(TranspilerHelper_op_BangBang(actual_46), TranspilerHelper_op_BangBang(expected_47)):
                    Assert_AreEqual(actual_46, expected_47, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_47)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_46)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_30(__unit: None=None) -> None:
                wb_28: Workbook = Workbook_1(None)
                ws_24: Worksheet = wb_28.active
                cell_7: Cell = ws_24.cell(1, 1, some("Hello World"))
                actual_48: CellType = CellType_fromCellType_Z721C83C5(type(cell_7.value).__name__)
                expected_49: CellType = CellType(2)
                if equals(TranspilerHelper_op_BangBang(actual_48), TranspilerHelper_op_BangBang(expected_49)):
                    Assert_AreEqual(actual_48, expected_49, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_49)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_48)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_31(__unit: None=None) -> None:
                wb_29: Workbook = Workbook_1(None)
                ws_25: Worksheet = wb_29.active
                cell_8: Cell = ws_25.cell(1, 1, some(True))
                actual_50: CellType = CellType_fromCellType_Z721C83C5(type(cell_8.value).__name__)
                expected_51: CellType = CellType(3)
                if equals(TranspilerHelper_op_BangBang(actual_50), TranspilerHelper_op_BangBang(expected_51)):
                    Assert_AreEqual(actual_50, expected_51, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_51)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_50)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_32(__unit: None=None) -> None:
                wb_30: Workbook = Workbook_1(None)
                ws_26: Worksheet = wb_30.active
                cell_9: Cell = ws_26.cell(1, 1, some(False))
                actual_52: CellType = CellType_fromCellType_Z721C83C5(type(cell_9.value).__name__)
                expected_53: CellType = CellType(3)
                if equals(TranspilerHelper_op_BangBang(actual_52), TranspilerHelper_op_BangBang(expected_53)):
                    Assert_AreEqual(actual_52, expected_53, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_53)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_52)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_33(__unit: None=None) -> None:
                wb_31: Workbook = Workbook_1(None)
                ws_27: Worksheet = wb_31.active
                cell_10: Cell = ws_27.cell(1, 1, some(now()))
                t_3: str = type(cell_10.value).__name__
                to_console(printf("%A"))(t_3)
                actual_54: CellType = CellType_fromCellType_Z721C83C5(t_3)
                expected_55: CellType = CellType(4)
                if equals(TranspilerHelper_op_BangBang(actual_54), TranspilerHelper_op_BangBang(expected_55)):
                    Assert_AreEqual(actual_54, expected_55, "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_55)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_54)) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_34(__unit: None=None) -> None:
                table: Table = Table_1(displayName="NewTable", ref="A1:B2")
                actual_55: str = table.displayName
                if TranspilerHelper_op_BangBang(actual_55) == TranspilerHelper_op_BangBang("NewTable"):
                    Assert_AreEqual(actual_55, "NewTable", "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "NewTable") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_55) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_35(__unit: None=None) -> None:
                wb_32: Workbook = Workbook_1(None)
                ws_28: Worksheet = wb_32.active
                table_1: Table = Table_1(displayName="NewTable", ref="A1:B2")
                ws_28.add_table(table_1)
                table_get: Table = ws_28.tables["NewTable"]
                actual_56: str = table_get.displayName
                if TranspilerHelper_op_BangBang(actual_56) == TranspilerHelper_op_BangBang("NewTable"):
                    Assert_AreEqual(actual_56, "NewTable", "")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + "NewTable") + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + actual_56) + "") + ('[0m')) + " \n\b    Message: ") + "") + "")


            def body_36(__unit: None=None) -> None:
                wb_33: Workbook = Workbook_1(None)
                ws_29: Worksheet = wb_33.active
                t1_2: Table = Table_1(displayName="Table1", ref="A1:B2")
                t2_2: Table = Table_1(displayName="Table2", ref="C1:D2")
                ws_29.add_table(t1_2)
                ws_29.add_table(t2_2)
                actual_57: Table = ws_29.tables["Table1"]
                expected_58: Table = t1_2
                if equals(TranspilerHelper_op_BangBang(actual_57), TranspilerHelper_op_BangBang(expected_58)):
                    Assert_AreEqual(actual_57, expected_58, "Equal table 1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_58)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_57)) + "") + ('[0m')) + " \n\b    Message: ") + "Equal table 1") + "")

                actual_58: Table = ws_29.tables["Table2"]
                expected_59: Table = t2_2
                if equals(TranspilerHelper_op_BangBang(actual_58), TranspilerHelper_op_BangBang(expected_59)):
                    Assert_AreEqual(actual_58, expected_59, "Equal table 2")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_59)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_58)) + "") + ('[0m')) + " \n\b    Message: ") + "Equal table 2") + "")


            def body_37(__unit: None=None) -> None:
                wb_34: Workbook = Workbook_1(None)
                ws_30: Worksheet = wb_34.active
                t1_3: Table = Table_1(displayName="Table1", ref="A1:B2")
                t2_3: Table = Table_1(displayName="Table2", ref="C1:D2")
                ws_30.add_table(t1_3)
                ws_30.add_table(t2_3)
                tables_1: Array[Table] = list(ws_30.tables.values())
                actual_59: Table = tables_1[0]
                expected_60: Table = t1_3
                if equals(TranspilerHelper_op_BangBang(actual_59), TranspilerHelper_op_BangBang(expected_60)):
                    Assert_AreEqual(actual_59, expected_60, "equal t1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_60)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_59)) + "") + ('[0m')) + " \n\b    Message: ") + "equal t1") + "")

                actual_60: Table = tables_1[1]
                expected_61: Table = t2_3
                if equals(TranspilerHelper_op_BangBang(actual_60), TranspilerHelper_op_BangBang(expected_61)):
                    Assert_AreEqual(actual_60, expected_61, "equal t2")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_61)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_60)) + "") + ('[0m')) + " \n\b    Message: ") + "equal t2") + "")


            def body_38(__unit: None=None) -> None:
                wb_35: Workbook = Workbook_1(None)
                ws_31: Worksheet = wb_35.active
                t1_4: Table = Table_1(displayName="Table1", ref="A1:B2")
                t2_4: Table = Table_1(displayName="Table2", ref="C1:D2")
                ws_31.add_table(t1_4)
                ws_31.add_table(t2_4)
                tables_2: Array[tuple[str, str]] = ws_31.tables.items()
                actual_61: tuple[str, str] = tables_2[0]
                expected_62: tuple[str, str] = ("Table1", "A1:B2")
                if equal_arrays(TranspilerHelper_op_BangBang(actual_61), TranspilerHelper_op_BangBang(expected_62)):
                    Assert_AreEqual(actual_61, expected_62, "1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_62)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_61)) + "") + ('[0m')) + " \n\b    Message: ") + "1") + "")

                actual_62: tuple[str, str] = tables_2[1]
                expected_63: tuple[str, str] = ("Table2", "C1:D2")
                if equal_arrays(TranspilerHelper_op_BangBang(actual_62), TranspilerHelper_op_BangBang(expected_63)):
                    Assert_AreEqual(actual_62, expected_63, "2")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_63)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_62)) + "") + ('[0m')) + " \n\b    Message: ") + "2") + "")


            def body_39(__unit: None=None) -> None:
                wb_36: Workbook = Workbook_1(None)
                ws_32: Worksheet = wb_36.active
                t1_5: Table = Table_1(displayName="Table1", ref="A1:B2")
                t2_5: Table = Table_1(displayName="Table2", ref="C1:D2")
                ws_32.add_table(t1_5)
                ws_32.add_table(t2_5)
                del ws_32.tables["Table2"]
                tables_3: Array[Table] = list(ws_32.tables.values())
                Expect_hasLength(tables_3, 1, "lenght")
                actual_63: Table = tables_3[0]
                expected_64: Table = t1_5
                if equals(TranspilerHelper_op_BangBang(actual_63), TranspilerHelper_op_BangBang(expected_64)):
                    Assert_AreEqual(actual_63, expected_64, "1")

                else: 
                    Helper_expectError(((((((((((((("    Expected: " + ('[36m')) + "") + str(expected_64)) + "") + ('[0m')) + " \n\b    Actual: ") + ('[31m')) + "") + str(actual_63)) + "") + ('[0m')) + " \n\b    Message: ") + "1") + "")


            return singleton(Test_testList("Worksheet", of_array([Test_testCase("init title", body_9), Test_testCase("set title", body_10), Test_testList("Item", of_array([Test_testCase("get cell", body_11), Test_testCase("Create Cell", body_12), Test_testCase("get cell range", body_13), Test_testCase("get column", body_14), Test_testCase("get row", body_15), Test_testCase("get columns", body_16), Test_testCase("get rows", body_17), Test_testCase("iter_rows", body_18), Test_testCase("iter_cols", body_19), Test_testCase("rows", body_20), Test_testCase("columns", body_21), Test_testCase("values", body_22), Test_testCase("append", body_23), Test_testCase("delete", body_24), Test_testCase("tableCount", body_25)])), Test_testList("Cell", of_array([Test_testCase("get via _.cell", body_26), Test_testCase("set via _.cell", body_27), Test_testList("cellType", of_array([Test_testCase("int", body_28), Test_testCase("float", body_29), Test_testCase("string", body_30), Test_testCase("bool-true", body_31), Test_testCase("bool-false", body_32), Test_testCase("datetime", body_33)]))])), Test_testList("Table", of_array([Test_testCase("create", body_34), Test_testCase("add to ws", body_35)])), Test_testList("ws.tables", of_array([Test_testCase("Item get", body_36), Test_testCase("values()", body_37), Test_testCase("items()", body_38), Test_testCase("delete", body_39)]))])))

        return append(singleton(Test_testList("Workbook", of_array([Test_testCase("Create new", body_2), Test_testCase("create new worksheet", body_3), Test_testCase("create new worksheet at position", body_4), Test_testCase("sheetnames", body_5), Test_testCase("Item", body_6), Test_testCase("for sheet in wb", body_7), Test_testCase("copy_worksheet", body_8)]))), delay(_arrow20))

    return append(singleton(Test_testList("openpyxl", of_array([Test_testCase("Minimal Read", body), Test_testCase("init wb", body_1)]))), delay(_arrow21))


tests: Model_TestCase = Test_testList("main", to_list(delay(_arrow22)))

Pyxpecto_runTests([], tests)

