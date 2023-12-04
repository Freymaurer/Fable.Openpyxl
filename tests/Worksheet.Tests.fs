module Tests.Worksheet

open Fable.Pyxpecto
open Fable.Openpyxl

let private tests_general = testList "general" [
    testCase "init title" <| fun _ ->
        let expected = "New Sheet"
        let wb = Workbook.create()
        let ws = wb.create_sheet(expected)
        Expect.equal ws.title expected ""
    testCase "set title" <| fun _ ->
        let expected = "New Title"
        let wb = Workbook.create()
        let ws = wb.create_sheet("New Sheet")
        ws.title <- expected
        Expect.equal ws.title expected ""
    testCase "iter_rows" <| fun _ ->
        let wb = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let ws = wb.active
        ws.iter_rows(min_row=1, max_col=3, max_row=2, action=fun row -> for cell in row do cell.value <- 42)
        Expect.equal ws["A1"].value 42 "A1" 
        Expect.equal ws["A2"].value 42 "A2" 
        Expect.equal ws["B1"].value 42 "B1" 
        Expect.equal ws["B2"].value 42 "B2" 
        Expect.equal ws["C1"].value 42 "C1" 
        Expect.equal ws["C2"].value 42 "C2" 
    testCase "iter_cols" <| fun _ ->
        let wb = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let ws = wb.active
        ws.iter_cols(min_row=1, max_col=3, max_row=2, action=fun col -> for cell in col do cell.value <- 42)
        Expect.equal ws["A1"].value 42 "A1" 
        Expect.equal ws["A2"].value 42 "A2" 
        Expect.equal ws["B1"].value 42 "B1" 
        Expect.equal ws["B2"].value 42 "B2" 
        Expect.equal ws["C1"].value 42 "C1" 
        Expect.equal ws["C2"].value 42 "C2" 
    testCase "rows" <| fun _ ->
        let wb = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let ws = wb.active
        let rows = ws.rows
        Expect.hasLength rows 3 "hasLenght"
        Expect.equal rows.[0].[0].value "A1" "A1"
    testCase "columns" <| fun _ ->
        let wb = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let ws = wb.active
        let rows = ws.columns
        Expect.hasLength rows 3 "hasLenght"
        Expect.equal rows.[0].[0].value "A1" "A1"
    testCase "values" <| fun _ ->
        let wb = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let ws = wb.active
        let rows = ws.values
        Expect.hasLength rows 3 "hasLenght"
        Expect.equal rows.[0].[0] "A1" "A1"
    testCase "append" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb.active
        let treedata = [|[|box "Type"; box "Leaf Color"; box "Height"|]; [|box "Maple"; box "Red"; box 549|]; [|box "Oak"; box "Green"; box 783|]; [|box "Pine"; box "Green";box 1204|]|]
        for row in treedata do
          ws.append(row)
        Expect.hasLength ws.rows 4 "row count" 
        Expect.hasLength ws.columns 3 "column count"
        Expect.equal ws.["C4"].value 1204 "value C4"
    testCase "delete" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb.active
        let t1 = Table.create("Table1", "A1:B2")
        let t2 = Table.create("Table2", "C1:D2")
        ws.add_table(t1)
        ws.add_table(t2)
        ws.delete_table("Table2")
        let tables = ws.tables.values()
        Expect.hasLength tables 1 "lenght"
        Expect.equal tables[0] t1 "1"
    testCase "tableCount" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb.active
        let t1 = Table.create("Table1", "A1:B2")
        let t2 = Table.create("Table2", "C1:D2")
        ws.add_table(t1)
        ws.add_table(t2)
        let actual = ws.tableCount
        Expect.equal actual 2 ""
]

let private tests_Item = testList "Item" [
    testCase "get cell" <| fun _ ->
        let wb_obj = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let sheet_obj = wb_obj.active
        let cell_obj = sheet_obj["A1"]
        Expect.equal cell_obj.value (box "A1") "" 
    testCase "Create Cell" <| fun _ ->
        let wb: Workbook = Workbook.create()
        let ws = wb.active
        ws["A1"] <- box 42
        let cell = ws["A1"]
        Expect.equal cell.value (box 42) "" 
    testCase "get cell range" <| fun _ ->
        let wb = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let ws = wb.active
        let cols = ws[("A1","C1")]
        Expect.hasLength cols 1 "column lenght"
        Expect.hasLength cols[0] 3 "rows lenght"
        Expect.equal cols.[0].[0].value "A1" "A1-value"
    testCase "get column" <| fun _ ->
        let wb = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let ws = wb.active
        let cols = ws[Column.i "A"]
        Expect.hasLength cols 3 "column cell lenght"
        Expect.equal cols[0].value "A1" "A1 - value"
        Expect.equal cols[1].value "A2" "A2 - value"
        Expect.equal cols[2].value "A3" "A3 - value"
    testCase "get row" <| fun _ ->
        let wb = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let ws = wb.active
        let cols = ws[Row.i 1]
        Expect.hasLength cols 3 "column cell lenght"
        Expect.equal cols[0].value "A1" "A1 - value"
        Expect.equal cols[1].value "B1" "B1 - value"
        Expect.equal cols[2].value "C1" "C1 - value"
    testCase "get columns" <| fun _ ->
        let wb = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let ws = wb.active
        let cols = ws[(Column.i "A", Column.i "B")]
        Expect.hasLength cols 2 "hasLength"
        for col in cols do
          Expect.hasLength col 3 "inner has Length"
        Expect.equal cols[0].[0].value "A1" "A1 - value"
        Expect.equal cols[1].[2].value "B3" "B3 - value"
    testCase "get rows" <| fun _ ->
        let wb = openpyxl.load_workbook(TestPaths.testFilePath_Simple)
        let ws = wb.active
        let rows = ws[(Row.i 1, Row.i 2)]
        Expect.hasLength rows 2 "hasLength"
        for row in rows do
          Expect.hasLength row 3 "inner has Length"
        Expect.equal rows[0].[0].value "A1" "A1 - value"
        Expect.equal rows[1].[2].value "C2" "C2 - value"
]

let main = testList "Worksheet" [
    tests_general
    tests_Item
]
