module Tests.Tables

open Fable.Pyxpecto
open Fable.Openpyxl

let main = testList "Tables" [
    testCase "Item get" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb.active
        let t1 = Table.create("Table1", "A1:B2")
        let t2 = Table.create("Table2", "C1:D2")
        ws.add_table(t1)
        ws.add_table(t2)
        Expect.equal ws.tables.["Table1"] t1 "Equal table 1"
        Expect.equal ws.tables.["Table2"] t2 "Equal table 2"
    testCase "values()" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb.active
        let t1 = Table.create("Table1", "A1:B2")
        let t2 = Table.create("Table2", "C1:D2")
        ws.add_table(t1)
        ws.add_table(t2)
        let tables = ws.tables.values()
        Expect.equal tables[0] t1 "equal t1"
        Expect.equal tables[1] t2 "equal t2"
    testCase "items()" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb.active
        let t1 = Table.create("Table1", "A1:B2")
        let t2 = Table.create("Table2", "C1:D2")
        ws.add_table(t1)
        ws.add_table(t2)
        let tables = ws.tables.items()
        Expect.equal tables[0] ("Table1", "A1:B2") "1"
        Expect.equal tables[1] ("Table2", "C1:D2") "2"
    testCase "delete" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb.active
        let t1 = Table.create("Table1", "A1:B2")
        let t2 = Table.create("Table2", "C1:D2")
        ws.add_table(t1)
        ws.add_table(t2)
        ws.tables.delete("Table2")
        let tables = ws.tables.values()
        Expect.hasLength tables 1 "lenght"
        Expect.equal tables[0] t1 "1"
]

