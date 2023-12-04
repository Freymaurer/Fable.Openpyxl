module Tests.Table

open Fable.Pyxpecto
open Fable.Openpyxl

let main = testList "Table" [
    testCase "create" <| fun _ ->
        let table = Table.create("NewTable", "A1:B2")
        Expect.equal table.displayName "NewTable" ""
    testCase "add to ws" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb.active
        let table = Table.create("NewTable", "A1:B2")
        ws.add_table(table)
        let table_get = ws.tables.["NewTable"]
        Expect.equal table_get.displayName "NewTable" ""
]
