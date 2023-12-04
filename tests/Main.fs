module Fable.openpyxl.Tests
open Fable.Pyxpecto

let test_main = testList "main" [
    Tests.Cell.main
    Tests.Table.main
    Tests.Tables.main
    Tests.Worksheet.main
    Tests.Workbook.main
    Tests.IO.main
    Tests.Openpyxl.main
] 

[<EntryPoint>]
let main argv =
    Pyxpecto.runTests [||] test_main
