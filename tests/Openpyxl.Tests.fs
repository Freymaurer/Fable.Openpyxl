module Tests.Openpyxl

open Fable.Pyxpecto
open Fable.Openpyxl

let main = testList "openpyxl" [
    testCase "init wb" <| fun _ ->
        let wb = openpyxl.Workbook()
        let ws = wb.active
        Expect.equal ws.title "Sheet" "" 
]