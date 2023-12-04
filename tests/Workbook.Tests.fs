module Tests.Workbook

open Fable.Pyxpecto
open Fable.Openpyxl

let main = testList "Workbook" [
    testCase "Create new" <| fun _ ->
        let wb = Workbook.create()
        Expect.pass ()
    testCase "create new worksheet" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb.create_sheet("New Sheet")
        Expect.pass ()
    testCase "create new worksheet at position" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb.create_sheet("New Sheet", 2)
        Expect.pass ()
    testCase "sheetnames" <| fun _ ->
        let wb = Workbook.create()
        let _ = wb.create_sheet("New Sheet1")
        let _ = wb.create_sheet("New Sheet2")
        let _ = wb.create_sheet("New Sheet3")
        Expect.hasLength wb.sheetnames 4 "hasLength"
        Expect.equal wb.sheetnames [|"Sheet"; "New Sheet1"; "New Sheet2"; "New Sheet3"|] "equal"
    testCase "Item" <| fun _ ->
        let wb = Workbook.create()
        let ws = wb["Sheet"]
        Expect.equal ws.title "Sheet" ""
    testCase "for sheet in wb" <| fun _ ->
        let wb = Workbook.create()
        for sheet in wb do
        Expect.equal sheet.title "Sheet" "only 1 sheet with the default init title"
    testCase "copy_worksheet" <| fun _ ->
        let wb = Workbook.create()
        let source = wb.active
        let copy = wb.copy_worksheet(source)
        Expect.equal source.title "Sheet" "source title"
        Expect.equal copy.title "Sheet Copy" "copy title"
        Expect.hasLength wb.sheetnames 2 "hasLenght"
]
