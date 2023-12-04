module Tests.Cell

open Fable.Pyxpecto
open Fable.Openpyxl

let main = testList "Cell" [
    testCase "get via _.cell" <| fun _ ->
        let wb: Workbook = Workbook.create()
        let ws = wb.active
        ws["A1"] <- box 42
        let cell = ws.cell(1,1)
        Expect.equal cell.value (box 42) "" 
    testCase "set via _.cell" <| fun _ ->
        let wb: Workbook = Workbook.create()
        let ws = wb.active
        let cell = ws.cell(1,1, 42)
        Expect.equal cell.value (box 42) "" 
    testList "cellType" [
        testCase "int" <| fun _ ->
          let wb: Workbook = Workbook.create()
          let ws = wb.active
          let cell = ws.cell(1,1, 42)
          let actual = cell.cellType |> CellType.fromCellType
          let expected = CellType.Integer
          Expect.equal actual expected ""
        testCase "float" <| fun _ ->
          let wb: Workbook = Workbook.create()
          let ws = wb.active
          let cell = ws.cell(1,1, 42.)
          let actual = cell.cellType |> CellType.fromCellType
          let expected = CellType.Float
          Expect.equal actual expected ""
        testCase "string" <| fun _ ->
          let wb: Workbook = Workbook.create()
          let ws = wb.active
          let cell = ws.cell(1,1, "Hello World")
          let t = cell.cellType 
          let expected = CellType.String
          let actual = CellType.fromCellType t
          Expect.equal actual expected ""
        testCase "bool-true" <| fun _ ->
          let wb: Workbook = Workbook.create()
          let ws = wb.active
          let cell = ws.cell(1,1, true)
          let t = cell.cellType 
          let expected = CellType.Boolean
          let actual = CellType.fromCellType t
          Expect.equal actual expected ""
        testCase "bool-false" <| fun _ ->
          let wb: Workbook = Workbook.create()
          let ws = wb.active
          let cell = ws.cell(1,1, false)
          let t = cell.cellType 
          let expected = CellType.Boolean
          let actual = CellType.fromCellType t
          Expect.equal actual expected ""
        testCase "datetime" <| fun _ ->
          let wb: Workbook = Workbook.create()
          let ws = wb.active
          let cell = ws.cell(1,1, System.DateTime.Now)
          let t = cell.cellType 
          let expected = CellType.DateTime
          let actual = CellType.fromCellType t
          Expect.equal actual expected ""
        testCase "empty" <| fun _ ->
          let wb: Workbook = Workbook.create()
          let ws = wb.active
          let cell = ws.cell(1,1)
          let t = cell.cellType 
          let expected = CellType.Empty
          let actual = CellType.fromCellType t
          Expect.equal actual expected ""
    ]
]