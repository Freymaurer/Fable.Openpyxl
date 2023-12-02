#r "nuget: Fable.Core, 4.2.0"
#r "nuget: Fable.Pyxpecto, 1.0.0-beta.2"

open Fable.Core
open Fable.Core.PyInterop
open System.Collections.Generic

/// For example "A1"
type XlsxCellAdress = string
type Letters = string

[<Erase; RequireQualifiedAccessAttribute>]
type Row = i of int
[<Erase; RequireQualifiedAccessAttribute>]
type Column = i of string

type Cell =
  abstract member value: obj

type Worksheet =
  abstract member cell: ?row:int * ?column:int -> Cell
  abstract member title : string with get, set 
  [<Emit("$0[$1]")>]
  abstract member Item: XlsxCellAdress -> Cell with get
  [<Emit("$0[$1] = $2")>]
  abstract member Item: XlsxCellAdress -> obj with set
  /// Returns columns of rows of cells
  [<Emit("[list(inner_tuple) for inner_tuple in $0[$1[0] : $1[1]]]")>]
  abstract member Item: (XlsxCellAdress * XlsxCellAdress) -> Cell [] [] with get
  /// Returns row cells
  [<Emit("list($0[$1])")>]
  abstract member Item: Row -> Cell [] with get
  /// Returns column cells
  [<Emit("list($0[$1])")>]
  abstract member Item: Column -> Cell [] with get
  /// Returns array of columns of cells
  [<Emit("[list(inner_tuple) for inner_tuple in $0[('{start}:{end}'.format(start=$1[0],end=$1[1]))]]")>]
  abstract member Item: (Column * Column) -> Cell [] [] with get
  /// Returns array of rows of cells
  [<Emit("[list(inner_tuple) for inner_tuple in $0[$1[0] : $1[1]]]")>]
  abstract member Item: (Row * Row) -> Cell [] [] with get
  /// Can be used to get or set a cell. If `value` is provided the value will be set.
  abstract member cell: row:int * column:int * ?value:obj -> Cell
  abstract member iter_rows: min_row:int * max_col:int * max_row: int
  abstract member iter_cols: min_row:int * max_col:int * max_row: int

type Workbook =
  inherit IEnumerable<Worksheet>
  abstract member active: Worksheet
  /// ``ws1 = wb.create_sheet("Mysheet") # insert at the end (default)``
  /// 
  /// or
  /// 
  /// ``ws2 = wb.create_sheet("Mysheet", 0) # insert at first position``
  /// 
  /// or
  /// 
  /// ``ws3 = wb.create_sheet("Mysheet", -1) # insert at the penultimate position``
  abstract member create_sheet: string * ?position:int -> Worksheet
  abstract member sheetnames: string [] with get
  [<Emit("$0[$1]")>]
  abstract member Item: string -> Worksheet
  /// Create copies of worksheets within a single workbook. Adds " Copy" to any existing title.
  /// 
  /// Does not copy all elements, such as Images and Charts. Cannot copy between workbooks.
  abstract member copy_worksheet: Worksheet -> Worksheet


type WorkbookStatic =
  [<Emit("new $0($1)")>]
  abstract member create: unit -> Workbook

[<Import("Workbook", "openpyxl")>]
let Workbook : WorkbookStatic = nativeOnly


type OpenPyXL =
  abstract member load_workbook: string -> Workbook 

let openpyxl: OpenPyXL = importAll "openpyxl"

open Fable.Pyxpecto

let tests = testList "main" [
  let testFilePath_Simple = "C:\\Users\\Kevin\\Desktop\\BookTest.xlsx"
  testCase "Minimal Read" <| fun _ ->
    let wb_obj = openpyxl.load_workbook(testFilePath_Simple)
    let sheet_obj = wb_obj.active
    let cell_obj = sheet_obj.cell(row = 1, column = 1)
    Expect.equal cell_obj.value (box "A1") "" 
  testList "Workbook" [
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
  testList "Worksheet" [
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
    testList "Item" [
      testCase "get cell" <| fun _ ->
        let wb_obj = openpyxl.load_workbook(testFilePath_Simple)
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
        let wb = openpyxl.load_workbook(testFilePath_Simple)
        let ws = wb.active
        let cols = ws[("A1","C1")]
        Expect.hasLength cols 1 "column lenght"
        Expect.hasLength cols[0] 3 "rows lenght"
        Expect.equal cols.[0].[0].value "A1" "A1-value"
      testCase "get column" <| fun _ ->
        let wb = openpyxl.load_workbook(testFilePath_Simple)
        let ws = wb.active
        let cols = ws[Column.i "A"]
        Expect.hasLength cols 3 "column cell lenght"
        Expect.equal cols[0].value "A1" "A1 - value"
        Expect.equal cols[1].value "A2" "A2 - value"
        Expect.equal cols[2].value "A3" "A3 - value"
      testCase "get row" <| fun _ ->
        let wb = openpyxl.load_workbook(testFilePath_Simple)
        let ws = wb.active
        let cols = ws[Row.i 1]
        Expect.hasLength cols 3 "column cell lenght"
        Expect.equal cols[0].value "A1" "A1 - value"
        Expect.equal cols[1].value "B1" "B1 - value"
        Expect.equal cols[2].value "C1" "C1 - value"
      testCase "get columns" <| fun _ ->
        let wb = openpyxl.load_workbook(testFilePath_Simple)
        let ws = wb.active
        let cols = ws[(Column.i "A", Column.i "B")]
        Expect.hasLength cols 2 "hasLength"
        for col in cols do
          Expect.hasLength col 3 "inner has Length"
        Expect.equal cols[0].[0].value "A1" "A1 - value"
        Expect.equal cols[1].[2].value "B3" "B3 - value"
      testCase "get rows" <| fun _ ->
        let wb = openpyxl.load_workbook(testFilePath_Simple)
        let ws = wb.active
        let rows = ws[(Row.i 1, Row.i 2)]
        Expect.hasLength rows 2 "hasLength"
        for row in rows do
          Expect.hasLength row 3 "inner has Length"
        Expect.equal rows[0].[0].value "A1" "A1 - value"
        Expect.equal rows[1].[2].value "C2" "C2 - value"
    ]
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
  ]
] 

Pyxpecto.runTests [||] tests