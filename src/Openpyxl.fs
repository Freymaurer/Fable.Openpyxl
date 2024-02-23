module Fable.Openpyxl

open Fable.Core
open Fable.Core.PyInterop
open System.Collections.Generic

module Helper =
  let writeBytes (bytes:byte [], path: string) : unit = emitPyExpr (bytes, path) """
    # Write the bytes data to the output file path using shutil
    with open($1, 'wb') as output_file:
        output_file.write($0)
  """

/// For example "A1"
type XlsxCellAdress = string
/// For example "A1:E5"
type XlsxCellRangeAdress = string
type Letters = string
[<Erase; RequireQualifiedAccessAttribute>]
type Row = i of int
[<Erase; RequireQualifiedAccessAttribute>]
type Column = i of string
type CellValue = obj

module CellType =
  let [<Literal>] Literal_Float = "float"
  let [<Literal>] Literal_Integer = "int"
  let [<Literal>] Literal_String = "str"
  let [<Literal>] Literal_Boolean = "bool"
  let [<Literal>] Literal_DateTime = "datetime"
  let [<Literal>] Literal_Empty = "NoneType"

[<RequireQualifiedAccess>]
type CellType =
| Float 
| Integer
| String
| Boolean
| DateTime
| Empty
with
  static member fromCellType (cellType:string) =
    match cellType with
    | CellType.Literal_Integer -> Integer
    | CellType.Literal_Float -> Float
    | CellType.Literal_String -> String
    | CellType.Literal_Boolean -> Boolean
    | CellType.Literal_DateTime -> DateTime
    | CellType.Literal_Empty -> Empty
    | anyElse -> failwith $"Unknown cell type of type: '{anyElse}'"

type Cell =
  abstract member value: CellValue with get, set
  [<Emit("type($0.value).__name__")>]
  abstract member cellType: string with get

type Table =
  [<Emit("$0.displayName")>]
  abstract member displayName: string with get, set
  abstract member name: string with get, set
  abstract member id: int with get, set
  [<Emit("$0.headerRowCount")>]
  abstract member headerRowCount: int with get, set
  abstract member ref: string with get, set

type TableMap =
  [<Emit("$0[$1]")>]
  abstract member Item: string -> Table
  [<Emit("list($0.values())")>]
  abstract member values: unit -> Table []
  /// Returns array of tuples: (displayName * ref)
  abstract member items: unit -> (string * string) []
  [<Emit("del $0[$1]")>]
  abstract member delete: displayName: string -> unit

type Worksheet =
  abstract member cell: ?row:int * ?column:int -> Cell
  abstract member title : string with get, set 
  [<Emit("$0[$1]")>]
  abstract member Item: XlsxCellAdress -> Cell with get
  [<Emit("$0[$1] = $2")>]
  abstract member Item: XlsxCellAdress -> CellValue with set
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
  abstract member cell: row:int * column:int * ?value:CellValue -> Cell
  /// Due to f# -> fable python transpilation. As of fable 4.6.1 this cannot be bound to a variable!
  [<Emit("""for row in $0.iter_rows(min_row=$1, max_col=$2, max_row=$3): $4(row)""")>]
  abstract member iter_rows: min_row:int * max_col:int * max_row:int * action:(Cell [] -> unit) -> unit
  /// Due to f# -> fable python transpilation. As of fable 4.6.1 this cannot be bound to a variable!
  [<Emit("""for col in $0.iter_cols(min_row=$1, max_col=$2, max_row=$3): $4(col)""")>]
  abstract member iter_cols: min_row:int * max_col:int * max_row:int * action:(Cell [] -> unit) -> unit
  [<Emit("[list(inner_tuple) for inner_tuple in $0.rows]")>]
  abstract member rows: Cell [] []
  [<Emit("[list(inner_tuple) for inner_tuple in $0.columns]")>]
  abstract member columns: Cell [] []
  /// iterates over all rows but returns just the value.
  [<Emit("[list(inner_tuple) for inner_tuple in $0.values]")>]
  abstract member values: CellValue [] []
  /// Used to append rows
  abstract member append: CellValue [] -> unit
  /// The default is one row. For example to insert a row at 7 (before the existing row 7):
  abstract member insert_rows: int -> unit
  /// The default is one column. For example to insert a row at 7 (before the existing row 7):
  abstract member insert_cols: int -> unit
  abstract member delete_rows: start_index:int * count:int -> unit
  abstract member delete_cols: start_index:int * count:int -> unit
  abstract member add_table: Table -> unit
  abstract member tables: TableMap with get
  [<Emit("del $0.tables[$1]")>]
  abstract member delete_table: displayName: string -> unit
  [<Emit("len($0.tables)")>]
  abstract member tableCount: int

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
  /// This operation will overwrite existing files without warning.
  abstract member save: path:string -> unit
  abstract member save: bytesio_obj:BytesIO -> unit
  abstract member template: bool with get, set
  abstract member iso_dates: bool with get, set

and BytesIO =
  abstract member x: string
  [<Emit("""with open($1, "wb") as f: f.write($0.getbuffer())""")>]
  abstract member ToFile: path:string -> unit
  abstract member getbuffer: unit -> obj
  abstract member getvalue: unit -> byte []

type OpenPyXL =
  abstract member load_workbook: path:string -> Workbook
  abstract member load_workbook: bytes:byte [] -> Workbook
  abstract member load_workbook: buffer: obj -> Workbook
  abstract member Workbook: unit -> Workbook

// - - Static create helper - - //

type WorkbookStatic =
  [<Emit("new $0($1)")>]
  abstract member create: unit -> Workbook

type TableStatic =
  [<Emit("$0(displayName=$1, ref=$2)")>]
  abstract member create: displayName:string * ref:XlsxCellRangeAdress -> Table

type BytesIOStatic =
  [<Emit("$0($1)")>]
  abstract member create: obj -> BytesIO
  [<Emit("$0($1)")>]
  abstract member create: unit -> BytesIO
  [<Emit("$0($1)")>]
  abstract member create: byte [] -> BytesIO

// - - Access helper - - //

[<Import("Workbook", "openpyxl")>]
let Workbook : WorkbookStatic = nativeOnly
[<Import("Table", "openpyxl.worksheet.table")>]
let Table: TableStatic = nativeOnly
[<Import("BytesIO","io")>]
let BytesIO: BytesIOStatic = nativeOnly
let openpyxl: OpenPyXL = importAll "openpyxl"

type Xlsx =
    /// read from a file
    static member readFile (path:string) : Workbook = openpyxl.load_workbook path
    /// read from bytes
    static member read (bytes:byte []) : Workbook = openpyxl.load_workbook bytes
    /// load from a buffer
    static member load (buffer:obj) : Workbook = openpyxl.load_workbook buffer
    /// write to a file
    static member writeFile(wb: Workbook, path: string) : unit = wb.save(path)
    /// write to a stream
    static member write (wb: Workbook) = 
        let output = BytesIO.create()
        wb.save(output)
        let bytes = output.getvalue()
        bytes
    /// write to a new buffer
    static member writeBuffer (wb: Workbook) = 
        let output = BytesIO.create()
        wb.save(output)
        let buffer = output.getbuffer()
        buffer