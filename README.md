# Fable.Openpyxl

Fable bindings for the python xlsx reader/writer **openpyxl**.

> Based on openpyxl version `3.1.2`.

# Docs

Fable.Openpyxl follows openpyxl syntax as close as possible. Most documentation from openpyxl can be used for Fable.Openpyxl together with built in F# intellisense.

Checkout the [tests](/tests) for implemented functions.

```fsharp
// minimal read
open Fable.Openpyxl

let path_to_file = @"tests/TestFiles/MinimalTest.xlsx"

let wb = openpyxl.load_workbook(path_to_file)
let ws = wb.active
let cell = ws.cell(row = 1, column = 1)
printfn "%A" (cell.value = "A1") // true
```

# Development

1. `dotnet tool restore`
2. `py -m venv .venv`, creates python virtual environment.
3. `.\.venv\Scripts\python.exe -m pip install -r requirements.txt`, install local python dependencies

## Python Dependency Management

- Install new local dependencies with `.\.venv\Scripts\pip.exe install <PACKAGE_NAME>`
- Freeze local dependencies with `.\.venv\Scripts\python.exe -m pip freeze > requirements.txt` .
