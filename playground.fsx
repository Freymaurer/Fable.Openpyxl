#r "nuget: Fable.Core, 4.2.0"

open Fable.Core
open Fable.Core.PyInterop

let openpyxl : obj = importAll "openpyxl"

printfn "%A" openpyxl