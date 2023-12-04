module Tests.IO

open Fable.Pyxpecto
open Fable.Openpyxl
open TestPaths

let main = testList "io" [
    testCase "to bytesio" <| fun _ ->
        let wb = Workbook.create()
        let output = BytesIO.create()
        wb.save(output)
        output.ToFile("./TestFiles/MinWriteTest.xlsx")
    testCase "to buffer" <| fun _ ->
        let wb = Workbook.create()
        let output = BytesIO.create()
        wb.save(output)
        let buffer = output.getbuffer()
        Expect.pass()
    testCase "to bytes" <| fun _ ->
        let wb = Workbook.create()
        let output = BytesIO.create()
        wb.active["A1"] <- 42
        wb.active["A2"] <- 69
        wb.save(output)
        let bytes = output.getvalue()
        Helper.writeBytes (bytes, "./TestFiles/ByteTest.xlsx")
    testCase "Minimal Read" <| fun _ ->
        let wb_obj = openpyxl.load_workbook(testFilePath_Simple)
        let sheet_obj = wb_obj.active
        let cell_obj = sheet_obj.cell(row = 1, column = 1)
        Expect.equal cell_obj.value (box "A1") "" 
]

