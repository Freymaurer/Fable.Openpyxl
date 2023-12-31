﻿module TestTasks

open BlackFox.Fake
open Fake.DotNet

open ProjectInfo
open BasicTasks
open Helpers

let runTests = BuildTask.create "RunTests" [clean; build] {
    let py_folder_name = "py"
    let test_file_folder = "TestFiles"
    for test_proj in testProjects do
        run dotnet $"fable {test_proj} -o {test_proj}/{py_folder_name} --lang py --noCache" ""
        run python $"{test_proj}/{py_folder_name}/main.py" ""
}

