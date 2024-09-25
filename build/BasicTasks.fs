module BasicTasks

open BlackFox.Fake
open Fake.IO
open Fake.DotNet
open Fake.IO.Globbing.Operators

open ProjectInfo

let clean = BuildTask.create "Clean" [] {
    !! "src/**/bin"
    ++ "src/**/obj"
    ++ "tests/**/bin"
    ++ "tests/**/obj"
    ++ "pkg"
    |> Shell.cleanDirs 
}


let setPrereleaseTag = BuildTask.create "SetPrereleaseTag" [] {
    printfn "Please enter pre-release package suffix"
    let suffix = System.Console.ReadLine()
    prereleaseSuffix <- suffix
    prereleaseTag <- (sprintf "%i.%i.%i-%s" release.SemVer.Major release.SemVer.Minor release.SemVer.Patch suffix)
    isPrerelease <- true
}

let build = BuildTask.create "Build" [clean] {
    solutionFile
    |> DotNet.build (fun x ->
        {x with MSBuildParams = {x.MSBuildParams with DisableInternalBinLog = true}}
    )
}

