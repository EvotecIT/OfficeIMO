# Beforeunload navigation evidence

The worker retains the current document when a `beforeunload` handler cancels the
event or leaves a nonempty `returnValue`. Clearing the handler allows the same
navigation to commit. The captures show both outcomes at 360 and 720 pixels.
The current public behavior and limits are in the
[runtime README](../../../../../OfficeIMO.Html.Runtime/README.md).

The broad .NET 10 runtime run passed all 537 tests with one xUnit worker. The
combined provider suite passed 333 tests, the core event suite passed 10, and the
focused .NET 8 worker suite passed 15. The provider built for .NET Standard 2.0
with no warnings or errors.

`validation.json` identifies the exact source, standards reference, test outcomes,
review closure, and artifact hashes. These are focused original regression tests,
not a full web-platform-test qualification. Inert parsed documents and foreign
body handlers are checked for window ownership; standalone bindings also exercise
load targeting, synchronous DOM parsing, and worker initialization.

From the repository root, initialize submodules and use the installed .NET SDK:

```sh
dotnet test External/AngleSharp/src/AngleSharp.Core.Tests/AngleSharp.Core.Tests.csproj \
  -f net10.0 --filter FullyQualifiedName~DOMEventsTests
dotnet test External/AngleSharp.Js/src/AngleSharp.Js.Tests/AngleSharp.Js.Tests.csproj \
  -f net10.0 -p:AngleSharpTestProject="$PWD/External/AngleSharp/src/AngleSharp/AngleSharp.Core.csproj"
dotnet test OfficeIMO.Html.Runtime.Tests/OfficeIMO.Html.Runtime.Tests.csproj \
  -f net8.0 --filter FullyQualifiedName~RuntimeBeforeUnloadTests
```

`RenderProof.cs.txt` is the executable capture/render fixture. Use it as a console
program referencing `OfficeIMO.Html`, `OfficeIMO.Html.Runtime`, and
`OfficeIMO.Html.AngleSharp`; pass the built worker DLL and an output directory.
The checked-in HTML inputs contain no external resources.

The host ran SDK 10.0.303 directly because the repository's pinned SDK was not
installed. Earlier broad runs encountered wall-clock deadlines under concurrent
machine load; those attempts are not passing evidence. Runtime deadlines were
not increased for qualification.

For the broad runtime run, use a runsettings file containing:

```xml
<RunSettings><xUnit><MaxParallelThreads>1</MaxParallelThreads></xUnit></RunSettings>
```

Pass it with `--settings <path>` to `dotnet test` on the runtime test project with
`-f net10.0`. The recorded full run and later pin qualification share identical
production source; subsequent provider edits only affect documentation and test
conventions.
