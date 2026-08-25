# OfficeIMO Visio benchmarks

This opt-in project measures complete in-memory VSDX creation/save and
load/structural-inspection workflows over deterministic multi-page shape and
connector graphs. Validation reopens every generated package and checks page,
shape, connector, Shape Data, boundary text, and output-size contracts before
timing.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- --filter '*VisioBenchmarks*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Visio.Benchmarks -- evidence --repeat 3 --json .\.benchmark-artifacts\visio\evidence.json
```

The isolated evidence runner records elapsed time and managed allocation per
operation, retained managed heap, sampled managed-heap growth, absolute process
peak, package bytes, source commit, dirty-tree state, runtime, and operating
system. Each child process validates its input and result before reporting.

The project stays outside `OfficeIMO.sln`. A commercial comparison belongs in
a separate opt-in project and may be reported only with a valid license and an
equivalent package contract; evaluation-mode shape limits are not valid
comparison evidence.
