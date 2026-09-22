# Resizable history-state evidence

The report saves a resizable buffer and two views over it, reloads the page, and
returns through cross-document back navigation. Each restored graph retains the
buffer maximum, the shared buffer identity, and the saved byte. Growing it changes
the tracking view while the fixed-length view keeps its original length. The six
PNGs show these results at 360 and 720 pixels; all were visually inspected.

`RenderProof.cs.txt` is executable console source. Reference OfficeIMO.Html,
OfficeIMO.Html.Runtime and OfficeIMO.Html.AngleSharp, then pass the built worker
DLL and an output directory. `validation.json` records exact source revisions,
the Jint upstream contribution, selected tests, artifact hashes and limitations.

Worker regressions also cover DataViews, zero-length views, shrink/regrow behavior,
invalid views and detached/shared buffers, clone budgets, failure atomicity and
frame-message payloads. A review regression proves that authored numeric array
accessors cannot intercept copied bytes.

Intermittent startup deadlines occurred during validation. Isolated final .NET 8
and .NET 10 checks and the rendered workflow passed without changing test deadlines;
the manifest retains the earlier failures and diagnostic timings. Their cause was
not established by this run.
