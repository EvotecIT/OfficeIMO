# Native history-state graph evidence

The report retains a typed error, its trace, its shared cause and a Map value
through reload and cross-document back navigation. The six PNGs show saved,
reloaded and restored states at 360 and 720 pixels; each was visually inspected.

`RenderProof.cs.txt` is the executable console source. Reference OfficeIMO.Html,
OfficeIMO.Html.Runtime and OfficeIMO.Html.AngleSharp, then pass the built worker
DLL path and an output folder. `source.html` is the authored report. Use the SDK
and source revisions recorded in `validation.json`.

Worker regressions exercise supported native values in the new realm: boxed
primitives, sparse arrays, cycles, aliases, special numbers, Map/Set, dates,
regular expressions, buffers/views and typed error/cause graphs. They also check
error-message conversion, ignored message/cause/stack accessors, preserved native
or explicit string traces, failure atomicity and trace budget accounting. Frame
messaging uses the same error payload owner.

The acceptance rules derive from [HTML structured serialization](https://html.spec.whatwg.org/multipage/structured-data.html#structuredserializeinternal).
The tests are original focused regressions, not a full WPT conformance run.
Resizable buffers/views and platform-object state remain unqualified. Jint 4.16.0
keeps a view's length-tracking metadata internal; this work does not infer it by
mutating source buffers or expose a partial resizable-view contract. Other
operating systems and general structuredClone/transfer APIs remain outside this evidence.
