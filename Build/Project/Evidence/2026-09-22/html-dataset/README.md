# Dataset runtime qualification

This fixture uses `element.dataset` to update a report, delete an obsolete field,
and enumerate its current fields before freezing and rendering the document.
The retained PNGs show the resulting capture at 360px and 720px widths on macOS
arm64. `validation.json` records the provider revisions and test results.

The binding follows the [HTML DOMStringMap contract](https://html.spec.whatwg.org/multipage/dom.html#domstringmap)
and [Web IDL named-property algorithms](https://webidl.spec.whatwg.org/#legacy-platform-objects).
`RuntimeDatasetTests` exercises the worker process, both provider forks, live
attribute changes, property descriptors, deletion, value conversion, catchable
errors, and the independent frozen capture. Standalone `DatasetTests` in the
AngleSharp.Js fork cover the binding with the official AngleSharp dependency.

## Reproduce

Initialize the pinned provider sources, then run:

```sh
git submodule update --init --recursive
dotnet test OfficeIMO.Html.Runtime.Tests/OfficeIMO.Html.Runtime.Tests.csproj -f net10.0 --filter RuntimeDatasetTests
```

For the retained render, create a temporary `net10.0` console project with project
references to `OfficeIMO.Html.Runtime` and `OfficeIMO.Html`. Copy
`RenderProof.cs.txt` to its `Program.cs`. Build the runtime worker and run the
console project with two arguments: the built worker DLL path and an empty output
directory. It validates the capture and writes `dataset.html` and both PNGs.

The original unrestricted-concurrency full-suite attempt was interrupted after
unrelated runtime deadline failures during heavy concurrent machine activity.
The completed run uses xUnit `MaxParallelThreads=2`; the exact command and outcome
are recorded in the manifest. No application deadlines were changed.

## Separate upstream candidate

[AngleSharp.Js DOM identity candidate](https://github.com/EvotecIT/AngleSharp.Js/tree/feature/dom-node-identity)
contains only canonical node/window reuse in `SameObject` getters and one standalone
mutation-record regression. Against unmodified upstream `devel`,
`observedRecord.target.id` returned `undefined`; after the fix, the target has the
concrete element API, strict identity with `document.body`, and retained script
properties. It requires no OfficeIMO code or fork-only core hooks.

The dataset change is separate: complete camel-case and invalid-name behavior
uses the AngleSharp core fork. Its binding patch is not presented as a standalone
replacement for the core fixes. No upstream pull request or package was published.

## Limits

This run qualifies macOS arm64, the selected framework targets and these contracts.
It does not establish full Web IDL, DOMException-constructor, browser, or framework
compatibility. Windows and Linux were not exercised in this run.
