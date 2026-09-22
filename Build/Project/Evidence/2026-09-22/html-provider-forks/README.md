# HTML provider fork qualification

This run checks OfficeIMO's migration from embedded AngleSharp source to pinned
EvotecIT forks, together with fixes for shared frame storage, native message-event
trust, and nested selectors. The provider and consumer revisions, environment,
commands, and observed results are recorded in [validation.json](validation.json).

The runtime consumes the fork assemblies in its isolated worker. Static consumers
continue to use their existing official packages. Jint remains unmodified.

| Owner | Patch group | Contribution boundary |
| --- | --- | --- |
| AngleSharp | DOM ownership, dataset mapping, detached-node comparisons, event/lifecycle behavior, mutation notifications, script preparation, and integrity | General library behavior with standalone regression tests |
| AngleSharp | Optional synchronization, synchronous scripting, mutation notification/microtask, and stylesheet blocking services | Reusable extension points requiring API discussion with upstream |
| AngleSharp.Js | Engine configuration, DOM constants, canonical node identity, NodeList prototype methods, collection distinctions, and cancellable document-readiness waits | JavaScript binding behavior and host configuration |
| OfficeIMO | Resource policy, quotas, storage ownership, frame message delivery, modules, navigation, rendering, and nested-selector expansion | Product-owned behavior retained outside the forks |

The forks retain upstream history and their licenses. Their `FORK.md` files describe
maintenance boundaries. This evidence does not establish upstream acceptance,
package publication, full browser compatibility, or dependency independence.

## Reproduction

Initialize the pinned submodules using the [provider setup guide](../../../../../External/README.md).
Run the commands in the manifest from the indicated repositories. This macOS arm64
run used SDK 10.0.303 directly because the repository-pinned SDK was unavailable;
that does not change the checked-in SDK selection.

The public runtime tests cover execution and capture, storage, messaging,
collections, observers, lifecycle, fetch, navigation, and resource policies. The
standalone fork tests exercise the reusable contracts without OfficeIMO's host.
The deployed .NET 8 smoke checks script execution, storage, NodeList identity,
independent capture, and both copied dependency licenses.

## Rendered evidence

The [input](nesting.html) contains a complex parent combined with `.x &` and repeated
`& + &`, plus an explicitly expanded control. The outputs at [360 pixels](nesting-360.png)
and [720 pixels](nesting-720.png) were inspected: the nested/control rows are red,
the first sibling is blue, and the next sibling is red. Both layouts fit their
viewports. These images qualify the changed selector behavior; they are not a new
full-layout or cross-platform conformance claim.

## Review and limits

One independent read-only review found two inherited standalone-fork defects:
public request callers could bypass integrity requirements, and a throwing mutation
observer could starve later observers. Regression tests reproduced both. Targeted
confirmation also identified the empty-integrity sibling path; same-origin,
no-CORS, and CORS-enabled empty-metadata cases now pass.

Intermediate runtime runs exposed intermittent startup timeouts: one run passed
519 tests with one timeout, and another passed 514 with six timeouts in different
cases. Isolated repetitions could pass, so they did not establish correctness.
The binding's availability helper waited for a document `load` event, while the DOM
fork correctly dispatched that event on the window. Calls made after readiness
completed bypassed the faulty wait. Deterministic tests reproduced completion and
cancellation defects in the helper. The fix observes document readiness on its event
loop and propagates cancellation through all availability stages; no deadline was
increased. The final suite result is recorded separately in the manifest.

Windows/Linux execution, package publication, isolated-container qualification,
and broader performance/corpus gates were not repeated by this run.
