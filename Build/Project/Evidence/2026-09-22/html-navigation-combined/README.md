# Combined navigation and document-input evidence

This example keeps a popup alive while its opener reloads, then opens the
popup's document from the opener. It checks that the popup retains its history
object and current state, that `document.open()` gives the popup the opener's
HTTP URL and base, and that a same-origin `pushState()` followed by `back()`
restores the serialized state. The state contains a resizable buffer and two
views. Growth of the in-memory state to 16 bytes does not rewrite the stored
8-byte history snapshot; traversal restores that snapshot and its shared view
aliases.

The same session checks root history state before and after reload, replaces a
loaded `srcdoc` frame with `document.open/write/close`, and preserves a paragraph
across writes that split a character reference. The four PNGs show the checked
page at 360 and 720 pixels before and after root reload. The popup itself is
verified through runtime assertions and the opener's visible status text; it is
not rendered as a separate page.

`RenderProof.cs.txt` is executable console source. Copy it to `Program.cs` in a
small `net10.0` project referencing OfficeIMO.Html, OfficeIMO.Html.Runtime and
OfficeIMO.Html.AngleSharp, then pass the built runtime-worker DLL path and an
output directory. `validation.json` pins the source commits, focused test
selection and artifact hashes. This is an adapted combined scenario, not an
official web-platform-test run. It does not qualify child-defined `window`
functions called across realms, reverse `about:` child-to-parent input streams,
the Navigation API, or general browser compatibility. The existing
[document-input evidence](../html-document-popup/README.md) retains its own
earlier frame mutation-record observation; this run verifies frame replacement
by checking the removed node's connectivity and the new document URL.
