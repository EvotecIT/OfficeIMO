# Document replacement and auxiliary-window evidence

The combined workflow restores shared resizable-buffer views, replaces an iframe
DOM while retaining its document identity, and opens an initially blank popup
with the initiator's base and origin. It verifies the popup message source and
closes the popup before reloading the root. The four PNGs show the checked state
before and after reload at 360 and 720 pixels; all were visually inspected.

This is bounded integration evidence. The frame is rebuilt with DOM APIs after
`document.open()`. It does not prove incremental `document.write()` parsing or
popup survival across root navigation. The enabled `OpenWriteClose` regression
still fails; the validation manifest records that failure separately from the
passing selections. No complete navigation-milestone claim is made.

`RenderProof.cs.txt` is executable console source. Reference OfficeIMO.Html,
OfficeIMO.Html.Runtime and OfficeIMO.Html.AngleSharp, then pass the built worker
DLL and an output directory. `validation.json` records source revisions, selected
checks, review coverage, artifact hashes and remaining qualification limits.
