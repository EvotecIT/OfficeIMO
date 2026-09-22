# Document input and surviving auxiliary-window evidence

The combined workflow restores shared resizable-buffer views and replaces an
iframe through `document.open/write/close`. Two writes preserve the existing
paragraph object, including text split at a character reference. The popup stays
open across root reload: its counter advances from one to two, and its saved
opener reference accesses and messages the replacement root document.

The four PNGs show the checked state before and after reload at 360 and 720
pixels. Runtime assertions also check the frame URL, document-removal mutation
record, history buffer aliases, view lengths, popup base, message origin and
message source. These are bounded contracts, not a complete browser or
navigation qualification.

`RenderProof.cs.txt` is executable console source. Reference OfficeIMO.Html,
OfficeIMO.Html.Runtime and OfficeIMO.Html.AngleSharp, then pass the built worker
DLL and an output directory. `validation.json` records source revisions, selected
checks, review coverage, artifact hashes and remaining qualification limits.
