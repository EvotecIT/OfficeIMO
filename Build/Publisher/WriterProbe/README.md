# Publisher writer feasibility probe

This opt-in contributor tool inventories native publications and synthesizes
small, template-free PUB files. It stays outside the normal solution and does
not expose a public Publisher creation API. Public bounded source saving belongs
to [OfficeIMO.Publisher](../../../OfficeIMO.Publisher/README.md#save-a-native-publication);
this probe remains a contributor tool for codec qualification. It references the existing
Publisher, PDF and shared binary codecs without adding runtime dependencies.
Template-free creation never reads a source publication or directory template.

Build with the repository's pinned SDK:

```powershell
$env:MSBUILDDISABLENODEREUSE = '1'
$env:DOTNET_CLI_USE_MSBUILD_SERVER = '0'
dotnet build Build/Publisher/WriterProbe/WriterProbe.csproj -c Release
```

The assembly uses the existing `OfficeIMO.Publisher.Tests` friend identity to
access internal codecs. Its separate output directory keeps it apart from the
test assembly; it has no test-framework reference or production visibility hook.

Invoke the built assembly with an input and unused output paths:

```powershell
$probe = 'Build/Publisher/WriterProbe/bin/Release/net10.0/OfficeIMO.Publisher.Tests.dll'
dotnet $probe inventory ./reference.pub ./reference.inventory.json
dotnet $probe diff ./native-text.pub ./native-textlong.pub ./text-change.diff.json
dotnet $probe create blank ./candidate-blank.pub
dotnet $probe create text ./candidate-text.pub 'Template-free publication'
dotnet $probe create styled-text ./candidate-styled-text.pub
dotnet $probe create shape ./candidate-shape.pub
dotnet $probe create image ./candidate-image.pub ./picture.png
dotnet $probe create combined ./candidate-combined.pub ./picture.png
dotnet $probe render ./reference.pub ./managed-comparison
dotnet $probe preserve ./reference.pub ./exact-copy.pub
dotnet $probe edit ./reference.pub ./edited.pub 293 36 48 120 84 0000FF
dotnet $probe edit-text ./native-text.pub ./edited-text.pub 2 'Replacement text'
dotnet $probe edit-story ./native-text.pub ./styled-edit.pub 2 ./runs.json

dotnet $probe edit-paragraphs ./native-textparagraphs.pub ./edited-paragraphs.pub 2 ./paragraph-runs.json
```

Output parent directories must exist. Existing outputs are never overwritten.
Inventory input is bounded to 64 MiB, plain text to 4,096 characters and images
to 1 MiB. Reports can contain complete recovered text; use synthetic fixtures.

`inventory` records compound metadata and hashes, known Contents fields, Quill
directories, OfficeArt envelopes and a separate OfficeIMO reader result.
The reader result's `sourceStyles` explains recovered formatting through the
retained native style catalog and original story use ranges. Definition pair
ordinals preserve source identity; nullable values distinguish missing fields
from explicit overrides. `CharacterEntry`/`ParagraphEntry` cover complete indexed
STSH1 entries, while `CharacterDefinition`/`ParagraphDefinition` cover their
self-sized OPLs. All source spans use byte offsets in `Quill/QuillSub/CONTENTS`.
Story and use `NativeStart`/`NativeLength` count UTF-16 units from TEXT start,
including CR and NUL, separately from normalized public text. FDPP selectors
remain null when absent; pair zero can still supply resolver defaults. STSH0/2
entries and STSH1 prefixes/suffixes remain opaque and do not establish names,
builtin identities or inheritance semantics.

Unknown encodings and incomplete decoding remain visible. `diff` compares
streams, native chunk IDs and Quill tag/ID pairs. It distinguishes changed
payloads and known field values from record relocation, retaining occurrence
indices for repeated fields. Matching inspected values do not establish semantic
equivalence for opaque payloads.

`render` creates a fresh directory with `managed.pdf`, per-page SVG and
`managed.json`. The JSON contains normalized pages, frames, stories, tables,
diagnostics and artifact hashes. The PDF header is checked. These are managed
comparison artifacts; their reports leave native acceptance and visual fidelity
unassessed.

## Loaded-source candidates

`preserve` uses the internal preparation path to copy an unchanged imported
publication byte-for-byte. A semantic no-op retains the original native encoding.
`edit` takes a publication object identifier, page-local point bounds and an
optional six-digit RGB fill. It uses ordinary model mutations, then prepares a
separate candidate. Unsupported changes return diagnostics without creating the
output. Both commands leave the source and the model's native baseline unchanged.

The bounded editor guards the original FSP SPID and publication-ID binding,
patches complete existing F010 anchors with exact signed EMUs, and preserves
native endpoint direction and imported page/master registration. Existing local
solid RGB, opacity and line-width slots can change only when their source
representation is unambiguous and requires no new property. Paint visibility,
alpha, inherited-only properties and duplicate scalar slots are refused.
`Escher/EscherStm` is replaced through the shared compound writer; streams outside
the prepared replacements and directory metadata are retained. Physical CFB allocation can
change. Invalid or unreachable directory trees, noncanonical roots and nonempty
active-content, signature or protected streams are outside edited-source support.

`edit-text` replaces a single nonempty plain story with one paragraph and one
styled run, retaining that formatting and its final native CR convention.
`edit-story` reads an array of `Text`, `FontSize` and `Bold` values from a JSON
file bounded to 64 KiB. Other run and paragraph attributes come from the source:

```json
[
  { "Text": "Plain 12 | ", "FontSize": 12, "Bold": false },
  { "Text": "Bold 18", "FontSize": 18, "Bold": true },
  { "Text": " | Plain 12", "FontSize": 12, "Bold": false }
]
```

A single original run can split into styled runs. A styled source can be edited
again with the same run count, retaining each original run's other OPL fields.
Font sizes require exact native units and the supported range; ambiguous native
size/bold tokens and formatting-page overflow refuse output. The
original single-story modern profile has one frame, no table story, one Quill directory, the
original 15 tag/ID records, single character/paragraph formatting pages and
matching BTE indexes. The separately qualified multiple-story/page profile below
extends this admission. Growth moves the following
payload by whole 512-byte pages, retaining formatting-page alignment; shrinking
clears the former text tail. The shared Publisher formatting-page codec emits
changed run boundaries and OPLs. Unchanged OPL tokens and opaque style/library payloads stay unchanged while
TEXT, STRS, FDPC/FDPP endpoints and BTEC/BTEP references are updated. Accepted
content edits also remove the optional SummaryInformation thumbnail (PID 17)
through the shared OLE property-set editor. Other raw metadata values, sections
and the property-set header are retained. Unchanged/no-op preparation remains an
exact byte copy. Invalid property-set structure refuses edited output; other
cached-record semantics remain unassessed.

`edit-paragraphs` reads one run array per existing paragraph, also bounded to
64 KiB. It retains the paragraph count, ordering and paragraph attributes.
Within admitted ordinary-story profiles, text can
grow or shrink, existing empty paragraphs can remain empty or gain text, and
size/bold edits use each original native range. Character ranges can cross
paragraph boundaries; each paragraph retains its own FDPP definition and end.
Adjacent identical character OPLs coalesce. Making a paragraph empty is refused when its source CR has a different
OPL from the last visible run, because that would substitute an unqualified
typing style. Formatting-page overflow also refuses output.

The multiple-story/page path retains existing story identities and order, requires
one frame per changed story, and admits one directory with at most 20 descriptors
and no table stories. It verifies every FDPC/FDPP page against its BTE index,
retains the existing page counts and shares byte-identical OPLs within each page.
Story lengths, endpoints, page offsets and later-page header starts are remapped;
unrelated Quill payloads remain byte-identical. Chained directories and output
that cannot fit the existing formatting pages are refused. The original
single-story/one-page encoding path remains unchanged.

Other formatting changes, run merges, paragraph-count changes, tables, image replacement,
page/master relationships, new objects and source-free publications have no writer in this path. Every
current semantic family is compared with the import baseline before output;
source-byte retention does not permit dropping unsupported edits. Each new
candidate still requires native qualification; the recorded samples below
establish only their listed operations. The preparer reports repair-free opening
and opaque semantic preservation as unassessed. Managed reopening and stream
comparisons supply structural evidence.

When story text or formatting changes, the preparation result's
`styleDependencies` retains the affected original story keys, native ranges,
definition pairs and direct OPL spans. It refers to the imported source even
after normal model edits. Unsupported changes return no candidate bytes; the
bindings do not establish editable named styles or unknown metadata semantics.

The `sourceEditing` section of the [qualification record](qualification.json)
binds six edited files to their source fixture or earlier edited/native-resaved
file, run specifications, codec source and native artifacts. All six passed
read-only open, initial inventory/PDF export, native save-as, read-only reopen
and a second inventory/PDF export. Native text, frame bounds and size/bold states
match the requested values within 0.1 point, and native rendered pixels remain
identical across resave. This includes re-editing a file after Microsoft
Publisher resaved it. These six runs do not exercise native text mutation or
establish repair-free opening, other profiles or opaque record semantics.

The record's `paragraphEditing` section binds five multi-paragraph specimens to
the independent TextParagraphs control or a native-resaved derivative. They cover
growth, shrink, filling an empty paragraph, size/bold changes including a styled
empty paragraph, and re-editing after Publisher resaves. Native text, frame bounds
and visible size/bold states match the requests within 0.1 point; native PDF pixels
match before and after resave. Raw comparison retains unrelated streams,
directory metadata, opaque Quill payloads and unrelated OPL tokens. These files
do not qualify paragraph-count changes, general native layout or repair-free
opening.

The `previewInvalidation` qualification section records four new thumbnail-free
variants of previously qualified edits. All other logical streams are identical
to those earlier candidates, and the SummaryInformation change removes only
PID 17. Native open/export/save/reopen succeeds, with unchanged native inventories
and rendered pixels across ten physical page states. Publisher recreates the
thumbnail when resaving these files. That raw property observation does not
qualify the regenerated thumbnail's appearance or repair-free opening. The
preparation diagnostic reports invalidation; OfficeIMO does not render a new
thumbnail.

The record also binds nine placement/paint candidates to independent native
Rectangle, TwoPageMaster and PictureOwnership controls. Across 19 physical pages,
native inventories and PDFs retain the requested rectangle fill/bounds, physical
and master frame bounds, and physical/master picture bounds after native
save/reopen. Unaffected object geometry/text and shared image pixels stay
unchanged; native rendered pixels match before and after resave. Original
Contents AA/AB extent fields remain unchanged for the two qualified picture
edits, without qualifying their general semantics.

The native runner can also exercise whole-range text editing on a candidate with
one physical page and one unlinked text box:

```powershell
./Build/Publisher/Invoke-NativePublisherOperation.ps1 -Operation Validate -Stage Full -InputPath ./styled.pub -OutputDirectory ./native-edit -NativeEditText 'Native edit accepted'
```

Initial inventory/PDF artifacts retain the original text. The runner changes
only its writable snapshot, saves to a separate PUB and checks the requested text
after read-only reopen. The recorded styled specimen passes this operation,
including quoted text and a trailing backslash; independent native PDF text
matches and source bytes remain unchanged. A freshly assigned range can omit the
final paragraph CR that appears after save/reopen. This is one controlled native
object-model edit; character-level editing, manual UI editing, general formatting
and repair-free opening remain unassessed.

## Creation cases

All cases use a single Letter page, one master and the same ordered
`CandidatePublicationSpec`. Object IDs, story keys, layout identifiers and SPIDs
are assigned once and shared by the Contents, OfficeArt and Quill serializers.
Text boxes and rectangles each own an independent story; pictures have no story.
The current drawing specification supports one picture per publication.

| Case | Objects and content |
| --- | --- |
| `blank` | Empty page. |
| `text` | Text box at `(72,72)`, size `288×108` points; optional plain text argument. |
| `styled-text` | Text box at the same position, containing the three styled runs below. |
| `shape` | Solid red rectangle at `(72,72)`, size `288×108`, with a CR-only story. |
| `image` | PNG or JPEG at `(72,72)`, size `288×108`; image path required. |
| `combined` | Styled text box at `(72,72)`, size `288×108`; red rectangle at `(72,216)`, size `144×72`; picture at `(330,216)`, size `96×96`. Image path required. |

The styled text uses one shared `OfficeRichTextParagraph` containing Arial runs:
`Plain 12 | ` at 12 points/plain, `Bold 18` at 18 points/bold, and
` | Plain 12` at 12 points/plain. Each run explicitly selects generated font
index zero, its size and its bold state. A native CR follows the final run.
In the combined case, the rectangle owns a second CR-only story. Its formatting
positions follow the styled story in the global text stream.

## Native codec structure

`create` never reads a PUB or invokes Publisher. The Contents stream has a typed
header and class catalog, with native record identifiers starting at 256. The
canonical compound writer serializes a fresh directory model carrying the
Publisher root and QuillSub storage class IDs. Generated CompObj identification,
an empty root `Objects` storage and a typed 12-byte `\u0003Internal` version
header provide the native recognition metadata. The Internal header contains
version `0x00010000`, a reserved zero WORD, zero context flags at unaligned offset
6 and a zero following-byte count at offset 10.

The page graph has six ordered references: one master, one physical page and
four utility pages. `DpgMaster=1` and `DpgSpecial=5` select the physical interval
`[1, 6-5]`. Non-master pages reference the generated master. The master alone
references an `OhMgPage` margin-guide record (`0x4C`) with explicit dimensions,
four calculated coordinates `0,width,0,height`, two column guides, two row
guides and zero gutters. Physical pages use the master relation for geometry.

Document field `0x20` references the Quill document (`0x5B`). Its frame catalog
(`0x61`) binds each story key to its object; its story catalog (`0x65`) binds the
same key to an `IdLayout` identifier. `IdLayout` is separate from the frame
reference count. Fresh story keys and layout identifiers follow story order.
Publications without stories retain empty catalogs.

Every case has an eight-slot `OplSccm` color scheme (`0x5C`) through document
field `0x22`, `OhColorScheme`: automatic black, blue, orange, gray, gold, dark
blue, purple and white. Rows are constructed from named RGB components; optional
scheme identity, name and extended properties use native defaults. Document
field 3 references a print block (`0x4B`) with one empty typed `OplDnb` row. Its
fields use native defaults; no printer name, device-specific payload or DEVMODE
data is copied.

The Quill envelope writes version values 4 and 7, byte `+0x0F` equal to 3, a
first payload position of 512 and a truthful stream size at `+0x14`. The stream
has at least 1,024 bytes. Header and directory descriptor counts follow the
actual records; descriptors point to their calculated serialized positions.
Generated default libraries contain `INK ` version words `0x74/0x74`, one Arial
`FONT` entry, one built-in-default `STSH` name and two paired empty character and
paragraph OPLs in `STSH` ID1. No native style payload is copied.

For text-bearing objects, `TEXT` concatenates stories in binding order. `SYID`
stores their keys, and `STRS` stores one UTF-16 length per story, a final zero
and one empty, length-prefixed SDF per story. Character and paragraph ranges use
absolute text-byte endpoints across the whole stream. `FDPC` and `FDPP` are
generated 512-byte pages, aligned to 512 bytes, with auxiliary WORD 1 and a zero
reserved DWORD. Their parallel arrays contain increasing endpoints and
page-relative style offsets. Plain text has one character range per story;
styled text has one per run, including the final CR in the last run. Paragraph
ranges follow story boundaries. `BTEC` and `BTEP` use calculated absolute page
offsets and bounds `[0, absolute TEXT end]`. The current formatter emits one
character page and one paragraph page, and rejects ranges that exceed a page.

OfficeArt contains a calculated `F006` drawing-group record and one `F002`
drawing container. The root group uses SPID 1024; object SPIDs follow in order.
Saved counts, last shape ID and identifier-cluster values follow that topology.
Text boxes use FSP type 202, rectangles type 1 and pictures type 75. Text-bearing
objects carry a calculated `F00D` ClientTextbox OPL with their story key in field
1/type `0x20`. Images use the shared delayed FBSE and standalone BLIP writers,
placing the image record in `EscherDelayStm`. No native property table or extra
utility drawing is copied.

## Qualification

The [qualification record](qualification.json) binds the combined specimen to
its source files, repository image and native application build. Its native
inventories and both reference PDFs establish the expected text, character
styles, object geometry and embedded image pixels before and after native
resave. The original and resaved PDF renders were also inspected. Repair-free
opening remains unassessed; these observations do not establish it.

Successful native opening, usable text ranges, saving, reopening and PDF export
are separate observations. Retain the candidate SHA-256, application version,
initial inventory and PDF, saved copy, reopened inventory and PDF, and the source
integrity result. The native validation tool performs those operations on an
isolated byte-identical copy:

```powershell
./Build/Publisher/Invoke-NativePublisherOperation.ps1 -Operation Validate `
    -Stage Full -InputPath ./candidate-combined.pub `
    -OutputDirectory ./candidate-native
```

`Inspect` stops after open and inventory; `Resave` also saves and reopens;
`Full` includes both PDF exports. Partial stages prove only their listed
operations. A completed Full report does not establish that Publisher performed
no repair: repair status remains unassessed unless a separate native UI
observation supplies that evidence. Compare the original and reopened native
text, styles, geometry and image pixels against explicit expectations.
OfficeIMO read/render success establishes only the managed result. Evidence for
these bounded cases does not qualify arbitrary Publisher features or a public
creation/save contract.
