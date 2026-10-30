# Publisher writer feasibility probe

This opt-in contributor tool inventories native publications and synthesizes
small, template-free PUB files. It stays outside the normal solution and does
not expose a public Publisher creation or save API. It references the existing
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
```

Output parent directories must exist. Existing outputs are never overwritten.
Inventory input is bounded to 64 MiB, plain text to 4,096 characters and images
to 1 MiB. Reports can contain complete recovered text; use synthetic fixtures.

`inventory` records compound metadata and hashes, known Contents fields, Quill
directories, OfficeArt envelopes and a separate OfficeIMO reader result.
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
Only `Escher/EscherStm` is replaced through the shared compound writer; other
logical streams and directory metadata are retained. Physical CFB allocation can
change. Invalid or unreachable directory trees, noncanonical roots and nonempty
active-content, signature or protected streams are outside edited-source support.

Stories and formatting, tables, image replacement, page/master relationships,
new objects and source-free publications have no writer in this path. Every
current semantic family is compared with the import baseline before output;
source-byte retention does not permit dropping unsupported edits. Prepared edit
candidates report native acceptance, repair-free opening and opaque semantic
preservation as unassessed. Managed reopening and stream comparisons are
structural evidence, not Microsoft Publisher qualification.

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
