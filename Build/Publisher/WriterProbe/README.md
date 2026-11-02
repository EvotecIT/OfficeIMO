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
dotnet $probe create blank-pages ./candidate-two-blank-pages.pub 2
dotnet $probe create text-pages ./candidate-text-pages.pub 3
dotnet $probe create object-pages ./candidate-object-pages.pub ./picture.png
dotnet $probe create image-pages ./candidate-image-pages.pub ./first.png ./last.png
dotnet $probe create named-image-pages ./candidate-named-images.pub ./first.png ./last.png
dotnet $probe create cropped-image-pages ./candidate-cropped-images.pub ./first.png ./last.png
dotnet $probe create shared-image-pages ./candidate-shared-images.pub ./first.png ./last.png
dotnet $probe create mixed-object-pages ./candidate-mixed-objects.pub ./image.png
dotnet $probe create master-content ./candidate-master-content.pub ./image.png

dotnet $probe create newsletter ./candidate-newsletter.pub ./image.png

dotnet $probe source-newsletter ./candidate-newsletter.pub ./replacement.jpg ./newsletter-source-saves

dotnet $probe create multiple-masters ./candidate-multiple-masters.pub

dotnet $probe create paired-masters ./candidate-paired-masters.pub
dotnet $probe create linked-story ./candidate-linked-story.pub
dotnet $probe create linked-story-overflow ./candidate-linked-overflow.pub
dotnet $probe create table ./candidate-table.pub
dotnet $probe create table-edited ./candidate-table-edited.pub
dotnet $probe create frame-layout ./candidate-frame-layout.pub
dotnet $probe create frame-layout-edited ./candidate-frame-layout-edited.pub
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
The `blank-pages` case requires a physical page count from 1 to 16. Other creation
cases place their objects on a single physical page, except `text-pages`, which
requires 2 to 16 pages and places independent stories on the first and last pages,
and the three-page `object-pages`, `image-pages` and `named-image-pages` specimens.

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
64 KiB. This contributor command retains the paragraph count, ordering and paragraph attributes.
Within admitted ordinary-story profiles, text can
grow or shrink, existing empty paragraphs can remain empty or gain text, and
size/bold edits use each original native range. Character ranges can cross
paragraph boundaries; each paragraph retains its own FDPP definition and end.
Adjacent identical character OPLs coalesce. Making a paragraph empty is refused when its source CR has a different
OPL from the last visible run, because that would substitute an unqualified
typing style. Directory capacity and native style-mapping limits still refuse output.

The `characterReading` qualification section binds two independent native controls
to typed font/color/decorations and all eighteen automation underline choices.
Quill's stored underline values have their own mapping; unsupported weights,
long-dash and compound patterns carry approximation reports. Native and managed
PDF/SVG references establish recovered content and visible decorations separately
from exact font metrics or decoration geometry.

The separate `sourceEditing.fontTableExtensionExperiment` records one appended
Consolas FONT entry with zero hint bytes and eight updated body script selectors.
Static text-component inspection shows that loading resolves names and replaces
saved font IDs with local IDs. Native open/export/save/reopen retains the complete
text and Consolas body font, with identical native pixels and spans across resave.
The original paragraph-CR definition remains separate, including its Calibri
whitespace span and mixed whole-range font name. This isolated experiment remains
distinct from the production character writer and does not establish repair-free opening.
The `sourceEditing.characterEditing` record qualifies public source saves for
italic, representable underline kinds, single strike, baseline, font families and
opaque RGB colors alongside size/bold. It includes mixed runs, native-resaved
re-edits and explicitly changed paragraph-CR attributes. The contributor JSON
commands above retain their narrower size/bold input schema; use the public
rich-text model for the additional attributes.

The public paragraph model also supports source saves for alignment, before/after
spacing, margins, first-line or hanging indents, absolute/multiple line spacing
and ordinary left tabs. The `sourceEditing.paragraphFormatting` record qualifies
21 native save/reopen cases, including fractional points, per-paragraph values,
linked overflow, count changes, clearing tabs and formatting/directory allocation.
Its independent controls include two-tab ordinal arrays and paragraph-aligned
left/center/right text. Center/right native paragraphs require the existing
`WithParagraphAlignment()` tab option; fixed-grid requests are refused. The
[support contract](../../../OfficeIMO.Publisher/SUPPORT.md) defines the admitted
units, native field and tab-setting limits. These operations use the public
rich-text API; the contributor JSON commands retain their narrower input schema.

The multiple-story/page path retains existing story identities and order, admits
one frame or a complete resolved ordinary document-page chain per changed story,
and admits known directory chains with at most 20 descriptors per page and
65,535 in total. Table stories use the separately qualified fixed-cell profile.
It verifies every FDPC/FDPP page against its BTE index
and shares byte-identical OPLs within each page.
Story lengths, endpoints, page offsets and later-page header starts are remapped;
unrelated Quill payloads remain byte-identical. Through the public
`ReplaceStoryParagraphs` API, paragraph counts can change when projected and raw
source paragraph/body definitions are uniform. The native codec can allocate or
retire formatting/directory pages and relocate enlarged indexes and resources.
Unknown padding remains retained; growth uses zero descriptor slots or appends a
directory page. The established single-directory allocation path retains its
nonzero-header guard. Unknown page metadata, directory capacity and ambiguous
source definitions guard that operation. Zero-paragraph native stories remain
unavailable. The original accepted single-story/one-page encoding path remains first
for edits within its supported character fields.

Linked-story source saving preserves existing native links, frame identities,
physical-page ownership and geometry. It requires every chain member to be
available, ordinary, visible and untransformed. Unresolved, grouped, transformed
and master-page chains remain unsupported for changed stories. The
`sourceEditing.linkedStoryEditing` qualification record contains seven public-save
specimens from the independent two-page Newsletter control, including a re-edit
of Publisher-resaved output. Complete native-resaved text and body size/bold,
unchanged Contents/artwork and other storage, retained frame links/geometry and
native PDF pixel/span equality are verified across 28 physical page states.
Visible frame slices are checked in order; an empty trailing frame can report
a lone paragraph CR after the complete story fits the preceding frame. This
does not qualify other chain profiles, repair-free opening or renderer parity.

Unsupported formatting changes, run merges, arbitrary count-changing rewrites
from nonuniform sources, page/master relationships, new objects and source-free
publications have no writer in this path. The separately qualified image and
table profiles are described below. Every
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

All cases use Letter-size physical pages and the same ordered
`CandidatePublicationSpec`. Most use one master; the multiple/paired-master
specimens carry their explicitly authored definitions. Object IDs, story keys, layout identifiers and SPIDs
are assigned once and shared by the Contents, OfficeArt and Quill serializers.
Text boxes and rectangles each own an independent story; pictures have no story.
Page, picture and image-asset counts depend on the selected case.

| Case | Objects and content |
| --- | --- |
| `blank` | Empty page. |
| `blank-pages` | Requested count of empty physical pages, with one shared empty master; count argument required. |
| `text-pages` | Builds a source-free publication through `PublisherDocument.Create`, explicitly adds and applies its empty master, then projects the final model through the shared ordinary-content adapter. Independent Arial 12/18 stories occupy the first and final pages, with explicit zero column gap; intermediate pages are empty. Count argument required. |
| `object-pages` | Projects the final public model of three pages with an empty applied master. A red rectangle without outline is on the first page, the middle page is empty, and an uncropped unnamed picture is on the last page. The rectangle has no stored story. Image path required. |
| `image-pages` | Projects the final public model of three pages with an empty applied master. Two distinct uncropped PNG assets are on the first and last pages; the middle page is empty. Pictures retain the unnamed native profile. Two image paths required. |
| `named-image-pages` | Uses the same public model as `image-pages`, with owned native filename resources and complete OfficeArt basename properties for subsequent image replacement. Two image paths required. |
| `shared-image-pages` | Uses public existing-image authoring for first/last-page placements of one asset, plus a distinct last-page asset. Native FBSE counts and delayed payloads follow model identity. Two PNG paths required. |
| `mixed-object-pages` | Projects the final public model with overlapping rectangle/picture/text/rectangle on the first page and text/picture/rectangle/picture on the last, with an empty middle page. Three pictures share one asset; two text frames own independent stories. One PNG path required. |
| `master-content` | Projects the final public authoring model: one centered master owns a rectangle, picture and single-paragraph text frame and is applied to three Letter pages. Page-owned foreground overlaps the master; one page picture reuses the master asset. Current object order, geometry, paint, stories and referenced asset identity supply the codec inputs. One PNG path required. |
| `newsletter` | Projects one final public model combining a recurring master heading, logo and rule with linked body text, an independent sidebar, a two-column story, a borderless 2×2 calendar and a page picture sharing the master asset. Three physical pages and six complete stories; one PNG path required. |
| `linked-story`, `linked-story-overflow` | Projects the final public model with an empty applied master. One story flows through three frames on pages 1, 3 and 2, alongside an independent label. Fixed Arial text and 18-point line spacing; the longer specimen retains terminal overflow. |
| `frame-layout`, `frame-layout-edited` | Projects six final public-model text frames on two pages with an empty applied master: one/two/three columns, asymmetric and fractional insets, full intercolumn gaps and Top/Center/Bottom alignment. The edited case applies two columns, a 14-point full gap, 9/3/15/6-point insets and Bottom alignment. |
| `table`, `table-edited` | Projects the final public model of a borderless 2×2 table with unequal tracks, bold headers, an empty final cell and independent label, with an empty applied master. The edited specimen adds a first-cell paragraph and refills the last cell. |
| `table-grid`, `table-grid-empty`, `table-grid-edited` | Requires a `<rows>x<columns>` argument within the legible specimen range of 1–12 rows and 1–8 columns. Authors an unmerged borderless grid with a caption, fixed tracks and cell paragraphs. The empty variant retains every empty cell; the edited variant changes bounds and the final cell, then places the table before its caption in drawing/story order. |
| `cropped-image-pages` | Uses the named-image model and public `SetPictureCrop` on the first picture, with distinct left/top/right/bottom fractions `0.25/0.125/0.0625/0.1875`. The shared OfficeArt encoder owns the native values. Two PNG paths required for this qualified specimen. |
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

The shared character encoder accepts black Arial runs with size, bold and an
optional language identifier. It refuses other text colors, italic, decorations,
baseline shifts, backgrounds, explicit pair kerning, per-run indentation and
link annotations. The shared paragraph encoder also refuses labels. These
checks apply to every styled story, including table cells and model-projected
master content, so unsupported formatting cannot be silently dropped.

## Native codec structure

`create` never reads a PUB or invokes Publisher. The Contents stream has a typed
header and class catalog, with native record identifiers starting at 256. The
canonical compound writer serializes a fresh directory model carrying the
Publisher root and QuillSub storage class IDs. Generated CompObj identification,
an empty root `Objects` storage and a typed 12-byte `\u0003Internal` version
header provide the native recognition metadata. The Internal header contains
version `0x00010000`, a reserved zero WORD, zero context flags at unaligned offset
6 and a zero following-byte count at offset 10.

`CandidateModelProjection` derives the text-page, master-content, mixed-page, object/image-page,
linked-story, frame-layout, table, newsletter and multiple/paired-master codec inputs from the final `PublisherDocument`;
none maintains a second editable object list. Current picture bounds, source
crops, referenced image bytes and model asset identities supply the codec. Named
profiles derive their native filename and OfficeArt basename from the current
asset ID and PNG/JPEG content type; unnamed profiles omit both name records.
Unused model assets are not emitted, and independent assets with identical bytes
retain separate native bindings. New edited variants have artifact reload proof;
their native opening, editing, saving and appearance require separate qualification.
The projection accepts the existing source-free Letter/single-master profile.
The ordinary object profile rejects imported publications, tables, linked or
multicolumn text, insets, non-Top alignment, wrap references, multiple paragraphs
and non-rectangle artwork. Those features require their separate writer profiles. Native paragraph terminators
retain the existing encoding convention. This projection does not admit public native creation or broaden
the probe's native acceptance evidence. The text-only multiple-master projection
accepts either two centered definitions or one centered definition followed by
the named left/right pair. It derives current names, assignments, frame order,
bounds and complete single-paragraph stories, retaining the existing omission
of the default column-gap property. Unassigned physical pages, nonzero gaps,
other frame layouts and non-text content require separate qualification.

The linked, frame-layout and table adapters require an explicit empty applied
master; this is specimen authoring, while `PublisherDocument.Create` still creates
no master by default. Linked frames retain complete story identity and chain
ordinals independently of physical page and drawing order. They require the
ordinary unwrapped single-column Top-aligned layout. The separate frame-layout
profile accepts unlinked, unwrapped one-to-three-column frames. The table profile
accepts a complete unmerged borderless grid within the native 16-bit cell-count
field and ordinary captions; current
tracks, cell paragraphs and table/caption order supply the codec inputs. Explicit
story automatic-hyphenation settings use the same native flag encoder as source
saving, preserving false separately from the omitted default. These edited variants
have encoded-artifact reload evidence; native editing, saving and appearance
remain separate qualification requirements.

`ComposedContent` combines the ordinary master/object profile with physical-page
linked and unlinked layout frames and at most one qualified table. It caches
linked stories by model identity and uses the existing graph, Quill and OfficeArt
encoders for all objects. Master tables, multiple tables, linked columns/insets,
wrap references and broader table-grid composition remain unqualified.
The newsletter and a later edited four-page variant have normal file-association
opening evidence in Microsoft Publisher before any native resave, without recovery
consent. All three initial pages and all four edited pages were inspected.
The `nativePrintState` qualification records the shared writer's generated
color-separation record and its required PrintBlock reference. Both corrected
variants open normally, export native PDFs and reach Print preview with Microsoft
Print to PDF Ready. Their seven native PDF pages render identically before and
after the print-state correction; their OfficeIMO PDF/SVG bytes are unchanged.
A subsequent OfficeIMO source-text save changes only Quill, preserves the print
records and exports the updated recurring heading on all three native PDF pages.
The `source-newsletter` command exercises the public `Save` API on this generated
three-page profile. It writes linked-story, final-cell, picture-replacement and
combined edits to an empty output directory, together with their hashes and save
diagnostics. The default strict policy refuses these edits; explicit `Allow`
retains the current unassessed-loss reports. Native read-only opening and PDF
export qualify all four outputs: complete linked text and cell paragraphs survive,
one replaced picture leaves the shared master asset intact, and unaffected pages
and regions match the native baseline. These are source-save observations;
native editing and resaving remain separate requirements.
Actual Print dispatch reports the same operation error for the corrected specimen
and a native-authored control, so printing remains unqualified. Native
editing/resaving, recovery-path acceptance and general appearance parity also
remain unqualified.

The page graph has `physical page count + 5` ordered references: one master,
the physical pages and four utility pages. `DpgMaster=1` and `DpgSpecial=5`
select the inclusive physical interval `[1, ordered count-5]`. The one-page
cases retain their original six-reference encoding. Non-master pages reference
the generated master. The master alone
references an `OhMgPage` margin-guide record (`0x4C`) with explicit dimensions,
four calculated coordinates `0,width,0,height`, two column guides, two row
guides and zero gutters. Physical pages use the master relation for geometry.
Each object's page index selects both its native parent and its page's ordered
object-reference list. Ordinary objects from all physical pages share one
OfficeArt drawing; page ownership does not allocate separate drawing IDs.

PrintBlock (`0x4B`) field 3 references a generated ColorSepBlock (`0x4F`) child.
Native Print preview requires this relation even when the initial document loader
accepts its absence. The child contains one automatic-color plate and explicit
native trap-width defaults, without stored printer names, drivers or device-mode
data. It is appended after existing resources to preserve their identifiers.

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

The record's `blankPageCreation` section separately qualifies generated two- and
four-page blank publications through all seven native operations. Each physical
page retains its order, page number, side and empty master relationship before
and after resave. All 12 initial/resaved page renders match the independent
native-written `BlankPages` reference at 144 dpi; the native resaves also reopen
through OfficeIMO. Counts 1 and 16 have managed/structural checks in this record.
All six original one-page creation cases remain byte-identical. This evidence
does not qualify objects on later pages, additional masters or repair-free opening.

The `textPageCreation` section separately qualifies the closed first/final-page
text specimen at two and three pages. All ten initial/resaved page renders match
the independently authored `TextPages` reference at 144 dpi, including the empty
middle page. Visible non-whitespace character positions and fonts match; the
producer's implicit CR typing is separate from the generated CR that follows its
last body run. Sixteen pages has managed/structural coverage. The six original
one-page specimens and three qualified multiple-blank-page specimens remain
byte-identical.

This work also qualifies an omitted native column gap and a subsequent explicit
zero-gap edit. OfficeIMO exposes the full gap: the native automation value is half
that distance. The generated model's zero gap is serialized explicitly, while an
omitted native gap imports as 7.2 points. The retained-source two-column edit has
native field, ownership, text and initial/resaved render-stability evidence; its
paragraph defaults do not establish layout parity with a fresh native template.
General wrapping, native repair-free opening and public native creation remain
unqualified.

The `objectPageCreation` record qualifies a story-free rectangle and embedded
picture on separate physical pages. The source-free model explicitly creates
and applies the empty master used by this native profile. Native resave retains
zero stored stories, although Publisher's rectangle text-range getter returns a
virtual CR. The independently authored `ObjectPages` reference contains an empty
shape story; its visible result matches this story-free encoding.

The same generated file is loaded and edited through `SetShapeBounds`,
`SetShapePaint` and `SetPictureBounds`, then saved through the existing public
native source-save path. All twelve original/edited initial/resaved page renders
match their independent native references at 144 dpi. The PNG bytes, page/object
ownership, fill and outline state survive native resave, and unrelated source
streams remain unchanged by the managed edit. Current managed reports include
typed shape and picture views with their paint descriptors.

Managed PDF output preserves the image and placement but enables interpolation;
the native references disable it. Enlarged image edges therefore differ. This
is separate from the native file's qualified pixel appearance, and does not
establish general managed/native rendering parity. The closed profile does not
qualify arbitrary model serialization, interleaved object order, additional
paint states or master content. Repair-free opening remains
unassessed.

The `imagePageCreation` record qualifies two distinct embedded PNGs on the first
and last of three physical pages, with an empty middle page. The public authoring
model supplies separate assets; the probe calculates their one-based FBSE indices
and consecutive delayed BLIP offsets through the shared image codec. It does not
deduplicate equal image bytes. The separate `sharedImageAuthoring` record qualifies
explicit model asset reuse, derived reference counts, and selected replacement
before and after native resave against independent `SharedImagePages` controls.

The `mixedObjectOrderCreation` record uses overlapping text, rectangles and shared
pictures in different orders on the first and last pages. Its independent
`MixedObjectPages` controls explicitly disable wrapping, matching public creation
without text-wrap exclusions. Combined story, vector-paint/bounds and picture
bounds edits retain object order on generated, native-resaved and native-authored
sources. The compact Quill source profile supports these text edits without
requiring native resave first; unavailable font/color resource growth refuses
output. Native repair-free opening and generic model serialization remain
unqualified.

The `masterContentCreation` record qualifies one populated single-page master
applied to three physical pages. Its objects remain owned once, and the image
store counts two logical placements rather than four rendered appearances.
Independent `MasterContent` controls retain master/foreground layering. Master
text, rectangle bounds/fill and image replacement can be saved on generated,
native-resaved and native-authored sources; the page-owned image remains on its
original asset. Inheritance and generic model serialization remain unqualified.

The `multipleMasterCreation` record qualifies multiple centered masters and a
centered master beside a left/right pair. `AddMasterPagePair` supplies both typed
definitions atomically; the probe derives owner indices and direct assignments
from the public pages. The native graph generates each master's geometry record
and uses its side registration for OfficeArt anchors. Master and page text-frame
edits save from generated, native-resaved and native-authored sources and match
independent `MultipleMastersEdited` and `TwoPageMasterEdited` controls. All 28
earlier source-free profiles remain byte-identical. Managed save/reload pixels are
stable, but native/managed text origins differ by up to 0.07 pt horizontally and
1.38 pt vertically; this does not establish native/managed pixel parity. Master inheritance, automatic
parity selection, repair-free opening and generic public native creation remain
unqualified.

The `masterAssignmentEditing` record qualifies existing physical-page direct
assignment and master-formatting suppression through public native source saves.
The native `IgnoreMaster` control keeps its selected reference while hiding the
definition's applied content. Reassignment restores formatting and selects the
requested existing definition. Independent single/paired reassigned and ignored
controls distinguish this from deleting a native reference. Unknown relationship
encodings remain explicit gaps and are not rewritten. New topology and generic
native creation remain separate contracts.

The `verticalTextPlacement` record qualifies final paragraph-after spacing in
22 independent native top/center/bottom profiles over eight pages. The shared
flow owner omits spacing at the region boundary while preserving source margins,
internal paragraph gaps and complete stories. Explicit 0/6/12 pt spacing, 12/18 pt
Arial and requested terminal marks distinguish that contract from deleting empty
paragraphs or changing the source styles. The largest measured first-baseline
difference is 1.40 pt; native font metrics and line advances remain unqualified.

The `linkedStoryCreation` record qualifies a three-frame story whose chain order
differs from physical page order. Frame catalog rows are ordered by story and
ordinal; the story's `CRef` count is independent of `IdLayout`. Generated Quill
layout slots support every frame ordinal without copying a native layout cache.
Paragraph overrides use the existing production paragraph codec. Source saves
preserve opaque layout data on compact and native-normalized files, including a
normal native profile without an unused palette. Native creation, growth and
shrink saves retain both complete stories, chain identity and terminal overflow.
All 19 earlier creation profiles remain byte-identical. Native open/save/reopen
and stable resave pixels are qualified separately from independent appearance:
the generated files do not reproduce the control's hyphenation and frame breaks.
Language/CR defaults, repair-free opening and generic public native creation
remain unqualified.

The `tableCreation` record qualifies one borderless, unmerged table and its
independent label. Public `AddTable` and cell-editing snapshots supply the grid,
complete story and cell paragraphs. Generated `TCD` endpoints partition that
story; table story metadata stays outside the ordinary frame catalog. The table
frame, cell records and layout slots contain no copied native cache. The shared
formatting-page encoder handles repeated styles when the earlier unshared page
would overflow; all 21 earlier creation outputs remain byte-identical.

Both creation variants and public source edits on generated, native-normalized
and independently authored files match native controls across ten stage/page
comparisons. The compact source profile admits validated table descriptors;
the existing table writer remaps cell boundaries while preserving opaque layout
data. Table fills/borders, dynamic row sizing, generic public native
creation and repair-free opening remain unqualified.

The `rectangularTableCreation` record extends the separate table profile to
1×1, empty 1×2, 1×4, 4×1, 3×4 and 12×8 grids. Microsoft Publisher opens and exports
all six, an edited table-first 3×4 specimen, and a public source edit adding a
third paragraph to the final 12×8 cell. Native object inspection confirms all
cell texts, bounds and row/column tracks. All eight PDF pages were inspected for
visible text and clipping. The 96-cell specimen uses two character-formatting
and two paragraph-formatting pages; creation reuses the source writer's allocator,
page links and indexes. The two original 2×2 outputs and four newsletter source
saves remain byte-identical. Creation still requires one Quill descriptor
directory and refuses formatting that exceeds it. These read-only native
operations do not qualify native editing, resave/reopen or repair-free opening.

The `table-merged` specimen uses public `MergeTableCells` before filling a
three-by-four grid. It carries horizontal, vertical and rectangular spans in
seven logical cells through the existing cell-record and Quill story encoders.
The `mergedTableCreation` record compares initial native export and native
save/reopen with independent `TableCreationMerged` controls. Public source edits
clear and grow the merged owners and match `TableCreationMergedEdited` before
and after native resave. Span/track/text readback and four native PDF page pairs
are exact at 100 dpi. The public model retains styled paragraphs during merging;
imported topology changes remain refused. Native/managed PDF parity, decorated
tables, dynamic row sizing and the public source-free native writer remain
unqualified.

The `tableCellDecoration` record maps solid cell fills and shared grid borders
from independent native controls into the canonical table model and current
SVG/PDF output. Cell fills use native definition indices; border ranges use
grid-boundary endpoints. Utility shapes remain outside publication object
identity. Existing Core colors, border-side values and drawing primitives own
the projected paint. Public cell-text growth preserves complete Escher bytes;
native initial/resaved PDFs match an independent text-edited decorated control.
Native/managed pixel parity and decorated template-free encoding remain
unqualified. The comparison worker retains explicit unknown
fill/border observations, including unavailable diagonal observations, and its
extended inventory also completes on the 96-cell control.

The `tableCellStyleEditing` record extends existing-source saving to solid cell
recoloring and shared-border color/width edits through the public canonical model.
The complete saved Escher stream matches the independent control in both
directions. Native initial/resaved inventories retain all cell text, spans, tracks,
fills and edges; four PDF page pairs match at 100 dpi. The writer guards existing
local scalar slots and refuses utility-record insertion or
removal, ambiguous ownership and widths without an exact signed-EMU encoding.
Restoring original styles before saving returns the exact original PUB bytes.
This does not qualify native source utility insertion/removal, native cell
editing after opening these outputs or repair-free acceptance.

The `table-decorated` and `table-decorated-edited` specimens add source-free solid
cell fills and shared borders from public authoring snapshots. `tableDecorationCreation`
records separate utility drawing allocation, exact native instance prefixes
`2/1/0`, calculated SPIDs and table/cell/grid bindings. A two-byte prefix control
distinguishes refused zero-instance framing from successful native opening.
The qualified rectangular creations and scalar edits of generated/native-resaved
input complete Full validation: eight cell/grid inventories and eight native PDF
pairs match independent controls at 100 dpi. Four workflows cover three distinct
candidate byte sequences; a direct source scalar edit equals fresh edited creation.
The earlier blank, borderless-table and merged-table bytes remain unchanged.
Decorated merges, source utility insertion/removal, broader native visibility definitions,
larger/composed utility drawings, native cell editing and repair-free acceptance
remain open; public native creation remains gated.

The `tableCellFillVisibility` record qualifies hiding/activating existing solid
cell utilities and inserting a missing local RGB scalar through the existing
source property editor. Owning filled use-bits, unrelated masks, opaque records,
latent colors and drawing-instance prefixes remain retained. Three independent
controls cover changed/restored visibility and combined paragraph growth; their
complete Escher streams match public source saves. Five Full workflows include
independent, generated and native-resaved input: ten cell/grid inventories and
ten native PDF pairs match controls at 100 dpi. The unsupported complex-peer
regression is fixed and confirmed; post-fix output bytes match every native-
validated input. Non-solid/transparent paint, ambiguous/unsupported boolean
roles, missing local visibility ownership, utility insertion/removal, broader
native producers and repair-free acceptance remain outside this profile.

The generated file and a public `SetPictureBounds` edit each match independently
authored `ImagePages` and `ImagePagesEdited` controls before and after native
resave: twelve page pairs at 144 dpi. Public reload preserves both complete PNG
payloads and picture/page bindings. The edit changes only the second picture's
anchor values; the first picture, image store and all other streams remain
byte-identical. Generic object ordering and repair-free opening remain unqualified.

The `namedImageReplacement` record adds complete filename metadata without
changing the fourteen earlier creation profiles. The named specimen supports a
last-picture PNG replacement, a later PNG-to-JPEG replacement after managed
save/reload, and JPEG replacement after Publisher itself saves the PNG result.
Each file matches its independent native final-state control before and after
native resave: 24 whole-page pairs at 144 dpi. Public reloads retain active image
payloads, ownership and bounds; the first picture remains unchanged.

The last path covers Publisher removing the local filename reference while
leaving old resource chunks in place. Native source saving appends a fresh owned
resource for the qualified embedded-picture profile and preserves those older
chunks. Native resave can retire unused image slots, which the reader keeps
inert. The independent native controls can deduplicate image IDs; their final
painted content is the comparison contract. Shared-asset authoring, broader crop profiles,
generic model serialization and repair-free acceptance remain unqualified.

The `pictureCropEditing` record qualifies public crop set/clear, combined crop
with PNG replacement and resized bounds, an existing JPEG crop, and the internal
`cropped-image-pages` creation profile. Five files match independent native
references across 24 initial/resaved page pairs at 144 dpi; ten public reloads
retain encoded fractions, assets and other placements. All fifteen preceding
creation profiles remain byte-identical.

The model uses source-edge fractions rather than automation points. Native
16.16 rounding is reported separately, and output whose rounded source area
collapses is refused. The source writer inserts or overrides all four local
crop fields, including zero values when clearing, and composes them with the
existing image and bounds edits. This does not qualify transformed/custom-mask
crop edits, native image-sampling parity, repair-free opening or a public native
creation API.

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

`Inspect` stops after read-only open and inventory; `Export` also writes the
initial PDF without opening a writable document; `Resave` saves and reopens;
`Full` includes both PDF exports. Validation opens the copied candidate directly
without creating a blank publication. Partial stages prove only their listed
operations. A completed Full report does not establish that Publisher performed
no repair: that needs direct native UI evidence or an independently qualified
repair-path observation.

The controller runs in Windows PowerShell 5.1 or PowerShell 7 and uses an
installed PowerShell 7 worker for native COM operations. Direct Windows
PowerShell COM export stalled in the measured 96-cell case; the controller's
PowerShell 7 route completes that same export. Each operation has a total
`TimeoutSeconds` bound and a `PhaseTimeoutSeconds` inactivity bound. Monotonic
phase receipts record progress. The worker identifies its isolated automation
server before document work, records the executable version/hash, cross-checks
available COM windows and detects executable changes during qualification.
Cleanup targets only its recorded PID, start time and executable.

```powershell
./Build/Publisher/Invoke-NativePublisherOperation.ps1 -Operation Validate `
    -Stage Export -InputPath ./candidate.pub -OutputDirectory ./native-export `
    -TimeoutSeconds 180 -PhaseTimeoutSeconds 45
```

`nativeOracleLifecycle` records fresh observations on Publisher
16.0.20527.20112: native blank creation; text mutation/save/reopen on an
independent control and a template-free text specimen; complete story and table
preservation through newsletter save/reopen; and 96-cell export and native
save/reopen with every cell text and track retained.
All three newsletter initial/resaved page pairs are pixel-identical. These
results qualify their listed operations and profiles. Public template-free
creation, repair-free acceptance and broader lifecycle coverage remain separate.
Three recent scratch cohorts carried an earlier executable identity without
capturing it; their `binaryIdentityCorrection` entries withdraw that attribution
while retaining the recorded geometry, text and PDF observations.

The qualification record's `observations.diagnosticObservation`
captures six selected diagnostic/consent sites for the exact combined candidate.
A native blank preserves its inventories and both 144 dpi page comparisons under
capture; a one-byte signature-invalid copy produces a severity-4 diagnostic and
native rejection. The combined candidate produces no selected diagnostic or message
request during its initial open, writable open or resave reopen. This calibrates
diagnostic capture, but repair-positive controls and other recovery paths remain
unqualified, so PUB00 stays open. The `repairControlBoundaries` record adds
single-scalar native-text controls: zero font size and primary font index 65535
survive save without selected diagnostics; the former has no body paint in its
native PDF, while the latter renders through fallback. An unresolved text-frame
catalog kind `0x7FF` produces another severity-4 rejection. These distinguish
accepted malformed character data from strict structural rejection; none is a
successfully repaired publication. Compare the original and reopened native
text, styles, geometry and image pixels against explicit expectations.
OfficeIMO read/render success establishes only the managed result. Evidence for
these bounded cases does not qualify arbitrary Publisher features or a public
creation/save contract.

`observations.laterLoadRecovery` extends the capture to the object parser return,
object loader return and three decoded suppression/retry boundaries. The native
Text control records 35 successful parser and loader results on each open. The
exact generated specimen records 16 on each original open and 21 after native
resave. Neither enters the selected suppression or retry paths. Current-schema
plain/captured inventories agree, as do four 144 dpi page pairs. The generic
recursive parser hook exceeded the 4,096-event cap in an earlier compatibility
attempt; that incomplete run is excluded. The decisive capture uses the exact
object-parser return site with the same cap and completeness guards.

Static instructions identify a document-bit suppression of `0x80040043`, a
kind-`0x6A` suppression of two decode errors and a later `0x66C2` consent/retry.
None of the 136 parsed repository fixtures contains kind `0x6A`; two additional
fixtures are outside that scan's parser. These branches remain without a positive
recovery control. Public native creation remains gated by PUB00.

`observations.validationConsentControl` supplies a positive request-path control.
A copy of the independent native Text fixture changes only the encoded type byte
of frame field 183 from `0x20` to unsupported type `0xF8`, preserving the field ID,
all lengths and every other byte. Native validation reports severity 2, retains a
zero strict-failure flag and enters the consent branch. Message `0x66C2` returns
`2` from caller RVA `0x191279`; validation returns `0x80040052` and loading stops
before the object parser. Eleven contiguous, same-thread events pair the load,
validator, diagnostic and message entries/returns. The original and copy remain
unchanged and both owned processes close. The Full attempt remains incomplete;
no post-open drain, successful repair or native resave is claimed. Accepted
recovery and the later suppression paths still need positive qualification.

### Direct native opening

`nativeUiOpening` records normal Windows file-association opening of the exact
combined candidate, before native resave, in Publisher 16.0.20527.20088. Its styled
text, red rectangle and embedded image appear on one page without accepting a
file-problem, repair or recovery prompt. The independent native Text control also
opens. Both source hashes remain unchanged after closing their windows.

The modified Text control used by `observations.validationConsentControl`
produces a visible file-problem and trusted-source warning on that same route.
This is a positive warning-path observation; successful recovery remains
unqualified after the warning was cancelled manually. The installation reports
read-and-print-only mode, so this UI observation establishes opening behavior
only. It does not establish native
editing, saving, silent repair coverage or acceptance of other specimens.
PUB00 and public source-free native creation remain gated by the recovery
qualification in the roadmap.

The `sourceEditing.objectOrdering` record covers existing document/master object
stacking through `ObjectIds` and `SetPageObjectOrder`. Independent native
`MixedObjectOrder`/`MixedObjectOrderReversed` controls establish drawing-record
order while the Contents ownership list stays fixed. Whole-record permutation
composes after supported property/image-store edits; selected IDs must occupy one
complete sibling scope. Nine actual public outputs complete native save/reopen,
including linked frames and master content, with 105 separately classified page
comparisons and PowerShell workflows. Groups, native wrap-exclusion lists and
split scopes remain outside this source profile. This does not promote the
source-free native writer or close PUB00.

The `sourceEditing.wrapOrderingAdmission` record qualifies six native wrap/order
controls: both obstacles in front, both behind, explicitly disabled wrapping,
either obstacle in front and swapped foreground obstacles. Native frame lists
follow participating foreground IDs in stacking order. The public edit and native
writer refuse an enabled or unresolved wrapping object crossing a text frame,
including initially empty exclusion lists. The paired public regression and
disabled-wrap native save/reopen distinguish refusal from a blanket ordering ban.
This establishes admission and producer semantics; rebuilding native/model wrap
relationships remains open.

The `sourceEditing.wrapRelationshipEditing` record covers canonical/native
exclusion regeneration during existing ordering and frame/shape/picture bounds
edits. Five source-codec wire candidates and two geometry controls precede eleven
actual public saves, each compared with independently authored native references.
Default-type controls distinguish implicit square pictures from ordinary
rectangular through wrapping; non-default sides remain outside the profile.
The `wrap-order` command exercises count/list replacement and whole-record
ordering directly through the shared codecs before a public profile is admitted:

```text
wrap-order <input.pub> <unused-output.pub> <specification.json>
```

The JSON specification supplies `ObjectIds` as a complete existing page
permutation and `FrameWraps` as an object mapping existing frame IDs to ordered
obstacle-ID arrays. Empty arrays remove the native count/list fields. This is
wire feasibility tooling, not another public editing model.

The `sourceEditing.imageReplacement` record covers public native saves for
existing PNG/JPEG picture placements, including shared assets, master pictures,
cropping, repeated replacement, combined frame changes and re-editing a native
resave with retired image slots. Its recipes bind input and replacement hashes
to independent raw image/filename/crop checks and native before/after PDFs.
Qualification covers delayed image streams; no-delay embedded-BLIP appends,
exact cropped rendering parity and repair-free opening remain separate gaps.

The `frameLayoutCreation` record qualifies six public-model frames on two pages
with one/two/three columns, full intercolumn gaps, asymmetric/fractional insets
and Top/Center/Bottom alignment. Creation writes those current snapshots into
OfficeArt properties and Contents alignment fields. Public layout edits save
across generated, normalized and independent producer profiles. Five native
Full runs complete 35 operations; 60 getter states and ten public reloads retain
the layouts and six complete stories. Eighteen of 20 independent page comparisons
match at 144 dpi. The baseline page-one difference is native hyphenation and its
resulting column breaks; all edited files match the edited native reference.
Column balancing, general appearance parity and repair-free opening remain
separate qualification work.

The `nativeLanguageFeasibility` experiment distinguishes the missing language
field from the opaque companion character field. Ordered `0x12=1033` alone
matches all ten initial/resaved page pairs for the column and linked-story
controls; `0x3E=1033` alone leaves eight of those pairs different. Both fields
also match all ten pairs. Native French controls independently associate
`LanguageID=1036` with `0x12=1036`. The `runLanguageLifecycle` record carries the qualified primary value through
the shared public run model and native codec. `frame-layout-language`,
`linked-story-language` and `linked-story-language-overflow` use
`WithLanguageId(1033)` on their public-model runs. Native source edits qualify
English/French set/reset and unspecified-language preservation; 12 files pass
84 native operations and all 52 independent page comparisons. The companion
field remains opaque. Broader language/script, native override clearing and
managed layout behavior remains separate.

The `sourceEditing.tableText` record covers existing cell text saves with fixed
grid, cell order and per-cell paragraph counts, including ordinary stories in
the same publication and re-editing a native resave. Independent checks bind
relative `TCD ` cell endpoints, native cell text/style/positions and unchanged
PDF pixels/spans across resave. The package support contract lists native typing
style, print-layout, rendering and broader table-profile limits.

The `sourceEditing.emptyTyping` record covers empty existing paragraphs/cells
and trailing empty paragraphs in admitted uniform count changes. Unspecified
typing retains the source CR; explicit empty runs use its resolved native
inheritance as the edit baseline. Source/native-resaved refill specimens confirm
the resulting visible font, style and color. Independent early/fresh-range
controls distinguish a body's style from its CR's defaults. Zero-paragraph
stories, cell-count changes and broader native/rendering qualification remain
separate limits in the package support contract.

The `sourceEditing.cellParagraphs` record covers fixed-grid cell paragraph-count
changes through the canonical per-cell Quill planner. Eleven public specimens
qualify first/final/all growth, shrinkage, balanced boundaries, empty typing and
refill, combined ordinary-story edits and native re-editing. A boundary-only
case moves a paragraph between cells while leaving the complete styled story
unchanged. Independent TCD/formatting/storage checks and fitting native PDFs
establish those profiles; a retained wrapped-text cohort separately records
Publisher's row expansion and changed resave rendering. The native empty-string
mutation control saves a retained CR, not a zero-paragraph story. Native repair,
general overflow layout and broader producer profiles remain unassessed.

The `sourceEditing.paragraphRetention` record covers nonempty paragraph retention
from nonuniform native definitions. Eight public outputs retain, delete, reorder
or duplicate complete source paragraphs, including explicit selection between equal
projected values with different native language metadata. Twelve initial/resaved
PDF pairs match independently authored final-state pages in pixels and text spans.
Raw checks preserve complete PC/PP definitions, final/intermediate CR roles,
neighboring stories and unrelated storage. Detached values require unambiguous
native templates. The public regression tests cover both explicit identity and
the refusal of ambiguous copies. The existing 116 save specimens remain
byte-identical. These controls qualify ordinary stories. The separate table-cell
controls are described below; text/formatting changes
during nonuniform count changes and repair-free opening remain unqualified.

The `sourceEditing.cellParagraphRetention` record qualifies the same shared codec
inside fixed table cells. Ten public outputs exercise first/final cells together,
each boundary cell on its own, copied values and explicit native language
selection. Complete PC/PP definitions, cell CR roles, TCD endpoints, tracks and
unrelated storage stay intact. Twelve full-page and eight row-region stage
comparisons match independently authored native controls in pixels and text spans.
Twenty public reloads retain complete text and cell ownership after native saving.
The shared authoring helper also reproduces all seven ordinary-control pages
without pixel/span changes. No writer logic or public API was added for these
fitting table profiles; general row reflow remains unqualified.

The `sourceEditing.emptyParagraphRetention` record covers empty paragraphs in
ordinary stories and first/final table cells. Retained empty paragraphs keep
their complete source CR in interior and final positions. Exact source snapshots
resolve opaque native language differences; ambiguous copied values are refused,
including when other paragraphs contain newly authored text in uniform sources.
Native final-state controls qualify deletion, reordering, duplication and mixed
retained/authored operations. Western font size, bold and italic use their own
fields; right-to-left fields do not overwrite them. Mixed-script rendering and
repair-free opening remain unqualified.

The `sourceEditing.nonuniformTextReplacement` record extends count changes to
text replacement from nonuniform native sources. Complete editable paragraph/run
formatting stays unchanged; exact retained selections take priority, and all
format-matched templates must agree on complete native definitions. Empty or
unsupported matching bodies prevent selection. Twenty ordinary/first-final-cell
outputs cover growth, shrinkage, reorder, duplication, sole paragraphs, equivalent
templates, retained empties and language selection alongside rewritten text.
Forty selected-page and 520 neighboring-page stage comparisons match independent
native controls. This does not qualify simultaneous formatting changes,
mixed-run or empty-source replacements, general row reflow or repair-free opening.

The `sourceEditing.mixedRunTextReplacement` record qualifies multiple ordered
runs within that replacement contract. All runs stay nonempty and retain their
complete editable formatting; every source run must map to a uniform complete
native definition, with agreement across all matching templates. Twenty native
outputs cover three visible styles per paragraph, different old/new run lengths,
paragraph reorder/duplication and hidden middle-run language ambiguity. Forty
selected-page and 480 neighboring-page stage comparisons match independent
native controls. Native resaving preserves effective styles and language at
every run/CR boundary. Run splitting/merging, empty-source refill and simultaneous
formatting rewrites remain outside this profile.

The `sourceEditing.retainedTextFormatting` record qualifies formatting changes
during nonuniform count changes through retained ordered run text. Exact
references/values take priority, then retained text, then unchanged formatting
for text replacement. Ambiguous text cannot fall back to another paragraph's
native definition merely because its style matches the request. Existing
character and paragraph codecs rewrite only supported fields; destination CR
typing, opaque metadata and per-cell boundaries remain retained.

Twenty-six native outputs cover size/emphasis/color/font/double underline and
paragraph layout alongside deletion/reorder/duplication, including destination
style collisions, hidden language ambiguity and source font/color pool growth.
Fifty-two selected-page and 768 neighboring-page stage comparisons match
independent native controls. Requests preserving neither run text nor complete
formatting, run splitting/merging and empty-source refill remain outside this
profile; broader native layout and repair-free opening remain unassessed.

The `sourceEditing.nonuniformEmptyRefill` record qualifies single-empty-run
refill with unchanged complete formatting during nonuniform count changes.
Complete source CR typing supplies the body and remains retained at the
terminator. Matching native definitions must agree; hidden language conflicts
still refuse output. Fourteen native outputs produce 28 selected-page and 280
neighboring-page stage comparisons against independent controls. The Unicode
control is separate mixed-script evidence; broader producer/reflow and repair-free
acceptance remain unqualified.

The `sourceEditing.retainedTextRunSegmentation` record qualifies split/merge
edits with complete paragraph text and preserved fields retained at each position.
Complete native paragraph/body templates must agree; each source/requested/native
character-range intersection goes through the existing formatting codec. Opaque
fields stay at their source text positions and destination CR roles remain
retained. Same-count ordinal bindings and the established uniform single-run
path keep their respective contracts.

Thirty native outputs cover ordinary stories and first/final table cells, with
60 selected-page and 1,140 neighboring-page stage comparisons against independent
controls. Public reload compares complete text and per-position styling because
native boundaries can split visibly identical runs. Changed full text plus new
formatting, unsupported preserved fields, labels, broader producer/mixed-script
profiles and repair-free acceptance remain unqualified.

The `sourceEditing.tableTrackGeometry` record covers saved table endpoints and
native row-growth controls. The reader recovers laid-out tracks from cumulative
coordinates, retaining reported nominal fallback for incomplete axes. Independent
short/wrapped label fixtures include native row-height and column-width getters.
They also reproduce the wrapped specimen's native open/resave drift: twelve
stage-matched PDF pixel/span comparisons qualify corresponding native and managed
outputs without requiring Publisher's two stages to render identically. Source
growth/shrink edits retain original stored tracks and native metadata; current
native reflow and managed cell styling/text metrics remain distinct limits.
