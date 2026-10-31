# Publisher support contract

OfficeIMO.Publisher reads native publications, supports bounded in-memory edits
and converts the current model. Bounded native source saving supports exact copies
and the documented edit profiles with operation-level fidelity reports. Its public model owns
publication pages, master definitions, placed pictures, supported vector artwork,
styled source stories, native table grids, embedded assets
and a recovery report. Shared `OfficeDrawing` owns the reconstructed visual scene;
the optional PDF package consumes that scene without another native decoder.

## Native profiles

| Profile | Read behavior | Evidence boundary |
| --- | --- | --- |
| Publisher 2002-and-later Contents/Quill/OfficeArt generation (`E8 AC 2C 00`) | Bounded structured decoding and positioned reconstruction | Apache POI Simple, Sample and Sample_2010 publications; brochure and newsletter corpus fixtures |
| Earlier Contents generation (`E8 AC 22 00`) | Explicit unsupported-profile exception; the shared signature and Quill presence do not identify the exact 97/98 or 2000 release | Apache POI Sample2000 and Sample98 fixtures |
| Encrypted, damaged or other publication generations | No decryption or salvage fallback | Required streams, lengths, references and encodings are validated; unknown versions are rejected |

The generation signature does not establish qualification for every Publisher
release or object type. Independent producer files exercise real native storage;
they are not produced by an OfficeIMO writer.

Two independent native controls qualify twelve whole-frame character-attribute
combinations and all eighteen automation underline choices across five physical
pages. Saved Quill underline values differ from the automation enum. Seven native
kinds map directly to shared patterns; the remaining eleven retain an explicit
approximation. The `characterReading` section of the
[qualification record](../Build/Publisher/WriterProbe/qualification.json) identifies
the fixtures, native values and reader source. Managed SVG/PDF inspection confirms
that the recovered decorations reach output; it does not establish matching native
decoration geometry, font availability or script metrics.
Both controls exercise two chained Quill directories. Separate source-save
qualification below covers character edits in these layouts.
Source-built PSWriteOffice converts both controls through `ConvertTo-OfficeDocument`
and the in-process `publisher-pdf` route on PowerShell 7. The five output pages
reopen, and all eleven underline approximation locations survive in workflow
evidence. This qualifies the conversion path, with packaging and native writing
remaining separate boundaries.

Inline OfficeArt picture data stays in the preserved source and reports omitted
artwork; positioned reconstruction uses image-store references. Cleared picture
references do not report missing images.
Native resaving can retain zero-count image-store descriptors whose delayed
offsets now point at another image. After collecting native BLIP references,
the reader leaves an unreferenced zero-count truncated or UID-mismatched slot
inert and reports `PUB_IMAGE_RETIRED_SLOT`; its identity remains reserved.
Valid unused assets remain available. Referenced or positive-count truncated
payloads still fail the read. UID comparison identifies retired slots; it is
not an integrity check over every active payload. Inspected retired payloads
still consume the image-processing byte budget.

## Reconstruction

| Content | Current behavior | Limit |
| --- | --- | --- |
| Document pages | Native order, physical size and page clipping, including empty physical pages | Master and utility roles follow the ordered native reference array and its role boundaries; malformed boundaries are rejected |
| Master pages | Recovered definitions applied to referring pages | Missing references report omission; nested master inheritance is unassessed |
| Text | UTF-16 stories, styled paragraphs, referenced native style values, explicit reciprocal links or complete native story/frame catalog bindings, sequential columns, continuation and overflow reporting | Shared measurement approximates native frame breaks and column balancing; source style pairs and use ranges are retained internally, while names, builtin selectors and auxiliary inheritance metadata remain unassessed |
| Character formatting | Native font names, size, referenced color, bold/italic, strikethrough, superscript/subscript and none, single, words-only, double, dotted, dashed and wavy underlines | Native underline weight, long-dash, compound and unknown patterns use the closest shared pattern with `PUB_TEXT_UNDERLINE_APPROXIMATED`. Font availability, exact metrics and decoration geometry remain separate rendering limits |
| Lists and tabs | Native Unicode bullet labels, qualified Symbol/Wingdings marker normalization, hanging indentation, text position and declared left tabs; complete tabbed lines follow recovered center/right paragraph alignment | Numbering sequences, additional per-stop alignments/leaders and drop caps are unassessed; undeclared tab stops use the shared 36-point interval. Native font metrics and exact tabbed-line positions remain separate rendering limits |
| Tables | `PublisherPage.Tables` exposes current tracks, spans, styled paragraphs, text and placement; qualified anchor-cell text edits preserve complete story partitions, and qualified bounds edits scale tracks proportionally; master tables remain on their master | Separate `CanEditCellText`/`CanEditBounds` and limitation properties identify supported operations. Unavailable frame geometry retains the grid and mapped text with `HasPlacement = false`; unresolved text mappings use `HasTextMapping = false`; individual cell borders, fills and padding are unassessed |
| Shapes | Rectangles, rounded rectangles, ellipses, diamonds, triangles and lines; solid paints and basic shadows | Other geometries and unsupported fill types report approximation; compound and non-solid strokes remain unassessed |
| Custom paths | Literal eight-byte signed and compact unsigned vertices in declared geometry space; bounded SG guide formulas and backward references; implicit open/closed line or cubic paths, explicit move/line/cubic/close/end commands, subpaths and no-fill/no-line controls | Device-pixel guide operands, limousine scaling, advanced commands and separate paint groups retain an explicit fallback; guide rounding, shared nonzero winding and native appearance remain unqualified |
| Linear fills | Native color stops, fixed-point angles, focus ramps and foreground/background opacity for fill types 4 and 7 | Custom anchors, native shading corrections and unrepresentable alpha ratios report approximation; native Publisher rendering remains unqualified |
| Line details | Native caps, joins, miter limits and dash order; triangle, stealth, diamond, oval and open-arrow ends on lines | Dash spacing and marker dimensions use stroke-relative approximations; unknown values report loss; marker shadows and decorations on unsupported open geometry are omitted with diagnostics |
| Groups | Native child coordinate spaces, non-uniform scale after child rotation, nested rotation/reflection and hidden-descendant suppression | Missing anchors report unresolved transforms; comparison against native Publisher rendering remains unqualified |
| Text wrapping | Native frame exclusion references and object wrap distances; transformed rectangular exclusions | Widest available interval per horizontal band; side selection, tight/through outlines and native font metrics can differ |
| Pictures | Embedded and delayed OfficeArt payload extraction, Publisher GIF envelopes, positioned images, source crop and supported custom-path masks | WMF/EMF use a supplied application codec or reported placeholder; native custom-mask appearance remains unqualified |
| Placed-picture model | `PublisherPage.Pictures` separates page/master ownership, stable object IDs, asset references, page-point bounds, transforms and normalized crop from embedded bytes; unsupported placements expose an edit limitation | Unresolved image references can remain unresolved; inspection does not qualify native save-back or editable effects |
| Picture creation | `AddPicture` appends an owned static PNG/JPEG asset and uncropped rectangular placement on a qualified document/master page, with explicit point bounds, no fill/outline/effects/transform or story/wrap references; shared object and separate image allocators, current exports | Non-normal EXIF orientation, animated PNG, other formats and invalid content reject creation. No automatic DPI sizing, crop/effect authoring, asset sharing API, external links or native PUB records; existing replacement eligibility remains unchanged |
| Picture controls | Source-RGB transparency keys, encoded-RGB brightness/contrast, preserve-grays tone handling, BT.709 grayscale and a 50-percent two-color threshold; original assets remain available | Color space, effect ordering and native Publisher pixels remain unqualified; recoloring and extended color controls are unassessed; invalid tone values or unresolved keys report omission |
| Fields, links and active content | Cached story characters are retained; source active content stays inert | No field evaluation, hyperlink reconstruction, macro execution, link refresh or object activation |
| Editable vector artwork | `PublisherPage.Shapes` records owner, stable ID, point bounds and placement availability; bounds and solid-paint edits for supported basic geometry; shared text-frame placement and current wrap obstacles | Grouped, transformed, hidden, custom-path, gradient, effected and unresolved artwork is inspection-only; detached editable descriptors do not establish native writable support |
| Primitive creation | Rectangle, rounded rectangle, ellipse and nondegenerate line descriptors append to qualified document/master pages; shared publication allocator and current SVG/PDF scenes | No native records; unsupported geometry/effects are rejected; new shapes have no story or wrap references |
| In-memory edits | Document-owned replacement of complete non-table story paragraphs and mapped table-cell paragraphs, bounds changes for qualified frames/tables, and image replacement/bounds changes for supported rectangular PNG/JPEG picture placements; stable IDs, ownership and order, fresh scenes, text ranges and overflow | Picture and table bounds edits reject unsupported graphic placement. Cell edits reject conflicting grids, unsafe partitions and stories also used by ordinary frames; no table topology/styling changes or arbitrary relinking. Native saving accepts only the separate source-edit profile; old references remain snapshots |
| Table creation | Empty unmerged borderless grids on qualified document/master pages; copied positive finite tracks, one new complete story, shared table/story identifier allocation and current SVG/PDF/Reader projection | No merged-cell authoring, cell decoration or native PUB records; text/model/projection limits and transactional failure behavior apply |
| Page/master authoring | Source-free `Create` with one empty physical page; appended same-size physical pages and centered single-page masters; explicit qualified direct master assignment or clearing; immutable native-source/page provenance | No resize, deletion, reorder, object moves, paired-master creation or master inheritance; unassessed backgrounds and registrations reject assignment; these changes cannot be saved as native PUB |
| Ordinary text creation | `AddTextFrame` creates a complete independent styled story and resolved one-frame chain on a qualified document/master page; one column, zero gap/insets, top alignment, no paint/transform/wrap; existing continuation API handles document-page chains | Authored ordinary stories have no implicit final terminator; full overflow is retained. Frame/story IDs reserve atomically; master-owned stories cannot continue across physical pages; no native text records are authored |
| Text-frame layout edits | `SetTextFrameLayout` replaces equal-width columns/full intercolumn gap, Core padding and Top/Center/Bottom alignment on qualified resolved document/master frames, including mixed text artwork. Imported frames can use the bounded native source-save profile; snapshots expose operation-specific eligibility | No table stories/aliases, grouped/transformed/unresolved placement, native balancing, distributed alignment or vertical writing. Unqualified native layout overrides remain read-only for this operation; bounds/read behavior stays separate. Native saving requires exact distances and safely editable local properties/chunk-directory records; fidelity remains Unassessed |
| Continuation-frame creation | Append an ordinary untransformed text frame to a resolved non-table story's tail on an existing document page; one complete styled story, current links/ranges/overflow, copied columns/padding/alignment, no fill/outline, and shared SVG/PDF/Reader projection | The source chain must contain ordinary visible untransformed document-page frames; master chains, ambiguous links and non-tail insertion are rejected; this is in-memory model creation, not native PUB creation |
| Native source saving | `ToPubBytes(Result)`, `Save` and `SaveAsync` with detached bytes, explicit `.pub` paths or caller-owned streams; exact unchanged source copies and bounded existing-source edits described below | Default strict loss policy blocks unassessed edits. Explicit `Allow` does not enable refused profiles. No source-free writer, general image/table/page topology changes or earlier-generation output |
| Rendered output | SVG per document page and multi-page PDF through shared engines | Reconstructed output carries source and rendering losses independently of native saving |

`OfficeIMO.Workflows` uses the same native codec for its `publisher-pdf` route,
file batches and document previews. Source recovery diagnostics survive output
publication and checkpoint reuse. Strict acceptance and resource-limit failures
preserve existing destinations.

`OfficeIMO.Reader.Publisher` projects complete stories once, mapped native table
rows, bullet markers, physical page inventory and optional original image bytes.
Ordinary stories use paragraph blocks; mapped table stories use one complete
table block alongside their structured rows and native page citation. Generic
column names preserve the first row as data. Merged spans flatten to one anchor
value with approximation evidence; row limits retain total counts and omission
reports. Master tables are extracted once without a physical-page citation and
use document-wide export ordinals. Repeated objects sharing a story retain
distinct table anchors; the complete story correlates with its first mapped table.
Shared-story chunks carry a page citation only when all occurrences agree,
preserving each table's native citation during extraction and JSON transport.
Unresolved cell text remains in complete stories rather than an empty dataset.
Native losses survive Reader transport and semantic PDF projection, which
correlates table blocks and rows to avoid repeated text. Typography, page artwork
and frame flow remain explicitly omitted from the semantic projection. The same
adapter is included in `OfficeIMO.Reader.All`.

`SourceObjectCount` measures the original native descriptors.
`CurrentObjectCount` includes retained imported objects and authored
tables, continuation frames and shapes; `ProjectedObjectCount` counts distinct current objects
represented in document or master scenes. Repeated master use counts once.
Authored tables, frames and shapes do not add native inspection records, and `SourceReadReport`
retains the original decoding evidence. Counts do not measure pixel fidelity
or prove that all content of an object survived. Complete stories remain
inspectable even when their visual placement is incomplete.

Every successful edit replaces the current projection and its `ReadReport`.
`SourceReadReport` remains the immutable decode-only evidence from the initial
load. Validation, cancellation and projection failures leave the current model
unchanged. Native bytes and compound metadata are retained privately and remain
inert. The import-time semantic baseline and exact source OfficeArt bindings
support the native source-save path: unchanged content copies the
original bytes; bounded existing anchor and solid-scalar edits replace the
OfficeArt stream. Frame-layout edits can grow local property tables and Contents
shape records while relocating dependent native offsets. Admitted ordinary-story profiles can
replace text of different
lengths and edit supported character attributes through retained per-run OPLs, Quill text, story lengths,
formatting endpoints and page indexes. A single original run can split into
styled runs, and matching run counts can be edited again. Each changed story has
one existing frame or a complete resolved chain of ordinary, visible,
untransformed document-page frames. Linked-story saving retains the native
chain, frame identities, page ownership and geometry; it does not relink frames.
Unresolved, unavailable, grouped, transformed and master-page chains refuse
changed-story output. Story identities and order remain fixed. Known source layouts
can have chained directory pages, with at most 20 descriptors per page and 65,535
descriptors in total. Existing table stories use the fixed-cell profile below. Repeated
byte-identical OPLs can share native page storage, with all story lengths, range
endpoints, page indexes and later-page start headers remapped. The original
single-story/one-page path retains its accepted encoding. In multi-paragraph stories, existing empty paragraphs can remain empty or
gain text. Each paragraph's
native definition and endpoint are retained independently; character ranges can
cross paragraph boundaries. Empty paragraph typing follows the source CR or an
explicit empty styled run, as described below. Unsupported nonempty run merges,
unsupported formatting and unknown directory metadata refuse
output. Edited output removes the optional [SummaryInformation thumbnail](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/f7933d28-2cc4-4b36-bc23-8861cbcd37c4) so it
cannot retain that stale content preview. The shared OLE property-set editor
retains the other raw property values and header; unrelated logical streams and
directory metadata remain unchanged. An unchanged semantic model still copies
the original native bytes exactly, including its preview. Malformed
SummaryInformation property-set structure refuses edited output. Other cached-record semantics
remain unassessed.
Unsupported semantic changes refuse output. This path does not establish
repair-free opening, general native edit acceptance or opaque semantic preservation.
The [writer probe](../Build/Publisher/WriterProbe/README.md#loaded-source-candidates)
records the independently qualified source-edit specimens.

Paragraph counts can increase or decrease from a uniform source: every source
paragraph has one run, matching projected formatting, byte-identical native
paragraph definitions and matching body/intermediate-CR definitions. Added
paragraphs use those definitions as a base, with qualified character and paragraph
edits applied independently; the final source CR definition is retained.
New runs can change size, bold, italic, representable underlines, single
strikethrough, baseline, font family and opaque RGB text color. Unchanged
attributes retain their source tokens; explicitly changed attributes also reach
the paragraph CR. None, single, words-only, double, dotted, dashed and wavy
underlines are writable. Double strikethrough and nonopaque colors are refused.
Font changes require a known uniform eight-script source selector or no direct
selector; mixed-script source mappings are refused. Names must contain 1–32
UTF-16 code units, with no control characters or malformed encoding. Existing
FONT entries and suffixes remain byte-identical; new names append indexed entries
with zero hints. Font availability and substitution remain native-environment limits.
Color edits require an absent or ordinary paired solid-color representation.
They update both color references and retain existing palette entries; ordinary
RGB entries are reused or appended. Other fill, opacity or ambiguous source
representations refuse output.

Paragraph saves support left, center, right and justified alignment, spacing
before/after paragraphs, left/right margins, first-line or hanging indentation,
and explicit absolute or multiple line spacing. Point values must be exactly
representable in the admitted native units. Use `lineHeightFactor: 1` for explicit
single spacing; clearing both line-height settings does not establish a native
default. Simultaneous nonzero first-line and continuation offsets are outside
the canonical native indent profile.

Ordinary explicit left tabs can be added, replaced or cleared. The default
interval is 36 points and the origin is `-paragraph.Margins.Left`. Native center
and right paragraphs require `TabStops.AlignWithParagraph` to be true, using
`WithParagraphAlignment()`; this includes paragraphs that currently contain no
tabs so later text insertion retains the same behavior. Left and justified
paragraphs use false. A fixed-grid center/right request is refused. Custom tab
origins, repeating intervals, per-stop alignment and leader settings are outside
this write profile. Existing unknown tab metadata is retained when only paragraph
alignment changes. Native field widths and formatting-page capacity remain
enforced.

Nonuniform source definitions refuse count changes, even when their projected
formatting looks equal. An empty final paragraph is refused when it would replace
a different native typing definition. A zero-paragraph story remains unsupported
for native saving; in-memory clearing and rendering remain available.

When ranges no longer fit the original allocation, the codec can append aligned
formatting pages and directory pages, relocate enlarged BTE indexes and resources,
and retire unused pages. Superseded known pages/indexes and relocated resource
payloads are cleared, while unrelated payloads and gap bytes remain retained.
Directory growth uses wholly zero unused descriptor slots or appends a page;
unknown padding is retained. The established single-directory allocation path
still refuses nonzero unassigned header bytes, and unknown page-descriptor
metadata is refused. Directory
capacity, output-byte limits and cancellation remain enforced.

The `sourceEditing.characterEditing` qualification record binds 24 public-save
specimens to the writer source. They cover adding and clearing attributes,
mixed runs, existing and new font/color entries, linked overflow, paragraph-count
changes, chained source layouts, directory growth and retirement, and re-editing
native-resaved output. Independent decoding verifies complete native-resaved text,
body attributes and explicitly changed CR attributes. Existing resource entries,
unrelated OPL tokens, logical streams and metadata remain byte-identical. Native
PDF pixels and spans match across save/reopen over 58 physical page states.
Eighteen earlier linked-story and paragraph/topology saves remain byte-identical.
Source-built PSWriteOffice converts four edited publications through its existing
in-process `publisher-pdf` route; all five output pages reopen. This qualifies
these specimens and operations, with repair-free opening, general opaque semantics,
packaged consumers and managed/native rendering parity still separate.

The `sourceEditing.paragraphFormatting` qualification record covers 21 public
save cases across 62 native physical page states. It includes per-paragraph
values, clearing tabs, resetting a formatted chained source, linked overflow,
uniform-source count changes, fractional point values, and 240 distinct paragraph
formats that allocate formatting and directory pages. Independent decoding of
style-pair inheritance and direct paragraph records verifies the native-resaved
layout values, complete text and retained body/CR character attributes. Unrelated
tokens, resources and storage remain unchanged; native PDF pixels and spans
match before/after resave. Forty-two earlier saves remain byte-identical.
Source-built PSWriteOffice converts four paragraph-edited publications through
`publisher-pdf`, with eight reopened output pages.

Frame-layout saving retains frame/story identities, links, page/master ownership,
unrelated properties and complete text. Columns, full intercolumn gap and insets
use the local OfficeArt property tables; vertical alignment uses the Contents
shape record. Missing default fields can be added. Unrelated rows, complex-data
tails, native separators and chunk payloads are retained, and enclosing lengths,
chunk-directory offsets and mirrored header pointers are updated when records
grow. Existing bounds and solid paint edits can be combined with layout edits.
Changed distances must have exact nonnegative signed-EMU representations
(12,700 EMUs per point). Publisher's automation
[`ColumnSpacing`](https://learn.microsoft.com/en-us/office/vba/api/publisher.textframe.columnspacing)
is half the full gap used by this API and the native OfficeArt property. Ambiguous local overrides,
flagged scalar targets, incomplete complex payloads and unqualified Contents
relocation headers refuse output, including with `Allow`.

The `sourceEditing.frameLayout` qualification record covers eleven public-save
specimens over 44 native physical page states. They include default-field growth,
fractional insets/gap, multiple frames, linked overflow, master frames, mixed
text artwork, reset and combined text/bounds/paint operations. Independent native
getters and raw decoding verify layout after resave; complete native Quill text,
unrelated records, streams and compound metadata remain retained. Native PDF
pixels and spans match before/after resave. Sixty-three earlier saves remain
byte-identical. Source-built PSWriteOffice converts five edited publications
through `publisher-pdf`, with twelve reopened output pages. Native column
balancing, managed/native rendering parity, repair-free opening and broader
producer/profile coverage remain unqualified.

Existing editable rectangular picture placements can save a replacement PNG or
JPEG through `ReplacePictureImage`. The native writer retains original image
payloads, changes only the selected placement's BLIP binding and filename, and
appends owned image slots and filename resources. Shared source assets, other
placements, master ownership, crop values and bounds remain retained; bounds
and qualified frame-layout changes can be combined with replacement. Repeated
edits can leave gaps in document-scoped asset IDs; saving maps new assets to
native slots and reports renumbering with `PUB_NATIVE_IMAGE_BINDINGS_RENUMBERED`.
An unresolved source reference cannot be satisfied accidentally by a newly
appended slot. Ambiguous picture properties, filename ownership, allocation
counters, image stores or reference-count underflow refuse output even with
`Allow`.

The `sourceEditing.imageReplacement` qualification record binds eight public
saves over 30 native physical page states: PNG/JPEG, one/all shared placements,
master ownership, repeated edits, combined frame/bounds edits, cropped pictures
and re-editing a Publisher-resaved file with retired image slots. Independent
decoding verifies image pixels, filenames, reference counts, retained source
payloads and crop fields; native getters verify physical bounds. Native PDF
pixels and spans match before/after resave. Seventy-four earlier saves remain
byte-identical. Source-built PSWriteOffice exercises the same edited files
through `publisher-pdf`. Independent native qualification covers the delayed
image stream; the embedded-BLIP append branch without that stream remains
unqualified. Cropped raster-matrix rounding has a measured residual of about
0.0074 points against the continuous model, so exact managed/native crop
rendering parity remains unqualified. Repair-free opening and broader producer
coverage also remain unassessed; changed native image output reports
`PUB_NATIVE_IMAGE_REPLACEMENT_UNQUALIFIED`.

Native table-cell text saving retains existing tables, tracks, merged ownership,
cell order, stored bounds and paint. Each cell retains its final terminator
convention. Paragraph counts can change when projected and raw source definitions
are uniform within that cell; neighboring cells may have different styles.
Supported run/paragraph formatting uses the same Quill
writer as ordinary stories; cell partitions use separately validated `TCD `
records. Their descriptor uses flag 1 and secondary kind `PLC `; the payload
has the fixed cell count, retained header and story-relative UTF-16 endpoints.
Nonfinal endpoints include their terminating character; the final endpoint is
exclusive. Source endpoints must exactly match the recovered cell partitions.
After editing, endpoints, complete story lengths, formatting pages and directory
locations are rewritten together. Ordinary stories in table-bearing documents
can also change length without changing relative cell ownership.

Unknown partition headers/descriptors, unresolved or conflicting ownership,
table stories also used by ordinary frames and count changes from nonuniform
definitions within the edited cell
refuse edited output. Table creation,
bounds/grid changes, cell merges and table styling have no native save profile.
These limits apply with `Allow`; unchanged content still returns an exact copy.

The `sourceEditing.tableText` qualification record covers six public saves over
twelve native document-page states: first/final/all cell replacements, bold/italic
retention, a preceding title-story length change, combined image replacement and
re-editing a native-resaved table. Independent decoding verifies complete stories,
cell endpoints and direct formatting; native PDF text occupies the original cell
positions. Native table bounds and before/after PDF pixels/spans remain equal.
The TableLabel control retains Publisher's multiple-label print layout in its
native PDF; its document page and table geometry are recorded separately. Native
PDF cell sizes retain the source's 9.984-point spans for nominal ten-point text;
exact managed/native font metrics and rendering parity remain unqualified.
Broader table profiles and repair-free opening remain unassessed. Changed Quill
output reports `PUB_NATIVE_QUILL_TEXT_UNQUALIFIED` and is blocked by default.

Empty existing paragraphs and fixed cells can save without requiring byte-equal
body/CR definitions. A paragraph with no runs leaves typing unspecified: the
native CR definition is retained byte-for-byte. A supplied empty styled run
requests its font, size, color and supported decorations. The canonical Quill
reader resolves the actual CR inheritance and original palette in an
operation-local context before applying those changes. Unrelated CR fields remain
retained; no meaning is guessed for body-only metadata removed with its text.
Native reload represents an empty CR as one styled empty run, including when the
input paragraph supplied no runs.

Clearing a cell with an empty paragraph collection keeps one empty paragraph and
the final paragraph's alignment, spacing, margins, indents and tabs. Other cells
and the final terminator convention remain retained. Supported uniform-source
paragraph-count changes can also end in an empty paragraph. Removing every
paragraph from a native story remains unsupported, and table grid changes still
have no native save profile. Unsupported CR encodings remain
refused when an explicit typing change would require rewriting them.

The `sourceEditing.emptyTyping` qualification record covers fifteen public saves
over forty native document-page states: first/final/all cell clearing, explicit
and unspecified typing, sources whose body and CR have different inherited
styles, single/multiple/trailing empty paragraphs and visible refill after native
resave. Independent STSH/formatting/resource decoding verifies effective typing,
exact retained CR and paragraph bytes, complete stories, cell endpoints and
frame bounds. Native refill text confirms font, bold/italic and color in the
target region; native initial/resaved PDFs have identical pixels and spans.
Eighty-eight earlier saves remain byte-identical. Seventeen native control
profiles separately establish early-range and fresh whole-cell behavior. Their
language setter changes field `0x12`; the exact role of body-only `0x3E` remains
unassessed. Source-built PSWriteOffice converts all fifteen specimens through
`publisher-pdf` with Unicode fallback settings. Repair-free opening, broader
producer coverage and exact managed/native rendering parity remain unqualified.

The `sourceEditing.cellParagraphs` qualification record covers eleven public
saves over twenty-two native document-page states: first/final/all-cell growth,
native multi-paragraph shrinkage, balanced count changes, a paragraph moved
between cells without changing the complete styled story, empty typing and
visible refill, preceding-story edits and re-editing a native resave. Each cell's
formatting is planned from its own source ranges, including its final CR.
Independent decoding verifies cell endpoints, paragraph definitions, effective
typing and unrelated storage. The fitting specimens retain native geometry and
have identical initial/resaved PDF pixels and spans; 103 earlier saves remain
byte-identical. Four independently authored native table profiles supply the
multi-paragraph controls.
Source-built PSWriteOffice converts all eleven specimens through `publisher-pdf`
with Unicode fallback settings and recovers each nonempty cell paragraph.
These results do not qualify arbitrary text fitting:
Publisher can expand rows when new text wraps or exceeds a cell, even though the
stored grid is retained. A separate wrapped-text cohort records changed native
heights and resave rendering; fixed rendered heights and overflow layout remain
unqualified.

Clearing a native whole-frame `TextRange.Text` to an empty string saves one CR,
with the story and frame retained. This control does not establish a zero-length
or zero-paragraph native story encoding; removing all semantic paragraphs still
refuses native output.

Two independent reading controls cover 17 whole-frame profiles over five pages,
including indexed arrays with two tabs and left/center/right tabbed lines. The
reader accepts canonical ordinal tab entries and retains the older repeated-zero
profile. Native/managed inspection verifies the recovered alignment behavior;
the controlled horizontal word-start measurements retain a maximum 0.148-point
residual with native ArialMT and managed Helvetica. That measurement does not
qualify exact font metrics, vertical positions or general rendering parity.

Public native saving prepares and bounds the complete candidate before destination
publication. `PublisherSaveOptions` defaults to a 64 MiB ceiling, `Block` loss
policy and `FailIfExists` file conflicts. `Block` also rejects unassessed fidelity;
`Allow` permits the documented candidate profiles but never bypasses preparation
refusal. `PublisherSaveReport` separates save diagnostics from its original
`SourceReadReport`, reports prepared byte count and identifies exact source copies.
Saving neither changes the semantic model nor rebinds its original source baseline.

Path saves require `.pub` and use shared strict atomic publication, including
explicit `Replace`. Existing destinations survive preparation, limit, policy and
pre-publication cancellation failures. Streams remain caller-owned; seekable
streams are replaced/truncated/rewound, while non-seekable streams receive bytes
at their current position. Preparation failure preserves bytes and position;
publication-time stream I/O or cancellation can leave partial output. Async saves
prepare synchronously before asynchronous destination I/O. Source-free models and
unsupported changes fail with a save report and no prepared bytes. Byte limits
and cancellation raise their own exceptions. Exact copies preserve source content
inertly and do not newly qualify active content, signatures or protected storage;
edited output refuses unsafe carriers.

Two specimens saved through a compiled public consumer qualify rectangle bounds
and different-length single-story text with 18-point bold formatting. Byte,
sync/async file and sync/async stream routes produce identical candidates. Both
pass native open, PDF export, save-as and reopen; native inventories retain the
requested content, and native PDF pixels match across resave. Independent compound
inspection preserves unrelated logical streams, directory metadata and other raw
SummaryInformation values. The `publicSaving` section of the
[qualification record](../Build/Publisher/WriterProbe/qualification.json) identifies
the exact source, codec and artifact hashes. This evidence does not establish
repair-free opening or broader native editing.

Eight multiple-story/page specimens qualify separate first, middle and last story
edits, simultaneous edits, 96-paragraph growth/shrinkage across two paragraph-formatting
pages, and re-editing after native resave. All pass native open, PDF export,
save-as and reopen. Complete stories, including overflow, survive in Publisher's
resaved Quill storage; native frame inventories expose the visible prefix.
Native PDF pixels match across resave, and independent checks retain paragraph
OPLs, other character tokens, unrelated Quill payloads, logical streams and
directory metadata. The `multipleStoryEditing` qualification record identifies
these exact specimens and their scope.

Seven linked-story specimens from the native two-page newsletter qualify
same-count growth/shrinkage, paragraph-count increases/decreases, formatting-page
allocation/retirement, size/bold changes, a neighboring title edit and re-editing
after native resave. Complete stories and body size/bold values survive in
Publisher's resaved Quill storage, including overflow. Native inventories retain
the two-frame chain, physical-page ownership and geometry; visible frame text is
checked in chain order. When the complete story fits the first frame, Publisher
reports a lone paragraph CR for the empty trailing frame. Native PDF pixels and
text spans match before and after resave across 28 physical page states.
Independent storage comparison retains native Contents, artwork, unrelated
logical streams and Quill payloads, directory metadata and other SummaryInformation
values. The `linkedStoryEditing` qualification record binds these exact specimens
to their source and codec hashes. Other producers and chain profiles, repair-free
opening and general opaque semantics remain unassessed.

Native client anchors can omit zero-valued coordinates. The independently
produced ZeroAnchors control qualifies all four coordinate positions: the reader
recovers all four text frames with their native point bounds. This does not create
missing physical slots for native bounds patching; that operation still requires
the complete original slot set.

Eleven paragraph/allocation specimens qualify paragraph append/removal, 240-paragraph
growth and retirement to one paragraph, neighboring-story preservation, distinct
character styles, shared styled paragraphs, simultaneous character/paragraph index
growth, and two re-edits after native resave. Native-resaved files retain complete
text and body size/bold values; native pixels and text spans match across resave.
The `paragraphTopology` qualification record separates these public operations
from the three preceding page-relocation experiments. Repair-free opening and
general opaque semantic preservation remain unassessed.

Six exact single-paragraph specimens of this profile have native
read-only open, PDF export, save-as and read-only reopen evidence in Microsoft
Publisher 16.0.20527.20088. They include OfficeIMO re-editing a native-resaved
styled file. Native text, frame bounds and size/bold states match the requests
within the recorded 0.1-point font-size tolerance; native PDF pixels match before
and after resave. These observations do not establish repair-free opening,
native text mutation, other profiles or opaque-record semantics. The
[qualification record](../Build/Publisher/WriterProbe/qualification.json)
keeps source provenance, exact file hashes and operation boundaries.

Five multi-paragraph specimens also pass native open, PDF export, save-as and
reopen checks: growth, shrink, filling an empty paragraph, size/bold changes
including a styled empty paragraph, and OfficeIMO re-editing after Publisher
resaves. Native text, frame bounds and visible size/bold states match the requests
within 0.1 point, with identical native PDF pixels across resave. These observations
qualify the listed specimens, not general paragraph editing, layout or repair-free
opening.

Four thumbnail-free edited specimens cover text sizing, rectangle bounds/fill,
master-picture placement and multi-paragraph formatting. All pass native open,
PDF export, save-as and reopen checks. Native inventories and rendered pixels
match the earlier content-qualified specimens across ten physical page states.
Publisher recreates a thumbnail on native resave in these cases; its appearance
is not decoded or qualified by the raw-property comparison. OfficeIMO invalidates
the old thumbnail without generating a replacement.

Nine additional placement/paint specimens cover a rectangle's bounds and solid
fill, a physical text frame, default and paired-master frames, and physical/master
picture bounds. Native inventories and PDFs retain the requested values across
19 physical pages before and after native save/reopen; unaffected object geometry,
text and shared image pixels remain unchanged. The two picture edits retain the
original Contents AA/AB extent fields; their general semantics remain unassessed.

A separate native object-model check replaces the whole text range of the styled
specimen's sole unlinked text box. The exact text survives native save/reopen and
PDF export, while frame/master state and original source bytes remain unchanged.
This qualifies one controlled text replacement, not character-level editing,
manual UI editing, general formatting changes or repair-free opening.

Empty paragraphs use the character formatting on their native paragraph CR,
including its font family and size, rather than the publication's global defaults.
The native TextParagraphs control qualifies this distinction with an Arial
12-point empty paragraph between two text paragraphs.

One retained Quill style catalog supplies the existing formatting resolver.
It keeps native character/paragraph pair ordinals, nullable decoded values,
whole indexed STSH entries and their inner definition spans, including when
the publication has no TEXT descriptor. Original story keys and ordinals bind
FDPP selections and direct FDPC spans in native UTF-16 coordinates, with CR and
NUL retained in that coordinate domain. Public story text normalizes CR to LF
and removes NUL. An absent FDPP selector stays absent even when the resolver
uses pair zero. In-memory edits retain these original bindings; changed-story
preparation reports identify their source dependencies. Unsupported changes
return no candidate bytes.
STSH0/2 entries and STSH1 prefixes/suffixes remain opaque. This catalog does not
expose an editable named-style library or qualify native Quill writing.

`PublisherPage.CreateDrawing()` returns a detached scene for callers
that need render-only annotations. Document SVG/PDF exports use the canonical
publication state, and the Reader extension projects its current stories,
tables, physical pages and assets without serializing or reopening it.

Page/master dimensions, registration and object ownership, picture placement,
complete text stories, table grids/placement, frame definitions/chains, supported shape geometry/paint and embedded asset state are
retained independently of the derived drawing.
Source-free publications start directly from that semantic state. One optional
read-only import-residue owner retains native graphics, bindings and original
storage for loaded documents; creation does not manufacture empty native
Contents, Quill or OfficeArt objects. `HasNativeSource` describes document byte
provenance and `PublisherPage.HasNativeOrigin` describes each definition.
Imported definitions keep their provenance after editing; appended definitions
are authored. Source-free source counts and inspected native records are zero.
Pages and masters share the object/story allocator, but do not inflate
`CurrentObjectCount` or `ProjectedObjectCount`; their counts come from the page
collections. `MaximumPages` applies to document and master pages together.
All supported pages have one publication size. Topology changes preserve the
exact imported registration and coordinate edit envelopes; new loaded-document
pages use the physical page's centered envelope. Source-free geometry requires
finite point coordinates and extents and has no native-generation storage range
claim. Frame/story creation preserves table-story references and shares the
existing semantic text validation, complete-story owner and Core layout.
Modern native bindings and coordinate interpretation are mapped at import;
supported picture edits and rendering consume the typed placement directly.
Text-frame story bindings, link/catalog order, columns, spacing, padding,
alignment and geometry are imported into semantic state once. Bounds edits and
new continuation frames do not synthesize native drawing or text records.
Layout edits replace only semantic frame settings; the existing Core flow
recomputes assigned ranges and complete overflow. A frame's graphics, transform,
owner, story, links and wrap references remain unchanged. Continuations copy the
current tail layout. Padding and gaps must leave a positive finite content area;
an already collapsed area does not prevent correcting the layout. Column count
shares the imported ceiling of 256 and the configured item limit; all existing
measurement and projection budgets remain active.
The imported layout edit family admits the absent defaults and exact ordinary
forms present in the current native controls: `WrapText` value 1 and the
`TextBooleanProperties` word `0x00020000`, with no explicit OfficeArt anchor or
text-flow override. Contents vertical alignment admits an absent field or the
observed four-byte scalar with values 0, 1 or 2. Other text-property forms and
unknown vertical-alignment encodings/values keep their existing render fallback but expose
`LayoutEditLimitation` and `PUB_TEXT_FRAME_LAYOUT_EDIT_UNSUPPORTED`. This is managed
layout qualification; the controls do not establish native column balancing or
exact Center/Bottom metrics. Internal native preparation compares all current
layout values and refuses changed layouts instead of returning unchanged source
bytes as a successful edit.
The one publication object allocator reserves imported identities and dangling
references; failed mutations do not consume an identity.
Vector artwork uses the shared `OfficeShape` descriptor, copied at its public
boundaries. Supported native geometry and paint are mapped by the same adapter
used for retained native artwork; supported projection consumes the current
descriptor directly. Text-capable artwork takes current placement from its
frame owner, including after `SetTextFrameBounds`. Other artwork retains its
private native projection and feature-level losses. Unsupported shape edits add
`PUB_SHAPE_EDIT_UNSUPPORTED` with no additional loss classification.
Horizontal and vertical lines retain zero logical extents independently of the
positive canvas used for rendering. Basic solid paint validates finite widths
and opacities; creation rejects clipping, transforms, effects, dashes, markers
and unsupported geometry. Shape and continuation creation share the same object
and projection limits and leave the allocator unchanged on failure.
Tables import their placement, tracks, cell ownership and frame padding once.
Qualified rectangular frame paint uses the existing shared shape mapper.
The TableLabel, Sample and Sample_2010 controls qualify OfficeArt host type 201
as a rectangular frame only when paired with a recovered native table grid.
This does not admit type 201 as independently editable vector artwork, and
custom paths, effects and unsupported paint still reject table bounds edits.
Unsupported graphics retain the private native projection adapter and reject
bounds edits. Tables remain excluded from independent shape editing. Current
table placement also supplies existing wrap exclusions.
Quill cell endpoints become ordered normalized semantic story segments before
their native offsets are discarded. Every segment retains its terminator
convention, and cell snapshots derive paragraphs from that shared partition.
An edit rebuilds the complete story and all compatible table occurrences
atomically; other cell content and merged ownership remain intact. Conflicting
table grids and stories also used by ordinary frames remain inspection-only
for cell text. Creating a table reserves both its object and story identity
through the publication allocator; its story is counted once and its page
object is counted once. No native records are synthesized.
Other object features retain their bounded imported representation while the
editable model develops. This is not a complete authoring or generation-neutral
native codec contract. Replacing one picture leaves shared-asset siblings
unchanged. Original native assets stay available, while superseded edit-owned
assets are released when no current picture references them. Replacement uses
the shared bounded PNG/JPEG decoder, and projection retains byte, element and
pixel-work limits. Unsupported picture edits report
`PUB_PICTURE_EDIT_UNSUPPORTED` without changing the existing reading or
rendering loss classification.

`PublisherPage.TextFrames` records recovered frame geometry and links. Assigned
UTF-16 ranges refer to normalized `PublisherTextStory.Text`; they are derived
from managed measurement, not claimed as native cached break positions. Later
empty frames have zero-length ranges. Explicit native chains are checked for
reciprocal links, matching story ownership, consecutive ordinals, cycles and
disconnected frames. When explicit links are absent, complete native catalog
bindings can establish a chain only when their declared frame count, object
identities and consecutive ordinals agree with the source frames. Contradictory
catalog evidence is rejected. Without complete catalog evidence, ambiguous
unlinked roots report omission instead of guessed ordering.
A missing printable frame stops subsequent assignment for that story. Frame
`X`, `Y`, `Width` and `Height` describe its unrotated rectangle; `PageTransform`
maps frame-local points into page-local points after object and ancestor-group
rotation/reflection and scale. Child-anchor rotations use `relRotation` when
explicitly stored, including zero; otherwise they use the ordinary rotation
property. The same transforms reach artwork and rectangular wrap
exclusions. Hidden group descendants retain their source stories and assets but
do not paint or exclude printable text.

## Qualification and limits

[Fixture provenance](../OfficeIMO.Publisher.Tests/Fixtures/provenance.json)
records the imported Apache POI samples' source revision and SHA-256 hashes.
Those samples use the Apache POI license and notice included beside the fixtures.
The separately authored native controls record their producer, MIT license and
file hashes in their own contracts, described in the
[fixture guide](../OfficeIMO.Publisher.Tests/Fixtures/README.md). Format descriptions
include [Apache POI's Publisher notes](https://poi.apache.org/components/hpbf/file-format.html)
and Microsoft's [Office Drawing binary specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-odraw/).

Managed tests check page inventories, frame coordinates, source style values,
table cell boundaries, grouped placements, raw image recovery, SVG dimensions,
PDF page/text output and carried fidelity evidence. They also cover malformed
native offsets, cyclic Quill directories, invalid UTF-16, resource ceilings,
cancellation and caller stream ownership. The native newsletter exercises ten
linked stories, including chains whose object order differs from story order,
and declared picture exclusions. Multi-column property and malformed-link
mutations protect the native codec boundary; they are synthetic evidence.
The table sample exposes six distinct cells through the native public model and
Reader rows. Cell-span, missing-text-map and master-ownership mutations protect
merged geometry, explicit unresolved text, and single-owner extraction. Reader
checks cover generic headers, row limits, canonical table views, CSV output, JSON
transport and semantic PDF text deduplication. These mutations do not qualify
individual cell styling or Publisher-rendered table appearance.
Group rotation/reflection mutations verify text-frame corners, picture placement,
hidden-child wrapping and scene copying. A reflected picture starting above the
page retains its pixels through the group transform and final page clip.
Constructed nested native records check
composition about distinct centres, quarter-turn anchor dimensions and the
projected-item ceiling. These tests do not establish Publisher-rendered fidelity.
Native line-property mutations check caps, joins, fixed-point miter limits,
dash order and spacing, all five supported marker kinds, disabled strokes,
unknown values and scene copying. SVG, raster and PDF retain open arrows as
outlines and other markers as filled geometry. Marker dimensions and native dash
rendering use approximation reports. The Publisher-authored fixture independently
confirms the visible dash and endpoint decorations; exact marker sizing and
device-dependent dash spacing remain unqualified.
Custom-path record mutations check coordinate offsets, compact vertices,
line/cubic command consumption, subpaths, paint controls, malformed arrays,
and counted AutoLine, CornerLine, AutoCurve, CornerCurve and SymmetricCurve
editing runs. Zero-count joint markers remain inert. Nonzero SmoothLine,
SmoothCurve, SymmetricLine and Freeform runs retain an unsupported-command
fallback; their drawing semantics remain unqualified. These record checks cover
guide/command fallback reports and inset picture masks. One-point and move-only
paths retain empty geometry and empty picture masks rather than revealing the
unmasked image. SG guide checks cover
all 17 stored formula identifiers, unsigned constants, signed adjustment values,
geometry-space centres and dimensions, inherited document defaults, local scalar
and stroke-use overrides, physical frame EMUs,
backward references, fixed-point degrees and the 128-record ceiling.
Only guides referenced by path vertices and their backward dependencies are
evaluated; unrelated formulas do not force otherwise valid paths or picture
masks to fall back. Referenced device-pixel operands require an output-device
context and retain a reported fallback. Referenced malformed dependencies,
division by zero, invalid square roots and 32-bit overflow also retain fallback
evidence without partial custom artwork.
Integer rounding follows the corresponding [VML formula contract](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.formula):
products round to nearest with ties toward positive infinity, averages truncate
toward zero, and inexact operations floor their results. Exact half-angle sine
and cosine products retain their integral values; adjacent fractional angles
retain flooring. These managed checks
do not qualify Publisher's rounding or discontinuous guide behaviour.
Trigonometric calculations use directed bounds to preserve the floor near
integer boundaries. An unresolved bound retains the reported guide fallback.
High-coordinate
shape/mask checks retain the full unsigned compact range;
eight-byte coordinates retain signed negative values. Shared array decoding
preserves following complex properties when native lengths exclude their
six-byte headers, including empty arrays. Path, guide and gradient-stop consumers
reject extra element bytes beyond their declared counts. Document-wide primary
and tertiary defaults apply before local properties, with native record limits
bounding property decoding and inherited copies. Cumulative limits account for decoded vertices/segments,
declared guides, expanded commands and copied shape/mask paths. The producer corpus has no custom
paths; these managed scene and SVG/raster/PDF checks do not establish native
Publisher winding or mask fidelity.
Native linear-fill mutations check physical and aspect-scaled angles, positive
and negative focus, reversed ramps, duplicate color positions, transparency,
fill-use masks and picture layering. Foreground/background opacity uses a
common shape opacity and relative stop alpha; ratios requiring RGBA rounding
report loss. Multi-color fills retain foreground opacity and report distinct
background opacity as an approximation. Native shading corrections use reported
sRGB interpolation. Custom fill rectangles and view-origin anchors retain their
primary color with a mapping diagnostic; non-rotating fills on transformed
frames report approximation. Other gradient, pattern, texture and background
fill modes retain their current approximation boundary. The five supported
original producer fixtures contain no declared non-solid fills. The additional
Publisher-authored fixture confirms a horizontal two-color fill's direction and
endpoint colors against Publisher's PDF export. Its interpolation differs and
retains the approximation report; other gradient cases remain managed evidence.
Picture-control mutations exercise transparency keys before tone changes,
brightness endpoints, contrast, preserve-grays, grayscale and two-color
precedence. Controlled raster output in original native picture frames is
checked through SVG, raster and reopened PDF output; original embedded payloads
remain unchanged. Invalid or unqualified controls retain their diagnostics when
other valid controls apply. Per-image and cumulative pixel ceilings account for
inspected GIF frames and repeated references; cumulative encoded-byte checks
reject repeated large-payload decoding, and codec cancellation aborts the read. These are managed
record and artifact checks; the supported producer fixtures contain no enabled
visible picture controls, so native Publisher pixel equivalence remains open.
The brochure and newsletter also verify referenced style defaults, direct
formatting precedence, native bullet labels and tab-array positions. Those
checks compare decoded values with the native records, not Publisher-rendered
line breaks. Bullet labels do not change source-story character ranges.

The libmspub 0.1.5 reader independently confirms page dimensions, text-frame
coordinates, text styling and table semantics for the simple/table sample
family. It produces no output for the brochure/newsletter fixtures in the
qualification environment; those corpus tests therefore do not claim an
independent-reader rendering comparison. Microsoft Publisher 16.0 opens all
seven supplied corpus files and exports PDFs with their expected page counts.
Visual comparisons of supported publications confirm page geometry and expose
the reported metafile placeholders, text-fitting and layout differences; they
do not establish pixel fidelity. A Publisher-authored fixture and reference PDF
also qualify modern page geometry, story recovery and borderless text/table
frames. The opt-in producer script and fixture hashes are retained with the
tests. No external program is required by normal builds or runtime conversions.

Twelve additional [native comparison controls](../OfficeIMO.Publisher.Tests/Fixtures/NativeComparison)
cover blank and mixed-content pages, changed-length text, a rectangle, embedded
pictures, a linked two-page newsletter, a custom-size multilingual table/label,
a two-sided landscape brochure and distinct single-page and paired masters.
Each retains a Publisher-authored PUB, native reference PDF and hashed contract
with physical page IDs, master relationships and authored expectations. These
controls qualify physical-page order including empty pages, selected geometry,
complete stories, native catalog-backed chains, table cells, image pixels,
brochure panel geometry and PDF page sizes/counts. Master controls qualify the
distinct definitions and assignments, definition-local frame geometry, and
applied marker origins measured from native PDF text spans. Those selected
origins agree within 0.1 point horizontally and 1 point in baseline position;
the baseline allowance accounts for the qualified difference in font metrics.
Each native and OfficeIMO PDF page contains only its assigned master's markers.
These controls do not qualify native frame breaks or broad pixel fidelity.
The TableLabel PUB has one 288-by-216-point publication page and a 2-by-2 table
at (18, 66), sized 180 by 72 points, with 90-point columns and 36-point rows.
Its native reference PDF is an imposed 595.32-by-841.92-point label sheet with
six copies in two columns and three rows. TableLabel qualification uses the PUB
inventory and an individual copy's content and geometry; it does not claim
whole-sheet page or rendering equivalence. The managed publication retains its
single label-sized page rather than reproducing the native print imposition.
The multilingual PDF check uses explicit Unicode font configuration; default
standard-font output is a separate contract. Native page enumeration also has
instruction and decompiler evidence for the qualified Publisher build. Its
serialized role-boundary interpretation is corroborated by these controls and
the original corpus; other producer builds and nested master inheritance remain
unassessed.

The default resource limits are 64 MiB input, 4 million source/projected text
characters, 1 million inspected records, 250,000 source/projected items,
512 compound streams, 1,024 document/master pages, nesting depth 64,
16 MiB per extracted/projected image and 64 MiB aggregate image accounting.
Decoded picture-effect rasters and application codec output are limited to
8 million pixels per image. Cumulative picture-effect work is limited to
64 million pixels, charging the selected raster once for decoding and once for
filtering, plus each inspected GIF, WebP or icon frame across all references.
Image extraction and each actual picture decode also charge their encoded bytes
against `Limits.MaxInputBytes`. Effect decoding fails the read when the image or
selected codec cannot produce a raster within those bounds.
Multi-frame effect sources use their first frame and report omitted frames;
`Images` retains the complete payload. Limits are
caller-configurable, reject oversized input and also account for repeated
projection work. The input byte ceiling also bounds cumulative encoded image
processing, and unavailable image-store entries consume the item ceiling.
The item ceiling also bounds cumulative projected gradient-stop work, including
focus expansion and repeated master-page projection, and recovered table tracks
and cells. Copied table text consumes a cumulative text ceiling. A native color-stop table
supports at most 256 entries; malformed or larger tables retain primary paint
with an invalid-gradient diagnostic.
Delayed decoding is cached and cancellation is checked between image entries.
The text ceiling also bounds cumulative continuation measurement, and the record
ceiling bounds exclusion-region inspection. Native frames support at most
256 columns within the configured item ceiling.
Finite text/table bleed coordinates survive projection and copying until page
clipping is applied. Each source text story currently has the shared drawing
ceiling of 100,000 characters and 4,096 rich-text runs/paragraphs.

Open implementation and qualification work belongs to the
[Publisher roadmap](../Docs/ROADMAP.md#publisher-publication-recovery).
