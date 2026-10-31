# Publisher native fixtures

The seven corpus publications come from Apache POI's
[`test-data/publisher`](https://github.com/apache/poi/tree/66109187d12b72696ea02101cec765d77d9bbdc9/test-data/publisher)
corpus. Source revision and file hashes are recorded in `provenance.json`.
`LICENSE-APACHE-POI.txt` and `NOTICE-APACHE-POI.txt` retain the upstream terms.

Simple, Sample and Sample_2010 exercise positioned text, styles and tables.
SampleBrochure and SampleNewsletter exercise grouped artwork, GIF/JPEG/WMF
resources, page order and linked stories. Sample2000 and Sample98 protect the
explicit earlier-generation rejection contract.

The fixtures are test inputs, not native output authored by OfficeIMO. Expected
simple/table geometry and style values are checked against libmspub 0.1.5's
`pub2raw` output. Native Publisher PDF renderings are not supplied by this corpus.

`native-authored.pub` and its reference PDF were created in Microsoft Publisher
16.0 using [the opt-in Windows producer script](../../Build/Publisher/New-NativeQualificationFixture.ps1).
The fixture contains a text frame, two-color gradient, decorated dashed line,
rotated diamond and 2-by-2 table on a 612-by-792-point page. Its producer metadata
and SHA-256 hashes are in `native-authored.json`. These authored fixture files
use OfficeIMO's MIT license; the Apache POI terms apply to the upstream corpus.

## Produce native comparison controls

[New-NativeComparisonCorpus.ps1](../../Build/Publisher/New-NativeComparisonCorpus.ps1)
creates synthetic publications directly in Microsoft Publisher and exports their
native PDF references. Run it in PowerShell 7 on Windows with Publisher installed and no existing
Publisher session. Choose an empty task-owned output directory:

```powershell
./Build/Publisher/New-NativeComparisonCorpus.ps1 -OutputDirectory ./native-controls `
    -Case Blank,Text,TextLong,Rectangle,Image
```

The minimal cases isolate a blank publication, one text frame, a longer version
of the same story, a solid rectangle and an embedded synthetic picture. BlankPages
has two empty physical pages; TextBlank has a text page followed by an empty page.
They distinguish blank document pages from utility and master definitions. Text and
TextLong use identical page geometry, frame geometry and Arial formatting; their
only authored difference is the text. TextParagraphs has one Arial 12-point story
with an empty paragraph between two text paragraphs. Its retained character
formatting covers that empty paragraph's CR; the final story CR has a separate
native character range. TextNonuniformParagraphs has seven independent final-state
pages: a four-paragraph baseline, deletion of the first/middle/final paragraph,
reordering, duplication and equal visible paragraphs with distinct native language
definitions. Alignment, paragraph spacing and indentation distinguish the source
paragraphs. Each frame uses Arial 12-point body text; the final CR has its own
inherited typing style. Its contract records requested and observed paragraph
formatting and language values.
TableNonuniformParagraphs uses the same seven paragraph profiles in the first
and final cells of a two-by-two table, with two unchanged neighboring cells.
Each 468-by-432-point table fits on one page. The control distinguishes native
per-cell formatting and language metadata from the complete table story and
its TCD partition endpoints; it does not measure general row reflow.
TextNonuniformEmptyParagraphs and TableNonuniformEmptyParagraphs add interior
and final empty paragraphs, independent final-state retention controls and
equal projected empties with distinct native language metadata. The ordinary
control also has five visible Western font witnesses with distinct right-to-left
size and emphasis fields. TextUniformEmptyParagraphs and
TableUniformEmptyParagraphs distinguish exact empty-paragraph identity from
uniform formatting, including retained empties mixed with newly authored text.
Their native contracts record the requested styles and observed native values.
TextNonuniformTextReplacements and TableNonuniformTextReplacements provide
fourteen independently authored pages each. Their final states combine text
growth/shrinkage with paragraph deletion, reordering and duplication, sole
paragraph replacements, equal native templates with different source text,
retained empty paragraphs and exact native language selection. The table control
edits first/final cells and keeps two neighbors unchanged. Fitting Arial 12-point
bodies isolate native paragraph definitions and CR roles from general row reflow.
TextMixedRunTextReplacements and TableMixedRunTextReplacements each provide
thirteen independent pages with three styled runs per paragraph: blue 14-point
bold text, red 12-point italic text and black 12-point regular text. Documented [native word ranges](https://learn.microsoft.com/en-us/office/vba/api/publisher.textrange.words) author their boundaries. Final states vary run lengths while
deleting/reordering/duplicating paragraphs and include equivalent templates,
middle-run-only edits and opaque language differences confined to the middle
run. The table uses fitting first/final cells and unchanged neighbors.
TextFormattingTopology and TableFormattingTopology each provide sixteen native
pages with retained three-run text and independently authored formatting/count
changes. The controls include emphasis, size, colors, fonts, double underline,
paragraph alignment/spacing/indentation and English/French middle-run metadata.
Native word ranges change visible runs while retaining CR typing. Destination
style collisions distinguish text correspondence from borrowing another native
paragraph's metadata. The Consolas/green final state also compares an edit of the
earlier mixed-run source whose font/color pools lack those entries.
TextStories has three separate stories, with an empty
paragraph inside its middle story. TextFormattingPages retains 96 paragraphs,
including overflow beyond its sole text frame, across two native paragraph-formatting
pages. ZeroAnchors places each client-anchor coordinate at native zero: native
storage omits the corresponding property instead of writing an explicit zero.
Its four text frames qualify recovery of those omitted coordinates.
TextFormattingMany has 240 paragraphs across three native paragraph-formatting
pages, retaining overflow. TextCharacterPages has four independent stories in
Arial 12, 14, 16 and 18 points, producing two native character-formatting pages.
TextCharacterAttributes has twelve whole-frame attribute controls across two
pages: regular/italic, single/double underline, strikethrough, superscript/subscript,
Arial/Cambria/Courier New/Consolas, red text and combined blue bold/italic text.
TextUnderlinePatterns records all eighteen automation underline choices across
three pages. Its saved Quill values differ from the automation enum; native
weight, long-dash and compound patterns are distinct from the shared model's
representable patterns. The fixture contracts record each requested attribute,
and their native PDFs are independent appearance references.
Both files use two chained Quill directories, with 22 and 26 descriptors.
They also provide independent inputs for the separately qualified character and
chained-directory source-save profiles in the writer qualification record.
TextParagraphAttributes covers fourteen whole-frame paragraph profiles across
four pages: alignment, before/after spacing, margins, first-line/hanging indents,
exact/multiple line spacing, two indexed left tabs and combined settings. Each
frame contains two paragraphs. TextAlignedTabs adds a one-page left/center/right
control for complete tabbed-line alignment. The native API may canonicalize a
multiple of 1.5 to its preset rule; the contracts retain the observed values.
These controls qualify paragraph recovery and the separately recorded save
operations, with exact native font metrics and rendering parity still unassessed.
TextFrameLayout supplies nine profiles across three pages: zero/default insets,
explicit insets, two/three columns, top/center/bottom alignment, combined settings,
mixed text/rectangle artwork and fractional point values. Native whole-frame
getters and independently decoded records establish that automation column
spacing is half the full native gap. Top alignment is omitted; center/bottom use
Contents field `0x35`. Its PUB, PDF and contract support the separately recorded
source-save qualification, including growth from omitted default fields.
TableCellTyping supplies nine table profiles on three pages: first/final/all
clearing, styled empty CRs, French language, refill and styling after clearing.
TableCellTypingEarlyRange supplies eight profiles on two pages with ranges
captured before assigning text. That range can exclude the subsequently
materialized CR, leaving its default typing different from the visible body.
The historical early-range fixture retains its original `native-tablecelltyping`
filenames; the current EarlyRange case reproduces its normalized cell/CR styles
and language observations. Both controls record whole-cell text and the documented
[`TextRange.LanguageID`](https://learn.microsoft.com/en-us/office/vba/api/publisher.textrange.languageid)
before and after editing. Original TableLabel authoring uses fresh whole-cell
ranges. These controls qualify the separately recorded empty-typing source saves;
they do not establish general language metadata editing or renderer parity.
TableCellParagraphs supplies four independently authored profiles on one page:
baseline, first-cell two paragraphs, final-cell three paragraphs and all-cell
two paragraphs. Fresh whole-cell styling keeps definitions uniform within each
cell while the top row remains bold and the bottom row regular. Its PUB, PDF,
native contract and provenance support separate cell paragraph-count save-back
qualification; they do not qualify fixed native heights for arbitrary overflow.

TableLabelShortParagraphs and TableLabelWrapped retain the original label's
page, table tracks and artwork while assigning short or wrapped two-paragraph
text to its first cell. Their contracts record the documented native row heights
and column widths in points. The wrapped control distinguishes saved cumulative
track endpoints from unchanged nominal sizes and independently reproduces native
open/resave row-height drift. Corresponding-stage output, stored geometry and
current native reflow are qualified separately.

Newsletter adds a two-page linked story
and a picture. TableLabel adds a custom page size, styled 2-by-2 table, picture
and Polish/Greek characters. Brochure adds two 792-by-612-point landscape sides,
each with three 264-point panels, explicit text frames and panel markers.
MultipleMasters applies two single-page masters in A/B/A order across three
physical pages. TwoPageMaster retains the default first-page master and applies
the distinct members of a native two-page master to the next left and right
pages. Distinct marker text and geometry identify each master in the reference
PDF. PictureOwnership places two physical pictures on its first page and one
picture on a master shared by both physical pages. The second physical picture
uses Publisher's [Shape.Duplicate method](https://learn.microsoft.com/en-us/office/vba/api/publisher.shape.duplicate).
Its saved native image-store references must be inspected separately; duplicating
a shape does not establish that Publisher shares the embedded image asset.
The retained PictureOwnership PUB has three placements referencing one native
image-store entry: one master picture and two pictures on the first physical page.
PictureReplacement replaces only the original physical placement with a distinct
64-by-32-pixel embedded PNG, retaining the duplicated and master placements.
The replacement target is unlocked before restoring its original 96-point
square bounds; the native contract asserts every picture's bounds and crop.
PictureCropped applies nonzero left/top/bottom crop to that original placement
while retaining the shared asset and physical frame. These two controls retain
their source PNGs, native PUB/PDF and authored/observed contracts.
ImageRetiredSlot contains a separate lifecycle regression: a public source-save
candidate independently resaved in Publisher, which retains stale zero-count
image-store metadata. Its provenance records the candidate hash and native
operations; it does not qualify independent template-free creation.
Omitting `-Case` selects every available case.

Each case has a PUB, a Publisher-exported PDF and JSON that separates its authored
contract from the observed native page/shape inventory. The root manifest records
the case contracts and file hashes. Master controls also record authored
assignment IDs and the observed shapes on each native master definition.
PictureOwnership also records its authored zero crop values and native
observations of picture cropping, embedding and page ownership. Numeric crop
values follow [Publisher's original-image point units](https://learn.microsoft.com/en-us/office/vba/api/publisher.pictureformat.cropleft).
Text fonts, margins and object colors are explicit; other native publication defaults
remain recorded as defaults. Producer
metadata records the Publisher executable version, PowerShell version, culture
and Arial font-file hash when available. Native file bytes can vary between runs;
hashes identify a particular produced reference rather than promise deterministic
Publisher serialization.

Master definition geometry uses that master's local coordinate system. Applied
marker placement on physical pages is checked separately against text origins
in the native reference PDF; local definition coordinates do not establish
their applied positions.

The master controls use unique single-character abbreviations. On the qualified
Publisher 16.0 build, the paired-master abbreviation `P` succeeds and `TP` returns
an invalid-argument error; the API documentation does not specify a maximum
abbreviation length. The producer keeps that observed constraint within its
synthetic controls.

The shared [native operation controller](../../Build/Publisher/Invoke-NativePublisherOperation.ps1)
also validates an OfficeIMO-written candidate:

```powershell
./Build/Publisher/Invoke-NativePublisherOperation.ps1 -Operation Validate `
    -InputPath ./candidate.pub -OutputDirectory ./candidate-native
```

Validation takes a byte-identical task-owned copy, opens it read-only without
adding it to recent files, and captures its inventory and native PDF before any
resave. It then opens only the task-owned copy writable, saves a separate PUB and
reopens that result. The original candidate's hash must remain unchanged.
Frame inventories read the visible `TextFrame.TextRange` prefix; they do not
establish that overflow was retained. Complete overflowing stories require a
separate check of the native-resaved Quill text and story-length tables. COM
Open success does not expose whether Publisher repaired a document, so the report
keeps repair status unassessed. Native UI observation is separate evidence.

The controller runs a hidden STA worker with a bounded deadline. It disables
macros for its owned application and records the Publisher process identified by
the COM window and exact process start time. Cleanup never stops processes by
name alone. A startup failure before that identity is recorded leaves uncertain
processes untouched and reports the gap. `.run` contains task-local process
receipts and logs; it is not a fixture input. An explicit completed-operation
receipt separates completed evidence from COM host teardown. The controller can
stop its exact worker after that receipt and records this disposition separately;
an ordinary timeout without a completion receipt remains a failed operation.

These controls are independent producer evidence. They do not establish native
writing, rendering parity or supported editing until the corresponding library
operation is checked against the exact retained files. Promote only reference
files used by executable tests or a maintained qualification contract; keep
intermediate runs and process logs outside the source fixture directory.
