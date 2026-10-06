# EPUB validation evidence

`validate_epub.py` runs an explicitly supplied EPUBCheck distribution over captured
publication bytes. It is an opt-in validation tool, outside normal builds and product
runtime dependencies. It does not install Java, EPUBCheck or any accessibility tool.

Download an official [EPUBCheck distribution](https://www.w3.org/publishing/epubcheck/)
and retain its adjacent library directory. Supply its JAR and a new output directory:

```sh
python3 Build/Epub/validate_epub.py book.epub edited-book.epub \
  --epubcheck-jar /path/to/epubcheck/epubcheck.jar \
  --output /path/to/task-evidence/epubcheck
```

The runner records the tool version and JAR hash, captures each input, hashes the
captured bytes, and retains the validator's log and JSON report. `summary.json`
identifies each input and outcome. A failed validator, timeout or missing report
produces a nonzero exit. Existing output directories are rejected to prevent mixing
reports from different runs. Inputs are limited to 256 publications, 128 MiB each;
`--timeout` controls the per-process limit in seconds.

To also capture automated accessibility checks, supply `--ace /path/to/ace-puppeteer`
from an independently installed [DAISY Ace](https://daisy.github.io/ace/) distribution.
Its browser runtime must already be available. The runner records its version, log,
JSON report, exit code, and reported outcome for the same captured EPUB bytes. Missing,
malformed, failing, or timed-out evidence produces a nonzero exit. Without `--ace`,
automated accessibility is recorded as unchecked.
The `automatedAccessibility.coverage` object separately records SVG coverage and
known limitations. Ace 1.4.6 reports containing SVG spine documents are marked
`not-checked` for that content; other versions remain `not-established` until their
coverage is verified. Coverage annotations never change the validator outcome or
turn a failing run into a pass.
On interruption or timeout, the runner terminates Ace's process tree, including detached
browser groups on POSIX. Cleanup has a bounded grace period beyond the requested audit
timeout. Run the POSIX timeout contract with
`python3 -m unittest discover -s Build/Epub`; it exercises ordinary and detached children.

To check explicit SMIL clip bounds against encoded audio, supply an installed
[ffprobe](https://ffmpeg.org/ffprobe.html). Add `--ffmpeg` to require a successful
full decode of every referenced audio resource:

```sh
python3 Build/Epub/validate_epub.py revised-narration.epub \
  --epubcheck-jar /path/to/epubcheck/epubcheck.jar \
  --ffprobe /path/to/ffprobe --ffmpeg /path/to/ffmpeg \
  --output /path/to/task-evidence/narration
```

The `encodedAudio` result records clip offsets, resource hashes, probed durations,
probe reports and decode outcomes for the captured publication. The audio stream's
duration takes precedence over the container duration. A clip beyond that duration
fails without a rounding allowance. Without `--ffprobe`, this scope remains
`not-checked`; without `--ffmpeg`, decoding remains `not-checked`. An explicitly
requested but unsupported check returns `not-checked` and a nonzero runner exit.
The runner records executable versions and hashes and does not install either tool.

This lane supports one package rendition, one audio stream per resource, explicit
clip ends, and local MP3, MP4, WAV, Ogg and FLAC resources. Missing clip ends, XML
base addressing, external/query/fragment references, encrypted or obfuscated
publications, and unsupported media types require separate qualification. Limits
are 10,000 ZIP entries and clips, 256 overlays and audio resources, 2 MiB per XML
member, 128 MiB per audio resource, and 256 MiB of expanded audio per publication.
Only referenced resources are materialized under temporary generated filenames;
fixed demuxers and a local protocol allowlist prevent playlist/network probing.
Temporary audio copies are removed after each check. Logs and compact probe reports
remain with the evidence.

Probed duration and successful decoding do not establish speech alignment, audible
endpoints, synchronized highlighting, seeking or native reader behavior. Those
scopes remain unchecked. In particular, codec delay and padding can differ from
the audible duration; the result reports encoded timing, not perceptual alignment.

A passing result establishes only the checks performed by the recorded tool versions.
Review warnings and retain the versions with the evidence. The summary explicitly
leaves comprehensive accessibility assessment and reader presentation unchecked even
when Ace passes. Complete human accessibility review, then inspect the same publication
bytes in the chosen reading systems. Never infer those results from validator exit codes.

Use a task-owned output location. Retain compact reports and decisive publication
fixtures deliberately; remove superseded captured publications and downloaded tools
when they are no longer needed. Source fixtures need producer, version and license
provenance before entering the maintained corpus.

## Independent media-overlay source

The W3C EPUB sample collection's
[Moby-Dick media-overlay source](https://github.com/IDPF/epub3-samples/tree/7651e2002b631e6577fadf7e9e0692fa6efb8746/30/moby-dick-mo)
provides an independent preservation and narration-inspection case. Pin revision
`7651e2002b631e6577fadf7e9e0692fa6efb8746`, retain its copyright and licensing notices,
and package the source with `mimetype` first and uncompressed. Keep the source
payloads unchanged and record the resulting archive hash with the validation reports.

The checked sample has two overlays. It passes EPUBCheck 5.4.0 and the native
`media-overlays` preflight check, and unchanged load/save preserves its archive bytes.
Its package lacks modern accessibility discovery metadata, so the overall native
preflight reports that gap. These results do not establish audio decoding, reader
synchronization, assistive-technology behavior or comprehensive accessibility.
The opt-in `ProducerEditQualification` runner exercises edits against this sample:

```text
dotnet run --project Build/Epub/Producer/ProducerEditQualification.csproj -- sample.epub EXPECTED_SHA256 new-output-directory
```

Supply the archive hash from the retained source-provenance record. The runner checks
that hash before editing; it does not authenticate the upstream Git revision or obtain
the sample. It rejects existing output directories and inputs over 128 MiB. It retains
an unchanged input baseline, then relocates a narrated chapter, its SMIL resource and
shared audio, splits the introduction, and merges it again. Edited outputs use a fixed
modification timestamp for repeatability. Run `validate_epub.py` separately over the
baseline and edited EPUBs; validator failures must remain visible.

The runner reopens each output and checks all 143 non-navigation chapter bodies,
141 original TOC entries (labels, order, nesting and targets), seven non-XML assets,
and both SMIL trees after normalizing moved references. Audio bytes, clip timing and
narration targets are preserved. Its JSON identifies every changed source payload,
records native preflight findings, and marks a run complete only after all four stages.
The split and merged editions retain an additional TOC entry for the new boundary.

All four outputs pass EPUBCheck 5.4.0 without warnings. Ace 1.4.6 reports 296 rule
failures in the unchanged, renamed and merged editions: missing discovery metadata,
missing document languages, EPUB-type/ARIA-role mappings and heading order. The split
edition has 297 because its cloned document adds one missing-language finding. These
are retained failures, not an accessibility pass. Native preflight also reports the
source's missing modern accessibility metadata. Keep those results separate from
preservation and EPUBCheck. Independent encoded-audio
qualification of the retained MP4 found all 40 clips within its 1436.43-second duration
and a successful full decode. Repeat this check on captured editions with the optional audio lane above. These results
do not establish synchronized highlighting, seeking, pause/resume or assistive-technology
acceptance in a reading system.

The sample stays outside shipped packages and is not a runtime dependency.

## ONIX fixtures

The opt-in `OnixFixtureGenerator` exercises the workflow's bibliographic ONIX export
with early, advance and confirmed notifications, person and organization credits,
explicit no-contributor metadata, a subtitle, publication date and English, Polish
and French language codes. It writes each ONIX record with its exact EPUB, records
their hashes and the three schema hashes, and checks that the schema rejects an
invalid language code. Commercial fixtures cover worldwide exclusions, distinct
Polish and British markets, all seven supported price bases, decimal scale, effective
dates, free and unannounced pricing, an unknown supply date, and withdrawal after
rights are lost. The `taxed` fixture covers all three tax types, six rate classifications,
multiple components, percentage-only and amount-only assertions, decimal scale,
zero rating and exemption. Its rates and classifications are synthetic serialization
inputs. The `discounted` fixture covers all four discount types, closed and open-ended
quantity ranges, percentage-only and amount-only values, and explicit zero discount.
It validates serialization, not trading-partner eligibility or tier calculation.
The `discount-coded` fixture covers all seven code schemes, named proprietary
schemes, and structurally formatted BIC and ISNI codes. These synthetic codes do
not assert allocation, ownership or any actual trade agreement. The `discoverability`
fixture selects an alternate EPUB title and covers all supported subject schemes,
main-subject markers, version assertions and multilingual headings. Classification
assignments are synthetic and do not establish vocabulary membership or suitability.
The `title-sorting` and `title-no-prefix` fixtures cover explicit prefixes (including
spaces and apostrophes), no-prefix assertions, product and collection title paths,
per-element language, part-only hierarchy elements and byte-preserving ONIX message
composition. The schema must reject simultaneous prefix and no-prefix assertions.
The `alternative-titles` fixture covers all supported book title classifications,
multilingual text and subtitles, independent sorting declarations and retained order
through message composition. Negative controls require rejection of an unknown title
type and an unknown language code on an alternative title.
The `collection-identifiers` fixture exercises national catalog identifiers, DOI,
URN, Japanese magazine IDs, ARK resolver URLs and ISSN-L, including message
composition and a schema negative control for an unknown identifier scheme. Its
synthetic values do not establish allocation, ownership or resolver availability.

The `collection-brand-universe` fixture covers standalone master brands and fictional
universes, their combination with all three series levels, scoped identifiers and
message composition. Its negative control requires rejection of an unknown identifier
level. Names and associations are synthetic and do not assert ownership or licensing.
The `edition` and `no-edition` fixtures cover numbered minor revisions, multilingual
plain-text statements and explicit absence of edition information. The generator
also validates each supported edition type against the schema. The `collection`
fixture covers all supported collection, identifier and sequence types, ordered
person and organization credits, every supported contributor role, and omitted
versus explicitly absent collection credits. The schema must reject misplaced
credits and conflicting contributor assertions. The `no-collection` fixture distinguishes explicit absence from omitted membership.
The `collection-hierarchy` fixture covers three hierarchy levels, display order
independent of level, multilingual title and part designations, and identifiers
scoped to each level. The `collection-frequency` fixture covers all supported
frequency codes. The schema must reject duplicate title sequence numbers and an
unknown frequency code. Collection identities and schedules are synthetic assertions,
not evidence of allocation, ownership or a publisher's actual publication schedule.
The audience fixtures cover every supported category, a main-audience marker, multilingual
descriptions, exact/open/closed age ranges and the 36–42 month boundary. Their combined
assertions exercise serialization and do not establish readership or suitability.
The `complexity` and `complexity-audience` fixtures cover all supported list 32
schemes, complexity-only metadata, ordering after audience descriptions, distinct values
in one scheme and preserved decimal spelling. Their values are synthetic publisher
assertions, not independently assigned levels or evidence of scoring accuracy.
The `adult-unrated`, `adult-general` and `adult-advice` fixtures cover explicit unrated
and unrestricted-adult assertions, every supported content-advice code, translated
headings and main flags scoped independently from general audience categories.
The code mapping follows ONIX list 203 issue 74, including death/grief and suicide.
`AudienceCodeValue` is a string in the structural schema: XSD validation verifies
structure, not current rating-code membership or the publisher's suitability judgment.
The `audience-headings` fixture covers translated general categories, code-plus-heading
and heading-only assertions, single unspecified-language headings, scoped main-audience
flags and whitespace/XML escaping. Its single-record message composition must preserve
the original bytes; it does not assess translation quality or recipient presentation.
The `audience-codes` fixture covers all supported additional audience scheme identifiers,
proprietary scheme-name escaping, main-audience flags across schemes and retained code values.
External code values are synthetic; schema validation does not establish their membership
in the scheme owner's current vocabulary or recipient acceptance.
The `audience-grades` fixture covers US preschool-to-kindergarten, Canadian grades
9–12 and an open-ended Chinese tertiary range. The generator also validates every
supported grade code in all three systems against the supplied schema. Grading
systems remain explicit; these synthetic assertions do not establish age equivalence
or educational suitability.
The collateral fixtures cover all supported text/recipient types, multilingual plain
text, attribution, source links, territory, usage dates, long descriptions and the
350-scalar Unicode boundary. All quotations and attributions are synthetic. The `collateral-xhtml` fixture exercises
namespace normalization, lists, definitions, quotations, tables and links, plus a
short description containing 350 decoded supplementary characters. Its generator
also requires the full schema to reject invalid paragraph nesting. Accessibility fixtures exercise unknown status, publisher contact
and information pages, feature codes, and explicit EPUB Accessibility 1.1/WCAG
declarations, including distinct certifier, credentialling organization, independent
report, intermediary and compatibility-report roles. Those assertions are synthetic
serialization inputs, not certifications of the accompanying EPUBs. The schema must also reject unknown country and currency codes.
The runner also writes `catalog.onix`, combining two distinct edition ISBNs under
one header, and `catalog-priced.epub` for the second product. The message evidence
retains each source ONIX and EPUB hash. It is outside normal builds and shipped packages.

Obtain the ONIX 3.1 reference XSD from [EDItEUR](https://www.editeur.org/93/Release-3.0-and-3.1-Downloads/)
and retain its unchanged adjacent code-list and XHTML schemas and their license
notices. The local schema directory must contain `ONIX_BookProduct_3.1_reference.xsd`,
`ONIX_BookProduct_CodeLists.xsd` and `ONIX_XHTML_Subset.xsd`. The runner's resolver
permits only those local files; it does not fetch network dependencies.

```sh
dotnet run --project Build/Epub/Onix/OnixFixtureGenerator.csproj -- \
  /path/to/onix-schema /path/to/new-onix-evidence
xmllint --nonet --noout --schema /path/to/onix-schema/ONIX_BookProduct_3.1_reference.xsd \
  /path/to/new-onix-evidence/early.onix \
  /path/to/new-onix-evidence/advance.onix \
  /path/to/new-onix-evidence/confirmed.onix \
  /path/to/new-onix-evidence/priced.onix \
  /path/to/new-onix-evidence/withdrawn.onix \
  /path/to/new-onix-evidence/catalog.onix \
  /path/to/new-onix-evidence/taxed.onix \
  /path/to/new-onix-evidence/discounted.onix \
  /path/to/new-onix-evidence/discount-coded.onix \
  /path/to/new-onix-evidence/discoverability.onix \
  /path/to/new-onix-evidence/edition.onix \
  /path/to/new-onix-evidence/no-edition.onix \
  /path/to/new-onix-evidence/collection.onix \
  /path/to/new-onix-evidence/no-collection.onix \
  /path/to/new-onix-evidence/audience.onix \
  /path/to/new-onix-evidence/audience-months.onix \
  /path/to/new-onix-evidence/audience-open.onix \
  /path/to/new-onix-evidence/complexity.onix \
  /path/to/new-onix-evidence/complexity-audience.onix \
  /path/to/new-onix-evidence/adult-unrated.onix \
  /path/to/new-onix-evidence/adult-general.onix \
  /path/to/new-onix-evidence/adult-advice.onix \
  /path/to/new-onix-evidence/audience-headings.onix \
  /path/to/new-onix-evidence/audience-codes.onix \
  /path/to/new-onix-evidence/audience-grades.onix \
  /path/to/new-onix-evidence/collateral.onix \
  /path/to/new-onix-evidence/collateral-unicode.onix \
  /path/to/new-onix-evidence/collateral-xhtml.onix
```

Retain the independent validator version, exit code and log alongside `evidence.json`.
The generator records independent validation and retailer acceptance as unperformed;
its own successful schema check does not stand in for either. Run the generated EPUBs
through the EPUB validation runner separately. ONIX schema conformance is narrower
than trade business-rule validation or recipient acceptance. When schema files come
from a mirror, retain its immutable revision and the original schema headers and
distinguish that provenance from a fresh official download.

## Typography fixtures

The opt-in generator uses the current EPUB owner to create Basic, Prose, and
Technical publications from one manuscript containing tables, code, long links,
an accessible SVG, Arabic, and Japanese. It also creates `glossary.epub` with two
terms, repeated references and localized return links, and `bibliography.epub` with
two entries formatted and title-sorted by `OfficeIMO.Bibliography`, three citations,
and return links. The bibliography fixture checks that order, escaped title text,
and italic formatting survive EPUB reopening. Its CSL style is an illustrative
local style, not a claim of qualification for every publisher's citation style.
It has no third-party test dependency and
is outside the normal solution and shipped packages.

`merge-context-selectors.epub` and its unmerged source exercise map and output
`name` values that converge during identifier repair. Explicit type selectors keep
the map content green and bold, and the unrelated output content blue at normal weight.
The map keeps its default inline layout; its block paragraph and the output retain their respective borders. The output is static and has no interactive controls.
The third chapter keeps the original shared stylesheet.

`merge-id-selectors.epub` and its unmerged `merge-id-selectors-source.epub` qualify
partial ID comparisons and selectors naming a
newly assigned ID. The renamed paragraph remains green, bold and bordered; the
retained paragraph is green and bordered; the unselected paragraph is green without
a border. The third chapter retains the original shared stylesheet. Check these
states in intended readers, including support for generated `:is(...)` alternatives
and `:where(...)` filters.

`manuscript-relationships.epub` imports a labelled section and a forward description
reference across proposed heading boundaries. The connected sections share one
content document; a separate chapter and cross-document heading links remain.
Use the fixture to qualify heading navigation and accessible descriptions in a reader.

`notes-and-pages.epub` combines a same-document footnote and a cross-document
endnote with labelled return links, front/body/back matter, and roman/Arabic
print-page labels in the page list. Its print source is explicitly synthetic.
Use the exact EPUB to check forward and return links, note presentation, matter
navigation and page-list activation in a reading system; automated validator
results do not establish those interactions.

`index.epub` exercises nested terms, multiple locators and an intra-index
cross-reference. Its labels target sections and chapters rather than inferred
screen page numbers. `renamed-index.epub` moves both chapters and the navigation
document into new directories, exercising locator, cross-reference and TOC repair.
Run both artifacts through the independent validators to compare their conformance.

`split-chapter.epub` divides a sectioned chapter into consecutive reading positions,
with forward, return and incoming links repaired across the new resource boundary.
Use it to check TOC order and link destinations in independent readers as well as
with the automated validators.

`fixed-layout-ltr.epub` and `fixed-layout-rtl.epub` contain an 800×600 landscape
canvas and a 600×800 portrait canvas. Typed regions position headers, sections and
footers while preserving the semantic document order. `fixed-layout-escaped-id.epub`
exercises fractional coordinates and an identifier containing punctuation and a
supplementary Unicode character. Verify its region rectangle in a browser as well
as checking its EPUB package. The page fixtures exercise typed orientation, spread and
page-side requests, logical DOM order and semantic links with a package-level
fixed-layout default. `fixed-layout-item-overrides.epub` deliberately omits that
default to expose readers that do not honor item-only fixed-layout declarations.
Validate their exact EPUB
bytes with EPUBCheck, then inspect spread placement, rotation and scaling in target
readers. Extracted XHTML can verify CSS canvas geometry in a browser; it does not
qualify a reader's spine interpretation or assistive-technology behavior.

`fixed-layout-svg.epub` contains an SVG spine page with two linked panels, a title
and description, and an explicit 800 × 600 viewport. Inspect the actual SVG
rendering and links; EPUB schema validation does not qualify native reader scaling
or assistive-technology navigation. Ace 1.4.6 ignores SVG spine documents and can
report their valid TOC targets as missing (`epub-toc-order`). Retain that result as
an explicit tool-coverage limitation; do not treat it as an automated accessibility
pass or remove valid navigation to satisfy the checker.

`read-aloud.epub` pairs two text paragraphs with recorded narration, explicit SMIL
clip intervals, package durations and active-text styling. The test-only MP3 and
its source/encoding notes live in `Fixtures/Assets`. Its accessibility summary
identifies the unnarrated heading and unqualified reader playback. Verify cue
synchronization, highlighting, seeking and pause/resume in a media-overlay reader;
successful decoding and EPUBCheck/Ace results do not establish those behaviors.
`read-aloud-revised.epub` exercises typed overlay replacement by moving the cue
boundary within the recorded inter-sentence silence while retaining both cue IDs.
`read-aloud-nested.epub` groups the same recorded sentences as list-item cues under
a list sequence and a containing section sequence. `read-aloud-nested-revised.epub`
updates its timing through nested replacement, preserving sequence and cue IDs.
Reader skipping and escaping remain separate native acceptance checks.
`read-aloud-svg.epub` and `read-aloud-inline-svg.epub` pair the same recording with
SVG text targets and a group sequence, including a timing replacement. They cover
standalone SVG and SVG inside XHTML respectively. Ace 1.4.6 does not inspect the
standalone SVG spine page; keep that coverage gap separate from EPUBCheck results
and native highlighting/playback acceptance.

`merged-chapters.epub` merges those split reading positions back into one resource,
retaining both TOC entries and repairing the incoming reference. Check its chapter
order, section links and return links alongside the split fixture.
`merged-identifiers.epub` starts with repeated chapter-local heading and description
IDs, then merges with an explicit replacement map. Verify distinct section labels
and links in both directions; its stylesheet deliberately covers both heading IDs.
`css-preservation.epub` merges chapters and moves a decorative SVG background. Its
external and embedded styles use comment-separated compound selectors. Verify the
first heading retains its blue color and border, the second retains its green color
and border, and both retain the background after URL repair. EPUBCheck 5.4.0 reports
`CSS-008` on these comment-separated selectors, while browser inspection retains
their compound-selector behavior and Ace 1.4.6 passes its automated checks. A
diagnostic copy with only the empty selector comments removed passes EPUBCheck;
retain the original rejection separately rather than treating the control as
qualification of the original bytes.

`merge-selectors.epub` merges two chapters with colliding heading IDs and shared
stylesheets. The second heading receives a new ID and green styling; the first and
unmerged third headings remain blue. Verify all three borders and the links between
chapters. The merged chapter uses separate private copies of each chapter's stylesheet
and import; the third chapter retains the original stylesheet paths.

`merge-nested-selectors.epub` exercises the same colors, borders and links with
nested heading rules, a nested media condition and interleaved declarations. The
custom property `--heading-data` retains its literal `#heading` value. Assess validator
acceptance and reader support separately; the writer preserves nesting rather than
flattening these rules.
EPUBCheck 5.4.0 reports `CSS-008` for this fixture's nested media rule and
custom-property block syntax. Ace 1.4.6 passes its automated checks. Verify the
expected styles and literal custom-property value separately in reading systems;
automated checks do not qualify the nested fixture for EPUBCheck-gated delivery
or native readers.

```sh
dotnet run --project Build/Epub/Fixtures/EpubFixtureGenerator.csproj -- \
  /path/to/task-evidence/typography
```

Supply a new directory. The generator writes EPUB files, a manifest with SHA-256
hashes and native preflight results, and expanded content with standards-mode HTML
previews. A native preflight error fails the run; unchecked scopes remain explicit
in the manifest. Fixture accessibility metadata describes the actual content and
known limits without declaring certification. The previews keep
the generated content and profile stylesheet but add simulated reader CSS for
light, dark, and enlarged text. Serve this directory on loopback when inspecting
the preview paths listed in `manifest.json`; for example:

```sh
python3 -m http.server 8766 --bind 127.0.0.1 \
  --directory /path/to/task-evidence/typography
```

Check narrow and wide widths, wrapping, direction, text enlargement, and inherited
colors. Run the EPUB files through `validate_epub.py`, and use the same captured
bytes for native reader checks. Browser previews do not reproduce EPUB pagination,
reader preferences, font handling, or assistive-technology behavior. Retain their
results separately from native reading-system acceptance.

The `merged-styles` fixture exercises opt-in chapter stylesheet concatenation. The
second chapter's paragraph color and border style apply to both chapters, while the
first chapter's border radius remains. Validate the EPUB, inspect wide and compact
rendering, and exercise both repaired chapter links. Browser rendering of extracted
XHTML proves this cascade example; it does not qualify native reader presentation.

`merge-relationships.epub` exercises exact and whitespace-token selectors for
ARIA labels, table headers and microdata `itemref` values during chapter merge.
The first and unmerged third sections and table cells stay blue; the second section
and its cell are green. All sections retain their six-pixel left and three-pixel
top borders. Follow the third-chapter link and return to the renamed second heading.
The linked stylesheet import is privately cloned, while the third chapter keeps
the original rules. Independent validator and reader results are separate from
native preflight and source-preservation checks.

`merge-resource-selectors.epub` moves the second chapter out of a subdirectory
and repairs exact and partial selectors for local links, citation URLs and an image source.
The two detail references exercise multi-value `:is(...)` expansion and
should retain double underlines. This requires reader support for `:is(...)`.
The first and third local links are blue; the second local link and quotation are
green. Local links retain three-pixel bottom borders, and the image retains its
four-pixel green border. The first chapter's forward link remains brown with a
two-pixel dotted border after its target is repaired. Follow the third-chapter link
and return to the renamed second heading. Shared CSS imports receive separate
private copies for both chapters. These are expected reader
checks; passing native preflight or an independent validator does not prove rendering.


`merge-body-scopes-source.epub` and `merge-body-scopes.epub` retain the two
body IDs, classes, inline styles and data attributes on separate containers. Check
navy text on pale blue in the first chapter, maroon text on pale pink in the second,
borders, wrapping at narrow widths, reciprocal body-ID links and both TOC entries.
The managed rendering check does not establish native reader or screen-reader
behavior, and moving body styles to containers can change viewport box layout.

`merge-metadata-source.epub` and `merge-metadata.epub` exercise explicit package
refinement retargeting. Inspect manifest and spine targets, the transferred spine
ID, retained metadata IDs, nested refinements, linked descriptions and both TOC entries.
Both descriptions intentionally apply to the merged chapter. Validators can check
the package contract; they cannot decide whether those assertions are appropriate
for a publisher's combined chapter.

`merge-matter-source.epub` and `merge-matter.epub` exercise front/body-matter
partition preservation through chapter merging. Compare the body declarations in
the source with the two enclosing sections in the merged chapter, reciprocal links,
and both navigation entries. EPUBCheck and Ace check package and automated
accessibility rules; intended-reader presentation and assistive-technology behavior
need separate qualification.

`merge-language.epub` merges an English chapter with an Arabic chapter while
retaining an explicit French/LTR passage. Check Arabic right-to-left paragraph
layout, French left-to-right layout, reciprocal links and both TOC targets. Inspect
screen-reader language switching separately. The wrapper carries source language
and direction; package conformance and automated accessibility checks do not prove
native presentation or speech behavior.
