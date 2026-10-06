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
and a successful full decode; the runner itself does not decode audio. These results
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
inputs. Accessibility fixtures exercise unknown status, publisher contact
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
  /path/to/new-onix-evidence/taxed.onix
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
