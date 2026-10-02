# iWork reader corpus

These package fixtures prove the bounded Pages, Numbers, and Keynote reader against files produced by multiple iWork generations. They remain read-only test inputs; OfficeIMO does not rewrite them.

| Folder | Upstream | Revision | Producer evidence | License |
|---|---|---|---|---|
| `nim-iwork` | [nim-iwork](https://github.com/halcyon-oss/nim-iwork) | `60a8e875692ac934956cdb5b39f88e557c97cca4` | Pages, Numbers, and Keynote 14.5 fixtures | MIT, copyright Alfred |
| `iwork-converter` | [iwork-converter](https://github.com/obriensp/iwork-converter) | `be4828260466c9c022ec12f9dc9dbfeb15ab1dea` | Pages 14.1, Numbers 11.1, and Keynote 8.1 fixtures | MIT, copyright Steve Dunham |
| `numbers-parser` | [numbers-parser](https://github.com/masaccio/numbers-parser) | `1c6c5c3d2e29a9abb601596678089f0a6c85d64c` | Numbers 15.1 fixture plus independently produced formula and merged-range workbooks | MIT, copyright Jon Connell |
| `picodocs` | [PicoDocs](https://github.com/PicoMLX/PicoDocs) | `5c18743d3d8120a76da124bd512a3cf5bcc28e82` | Pages 14.4.1 package produced by Pages 14.5 with sections, headers, footers, an image, hyperlinks, and three editable tables | MIT, copyright Pico MLX |
| `keynotekit` | [KeynoteKit](https://github.com/memfrag/KeynoteKit) | `5e01f1e061b608e16c4480444d8e04790a625b34` | Keynote 15.2.1 image and editable-table fixtures independently maintained as parser/writer regressions | 0BSD, copyright Martin Johannesson |

The complete upstream license notices are reproduced in `OfficeIMO.IWork/THIRD-PARTY-NOTICES.md`. Fixture provenance and expected semantic assertions live beside the executable corpus tests in `OfficeIMO.IWork.Tests`.

## Pages inline attachment positions

`pages-inline-anchors.json` records five image/table attachment positions from the unmodified `iwork-converter/a.pages` and `picodocs/sample-v14.4.pages` fixtures. The existing provenance, hashes and license notices above apply. Positions are UTF-16 offsets in native text storage, before object-marker removal. The manifest is extracted through the independent numbers-parser 4.19.0 schemas:

```bash
python Build/IWork/extract-pages-inline-anchors.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/pages-inline-anchors.json
```

OfficeIMO compares native attachment/drawable identities and offsets with this evidence. Saved/reopened DOCX and Reader tests check the representative fixture's object order. This qualifies source attachment decoding and destination placement; native Apple export, pagination, wrapping and rendered equivalence remain unqualified.

## Pages automatic row sizing

`pages-table-sizing.json` records the native automatic-resize setting and individual row heights of all three tables in the unchanged `picodocs/sample-v14.4.pages` fixture. It uses the same pinned source, hash and MIT provenance above. Reproduce the manifest with independent numbers-parser 4.19.0 schemas:

```bash
python Build/IWork/extract-pages-table-sizing.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/pages-table-sizing.json
```

Saved/reopened DOCX retains the declared 16.5-point heights as minimum constraints, so wrapped cell content can increase row height. The embedded package preview provides a bounded visual counterexample to treating these rows as fixed. This qualifies the source setting and destination constraint; native Apple export, complete styling and pagination remain unqualified.

## Pages selected cell fills and layout

`pages-cell-fills.json` records 64 selected modern cell-style keys, parent chains, solid/no-fill declarations, four-sided padding and vertical alignment across the same three native Pages tables. It includes four empty cells whose explicit fills must survive sparse projection. Reproduce it with the pinned independent numbers-parser 4.19.0 schemas:

```bash
python Build/IWork/extract-pages-cell-fills.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/pages-cell-fills.json
```

Saved/reopened DOCX compares each selected fill, padding side and vertical alignment with this manifest. All 64 selected cells have middle alignment and four-point padding. Empty fill declarations clear parent colors; table-role defaults and banding are outside this evidence. Native Apple export and complete appearance remain unqualified.

## Cross-table rectangular formulas

`numbers-parser/cross-table-formulas.numbers` is the unmodified upstream `tests/data/create-formulas.numbers` at revision `1c6c5c3d2e29a9abb601596678089f0a6c85d64c`, covered by the existing numbers-parser MIT notice. Its adjacent JSON manifest records sixteen `COUNTA` formulas with cross-table rectangular references, all combinations of absolute and relative endpoint coordinates, native target UUIDs and cached numeric values. `Build/IWork/extract-cross-table-formula-fixture.py` reads the pinned package through numbers-parser 4.19.0 to reproduce the manifest without rewriting the source:

```bash
python Build/IWork/extract-cross-table-formula-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.json --whole-axis-output OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/whole-axis-formulas.json
```

OfficeIMO tests compare all sixteen reconstructed source ranges and saved/reopened XLSX expressions with the manifest, preserving mixed endpoint flags and numeric caches. Separate boundary tests cover forward targets, normalized colliding names, and missing, inactive or ambiguous identities; unresolved targets never become complete local formulas. This qualifies the finite rectangular subset through independent reference and cache evidence.

`whole-axis-formulas.json` uses the same hashed native package for 59 row/column references, including coordinate-backed header-name aliases. The extractor resolves axes and target header/footer metadata through numbers-parser, computes `SUM`, `COUNT` or `COUNTA` over the selected current body, and verifies every result against the native cache. Saved-output tests compare exact fixed body bounds, endpoint flags, caches and the typed approximation diagnostic. Synthetic cases cover footers, normalized colliding names, ambiguous metadata and empty bodies. Native labels and automatic expansion are not preserved in fixed XLSX ranges. Native Apple export, appearance, recalculation and cache freshness remain unqualified.

## Formula function identities

`numbers-parser/function-identities.json` records 56 native function identifiers from the pinned numbers-parser 4.19.0 function map, plus the expressions, function nodes and typed caches of three cases in the unmodified `cross-table-formulas.numbers` fixture. OfficeIMO checks every selected identity, rejects unknown identifiers and invalid argument counts, and saves/reopens the native nested `OR`, `POWER`, and cross-table `TEXTJOIN` cases. The `TEXTJOIN` descriptor records independently resolved coordinates and a join of current referenced values that agrees with its cache. Its saved-output test checks the XLSX compatibility prefix, typed string cache, and recalculation after changing a referenced value and testing both empty-cell policies. The additional `MINA`, `MAXA` and `AVERAGEA` identities use the independent map; synthetic native-format cases check saved output separately from the three unmodified document cases. Sample argument counts exercise reconstruction; they do not qualify evaluation for arbitrary argument types. Reproduce the evidence without rewriting the package:

```bash
python Build/IWork/extract-function-identities.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/function-identities.json
```

The existing numbers-parser MIT notice covers the provider and fixture. The manifest does not qualify Apple export, appearance, recalculation or cache freshness.

## Empty hidden-state extents and disabled filters

`empty-hidden-states.json` describes five table models in the unchanged `nim-iwork/simple.numbers`, `picodocs/sample-v14.4.pages` and `keynotekit/tabledeck-v15.2.1.key` packages. Their pinned provenance and license notices above apply. Independent numbers-parser 4.19.0 schemas verify empty base/summary hidden states, correct axis directions, zero hidden counts and ten disabled filter sets without rules. Reproduce the manifest without rewriting the packages:

```bash
python Build/IWork/extract-empty-hidden-states.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/empty-hidden-states.json
```

OfficeIMO checks the source hashes, selected model identities and absence of new visibility warnings or declaration failures. This qualifies the empty/disabled path. The base user-hidden column qualification below supplies independent positive selection evidence; active filtering, collapsed groups and Apple export equivalence remain unqualified.

## Individual table dimensions

`numbers-parser/individual-dimensions.numbers` is produced by numbers-parser 4.19.0 using `Build/IWork/create-dimension-fixture.py`. It declares three row heights (20, 10, and 30 points) and two column widths (40 and 20 points), with source labels in the first column. The generator reopens the package through the independent producer; OfficeIMO tests read the declared dimensions and save/reopen the XLSX result. This fixture also exercises the valid seven-field tile envelope with the wide-row flag. It qualifies declared dimensions, not Apple automatic row sizing or visual equivalence.

To recreate the semantic fixture in an isolated Python 3.10+ environment, install `numbers-parser==4.19.0`, then run:

```bash
python Build/IWork/create-dimension-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/individual-dimensions.numbers
```

The producer creates new package identifiers and timestamps, so regenerated bytes can differ. The committed fixture has the checksum recorded below. The existing numbers-parser MIT notice covers its bundled template and implementation.

## Numeric formats

`numbers-parser/number-formats.numbers` is produced by numbers-parser 4.19.0 using `Build/IWork/create-number-format-fixture.py`. Its adjacent generated JSON manifest records the package hash, thirteen numeric values, exact source coefficient/exponent text, selected format metadata and independent producer display strings. Cases cover decimal precision, automatic formatting, grouping, all four negative styles, percentage scaling and zero. The generator reopens the package through the producer before writing the oracle.

```bash
python Build/IWork/create-number-format-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/number-formats.numbers
```

OfficeIMO tests compare metadata, save/reopen XLSX numeric values and compare image-snapshot text with the producer. The fixture also exposes finite Decimal128 values above fifteen significant digits; their exact source text and explicit approximation survive reconstruction. This evidence does not qualify Apple native rendering or full style fidelity. Regenerated package identifiers and timestamps can change the bytes; compare the semantic manifest cases when reproducing the fixture.

## Currency formats

`numbers-parser/currency-formats.numbers` is produced by numbers-parser 4.19.0 using `Build/IWork/create-currency-format-fixture.py`. The adjacent generated manifest records eleven currency cases across six identifiers, precision, grouping, negative styles, accounting settings, and independent source display strings. Destination display strings specify the portable identifier-prefix contract, including its explicit approximation report; they are not Apple appearance oracles.

```bash
python Build/IWork/create-currency-format-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/currency-formats.numbers
```

The generator saves and reopens the package before recording metadata and source display text. OfficeIMO checks source semantics and saved/reopened XLSX numeric values, image text, and red styles. The existing numbers-parser MIT notice covers the template and implementation. Regenerated package identifiers and timestamps can change bytes; compare the semantic manifest cases.

## Scientific formats

`numbers-parser/scientific-formats.numbers` is produced by numbers-parser 4.19.0 using `Build/IWork/create-scientific-format-fixture.py`. Its generated manifest records fourteen scientific selections, exact source numeric text, producer-decoded values, correctly rounded portable values, and selected mantissa precision. Twelve explicit-precision display strings come from the reopened producer. Its automatic sentinel is rendered as 253 fractional places, so the two automatic cases qualify metadata only; destination expectations describe OfficeIMO's approximate display contract.

```bash
python Build/IWork/create-scientific-format-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/scientific-formats.numbers
```

The scientific and numeric-format generators share `fixture_number_values.py` to derive portable numeric values and precision evidence from the stored finite Decimal128 coefficient and exponent. OfficeIMO tests compare source metadata and exact numeric values, then save/reopen XLSX and check display text. This is independent-producer evidence, without an Apple native export or appearance claim. Regenerated identifiers and timestamps can change package bytes; compare semantic manifest cases. The existing numbers-parser MIT notice covers the template and implementation.

## Fraction formats

`numbers-parser/fraction-formats.numbers` is produced by numbers-parser 4.19.0 using `Build/IWork/create-fraction-format-fixture.py`. Its generated manifest records twenty-six cases across all nine denominator modes, stored numeric precision, source metadata, and reopened producer display strings. Eighteen display examples qualify the independent oracle. Eight destination alternatives cover negative whole parts, rollover, and midpoint ties where the producer's output loses information, remains unnormalized, or differs in rounding policy. Those alternatives qualify OfficeIMO's reported display approximation, without claiming producer or Apple appearance equivalence.

```bash
python Build/IWork/create-fraction-format-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/fraction-formats.numbers
```

The generator reuses `fixture_number_values.py` for exact Decimal128 text and portable numeric values. The near-quarter-midpoint example retains the value actually stored by the writer, which lies below the midpoint. Tests compare saved/reopened XLSX numeric types, values, format codes, and image-snapshot text. Regenerated identifiers and timestamps can change bytes; compare semantic manifest cases. The existing numbers-parser MIT notice covers the template and implementation.

## Independent Apple exports

`native-exports/numbers-formulas-v14.5.json` records exports of `numbers-parser/test-10-formulas.numbers` made with Apple Numbers 14.5 (build 7045.0.17). The unmodified XLSX and PDF references include artifact hashes, export settings, source licensing, font provenance, and qualification limits. Numbers exports one worksheet per table and inserts a title row; the manifest accounts for that row when comparing source coordinates and formulas.

`IWorkAppleExportQualificationTests` saves and reopens OfficeIMO's XLSX, compares all 28 formula expressions with the independent export, and compares 24 stable typed cached values. Two volatile `NOW` caches are excluded because Apple recalculates them; two error formulas have no native cached value. Generic source error text is retained with an approximation diagnostic and is not qualified as a native Excel error value. The PDF supplies a two-page visual reference; rendered equivalence is not yet qualified. Its axis-aligned table grids have 98-point columns and 20.07-point rows. A separate saved-output check compares declared column widths and bounds the difference between declared row heights, Apple XLSX heights and the measured PDF grid; it does not implement Apple automatic row sizing.

The `pdf.tableGeometry` manifest field is generated from the hashed PDF by the opt-in `pdfplumber` tool (version 0.11.9 for this evidence):

```bash
python Build/IWork/update-numbers-export-geometry.py OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-formulas-v14.5.json
```

The extractor accepts this pinned two-table fixture only, verifies the PDF hash and checks grid line counts against its source coordinate mapping before updating the manifest. It is not part of normal restore or runtime dependencies.

## Fixture checksums

| Fixture | SHA-256 |
|---|---|
| `iwork-converter/a.key` | `929347827a7478c123dd3e3828e9751b5cf2ae977d2edd2a5f0774014fc4fefc` |
| `iwork-converter/a.numbers` | `43afd5a01cff9283fea5a11716d96f696b5853b75f82a3fe4469f240dfb82947` |
| `iwork-converter/a.pages` | `8481e3071c8ea1cc9543354bcd1ff66f79e6c2c8686096fa3ae60ab96fd49ebf` |
| `keynotekit/imagedeck-v15.2.1.key` | `a9af589197588e04ee52388b0aa6c2dad1110e5d6db814b58afe543831cf2128` |
| `keynotekit/tabledeck-v15.2.1.key` | `384962b1fff18abc5a901b59dc5f8820c2a959977f18f90dc9cd10095bdd0a56` |
| `nim-iwork/simple.key` | `ba95755df82ceb0ca834e1e03e2777c34fad906320d8336b4f3fefc6b48607eb` |
| `nim-iwork/simple.numbers` | `d0b00d9cae5985cccaa3b2fb251fae92eb0e38360fb4b5df8b4350eb658f752b` |
| `nim-iwork/simple.pages` | `5aee6d03277d2db2104f593e64afe081dec539f0117b97124b6f99158124c93e` |
| `numbers-parser/individual-dimensions.numbers` | `ccecc5494a71b9943e9a37641f843d0e8ce4e7c3531d72a50031dd7c23b38762` |
| `numbers-parser/currency-formats.numbers` | `6f45fa942ab26b2e9a3c2487fec491d5c43e85cfe1c191520134d5b64ac61475` |
| `numbers-parser/number-formats.numbers` | `1fb277a6897c0fc4387cd50a51c57de0a08323b694b98a5926b9bc5190717b1f` |
| `numbers-parser/scientific-formats.numbers` | `23bfaa5cf394c46ec5192e9612aca75c29f00e84331f11ffc338ef338aa5847d` |
| `numbers-parser/fraction-formats.numbers` | `da140e7eeac3122505690af655a5896aa056573dd15826a2f84d035388fbff55` |
| `numbers-parser/issue-102-v15.1.numbers` | `88a9fa7be095d03004478393a87a4a97602d7468f839d067ec9118c524c55176` |
| `numbers-parser/cross-table-formulas.numbers` | `9371c5b1d6ee4dfa17569097f064eba9c67f804d88b48638efbbeeb459d07dd4` |
| `numbers-parser/test-10-formulas.numbers` | `dd85bad68898ce5b065f277c0b9be1f3c32d696e3baa6b09d3614bbd35a5249f` |
| `numbers-parser/test-9-merges.numbers` | `d640c0012d629834161827cb2f564d0966d24e159586f82426a69c12a8f334cf` |
| `picodocs/sample-v14.4.pages` | `4714477138d0a4090fc2ee2ba2ebb6adcd0fb6ce20a28897a6247a8e17d1ddce` |

## Keynote unbanded table-fill defaults

`keynote-table-fill-defaults.json` records the unchanged Keynote 15.2.1 table fixture's disabled banding and explicit no-fill role declarations through pinned independent numbers-parser 4.19.0 schemas. Reproduce it with:

```sh
python Build/IWork/extract-keynote-table-fill-defaults.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/keynote-table-fill-defaults.json
```

All nine saved/reopened PPTX cells preserve explicit no-fill and suppress destination theme backgrounds. This qualifies source declarations and saved properties; role intersections, banded appearance, Apple exports and complete styling remain unqualified.

## Native Numbers table banding

`native-exports/numbers-banding-v14.5.numbers` is an OfficeIMO-authored fixture saved by Numbers 14.5, covered by the repository MIT license. Three eight-row tables vary the header-row count from zero to two and retain one header column and one footer row. Row seven's “Footer” label is body content; the actual footer is the empty eighth row. No selected cell-fill overrides are stored.

The matching Apple XLSX and three-page PDF exports qualify the alternating body-row pattern and region intersections. The XLSX inserts a table-title row. The manifest records all 72 native position fills, source and export hashes, producer version, export settings and limits. Reproduce the source/schema and XLSX oracle extraction with pinned opt-in numbers-parser 4.19.0:

```sh
python Build/IWork/extract-numbers-banding.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-banding-v14.5.json
```

This evidence covers opaque sRGB region and band fills in this fixture. It does not establish complete appearance, selected override behavior in Apple exports, other producer versions, or Pages and Keynote banded exports.

## Native Numbers function exports

`native-exports/numbers-functions-v14.5.json` records two OfficeIMO-authored Numbers 14.5 fixtures and their Apple XLSX exports. The matching intake workbooks contain cache-free OOXML formulas built with Python standard-library ZIP/XML, independently of OfficeIMO. Numbers imports, evaluates and writes the native files and exports. All assets are covered by the repository MIT license.

The manifest records 26 expressions, native function IDs and argument counts, numeric cache comparisons, producer version and artifact hashes. It exercises `OFFSET` as a scalar and as a `SUM` range argument, `PROB` exact and interval bounds, zero probabilities and invalid distributions, and fixed and variable `RANDBETWEEN` bounds. Reproduce the manifest with pinned opt-in numbers-parser 4.19.0:

```sh
python Build/IWork/extract-numbers-functions.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/numbers-functions-v14.5.json
```

The fractional random cases returned zero in this native import and remain outside local evaluation qualification. Random caches are snapshots. Native error cells have no XLSX error cache, so this evidence does not qualify error-code equivalence. Saved/reopened conversions retain the expressions and diagnosed caches; fresh shared-owner evaluation and edits to referenced inputs are tested separately. The PDF is a native reference, not a qualification of complete layout, implicit intersection, dynamic spills or other producer versions.

## Base user-hidden columns

`numbers-parser/user-hidden-columns.json` describes six hidden columns across three tables in the unchanged `numbers-parser/cross-table-formulas.numbers` package. Its pinned upstream revision and MIT notice above apply. Build metadata records Numbers `M14.3-7042.0.76-4` and the Blank 11.2 template. Independent numbers-parser 4.19.0 schemas resolve selected UUIDs through nonidentity column permutations and verify their inverse vectors, zero legacy hidden counts and empty hidden cells. Reproduce the manifest:

```bash
python Build/IWork/extract-user-hidden-columns.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/user-hidden-columns.json
```

OfficeIMO checks the source checksum, recovered positions and every column’s saved XLSX hidden attribute. Synthetic format-boundary cases cover positive rows, populated hidden cells, Reader inclusion and DOCX/PPTX partial policy. This qualifies base user-hidden column decoding and XLSX metadata; Apple export/render equivalence, active filters, pivot hiding, summary states and collapsed groups remain open.

## Single-cell and endpoint formulas

`numbers-parser/single-cell-formulas.numbers` and `endpoint-formulas.numbers` are unchanged upstream `tests/data/test-all-formulas.numbers` and `test-extra-formulas.numbers` at revision `d3836ebda1110b5c13b8722642ca61111fe8e865`. The existing numbers-parser MIT notice applies. The first package records an XLSX import followed by Numbers 11.1 through 13.1 saves; the second records CSV import and Numbers 11.1/12.2 saves. Their adjacent manifests retain source checksums and build metadata. Reproduce each with independent numbers-parser 4.19.0:

```bash
python Build/IWork/extract-single-cell-formulas.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/single-cell-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/single-cell-formulas.json
python Build/IWork/extract-single-cell-formulas.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/endpoint-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/endpoint-formulas.json
```

The manifests resolve node-36 identities, signed relative offsets and mixed absolute flags through independent native schemas. Current-cell computations agree with all eight reference caches. The same extractor records six scalar expressions and native function nodes: `EXACT` (matching and case-different strings), `LOWER` and `TRIM` in the first package, and `NOT` plus `UPPER` in the second. Independent numeric/ASCII computations agree with their Boolean/text caches. OfficeIMO checks the first package’s five reconstructed reference expressions, four scalar expressions and typed caches. Its pre-1900 dates still require whole-workbook XLSX fallback; the fixture is not edited to bypass that safety rule. The first manifest also records 36 numeric expressions: nine `INT`, `MOD` and `SQRT` cases plus sixteen `SIGN`, `TRUNC`, `ROUNDUP` and `ROUNDDOWN` cases with negative literals and nested `ABS`; independent numeric computations agree with their reconstructed source formulas and numeric caches. Independent decimal computations cover signed digit arguments, omitted `TRUNC` digits and positive/negative/zero signs. Equivalent exponent and decimal literal spellings are normalized when comparing reconstructed formulas. Eleven additional `EXP`, `LN`, `LOG` and `LOG10` expressions cover nesting, exponentiation and omitted/explicit bases. Their exact producer caches agree with independent transcendental computations within the manifest’s recorded relative tolerance of 1e−14 and absolute tolerance of 1e−15; this does not replace the source cache with a freshly calculated value. Synthetic native-format inputs separately qualify editable XLSX caches and local recalculation for these eleven functions, including explicit `TRUNC` digits. The second package qualifies saved/reopened `COUNTBLANK`, `MAX`, single-cell `OFFSET`, `NOT` and `UPPER` expressions with numeric/Boolean/text caches. Recalculation follows edited references and logical/text literals. Synthetic packages cover all absolute/relative combinations, forward and normalized names, unresolved identities, malformed coordinates, endpoint whitespace and incompatible endpoint targets. Apple export, rendering and broader native recalculation remain unqualified.

## Native root cell comments

`cell-comments/native-roots.numbers` is the unchanged MIT-licensed `fixtures/olekristensen-v26.3-demo06-formulas-round8.numbers` from [cupertino-files](https://github.com/den-frie-vilje/cupertino-files) at revision `6879e6ed49e0eed7e7f393ff4ee558dc2dbec561`. The upstream attribution records a macOS Numbers save. The adjacent MIT notice retains copyright 2026 Ole Kristensen. The fixture contains the author's published test notes, display name and native author identifiers; it is retained deliberately as native comment evidence.

With independent numbers-parser 4.19.0 installed, reproduce its hash-pinned manifest:

```sh
python Build/IWork/extract-cell-comments.py OfficeIMO.TestAssets/Documents/IWorkCorpus/cell-comments/native-roots.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/cell-comments/native-roots.json
```

The three selected roots qualify exact text, display author, timestamps, source record identities and saved/reopened XLSX cell anchors. Synthetic packages cover empty commented cells in all three formats, unresolved references, duplicate keys/text, unsupported replies and text/catalog limits. The fixture contains no replies and is not an Apple export or rendered-appearance oracle.

## Table catalog type and kind evidence

`table-catalog-contract.json` records independent registry aliases, source hashes and declared catalog links across the checked-in native-format corpus. Reproduce it with opt-in numbers-parser 4.19.0:

```sh
python Build/IWork/extract-table-catalog-contract.py OfficeIMO.TestAssets/Documents/IWorkCorpus OfficeIMO.TestAssets/Documents/IWorkCorpus/table-catalog-contract.json
```

The 312 links across 24 packages all use type `6005` and explicitly declare the kind expected by the owning store field. The independent registry also maps `6201` to `TST.TableDataList`; synthetic selected-value cases exercise that alias for strings, formulas, rich text, styles, comments and number formats. There is no native `6201` sample in this corpus. Model records may be inactive, and the inventory does not qualify selection, Apple exports or rendered appearance.
