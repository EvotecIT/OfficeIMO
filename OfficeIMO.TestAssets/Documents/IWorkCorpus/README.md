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

## Cross-table rectangular formulas

`numbers-parser/cross-table-formulas.numbers` is the unmodified upstream `tests/data/create-formulas.numbers` at revision `1c6c5c3d2e29a9abb601596678089f0a6c85d64c`, covered by the existing numbers-parser MIT notice. Its adjacent JSON manifest records sixteen `COUNTA` formulas with cross-table rectangular references, all combinations of absolute and relative endpoint coordinates, native target UUIDs and cached numeric values. `Build/IWork/extract-cross-table-formula-fixture.py` reads the pinned package through numbers-parser 4.19.0 to reproduce the manifest without rewriting the source:

```bash
python Build/IWork/extract-cross-table-formula-fixture.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.json --whole-axis-output OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/whole-axis-formulas.json
```

OfficeIMO tests compare all sixteen reconstructed source ranges and saved/reopened XLSX expressions with the manifest, preserving mixed endpoint flags and numeric caches. Separate boundary tests cover forward targets, normalized colliding names, and missing, inactive or ambiguous identities; unresolved targets never become complete local formulas. This qualifies the finite rectangular subset through independent reference and cache evidence.

`whole-axis-formulas.json` uses the same hashed native package for 59 row/column references, including coordinate-backed header-name aliases. The extractor resolves axes and target header/footer metadata through numbers-parser, computes `SUM`, `COUNT` or `COUNTA` over the selected current body, and verifies every result against the native cache. Saved-output tests compare exact fixed body bounds, endpoint flags, caches and the typed approximation diagnostic. Synthetic cases cover footers, normalized colliding names, ambiguous metadata and empty bodies. Native labels and automatic expansion are not preserved in fixed XLSX ranges. Native Apple export, appearance, recalculation and cache freshness remain unqualified.

## Formula function identities

`numbers-parser/function-identities.json` records 32 native function identifiers from the pinned numbers-parser 4.19.0 function map, plus the expressions, function nodes and typed caches of two cases in the unmodified `cross-table-formulas.numbers` fixture. OfficeIMO checks every selected identity, rejects the displaced unqualified identifiers, and saves/reopens the native nested `OR` and `POWER` cases. Sample argument counts exercise reconstruction; they do not qualify evaluation for arbitrary argument types. Reproduce the evidence without rewriting the package:

```bash
python Build/IWork/extract-function-identities.py OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/cross-table-formulas.numbers OfficeIMO.TestAssets/Documents/IWorkCorpus/numbers-parser/function-identities.json
```

The existing numbers-parser MIT notice covers the provider and fixture. The manifest does not qualify Apple export, appearance, recalculation or cache freshness.

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
