# Entity recognition

`OfficeIMO.Core` provides `OfficeIMO.Entities` for local identifier validation, text recognition, and masking. Use a scheme when you know which field you have. Use the finder when you need to locate values in prose or labelled document segments.

## Validate a known field

```csharp
using OfficeIMO.Entities;

EntityValidationResult result = FinancialSchemes.Iban.Validate("GB82 WEST 1234 5698 7654 32");
if (result.IsValid) {
    string canonical = result.Normalized!;
    string printed = result.Scheme.Format(canonical);
    EntityDetails? details = result.Scheme.Decode(canonical);
}
```

`Validate` reports a typed status and the canonical spelling when the format parsed. `TryNormalize` and `Normalize` produce a value only after validation succeeds. A passing check establishes local format and checksum validity. Check `CheckDigit` to tell a passing checksum from a format-only result. `Check` identifies the applied algorithm and its source, and `CheckProvenance` distinguishes issuer-published and community-documented checks. For example, Spanish VAT applies the issuer-published NIF/NIE check to personal numbers and a community-documented check to company numbers.

`EntityCatalog.Default` includes the following schemes. Register selected packs with `EntityCatalogBuilder` or implement `EntityScheme` and `IEntityRecognizer` for your own identifiers.

| Pack | Schemes |
| --- | --- |
| Poland | PESEL, NIP, REGON, KRS, identity card and passport |
| EU VAT | The 27 member states and Northern Ireland. Use `EL` for Greece and `XI` for Northern Ireland when looking up the VIES scheme. |
| United Kingdom | National Insurance number, Companies House number and GB VAT |
| United States | SSN and EIN |
| Financial | IBAN, BIC, payment card, ISIN and LEI |
| Internet | Email and URL |

The stage 1 profile accepts a UK National Insurance number without its suffix in a known field. The Companies House scheme covers formats with numeric tails, BIC uses assigned ISO 3166 country codes, and LEI requires `00` in positions five and six. URL recognition covers dotted hosts and `localhost`. Unprefixed VAT recognition in text is limited to contextual digit runs.

Validation is offline. Registration, ownership and deliverability require a separate verification step.

## Find and mask entities

```csharp
var catalog = new EntityCatalogBuilder().AddPoland().AddFinancial().Build();
var finder = new EntityFinder(catalog);
string text = "PESEL: 44051401359";
EntityFindResult result = finder.Find(text);

string masked = EntityMasker.Mask(text, result.Matches);
// PESEL: [NationalId]
```

Set `EntityFinderOptions` to filter by kind, country or scheme ID. For typed data or table columns, supply an `EntityTextSegment` with the label and optional location. When the full segment is a known typed value, set `structuredSource: true` to score a valid whole-value match at 1.0.

Segments are independent by default. Set `continuesPrevious: true` on a segment when its first character directly follows the previous segment's last character. The finder searches the joined text and returns a `Parts` entry for each source segment covered. Every part has a zero-based segment index, a UTF-16 start offset and a length. The match's `SegmentIndex`, `Start` and `Length` describe its first part, and `Location` comes from that segment.

The default confidence threshold is 0.50. A bare PESEL has confidence 0.40. A nearby `PESEL` keyword raises it to 0.80. `BestOnly` returns winners with declared contained readings and same-span alternatives. `All` returns every eligible reading and can be used to mask the union. Limits default to 10,000 returned matches and 16,777,216 input characters. Check `Truncated` before treating a search as complete.

Pass the original segments and findings to `EntityMasker.Mask(segments, result.Matches)` to get one masked string per input segment. Fixed masking writes one mask character per matched UTF-16 code unit. Partial masking applies the scheme to the full value and distributes its result across the parts. A kind label or keyed pseudonym replaces the first part and clears the remaining matched parts.

Masking defaults to personal and financial data. Overlapping and adjacent selections in the same logical text are written as one replacement. A partial mask over several selections uses fixed characters. A pseudonym over such a region combines the selected canonical values. The map returned by the string overload of `Pseudonymize` records the first written spelling of that region.

For pseudonyms, set `PseudonymKey` to a caller-owned secret key. The same canonical value and key produce the same token. Use `Pseudonymize` when you need the token map. Treat the map, normalized values and decoded attributes as sensitive. Match descriptions and finder diagnostics contain no source values.
