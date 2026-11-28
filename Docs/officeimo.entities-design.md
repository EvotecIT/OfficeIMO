# OfficeIMO entity recognition design

This is the architecture and API direction for recognizing identifiers and personal data in text: validating a value someone typed, normalizing it to one canonical form, formatting it for display, and finding such values inside documents. The validation, recognition and masking API is documented in the [entity recognition guide](officeimo.entities.md). The consumer sections describe future integrations.

Redaction, the Reader, Excel and CSV cleaning, invoicing, email, mail merge, PSWriteOffice and Studio all need the same answers: is `123-456-32-18` a valid Polish NIP, is this 11-digit number a PESEL or an order number, where are the IBANs on this page. Today each of those would grow its own regular expressions and checksums. One engine in `OfficeIMO.Core` answers them once, and each host asks it.

## Goals and limits

- **One model for every kind.** An email address, a PESEL, an EU VAT number and an IBAN are all *entities*: a kind, an optional country, a scheme, a normalized value and, when found in text, a span with confidence and evidence.
- **Grows by country and kind without rework.** A new identifier is one scheme class and one registration line. A new kind is a new `EntityKind` value. Neither changes the finder, the masker or any consumer.
- **Rules first, data later.** Identifiers with published check digits, email and URL need no data files and live in Core. Phones, postal addresses and company legal forms need large, versioned metadata and go to an optional `OfficeIMO.Entities.Data` package.
- **Precision before recall.** An 11-digit number is not a PESEL because its check digit happens to match. Findings carry the evidence behind them, and the default threshold keeps unlabelled weak matches out.
- **Values never leak.** Matches never print their value through `ToString()`, diagnostics or exceptions, and the engine never logs.
- **Small and dependency-free.** Core stays dependency-free and trim/NativeAOT-safe: no reflection, no runtime code generation, no regular expressions in recognizers, explicit registration.

Out of scope: named-entity recognition of people, organizations and places in free prose (that is a statistical model, not a rule), online verification against registries (VIES, GUS BIR, Companies House API; a host may add those, the engine never calls a network), and anything that infers sensitive attributes beyond what an identifier's specification encodes.

### Context: what others cover

OneContact (`recursive-ai-oss/onecontact-csharp`, MIT OR Apache-2.0, .NET 10) parses, normalizes and finds phones, email, web addresses, postal addresses, person names and company names, with libphonenumber, WHATWG, UTS #46, Public Suffix List and libaddressinput data. It does not cover national identifiers, tax and VAT numbers or financial identifiers. OfficeIMO's scope is the union: contact data *and* identifiers, recognized inside documents and wired into redaction, cleaning and AI preparation. Nothing is taken from that project beyond knowing what its README says it covers.

## Ownership and placement

| Capability | Owner | Notes |
| --- | --- | --- |
| Kinds, schemes, validation, normalization, formatting, decoding | `OfficeIMO.Core`, namespace `OfficeIMO.Entities` | Pure algorithms and small constant tables. |
| Rules-based recognizers (identifiers, email, URL), finder, overlap resolution, masking | `OfficeIMO.Core`, `OfficeIMO.Entities` | One shared tokenizer, no regular expressions. |
| Context keywords per language | `OfficeIMO.Core`, inside each country pack | A few hundred short strings in total. |
| Phones, postal address formats, company legal forms | `OfficeIMO.Entities.Data` (later) | Generated tables from libphonenumber metadata, Google address metadata and GLEIF ISO 20275 lists, each recorded in `Build/third-party-material.json`. |
| Applying findings to documents (redaction, cleaning, chunk metadata) | Existing format owners | They take `EntityMatch` lists; they never re-implement recognition. |
| CLI, MCP, PSWriteOffice, Studio | Existing hosts | Map parameters to `EntityFinderOptions` and present results. |

Budget for Core: stage 1 should add well under 100 files and under 150 KB to the assembly. Each country pack is one folder; the finder, tokenizer and masker are shared.

## Model

```text
EntityKind        open value type: Email, Url, Phone, PostalAddress, BankAccount, BankIdentifier,
                  PaymentCard, SecurityId, LegalEntityId, NationalId, TaxId, VatId, CompanyId,
                  IdentityDocument, ... (new kinds are just new values)
EntityScheme      one identifier definition: id "PL.PESEL", kind NationalId, country PL,
                  Validate / Normalize / Format / Decode, metadata
IEntityRecognizer finds candidates of one or more schemes in tokenized text
EntityCatalog     immutable set of schemes, recognizers and context keywords, built explicitly
EntityFinder      runs recognizers, scores, resolves overlaps, applies options
EntityMatch       span, kind, scheme, normalized value, confidence, evidence, alternatives
EntityMasker      replaces matched spans: label, partial, fixed character or keyed pseudonym
```

### Kinds

`EntityKind` is a `readonly struct` wrapping a case-insensitive string, shaped like `ChatRole` in Microsoft.Extensions.AI: well-known static values, a public constructor for new ones, value equality, and `ToString()` returning the name.

| Kind | Stage 1 schemes | Later |
| --- | --- | --- |
| `NationalId` | PL PESEL, UK NINO, US SSN | other national IDs per country pack |
| `TaxId` | PL NIP, US EIN | US ITIN, other taxpayer numbers |
| `VatId` | EU VAT (27 states + XI), GB VAT | CH UID-MWST, NO MVA |
| `CompanyId` | PL REGON (9/14), PL KRS, UK company number | other registers |
| `IdentityDocument` | PL identity card, PL passport | ICAO 9303 machine-readable zones |
| `BankAccount` | IBAN | national BBAN formats |
| `BankIdentifier` | BIC | sort codes, routing numbers |
| `PaymentCard` | card number (Luhn) | |
| `SecurityId` | ISIN | CUSIP, SEDOL |
| `LegalEntityId` | LEI | |
| `Email` | email address | |
| `Url` | URL | |
| `Phone`, `PostalAddress`, `CompanyName` | | `OfficeIMO.Entities.Data` |

A kind answers "what sort of thing is this" for policies ("redact every `NationalId`"); the scheme answers "which identifier exactly". Countries are ISO 3166-1 alpha-2 codes, with `null` for international schemes (IBAN, LEI, email). The EU VAT scheme family uses the VIES prefixes (`EL` for Greece, `XI` for Northern Ireland) as its *prefix*, while `Country` stays the ISO code (`GR`, `GB`).

### Schemes

One `EntityScheme` subclass per identifier. Schemes are stateless singletons exposed through static classes (`PolandSchemes.Pesel`, `FinancialSchemes.Iban`) and through the catalog by id.

```csharp
namespace OfficeIMO.Entities;

public abstract class EntityScheme {
    public abstract string Id { get; }                // "PL.PESEL", "EU.VAT.DE", "IBAN"
    public abstract EntityKind Kind { get; }
    public abstract string DisplayName { get; }       // "PESEL"
    public virtual string? Country => null;           // ISO 3166-1 alpha-2
    public virtual string? Issuer => null;            // "Ministry of Digital Affairs"
    public virtual Uri? Specification => null;        // format or check documentation
    public virtual IReadOnlyList<EntityCheckInfo> Checks => [];  // every check the scheme can run
    public virtual int MinLength => 1;                // normalized length bounds
    public virtual int MaxLength => 64;
    public virtual EntityCheckKind Check => EntityCheckKind.None;  // Mod11, Mod97_10, Luhn, ...
    public virtual EntitySensitivity Sensitivity => EntitySensitivity.Personal;

    public abstract EntityValidationResult Validate(string? value);
    public bool IsValid(string? value) => Validate(value).IsValid;
    public bool TryNormalize(string? value, out string normalized);
    public string? Normalize(string? value);
    public virtual string Format(string normalized, EntityFormatStyle style = EntityFormatStyle.Display);
    public virtual EntityDetails? Decode(string normalized) => null;  // opt-in
    public virtual string MaskPartial(string normalized);
}
```

- **Validate** accepts the forms people write: surrounding whitespace, the scheme's own separators (spaces, hyphens, dots, non-breaking spaces), lower case where the scheme is case-insensitive, and an optional country prefix where one is customary (`PL` before a NIP). It returns a result, never throws for bad input.
- **Normalize** returns the canonical compact form: digits and upper-case letters only, with the prefix the scheme defines (`PL1234563218` for a Polish VAT number, `1234563218` for a NIP).
- **Format** renders a normalized value: `Compact`, `Display` (the issuer's printed grouping, such as `123-456-32-18` or `PL61 1090 1014 0000 0712 1981 2874`) and `Electronic` where it differs.
- **Check provenance** says where a check algorithm comes from. `IssuerPublished` means the issuer (or a ministry publishing the rules member states agreed to disclose) states it. `CommunityDocumented` means no issuer publication exists and the algorithm is taken from a documented description that matches real issued numbers. Provenance belongs to each check, not to the scheme, because one scheme can run different checks for different variants: `EU.VAT.ES` uses the Ministry of the Interior's NIF/NIE letter rule (issuer-published) for individuals and a community-documented check for companies. `EntityCheckInfo` describes each check (`Name` such as `"NIF letter"`, `Provenance`, `Source` URI). Validation results and matches name the check that actually ran and carry its provenance, so a host can treat a community-documented result as weaker or ignore it.
- **Decode** is opt-in and returns only what the specification encodes: PESEL birth date and sex, IBAN country and BBAN, card issuer category from the IIN's first digit, REGON 14's parent REGON 9. It never guesses.

```csharp
public sealed class EntityValidationResult {
    public bool IsValid { get; }
    public EntityValidationStatus Status { get; }      // Valid, Empty, InvalidCharacters, InvalidLength,
                                                       // InvalidFormat, InvalidCheckDigit, InvalidDate,
                                                       // NotAllocated, UnsupportedVariant
    public string? Normalized { get; }                 // set when the format parsed, even if the check failed
    public EntityScheme Scheme { get; }
    public EntityCheckOutcome CheckDigit { get; }      // Passed, Failed, NotDefined
    public EntityCheckInfo? Check { get; }             // the check that ran, null when none applies
    public EntityCheckProvenance CheckProvenance { get; }  // Check?.Provenance ?? None
}

public sealed class EntityDetails {
    public DateTime? BirthDate { get; }
    public EntitySex? Sex { get; }
    public string? Country { get; }
    public string? Variant { get; }                    // "REGON14", "Visa", "FR-old-format"
    public IReadOnlyDictionary<string, string> Properties { get; }
}
```

`NotAllocated` covers values the issuer states it never assigns (SSN area 000, 666, 900-999; NINO prefixes such as `BG`, `GB`, `NK`; EIN prefixes outside the IRS list). `InvalidDate` covers PESEL dates that do not exist. A PESEL dated in the future validates (the number is well formed) but the recognizer treats it as weak evidence, see [scoring](#scoring).

### Recognizers and the shared tokenizer

```csharp
public interface IEntityRecognizer {
    string Id { get; }
    IReadOnlyList<EntityScheme> Schemes { get; }
    EntityTokenShape Shapes { get; }                   // which token shapes it wants: DigitRun, AlphaNumericRun, At, Scheme
    void Recognize(EntityRecognitionContext context);  // reads tokens, calls context.Add(candidate)
}
```

The finder tokenizes each logical text chain once and shares its runs and anchors across recognizers.

- **Digit runs**: maximal sequences of digits joined by single separators (space, hyphen, dot, slash, non-breaking space, thin space), with the separator pattern recorded (`3-3-2-2`, `4 4 4 4`).
- **Alphanumeric runs**: contiguous ASCII groups and boundary information for IBAN, VAT, NINO, passport and ISIN shapes.
- **Anchors**: `@` positions for email and `://`, `mailto:` or `www.` positions for URLs. The email and URL recognizers expand outward under RFC limits.

The tokenizer identifies contiguous alphanumeric groups and maximal digit runs. Recognizers can extend an alphanumeric group across supported separators within the length allowed by the scheme. For example, the IBAN recognizer starts at `PL61` after the label and includes the following account groups. Candidates start and end at group boundaries. A digit-only candidate that leaves adjacent groups in the same run needs issuer formatting or a different separator to distinguish it from the adjacent grouping. Boundaries prevent a candidate inside a longer unbroken word or number. The finder also recognizes the declared NIP inside a Polish VAT number and the parent REGON inside a branch REGON.

Unprefixed VAT recognition in prose requires a numeric national part, a passing check and scheme-specific context. Known fields can validate an unprefixed alphanumeric value directly with the country scheme.

Recognizers are separate from schemes because one recognizer can serve many schemes (one EU VAT recognizer for 28 prefixes, one IBAN recognizer for every country) and because a scheme is useful without any finder (validating an Excel column, an invoice field or a PowerShell parameter).

### Candidates, evidence and matches

```csharp
public sealed class EntityMatch {
    public int Start { get; }
    public int Length { get; }                        // length of the first part as written
    public EntityKind Kind { get; }
    public EntityScheme Scheme { get; }
    public string Normalized { get; }                 // sensitive
    public EntityCheckInfo? Check { get; }            // the check that produced this match
    public EntityCheckProvenance CheckProvenance { get; }
    public double Confidence { get; }                 // 0..1
    public EntityEvidence Evidence { get; }           // flags
    public string? ContextKeyword { get; }            // "NIP", "nr rachunku"
    public string? ContextLanguage { get; }           // BCP 47: "pl", "en"
    public int SegmentIndex { get; }                  // input segment where the match starts (0 for a string)
    public EntitySegmentLocation? Location { get; }   // from the first segment covered
    public IReadOnlyList<EntityMatchPart> Parts { get; } // source segment index, start and length for each part
    public IReadOnlyList<EntityMatch> Alternatives { get; }  // other readings of the same span
    public IReadOnlyList<EntityMatch> Contained { get; }     // e.g. the NIP inside a PL VAT number
    public override string ToString();                // "TaxId/PL.NIP [12..25] 0.92", never the value
}

[Flags]
public enum EntityEvidence {
    None = 0,
    CheckDigitPassed = 1,
    NoCheckDefined = 1 << 1,
    IssuerFormatting = 1 << 2,      // grouping matches the issuer's printed form (NIP 3-3-2-2, IBAN groups of 4)
    CountryPrefix = 1 << 3,         // "PL", "DE", "GB" in front
    ContextKeyword = 1 << 4,        // keyword in the window before the value
    Label = 1 << 5,                 // host-supplied label: column header, form field, invoice element
    PlausibleDate = 1 << 6,         // PESEL date in the past and within 130 years
    KnownIssuerRange = 1 << 7,      // card IIN, EIN prefix
    AmbiguousReading = 1 << 8,      // another scheme also validates this span
    StructuredSource = 1 << 9,      // the value came from a typed field, not prose
    CommunityDocumentedCheck = 1 << 10  // the check passed, but its algorithm is not issuer-published
}
```

### Input segments

Hosts rarely have one flat string. The finder accepts either a string or a list of segments, each with optional location and label:

```csharp
public readonly struct EntityTextSegment {
    public EntityTextSegment(string text, EntitySegmentLocation? location = null, string? label = null,
        bool structuredSource = false, bool continuesPrevious = false);
}

public sealed class EntitySegmentLocation {
    public string? Part { get; }      // "word/document.xml", "Sheet1", "page 3"
    public string? Address { get; }   // "B7", "paragraph 12", "chunk 4"
    public int? Page { get; }
}
```

Excel passes each cell with its column header as `label`; the Reader passes chunks with their source location; invoicing passes a field value with the element name as `label` and the match gets `StructuredSource`.

Segments are independent by default. A segment with `continuesPrevious: true` joins the previous segment verbatim. Context and matches can cross those joins. The match's `SegmentIndex`, `Start` and `Length` give its first part's coordinates in UTF-16 code units, and `Location` comes from that segment. `Parts` lists the portion in every segment the match covers.

## Scoring

Confidence is a number from 0 to 1 built from the scheme's base strength and the evidence, with fixed, documented steps so results are explainable and stable across releases. The base reflects how likely a random string of that shape passes the check:

| Base class | Examples | Base |
| --- | --- | --- |
| Strong check over a distinctive shape | IBAN (ISO 7064 MOD 97-10 plus country length), LEI, ISIN, EU VAT with country prefix and check | 0.85 |
| Distinctive shape, no check | Email, URL, UK NINO pattern with valid prefix and suffix | 0.80 |
| Country prefix and VIES format, no documented check | an EU VAT variant for which no check is documented at all | 0.60 |
| Check digit over a plain digit run | PESEL, NIP, REGON, card number (Luhn) with a payment-card major industry identifier, EU VAT digits without prefix | 0.40 |
| Check digit over a short letter-digit shape | PL identity card and passport | 0.30 |
| Shape only | KRS, SSN, EIN, BIC, UK company number | 0.20 |

Checks use the same base across both provenance categories. Provenance is available in `CheckProvenance` and the `CommunityDocumentedCheck` evidence flag.

Adjustments, applied in order and clamped to [0, 1]:

- `Label` matching the scheme (`NIP`, `PESEL`, `IBAN`, `VAT ID`): +0.50. `ContextKeyword` within the window (default 48 characters before the value in the logical text, nothing but punctuation, spaces and short connecting words between): +0.40.
- `IssuerFormatting`: +0.10. `CountryPrefix` where optional: +0.10. `KnownIssuerRange`: +0.10.
- `PlausibleDate` missing for PESEL (future date or older than 130 years): -0.30.
- `AmbiguousReading` without disambiguating context: -0.15 for each competing reading.
- `StructuredSource`: the value is taken as given, so confidence equals 1.0 when valid.

The default `MinimumConfidence` is 0.50.

- `44051401359` alone in prose is not reported (0.40). `PESEL: 44051401359` is (0.80). The same number in a column headed `PESEL` is (0.90).
- `1234563218` alone is not reported as a NIP; `NIP 123-456-32-18` is (0.90 after formatting and keyword); `PL1234563218` is reported as an EU VAT number (0.85) with the NIP in `Contained`.
- A Luhn-valid 16-digit number starting with 3-6 (banking and financial major industry identifiers) reaches 0.50 and is reported; with `card` or `karta` nearby it reaches 0.90.
- `SSN 123-45-6789` is reported (0.60 + formatting); a bare `123-45-6789` is not (0.30).

Redaction hosts may lower the threshold (for example to 0.30) and accept more false positives; validation hosts that already know the column use `Label` or call `Scheme.Validate` directly.

## Overlaps and false positives

The finder resolves overlaps after scoring, deterministically:

1. **Containment declared by a scheme.** Some schemes wrap others: the PL VAT number contains the NIP, the REGON 14 contains a REGON 9, an IBAN's BBAN can contain a national account number. A scheme declares `Contains` relations; the outer match wins and the inner one moves to `Contained`. Consumers that redact use the outer span; consumers that index can still see the NIP.
2. **Same span, several valid schemes.** A 10-digit run can pass both NIP and a REGON-like check by chance, an 11-digit run both PESEL and a phone shape later. The highest confidence wins and the others go to `Alternatives`; each loses 0.15 for ambiguity unless context points to one (a `NIP` keyword lifts the NIP and not the others).
3. **Partial overlap.** Higher confidence wins; ties go to the longer span, then the stronger check class, then catalog registration order. The loser is dropped and counted in the result's diagnostics (count only, never the value).
4. **Policy.** `EntityOverlapPolicy.BestOnly` (default) returns the winners above. `EntityOverlapPolicy.All` returns every candidate above the threshold, for redaction tools that want the union of spans.

Additional false-positive guards: boundary rules on runs (above), dates and amounts (a run that parses as `dd.mm.yyyy` or carries a currency sign or decimal comma is not offered to digit-only schemes), repeated digits (`00000000000`, `1111111111`) rejected even if a check passes, and rejection of values the issuer marks as specimens.

## Context keywords

Each country pack registers keywords per scheme and language: `PESEL` (pl, en), `NIP`, `numer identyfikacji podatkowej` (pl), `tax identification number` (en); `REGON`; `KRS`; `nr dowodu osobistego`, `seria i numer dowodu` (pl); `IBAN`, `nr rachunku`, `numer konta` (pl), `account number` (en), `Kontonummer` (de); `USt-IdNr.` (de), `TVA intracommunautaire` (fr), `NIF-IVA` (es/pt), `partita IVA` (it), `BTW-nummer` (nl), and so on for each VIES language. Matching is case- and diacritic-insensitive on whole words, with an optional trailing `:`, `nr`, `no.` or `#`. Keyword lists are plain arrays inside the pack, so trimming removes packs nobody registers.

`EntityFinderOptions.Languages` limits which keyword languages count; null means all registered.

## Registration and catalog

Registration is explicit and AOT-safe: no assembly scanning, no attributes, no reflection.

```csharp
EntityCatalog catalog = new EntityCatalogBuilder()
    .AddPoland()          // PESEL, NIP, REGON, KRS, identity card, passport + pl keywords
    .AddEuVat()           // 27 member states + XI, with each state's keywords
    .AddUnitedKingdom()   // NINO, company number, GB VAT
    .AddUnitedStates()    // SSN, EIN
    .AddFinancial()       // IBAN, BIC, payment card, ISIN, LEI
    .AddInternet()        // email, URL
    .Add(new MyCustomerNumberScheme(), new MyCustomerNumberRecognizer())
    .Build();

EntityCatalog all = EntityCatalog.Default;        // every rules-based pack in Core
EntityScheme nip = catalog.GetScheme("PL.NIP");
EntityValidationResult r = catalog.Validate("PL.NIP", "123-456-32-18");
```

`EntityCatalog` is immutable and thread-safe; build it once and reuse it. Later, `OfficeIMO.Entities.Data` adds `.AddPhones()`, `.AddPostalAddresses()` and `.AddCompanyLegalForms()` as extension methods on the same builder, so registering data-backed recognizers never changes Core.

## Finder and options

```csharp
var finder = new EntityFinder(catalog, new EntityFinderOptions {
    Kinds = { EntityKind.NationalId, EntityKind.TaxId, EntityKind.BankAccount },
    Countries = { "PL" },              // empty = all; international schemes always pass this filter
    MinimumConfidence = 0.5,
    MaxMatches = 10_000,
    MaxInputCharacters = 16 * 1024 * 1024,
    ContextWindow = 48,
    OverlapPolicy = EntityOverlapPolicy.BestOnly
});

EntityFindResult result = finder.Find(text);
EntityFindResult cells = finder.Find(segments, cancellationToken);
// result.Matches, result.Truncated, result.DroppedOverlaps, result.Diagnostics (no values)
```

Options also take `IncludeSchemes` and `ExcludeSchemes` by id. Limits fail soft: when `MaxMatches` or `MaxInputCharacters` is reached the result says `Truncated` and the host decides what to tell the user. `Find` is synchronous and CPU-bound; it checks the cancellation token between segments and every 64 KB inside long segments.

The public API uses strings across all supported targets. Checksum arithmetic streams digits and letter expansions without allocating an expanded numeric string.

## Privacy

- `EntityMatch.ToString()`, `EntityValidationResult.ToString()` and every diagnostic describe kind, scheme, position and status only, never the value.
- The engine writes no logs and raises no events. Hosts that log results must log the `ToString()` form or counts.
- Exceptions thrown for bad arguments (null catalog, negative limits) never include input text.
- `Sensitivity` on each scheme (`Personal`, `Financial`, `Organizational`, `Public`) lets hosts treat a company's KRS differently from a PESEL. Defaults for masking and redaction select `Personal` and `Financial`.

## Masking

```csharp
string safe = EntityMasker.Mask(text, result.Matches, new EntityMaskOptions {
    Style = EntityMaskStyle.KindLabel          // "[NationalId]", "[BankAccount]"
});
// Other styles:
//   Fixed         "███████████"               (character and length policy configurable)
//   Partial       "PL61 •••• •••• •••• •••• •••• 2874", "•••• •••• •••• 4242"
//   Pseudonym     "[TaxId:7f3a9c]"           HMAC-SHA256 over the normalized value with a caller key,
//                                            so the same NIP gets the same token across documents
```

The segment overload `EntityMasker.Mask(segments, matches, options)` returns a string for each input segment. It uses every match part to mask the original source text. Fixed masking replaces each matched UTF-16 code unit. Partial masking applies the scheme to the full value and distributes its result across the covered parts. Kind labels and pseudonyms replace the first part and clear the matched text from subsequent parts.

Partial masking follows each scheme's rule (`Scheme.MaskPartial`): card numbers keep the last four digits as payment standards allow, IBANs keep the country code and the last four, emails keep the first character of the local part and the domain. Pseudonyms need a caller-supplied key; there is no default key, and the engine keeps no map. A host that needs to restore values after an AI round trip keeps its own `token → value` map from `EntityMasker.Pseudonymize`, which returns both the masked text and the map.

## Data-backed recognizers (later, `OfficeIMO.Entities.Data`)

| Recognizer | Source | Licence | Notes |
| --- | --- | --- | --- |
| Phones (parse, validate, format E.164/national/international, find) | libphonenumber `PhoneNumberMetadata.xml` | Apache-2.0 | Generated C# tables at build time from a pinned release; no XML parsing at run time. |
| Postal address formats and validation per country | Google address metadata (libaddressinput data service snapshot) | to be confirmed before use | Formatting for mail merge and invoices first, recognition in prose much later. |
| Company legal forms (ISO 20275 ELF codes) | GLEIF ELF code list | CC0 | Recognizes `sp. z o.o.`, `GmbH`, `S.A.` suffixes for company-name detection. |

Each source gets an entry in `Build/third-party-material.json`, its licence text packed and a section in the package's `THIRD-PARTY-NOTICES.md`; `python Build/validate_third_party_notices.py` must pass. Generators live in `Build/` and run from a pinned snapshot; the package never downloads data. Data releases follow the OfficeIMO release train, not upstream's schedule.

Core never references this package. The Data package references Core and plugs into `EntityCatalogBuilder` through extension methods.

## Ownership and consumer migration

The engine owns IBAN validation. The other consumer migrations below are proposed.

| Capability | Previous or current implementation | Integration |
| --- | --- | --- |
| IBAN normalization, SWIFT IBAN Registry BBAN structures and MOD 97-10 (`InvoiceBankAccountIdentity`, internal) | `OfficeIMO.SharedSource/Invoicing/InvoiceBankAccountIdentity.cs`, source-linked into `OfficeIMO.Invoicing` and `OfficeIMO.Pdf` | Moved into Core as the implementation of the public `FinancialSchemes.Iban`; the shared IBAN file is deleted. `OfficeIMO.Invoicing` references `OfficeIMO.Core`, and every former `IsValidIban` caller in Invoicing and Pdf uses the public scheme. The `OfficeIMO.SharedSource/Invoicing/*.cs` links stay, because they still supply the other shared invoice types (`InvoiceProfiles`, `InvoiceXml`, `InvoiceSyntax` and others). |
| NIP shape check without checksum (regular expression), `PL` prefix stripping | `OfficeIMO.Invoicing/Poland/Fa3InvoiceWriter.MappingChecks.cs`, `Fa3InvoiceWriter.Fields.cs`, `OfficeIMO.Invoicing.KSeF/KsefProtocol.cs` | Stage 2: call `PolandSchemes.Nip` for NIP and EU VAT values on FA(3) parties; keep the schema regular expression only where the schema itself defines it. |
| BIC shape check (8 or 11 characters) | `Fa3InvoiceWriter.MappingChecks.cs` | Stage 2: `FinancialSchemes.Bic`, which also checks the ISO 3166 country part. |
| Tax registration presence check only | `OfficeIMO.Invoicing.Validation/InvoiceModelValidator.TaxRegistrations.cs` | Stage 2: optional format and check-digit validation of `InvoiceTaxRegistration` with `VatId`/`TaxId` schemes, reported as warnings because invoices can legitimately carry foreign identifiers outside the catalog. |
| SMTP address normalization through `System.Net.Mail.MailAddress`, case kept, no IDN | `OfficeIMO.Email/Model/EmailSmtpAddress.cs`; dedup keys in `Contact/EmailContactConsolidation.cs` | Stage 2: `InternetSchemes.Email` gives RFC 5321/5322/6531 validation and a comparison key (domain lower-cased and IDNA-mapped to A-labels, local part kept as written because it is case-sensitive by RFC 5321). Email keeps its MIME address parser; the scheme validates the address it extracts. |
| PDF redaction search by literal and regular expression | `OfficeIMO.Pdf/Manipulation/PdfRedactionSearchOptions.cs`, `PdfRedactionPlanner.Search.cs`, `OfficeIMO.Workflows/PdfRedactionWorkflowContracts.cs` | Stage 2: entity criteria next to literal and regex, see below. |

`OfficeHiddenDataKind.PersonalIdentities` (author and people parts) is metadata, not body content. Entity findings in body text are a separate concern and should not become another `OfficeHiddenDataKind` flag, which would be confused with it. They belong to redaction plans and inspection reports.

`OfficeIMO.Core/TextAnalysis` (`UnicodeTextSegmenter`, `OfficeTextIndex`) is internal and segment-oriented; the entity tokenizer is a separate, smaller run tokenizer, but should reuse its Unicode category helpers rather than add new tables.

## Consumers

Every consumer goes through the same three calls: `Scheme.Validate` for a known field, `EntityFinder.Find` for text, `EntityMasker` for output. Hosts map parameters and present results; nothing below re-implements recognition.

### Redaction plans

PDF already has a reviewable plan: search criteria produce `PdfRedactionMatch` items with page geometry, which the caller reviews and applies. Entities become one more criterion:

```csharp
var search = new PdfRedactionSearchOptions()
    .AddEntities(EntityCatalog.Default, new EntityFinderOptions {
        Kinds = { EntityKind.NationalId, EntityKind.TaxId, EntityKind.BankAccount },
        MinimumConfidence = 0.3
    });
PdfRedactionPlan plan = pdf.Redactions.Search(search);
// each PdfRedactionMatch carries criterion "entity:PL.PESEL" (never the value) and the confidence
```

The planner already extracts page text with glyph geometry. The Stage 2 redaction criterion can send consecutive runs as continued segments and map each returned `EntityMatchPart` to the glyphs of its source run. The workflow recipe (`officeimo.pdf.redaction.recipe.v1`) gains `PdfRedactionRuleKind.Entity` with a kind or scheme list and an optional minimum confidence. Evidence carries the rule name and scheme IDs.

Word, Excel and PowerPoint have no text redaction yet. When they get it, they follow the same plan shape (find, review, apply, verify): Word over runs with paragraph locations, Excel over cells with the header as label, PowerPoint over shape text. Entities arrive as a criterion from day one rather than a later bolt-on. Until then, these formats offer inspection only (`Find-OfficeEntity` below).

### Inspection and hidden data

A document-level inspection report lists entity findings by kind, scheme, location and confidence, without values, next to hidden-data and content-safety findings: `OfficeDocument.InspectEntities(path, options)` in the cross-format package, built on each format's text projection (the Reader already produces it). This answers "does this file contain personal data before I send it" without changing the file.

### OfficeIMO.Reader and AI preparation

`ReaderChunk` has no metadata bag today. Add an optional entity pass as a Reader processor, enabled by one builder call, that fills a new `ReaderChunk.Entities` list (kind, scheme, offsets, confidence; normalized values only when the caller opts in) and can mask chunk `Text` and `Markdown` before they leave the process:

```csharp
var reader = new OfficeReaderBuilder()
    .WithEntities(EntityCatalog.Default, mask: EntityMaskStyle.Pseudonym, key: hostKey)
    .Build();
```

`OfficeIMO.AI` requests that allow remote processing can then require masked chunks (`OfficeAiRequest.MaskEntities`), and `OfficeAiFieldType` can gain identifier field types (`Iban`, `VatId`, `Email`) whose extracted values are validated by the matching scheme before they are returned.

### Excel and CSV cleaning

- **Profile a column:** run `Scheme.Validate` over a sample (header as label) and report which schemes fit and what share of values validate.
- **Validate:** `CsvColumnBuilder.AsEntity(PolandSchemes.Nip)` and an Excel equivalent report invalid rows with their status (`InvalidCheckDigit`, `InvalidLength`) through the existing `CsvValidator` messages.
- **Normalize or format:** rewrite a column to `Compact` or `Display` form (`NIP 1234563218` to `123-456-32-18`, IBANs grouped by four), leaving invalid values untouched and listed.
- **Mask:** replace a column with partial or pseudonymous values for sharing.

The engine is called per cell; the Excel and CSV packages own reading, writing and reporting.

### OfficeIMO.Invoicing

Party VAT and tax registrations, seller and buyer NIP on FA(3), IBAN and BIC in payment means all go through schemes. Postal address formatting for invoice PDFs comes later from the Data package. Validation stays advisory where invoices may carry identifiers outside the catalog.

### OfficeIMO.Email

Address validation and a comparison key for contact consolidation and dedup, from `InternetSchemes.Email`. Display names never decide identity, as today. Email bodies can be inspected or masked with the finder like any other text.

### Word mail merge

Later, with the Data package: format a postal address block per destination country from structured fields, and validate identifiers merged into templates (`ExecuteWithReport` reports invalid NIP or IBAN values in the data source rather than printing them).

### PSWriteOffice, CLI and MCP

```powershell
Test-OfficeEntity -Scheme PL.NIP -Value '123-456-32-18'           # status, normalized, formatted
Find-OfficeEntity -Path .\contract.docx -Kind NationalId,TaxId     # findings without values by default
Find-OfficeEntity -Path .\contract.docx -IncludeValue              # explicit opt-in
Get-OfficePdfRedactionPlan -Path .\scan.pdf -Entity NationalId,BankAccount -MinimumConfidence 0.3
Protect-OfficeText -Text $text -Entity Email,Phone -MaskStyle Partial
Get-OfficeEntityScheme -Country PL                                  # catalog listing
```

The cmdlets stay thin: they build options, call the engine and write the result objects. The CLI and MCP server expose the same three operations (validate, find, mask) with values excluded from output unless requested.

### Studio

An inspection panel listing entity findings per document with jump-to-location, a redaction review that preselects entity matches, and a mask-before-AI switch for the assistant. Studio interface text goes through the translation scripts like everything else.

## Sources and test vectors

Each scheme records its primary source in `Specification`, and its tests use the issuer's published example values where the issuer publishes any. Where an issuer publishes the algorithm but no example, tests compute vectors from the published weights and record the derivation next to the test.

Where no issuer publication exists, a community-documented algorithm may be used. The check's `EntityCheckInfo` then has `Provenance = CommunityDocumented` and cites the description in `Source`, and its tests use issuer artifacts when available or worked vectors with a documented derivation. The description is implemented from its prose; code from other projects is never copied or ported.

Two findings shape stage 1:

- The European Commission does not publish national VAT check algorithms ([VIES FAQ](https://ec.europa.eu/taxation_customs/vies/#/faq), question 11). The Austrian Federal Ministry of Finance publishes the construction rules that member states agreed to disclose, with a worked example for each ([BMF, *UID-Konstruktionsregeln*, November 2020](https://bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf)). For additional variants the engine follows the sources below. Formats without an applicable documented check report `CheckDigit = NotDefined`.
- Several Polish identifiers have a legal basis for their structure but no published check algorithm: REGON (structure in the REGON regulation, weights not published), KRS, identity card and passport numbers. Stage 1 uses community-documented checks for REGON (9 and 14 digits), the identity card and the passport, marked as such; KRS has no check digit and is validated by length. An issuer publication, if one turns up, replaces the community source and flips the provenance.

| Scheme | Primary source | Defines | Issuer example |
| --- | --- | --- | --- |
| PL PESEL | [Population Register Act, art. 15(2)](https://api.sejm.gov.pl/eli/acts/DU/2026/384/text/U/D20260384Lj.pdf) (Dz.U. 2026 poz. 384); [MSW regulation on assigning PESEL, § 7](https://api.sejm.gov.pl/eli/acts/DU/2015/1984/text.html) (Dz.U. 2015 poz. 1984) | Date of birth with month offsets +80 (1800-1899), +0 (1900-1999), +20 (2000-2099) only; serial; sex in the 10th digit (even = female); weights 1-3-7-9-1-3-7-9-1-3, check = (10 − sum mod 10) mod 10 | none |
| PL NIP | [BMF, *UID-Konstruktionsregeln*, November 2020](https://bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf), PL section | Weights 6-5-7-2-3-4-5-6-7, mod 11, remainder 10 is invalid | 5260001246 |
| PL REGON 9/14 | [REGON regulation, § 12](https://api.sejm.gov.pl/eli/acts/DU/2015/2009/text.html), [Numerik](https://numerik.slashlab.pl/guide/algorithms/) | Separate community-documented modulus 11 checks for the nine-digit parent and the 14-digit branch | |
| PL KRS | [Court register regulation](https://api.sejm.gov.pl/eli/acts/DU/2026/649/text.pdf) | Ten digits in register extracts. The finder pads shorter labelled values. | |
| PL identity card | [MSWiA ID card template](https://api.sejm.gov.pl/eli/acts/DU/2025/1031/text.pdf), [itMock](https://itmock.pl/blog/numer-dowodu-osobistego) | Community-documented check uses the first digit following the letter series | |
| PL passport | [MSWiA passport regulation](https://api.sejm.gov.pl/eli/acts/DU/2022/2050/text.html), [Romuald Dębski’s passport checksum notes](https://romek.info/ut/paszport.html) | Community-documented number check uses the first digit following the letter series. The legacy BM series uses format-only validation. | |
| EU VAT formats | [VIES FAQ](https://ec.europa.eu/taxation_customs/vies/#/faq), question 11 | Structure per state, `EL` for Greece, `XI` for Northern Ireland (including `GD`/`HA` forms) | none |
| EU VAT checks | [BMF, *UID-Konstruktionsregeln*, November 2020](https://bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf) | Rules and examples for AT, BE (old and 2020+), DE, DK, EE, FI, HU, IE, IT, LT, LU, NL (old and 2020+: mod 11 first, then mod 97), PL, PT, SE, SI, SK, GB/XI | AT U10223006, BE 0776091951, DE 111111125, FI 09853608, NL 010000446B01, PL 5260001246 |
| ES personal NIF/NIE letter | [Ministerio del Interior, control digit of the NIF/NIE](https://www.interior.gob.es/opencms/es/servicios-al-ciudadano/tramites-y-gestiones/dni/calculo-del-digito-de-control-del-nif-nie/) | Number mod 23 and letter table; NIE X/Y/Z replaced by 0/1/2. Applies to Spanish VAT numbers of individuals; company NIF control is not published | 12345678Z |
| BG and CZ company VAT | [VatDB Bulgaria](https://vatdb.com/guides/validate-bg-vat-number/), [VatDB Czechia](https://vatdb.com/guides/validate-cz-vat-number/) | Weighted company checks. Bulgarian ten-digit and Czech nine- and ten-digit variants are format-only. | |
| CY | [VatDB Cyprus](https://vatdb.com/guides/validate-cy-vat-number/) | Positional lookup, modulus 26 and reserved `12` prefix | |
| EL, FR, LV, MT, RO | [CDQ check-digit descriptions](https://meta.cdq.com/Public:Rule_Category/IDENTIFIER_CHECKDIGIT) | National checks, including French numeric and alphanumeric VAT keys. Latvia’s personal variant is format-only. | |
| ES K/L/M personal NIF | [Spanish Tax Agency technical note](https://www.caib.es/sites/civitasmanualsisuport/f/41519) | Applies the personal control letter to the seven digits following the prefix | |
| ES company VAT | [NIF/CIF control-character documentation](https://josegoval.github.io/nif-dni-nie-cif-validation/es/guides/control-character/) | Weighted modulus 10 and numeric or alphabetic control character | |
| HR | [Croatian central register OIB check](https://regos.hr/app/uploads/2018/07/KONTROLA-OIB-a.pdf) | Issuer-published ISO 7064 MOD 11,10 | |
| UK NINO | [HMRC NIM39110](https://www.gov.uk/hmrc-internal-manuals/national-insurance-manual/nim39110) | Two letters, six digits, suffix A-D; D, F, I, Q, U, V never first or second letter; O never second; BG, GB, KN, NK, NT, TN, ZZ not used | QQ 12 34 56 A (labelled specimen) |
| UK VAT | [BMF, *UID-Konstruktionsregeln*, November 2020](https://bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf), GB section; HMRC [VATAFRS2500](https://www.gov.uk/hmrc-internal-manuals/vat-agricultural-flat-rate-scheme/vatafrs2500) confirms a modulus 97 check | Weights 8..2 plus the last two digits; valid when (sum mod 97) or ((sum + 55) mod 97) is 0; excluded ranges; GD000-499 government departments, HA500-999 health authorities | 434031494 |
| UK company number | Companies House [URI customer guide](https://assets.publishing.service.gov.uk/media/5d08c0f340f0b6094a379078/uniformResourceIdentifiersCustomerGuide.pdf) | 8 characters, prefix table (SC, NI, OC, SO, SL, LP, FC, ...); no check digit. The guide is from 2012 and lacks newer prefixes such as OE | 02050399, SC002180 |
| US SSN | SSA [randomization](https://www.ssa.gov/employer/randomization.html) and [FAQ](https://www.ssa.gov/employer/randomizationfaqs.html) | Area 000, 666 and 900-999 never assigned; group 00 and serial 0000 invalid | none |
| US EIN | IRS [valid EIN prefixes](https://www.irs.gov/businesses/small-businesses-self-employed/how-eins-are-assigned-and-valid-ein-prefixes) | Valid two-digit prefixes; no check digit | none |
| IBAN | [SWIFT IBAN Registry](https://www.swift.com/standards/data-standards/iban-international-bank-account-number) ([TXT](https://www.swift.com/swift-resource/11971/download)); ISO 13616-1, ISO/IEC 7064 MOD 97-10 | Country, length and BBAN structure for every registered country; check digits 02-98, MOD 97-10 | one electronic example per country, such as AD1200012030200359100100, AT611904300234573201 |
| BIC | [SWIFT BIC page](https://www.swift.com/standards/data-standards/bic-business-identifier-code), ISO 9362 | 4-letter party prefix, ISO 3166 country, 2-character suffix, optional 3-character branch; no check digit | none |
| Payment card | ISO/IEC 7812-1 | Major industry identifier and Luhn check | none |
| ISIN | [ANNA identifiers](https://anna-web.org/identifiers/), ISO 6166 | Country code, 9-character national code, check digit by the mod 10 "double-add-double" method over letters mapped to numbers | none |
| LEI | [LEI ROC, "The LEI"](https://www.leiroc.org/lei.htm), ISO 17442-1 | 4-character LOU prefix, `00` reserved, 12-character entity part, 2 check digits by ISO/IEC 7064 MOD 97-10 | none |
| Email | [RFC 5321](https://www.rfc-editor.org/rfc/rfc5321) §4.5.3.1, RFC 5322, RFC 6531, RFC 5890/5891, [UTS #46](https://www.unicode.org/reports/tr46/) | Local part ≤ 64 octets, domain ≤ 255, path ≤ 256; SMTPUTF8; IDNA mapping | [IdnaTestV2.txt](https://www.unicode.org/Public/idna/latest/IdnaTestV2.txt) |
| URL | [WHATWG URL Standard](https://url.spec.whatwg.org/), [RFC 3986](https://www.rfc-editor.org/rfc/rfc3986) | Parsing and serialization | web-platform-tests `urltestdata.json` |

The scheme metadata links to the format or checksum documentation. Where a scheme has more than one check, `EntityCheckInfo.Source` identifies the documentation used for that check.

## Integration roadmap

The [entity recognition roadmap](ROADMAP.md#entity-recognition) tracks open document, host and country-pack work. Core provides the scheme catalog, finder and masker. Invoicing and PDF use the same Core IBAN implementation and registry table.
