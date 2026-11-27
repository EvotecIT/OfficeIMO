# OfficeIMO entity recognition design

This is the architecture and API direction for recognizing identifiers and personal data in text: validating a value someone typed, normalizing it to one canonical form, formatting it for display, and finding such values inside documents. Nothing here is implemented yet. The first stage is specified in the [implementation brief](#appendix-a-stage-1-implementation-brief) at the end.

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
    public virtual Uri? Specification => null;        // primary source for the format
    public virtual EntityCheckProvenance CheckProvenance => EntityCheckProvenance.None;
                                                      // IssuerPublished, CommunityDocumented, None
    public virtual Uri? CheckSource => null;          // where the check algorithm is documented
    public virtual int MinLength => 1;                // normalized length bounds
    public virtual int MaxLength => 64;
    public virtual EntityCheckKind Check => EntityCheckKind.None;  // Mod11, Mod97_10, Luhn, ...
    public virtual EntitySensitivity Sensitivity => EntitySensitivity.Personal;

    public abstract EntityValidationResult Validate(string? value);
    public bool IsValid(string? value) => Validate(value).IsValid;
    public bool TryNormalize(string? value, out string normalized);
    public virtual string Format(string normalized, EntityFormatStyle style = EntityFormatStyle.Display);
    public virtual EntityDetails? Decode(string normalized) => null;  // opt-in
}
```

- **Validate** accepts the forms people write: surrounding whitespace, the scheme's own separators (spaces, hyphens, dots, non-breaking spaces), lower case where the scheme is case-insensitive, and an optional country prefix where one is customary (`PL` before a NIP). It returns a result, never throws for bad input.
- **Normalize** returns the canonical compact form: digits and upper-case letters only, with the prefix the scheme defines (`PL1234563218` for a Polish VAT number, `1234563218` for a NIP).
- **Format** renders a normalized value: `Compact`, `Display` (the issuer's printed grouping, such as `123-456-32-18` or `PL61 1090 1014 0000 0712 1981 2874`) and `Electronic` where it differs.
- **Check provenance** says where a check algorithm comes from. `IssuerPublished` means the issuer (or a ministry publishing the rules member states agreed to disclose) states it. `CommunityDocumented` means no issuer publication exists and the algorithm is taken from a documented description that matches real issued numbers; the scheme cites it in `CheckSource`. Results and matches carry the provenance, so a host can treat a community-documented check as weaker or ignore it.
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
    public EntityCheckProvenance CheckProvenance { get; }  // copied from the scheme
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

The finder tokenizes the text once into runs, and recognizers read the runs rather than scanning characters themselves:

- **Digit runs**: maximal sequences of digits joined by single separators (space, hyphen, dot, slash, non-breaking space, thin space), with the separator pattern recorded (`3-3-2-2`, `4 4 4 4`).
- **Alphanumeric runs**: letters and digits with the same separator rules, for IBAN, VAT, NINO, passport and ISIN shapes.
- **Anchors**: `@` positions for email and `://` or `www.` positions for URLs; the email and URL recognizers expand outward under RFC limits.

Each run knows whether a letter or digit touches either end. A recognizer never reports a candidate inside a longer run of the same character class: the middle 11 digits of a 20-digit number are not a PESEL. Every recognizer has a fixed maximum candidate length, so scanning is linear in the text and there is no backtracking to exploit.

Recognizers are separate from schemes because one recognizer can serve many schemes (one EU VAT recognizer for 28 prefixes, one IBAN recognizer for every country) and because a scheme is useful without any finder (validating an Excel column, an invoice field or a PowerShell parameter).

### Candidates, evidence and matches

```csharp
public sealed class EntityMatch {
    public int Start { get; }
    public int Length { get; }                        // covers separators and prefix as written
    public EntityKind Kind { get; }
    public EntityScheme Scheme { get; }
    public string Normalized { get; }                 // sensitive
    public EntityCheckProvenance CheckProvenance { get; }
    public double Confidence { get; }                 // 0..1
    public EntityEvidence Evidence { get; }           // flags
    public string? ContextKeyword { get; }            // "NIP", "nr rachunku"
    public string? ContextLanguage { get; }           // BCP 47: "pl", "en"
    public EntitySegmentLocation? Location { get; }   // from the input segment, see below
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
    public EntityTextSegment(string text, EntitySegmentLocation? location = null, string? label = null);
}

public sealed class EntitySegmentLocation {
    public string? Part { get; }      // "word/document.xml", "Sheet1", "page 3"
    public string? Address { get; }   // "B7", "paragraph 12", "chunk 4"
    public int? Page { get; }
}
```

Excel passes each cell with its column header as `label`; the Reader passes chunks with their source location; invoicing passes a field value with the element name as `label` and the match gets `StructuredSource`. Matches never cross segment boundaries.

## Scoring

Confidence is a number from 0 to 1 built from the scheme's base strength and the evidence, with fixed, documented steps so results are explainable and stable across releases. The base reflects how likely a random string of that shape passes the check:

| Base class | Examples | Base |
| --- | --- | --- |
| Strong check over a distinctive shape | IBAN (ISO 7064 MOD 97-10 plus country length), LEI, ISIN, EU VAT with country prefix and check | 0.85 |
| Distinctive shape, no check | Email, URL, UK NINO pattern with valid prefix and suffix | 0.80 |
| Country prefix and VIES format, no documented check | an EU VAT variant for which no check is documented at all | 0.60 |
| Check digit over a plain digit run | PESEL, NIP, REGON, card number (Luhn) with a payment-card major industry identifier, EU VAT digits without prefix | 0.40 |
| Check digit over a short letter-digit shape | PL identity card, PL passport | 0.30 |
| Shape only, no check | KRS, SSN, EIN, BIC, UK company number | 0.20 |

Community-documented checks use the same base as issuer-published ones; they are no weaker mathematically. The difference is provenance, which stays visible in `CheckProvenance` and the `CommunityDocumentedCheck` evidence flag.

Adjustments, applied in order and clamped to [0, 1]:

- `Label` matching the scheme (`NIP`, `PESEL`, `IBAN`, `VAT ID`): +0.50. `ContextKeyword` within the window (default 48 characters before the value, same segment, nothing but punctuation, spaces and short connecting words between): +0.40.
- `IssuerFormatting`: +0.10. `CountryPrefix` where optional: +0.10. `KnownIssuerRange`: +0.10.
- `PlausibleDate` missing for PESEL (future date or older than 130 years): -0.30.
- `AmbiguousReading` without disambiguating context: -0.15 for each competing reading.
- `StructuredSource`: the value is taken as given, so confidence equals 1.0 when valid.

The default `MinimumConfidence` is 0.50. Consequences that tests pin:

- `44051401359` alone in prose is not reported (0.40). `PESEL: 44051401359` is (0.90). The same number in a column headed `PESEL` is (0.90).
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

Validation allocates only the normalized string; the finder allocates per match, not per character. On .NET 8+ the tokenizer and checks work on spans internally; the public API is string-based so the same surface compiles for `netstandard2.0` and `net472`.

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

Partial masking follows each scheme's rule (`Scheme.MaskPartial`): card numbers keep the last four digits as payment standards allow, IBANs keep the country code and the last four, emails keep the first character of the local part and the domain. Pseudonyms need a caller-supplied key; there is no default key, and the engine keeps no map. A host that needs to restore values after an AI round trip keeps its own `token → value` map from `EntityMasker.Pseudonymize`, which returns both the masked text and the map.

## Data-backed recognizers (later, `OfficeIMO.Entities.Data`)

| Recognizer | Source | Licence | Notes |
| --- | --- | --- | --- |
| Phones (parse, validate, format E.164/national/international, find) | libphonenumber `PhoneNumberMetadata.xml` | Apache-2.0 | Generated C# tables at build time from a pinned release; no XML parsing at run time. |
| Postal address formats and validation per country | Google address metadata (libaddressinput data service snapshot) | to be confirmed before use | Formatting for mail merge and invoices first, recognition in prose much later. |
| Company legal forms (ISO 20275 ELF codes) | GLEIF ELF code list | CC0 | Recognizes `sp. z o.o.`, `GmbH`, `S.A.` suffixes for company-name detection. |

Each source gets an entry in `Build/third-party-material.json`, its licence text packed and a section in the package's `THIRD-PARTY-NOTICES.md`; `python Build/validate_third_party_notices.py` must pass. Generators live in `Build/` and run from a pinned snapshot; the package never downloads data. Data releases follow the OfficeIMO release train, not upstream's schedule.

Core never references this package. The Data package references Core and plugs into `EntityCatalogBuilder` through extension methods.

## Existing code to consolidate

There is no shared validation code yet. What exists is small and moves into the engine rather than living beside it:

| Today | Where | Stage 1 action |
| --- | --- | --- |
| IBAN normalization, SWIFT IBAN Registry BBAN structures and MOD 97-10 (`InvoiceBankAccountIdentity`, internal) | `OfficeIMO.SharedSource/Invoicing/InvoiceBankAccountIdentity.cs`, source-linked into `OfficeIMO.Invoicing` and `OfficeIMO.Pdf` | Move into Core as the implementation of the public `FinancialSchemes.Iban` and delete the shared source file. `OfficeIMO.Invoicing` gains a reference to `OfficeIMO.Core` (any package may reference Core); both Invoicing and Pdf call the public scheme. |
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

The planner already extracts page text with glyph geometry; it passes each text run as an `EntityTextSegment`, and maps match offsets back to glyph boxes exactly as it does for regex matches today. The workflow recipe (`officeimo.pdf.redaction.recipe.v1`) gains `PdfRedactionRuleKind.Entity`, with `Value` holding a kind or scheme list such as `NationalId,PL.NIP` and an optional minimum confidence; evidence keeps the rule name and scheme ids, as it does now.

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

Where no issuer publication exists, a community-documented algorithm may be used. The scheme then sets `CheckProvenance = CommunityDocumented`, cites the description in `CheckSource`, and its tests include at least one value from a public issuer artifact (a specimen document, a register extract or the issuer's own sample) that the algorithm accepts. The description is implemented from its prose; code from other projects is never copied or ported.

Two findings shape stage 1:

- The European Commission does not publish national VAT check algorithms ([VIES FAQ](https://ec.europa.eu/taxation_customs/vies/#/faq), question 11). The Austrian Federal Ministry of Finance publishes the construction rules that member states agreed to disclose, with a worked example for each ([BMF, *UID-Konstruktionsregeln*, November 2020](https://bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf)). Ten states did not agree (BG, CY, CZ, EL, ES, FR, HR, LV, MT, RO); for them the scheme uses a community-documented check where one exists (`CheckProvenance = CommunityDocumented`) and otherwise checks the VIES format only (`CheckDigit = NotDefined`).
- Several Polish identifiers have a legal basis for their structure but no published check algorithm: REGON (structure in the REGON regulation, weights not published), KRS, identity card and passport numbers. Stage 1 uses community-documented checks for REGON (9 and 14 digits), the identity card and the passport, marked as such; KRS has no check digit and is validated by length. An issuer publication, if one turns up, replaces the community source and flips the provenance.

| Scheme | Primary source | Defines | Issuer example |
| --- | --- | --- | --- |
| PL PESEL | [Population Register Act, art. 15(2)](https://api.sejm.gov.pl/eli/acts/DU/2026/384/text/U/D20260384Lj.pdf) (Dz.U. 2026 poz. 384); [MSW regulation on assigning PESEL, § 7](https://api.sejm.gov.pl/eli/acts/DU/2015/1984/text.html) (Dz.U. 2015 poz. 1984) | Date of birth with month offsets +80 (1800-1899), +0 (1900-1999), +20 (2000-2099) only; serial; sex in the 10th digit (even = female); weights 1-3-7-9-1-3-7-9-1-3, check = (10 − sum mod 10) mod 10 | none |
| PL NIP | [BMF, *UID-Konstruktionsregeln*, November 2020](https://bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf), PL section | Weights 6-5-7-2-3-4-5-6-7, mod 11, remainder 10 is invalid | 5260001246 |
| PL REGON 9/14 | [REGON regulation, § 12](https://api.sejm.gov.pl/eli/acts/DU/2015/2009/text.html) (Dz.U. 2015 poz. 2009) | 9 digits (8 serial + check) and 14 digits (REGON 9 + 4-digit local unit + check); weights not published | none |
| PL KRS | KRS Act and [MS regulation](https://api.sejm.gov.pl/eli/acts/DU/2026/649/text.pdf) | Refer to the KRS number without defining its length; format (10 digits, leading zeros) is taken from register extracts | none |
| PL identity card | [MSWiA regulation on the ID card template](https://api.sejm.gov.pl/eli/acts/DU/2025/1031/text.pdf) (Dz.U. 2025 poz. 1031) | Field "series and number" only; MRZ check digits follow ICAO Doc 9303 Part 3 (7-3-1 weights) | none |
| PL passport | [MSWiA passport regulation](https://api.sejm.gov.pl/eli/acts/DU/2022/2050/text.html) (Dz.U. 2022 poz. 2050) | No number format; MRZ per ICAO Doc 9303 | none |
| EU VAT formats | [VIES FAQ](https://ec.europa.eu/taxation_customs/vies/#/faq), question 11 | Structure per state, `EL` for Greece, `XI` for Northern Ireland (including `GD`/`HA` forms) | none |
| EU VAT checks | [BMF, *UID-Konstruktionsregeln*, November 2020](https://bmf.gv.at/dam/jcr:d6794f8f-d321-43df-9840-1a841f9bf5dc/BMF_UID_Konstruktionsregeln_Stand_November%202020.pdf) | Rules and examples for AT, BE (old and 2020+), DE, DK, EE, FI, HU, IE, IT, LT, LU, NL (old and 2020+: mod 11 first, then mod 97), PL, PT, SE, SI, SK, GB/XI | AT U10223006, BE 0776091951, DE 111111125, FI 09853608, NL 010000446B01, PL 5260001246 |
| ES personal NIF/NIE letter | [Ministerio del Interior, control digit of the NIF/NIE](https://www.interior.gob.es/opencms/es/servicios-al-ciudadano/tramites-y-gestiones/dni/calculo-del-digito-de-control-del-nif-nie/) | Number mod 23 and letter table; NIE X/Y/Z replaced by 0/1/2. Applies to Spanish VAT numbers of individuals; company NIF control is not published | 12345678Z |
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

ISO standards are paid and are cited by number only; every issuer-published algorithm above is also stated in a free publication in the same row. Community-documented checks are listed in the brief with the source each scheme must cite.

## Delivery stages

1. **Stage 1, engine in Core:** kinds, scheme base, catalog and builder, tokenizer, finder with scoring and overlap resolution, masker, and the schemes listed in the [brief](#appendix-a-stage-1-implementation-brief). Move the IBAN helper into the engine; `OfficeIMO.Invoicing` references Core, and Invoicing and Pdf switch to the public IBAN scheme. No other consumer changes.
2. **Stage 2, first consumers:** PDF redaction entity criterion and workflow rule kind; PSWriteOffice `Test-OfficeEntity`, `Find-OfficeEntity`, `Get-OfficeEntityScheme` and `-Entity` on the PDF redaction cmdlets; invoicing and email adoption; CLI and MCP operations.
3. **Stage 3, documents and data:** cross-format inspection report, Reader processor with masking, AI request masking, CSV and Excel column profiling, validation, normalization and masking.
4. **Stage 4, `OfficeIMO.Entities.Data`:** phones from libphonenumber metadata, then postal address formats (invoices, mail merge), then ISO 20275 legal forms.
5. **Stage 5, more countries:** national identifiers for further EU states (DE Steuer-ID, FR NIR, ES DNI/NIE, IT codice fiscale, NL BSN, ...) as country packs, each with its own primary sources and keywords. Studio panels follow the engine stages they present.

## Decisions

- **Any package may reference `OfficeIMO.Core`.** `OfficeIMO.Invoicing` references Core and uses the public schemes; there is no linked source copy of entity code anywhere.
- **Community-documented checks are allowed** where no issuer publication exists, with the provenance of every check (`IssuerPublished` or `CommunityDocumented`) visible on schemes, validation results and matches.

## Open decisions

- **SWIFT IBAN Registry data.** The existing BBAN structure table is transcribed from the SWIFT IBAN Registry. Check whether its terms require an entry in `Build/third-party-material.json` when it moves into Core; add one if they do.
- **Payment card ranges.** Card brand detection by IIN range needs a maintained range table; stage 1 uses only the first digit (ISO/IEC 7812 major industry identifier) and Luhn. A brand table is a later, sourced addition.
- **Issuer sources for community-documented checks.** REGON weights, the identity card and passport check digits, and the checks for the ten EU states that did not disclose theirs ship as community-documented. If GUS, MSWiA or a national tax authority publishes them (for example in the GUS BIR service specification), the citation and provenance change; behaviour should not.

## Appendix A: stage 1 implementation brief

This brief is self-contained: it is what an implementer needs for the first PR. The sections above explain why.

### Goal

Add `OfficeIMO.Entities` to `OfficeIMO.Core`: an extensible model and engine to validate, normalize, format, decode and find identifiers and personal data in text, plus masking. One PR, titled as a plain sentence (for example `Add entity recognition for identifiers and personal data to OfficeIMO.Core`). Nothing is released before 4.0, so the API can still change.

### Constraints

- Warnings are errors, nullable is on, follow `.editorconfig`. Targets are Core's: `netstandard2.0`, `net8.0`, `net10.0`, plus `net472` on Windows. The public API is string-based; span-based internals under `#if NET8_0_OR_GREATER` are fine.
- No new package references. No reflection, no `Regex` in recognizers or schemes, no runtime code generation, no static constructors that do heavy work. Must stay trim- and NativeAOT-safe (Core is used by the AOT smoke projects).
- No code copied, ported or adapted from other projects (python-stdnum, libphonenumber, Presidio, OneContact or any other). Implement every algorithm from the source cited in the [sources table](#sources-and-test-vectors) or, for community-documented checks, from a public prose description that you cite in `CheckSource` and in an XML doc comment. Set `CheckProvenance` on every scheme that has a check.
- Never put input values into `ToString()`, exception messages or diagnostics.
- Update Core's public API tracking files the way recent Core API additions did.
- Tests: apply the test-audit authoring gate. Test contracts (validation status, normalized output, scoring thresholds, overlap outcomes, masking output), not implementation details. Use issuer-published examples where they exist; otherwise compute vectors from the published weights and add a one-line comment with the derivation. No timing or allocation assertions outside `[Trait("Category", "Performance")]`.

### Files

```text
OfficeIMO.Core/Entities/
  EntityKind.cs                      readonly struct, well-known values, case-insensitive equality
  EntityScheme.cs                    abstract base (see design "Schemes"), Contains relations, MaskPartial
  EntityValidationResult.cs, EntityValidationStatus.cs, EntityCheckOutcome.cs, EntityCheckKind.cs, EntityCheckProvenance.cs
  EntityDetails.cs, EntitySex.cs, EntitySensitivity.cs, EntityFormatStyle.cs
  EntityCatalog.cs, EntityCatalogBuilder.cs   immutable catalog; Default = every pack below
  EntityFinder.cs, EntityFinderOptions.cs, EntityFindResult.cs, EntityOverlapPolicy.cs
  EntityMatch.cs, EntityEvidence.cs, EntityTextSegment.cs, EntitySegmentLocation.cs
  EntityMasker.cs, EntityMaskOptions.cs, EntityMaskStyle.cs, EntityPseudonymResult.cs
  IEntityRecognizer.cs, EntityRecognitionContext.cs, EntityCandidate.cs, EntityTokenShape.cs
  Internal/EntityTokenizer.cs        digit runs, alphanumeric runs, @ and URL anchors, boundary flags
  Internal/EntityScoring.cs          base classes and adjustments exactly as in the design's Scoring table
  Internal/EntityOverlapResolver.cs  rules 1-4 from the design's Overlaps section
  Internal/EntityContextMatcher.cs   keyword window, case- and diacritic-insensitive whole words
  Internal/CheckDigits.cs            weighted mod 10/11, ISO 7064 MOD 97-10 (streaming, no BigInteger), Luhn
  Poland/PolandSchemes.cs + one file per scheme, PolandRecognizers.cs, PolandKeywords.cs
  EuropeanUnion/EuVatSchemes.cs + one file per member state, EuVatRecognizer.cs, EuVatKeywords.cs
  UnitedKingdom/..., UnitedStates/..., Financial/..., Internet/...  same layout
  EntityCatalogBuilderExtensions.cs  AddPoland, AddEuVat, AddUnitedKingdom, AddUnitedStates, AddFinancial, AddInternet
  Financial/IbanStructure.cs         internal; moved from SharedSource/Invoicing/InvoiceBankAccountIdentity.cs, which is deleted
OfficeIMO.Entities.Tests/          new xUnit project referencing only OfficeIMO.Core (net8.0, net10.0, net472 on Windows); add to OfficeIMO.slnx and OfficeIMO.Core.slnf
```

### Schemes and rules

Each scheme: `Id`, `Kind`, `Country`, `DisplayName`, `Issuer`, `Specification`, normalized length bounds, `Check`, `Sensitivity`, `Validate`, `Normalize`, `Format` (`Compact`, `Display`), `MaskPartial`, and `Decode` where noted.

| Id | Kind | Normalized form | Validation | Base class | Notes |
| --- | --- | --- | --- | --- | --- |
| `PL.PESEL` | NationalId | 11 digits | Weights 1-3-7-9-1-3-7-9-1-3, check = (10 − sum mod 10) mod 10; date must exist with month offsets +80/+0/+20 only (+40/+60 → `UnsupportedVariant`) | check over digit run | `Decode`: birth date, sex (10th digit even = female). Future date or > 130 years old: valid but no `PlausibleDate`. |
| `PL.NIP` | TaxId | 10 digits | Weights 6-5-7-2-3-4-5-6-7, sum mod 11 = check; remainder 10 invalid. Accept optional `PL` prefix and groupings 3-3-2-2, 3-2-2-3 | check over digit run | `Display` = `123-456-32-18`. Issuer example 5260001246. |
| `PL.REGON` | CompanyId | 9 or 14 digits | Community-documented: 9 digits weights 8-9-2-3-4-5-6-7, 14 digits weights 2-4-8-5-0-9-7-3-6-1-2-4-8; sum mod 11, remainder 10 gives 0. A REGON 14 must also have a valid REGON 9 prefix | check over digit run | `Decode` for 14 digits: parent REGON 9 in `Properties["parent"]`; `Contains` REGON 9. |
| `PL.KRS` | CompanyId | 10 digits | Format only, left-pad shorter digit runs only when labelled | shape only | Context or label required to be reported. |
| `PL.IDCARD` | IdentityDocument | 3 letters + 6 digits | Community-documented: letters A=10..Z=35, weights 7-3-1-9-7-3-1-7-3 over all nine characters (the first digit is the check digit), sum mod 10 = 0 | short letter-digit check (0.30) | Context or label required. The MSWiA specimen `ABA300000` must validate. |
| `PL.PASSPORT` | IdentityDocument | 2 letters + 7 digits | Community-documented: letters A=10..Z=35, weights 7-3-9-1-7-3-1-7-3 over all nine characters (the first digit is the check digit), sum mod 10 = 0 | short letter-digit check (0.30) | Context or label required. Test against a published specimen number. |
| `EU.VAT.{CC}` | VatId | VIES prefix + national part | VIES format for all 27 + XI; check rules from the BMF document for AT, BE, DE, DK, EE, FI, HU, IE, IT, LT, LU, NL, PL, PT, SE, SI, SK, XI; ES individuals via the NIF/NIE letter rule; community-documented checks for BG, CY, CZ, EL, ES companies, FR, HR, LV, MT, RO, each citing its source; `NotDefined` only where none is documented | strong check with prefix; prefix without any check 0.60 | One recognizer for all. Country = ISO code (EL → GR, XI → GB). EU.VAT.PL `Contains` PL.NIP. Use each issuer example as a test. |
| `GB.VAT` | VatId | `GB` + 9 or 12 digits, or `GBGD`/`GBHA` + 3 | Weights 8..2 plus last two digits; valid when sum mod 97 = 0 or (sum + 55) mod 97 = 0; excluded ranges per BMF; GD 000-499, HA 500-999 | strong check with prefix | Shares rules with EU.VAT.XI. Example 434031494. |
| `GB.NINO` | NationalId | 2 letters + 6 digits + A-D | HMRC NIM39110 letter rules and excluded prefixes; suffix optional on input, normalized without spaces | distinctive shape | HMRC's specimen `QQ 12 34 56 A` fails the letter rule (Q is never used), so it is rejected. |
| `GB.COMPANY` | CompanyId | 8 characters | Digits, or a listed 2-letter prefix + 6 digits | shape only | Context or label required. |
| `US.SSN` | NationalId | 9 digits | Area not 000, 666, 900-999; group not 00; serial not 0000 | shape only | `Display` = `123-45-6789`; the hyphenated form gets `IssuerFormatting`. |
| `US.EIN` | TaxId | 9 digits | Prefix in the IRS valid-prefix list (as a constant table with the page's update date in a comment) | shape only | `Display` = `12-3456789`. |
| `IBAN` | BankAccount | country + 2 check digits + BBAN, upper case | Registry length and BBAN structure, check digits 02-98, MOD 97-10 (reuse `IbanStructure`) | strong check | `Decode`: country, BBAN. Registry examples as tests. |
| `BIC` | BankIdentifier | 8 or 11 upper-case characters | 4 letters + ISO 3166 country + 2 alphanumerics + optional 3 alphanumerics | shape only | Context or label required (8 capitals look like words). |
| `CARD` | PaymentCard | 12-19 digits | Luhn; first digit (major industry identifier) recorded | check over digit run | `KnownIssuerRange` when the first digit is 3-6. `Sensitivity = Financial`. |
| `ISIN` | SecurityId | 12 characters | 2 letters + 9 alphanumerics + check digit, letters mapped A=10..Z=35, then Luhn | strong check | Sensitivity `Public`. |
| `LEI` | LegalEntityId | 20 characters | Characters 5-6 `00`; MOD 97-10 over the whole value = 1 | strong check | Sensitivity `Organizational`. |
| `EMAIL` | Email | local@domain, domain lower-cased | RFC 5321 lengths; RFC 5322 dot-atom or quoted local part; SMTPUTF8 allowed; domain labels ≤ 63, IDNA-mapped via `IdnMapping` for the comparison key | distinctive shape | `Properties["key"]` = comparison key. No obsolete syntax, no comments, no IP literals in the finder (validation accepts IP literals). |
| `URL` | Url | as written, scheme and host lower-cased | `http`, `https`, `ftp`, `mailto` and `www.` forms; host must have a dot or be `localhost`; trailing sentence punctuation and unbalanced closing brackets excluded | distinctive shape | Sensitivity `Public`; masking off by default. |

### Finder behaviour to implement and test

- Tokenize once; recognizers read runs. Separators inside runs: single space, hyphen, dot, slash, U+00A0, U+2009, U+202F. A candidate never starts or ends inside a longer letter/digit run of the same class.
- Scoring exactly as the design's table; default `MinimumConfidence` 0.50; context window 48 characters before the value within the segment; label match via `EntityTextSegment.label`.
- Overlap rules 1-4; `Contains` relations: EU VAT PL contains PL NIP; REGON 14 contains REGON 9; GB VAT is not related to anything in stage 1.
- Reject before checks: all-same-digit runs, runs that parse as dates (`dd.mm.yyyy`, `yyyy-mm-dd`, `dd/mm/yyyy`) for digit-only schemes, runs immediately preceded by a currency sign or followed by a currency code or decimal separator plus two digits.
- Limits: `MaxMatches` (default 10,000), `MaxInputCharacters` (default 16,777,216); `Truncated` flag; cancellation checked between segments and every 64 KB.
- `EntityMatch.ToString()` = `"{Kind}/{Scheme.Id} [{Start}..{Start+Length}] {Confidence:0.00}"`.

Acceptance tests that must exist (values below are computed from the published weights):

| Input | Expected |
| --- | --- |
| `44051401359` alone in a sentence | no match at default threshold; one PL.PESEL match at 0.40 with `MinimumConfidence = 0.3` |
| `PESEL: 44051401359` | PL.PESEL, confidence ≥ 0.80, `ContextKeyword = "PESEL"`, `ContextLanguage = "pl"`; `Decode` gives 1944-05-14, male |
| `44051401358` with `PESEL:` | no match (check digit fails); `Validate` returns `InvalidCheckDigit` |
| `NIP 123-456-32-18` | PL.NIP, confidence ≥ 0.80, `IssuerFormatting` set |
| `VAT: PL1234563218` | EU.VAT.PL as the match, PL.NIP in `Contained` |
| cell `1234563218` with label `NIP` | PL.NIP ≥ 0.80 |
| `REGON 123456785` | PL.REGON (9), `CheckProvenance = CommunityDocumented`, confidence 0.80 from the keyword; `123456784` fails |
| `NIP 5260001246`, `ATU10223006`, `DE111111125`, `GB434031494`, `BE0776091951` | each valid with its issuer example |
| `FR40303265045` | EU.VAT.FR valid through the community-documented key (12 + 3 × (SIREN mod 97)) mod 97, `CheckProvenance = CommunityDocumented`; `FR41303265045` fails |
| `NIP 5260001246` | `CheckProvenance = IssuerPublished` |
| `ES12345678Z` | EU.VAT.ES valid through the NIF letter rule |
| `IBAN PL61 1090 1014 0000 0712 1981 2874` | IBAN, normalized `PL61109010140000071219812874`, `Display` round-trips to the grouped form, `MaskPartial` = `PL61 •••• •••• •••• •••• •••• 2874` |
| `GB82 WEST 1234 5698 7654 32`, `DE89 3704 0044 0532 0130 00` | IBAN valid (SWIFT registry examples) |
| `4111 1111 1111 1111` with `card` nearby | PaymentCard ≥ 0.80; partial mask keeps last four |
| `US0378331005` | ISIN valid; `US0378331006` invalid |
| `5493001KJTIIGC8Y1R12` | LEI valid (ISO 7064 MOD 97-10) |
| a 20-digit number containing a valid PESEL in its middle | no match |
| `12.05.2023` | no digit-scheme match |
| `jan.kowalski@example.com`, `jan@bücher.example` | Email; comparison key lower-cases and IDNA-maps the domain only |
| `https://example.com/a?b=c` in prose with trailing `.` | Url without the trailing dot |

Masking tests: `KindLabel`, `Fixed`, `Partial` per scheme, `Pseudonym` stable for the same normalized value and key and different for different keys; `Pseudonymize` returns the map; overlapping and adjacent matches mask once.

### Consolidation in the same PR

- Move `OfficeIMO.SharedSource/Invoicing/InvoiceBankAccountIdentity.cs` into Core as `Entities/Financial/IbanStructure.cs` (internal) and build `FinancialSchemes.Iban` on it. Delete the shared source file and both `Compile Include` links.
- Add a `ProjectReference` to `OfficeIMO.Core` in `OfficeIMO.Invoicing.csproj`.
- Switch every caller to `FinancialSchemes.Iban.IsValid`: in Invoicing `Poland/Fa3InvoiceWriter.MappingChecks.cs`, `Poland/Fa3InvoiceReader.Fields.cs`, `Xml/InvoiceSerializer.cs`, `Xml/InvoiceParser.CiiParties.cs`, `Xml/InvoiceParser.UblParties.cs`; in Pdf `PdfCiiDocumentHeaderInspector.PaymentAccounts.cs` and `PdfComplianceAnalyzer.ElectronicInvoice.PaymentAccounts.cs`. Behaviour must not change; run the Invoicing and Pdf test projects.

### Done means

- `dotnet build OfficeIMO.Core.slnf -c Release` and `dotnet test OfficeIMO.Entities.Tests -c Release` pass on net8.0 and net10.0; Pdf and Invoicing tests still pass; the AOT smoke project for Core still publishes without new warnings.
- The design doc's stage 1 roadmap item is removed from `Docs/ROADMAP.md` and the delivered contract is described in a short `OfficeIMO.Core` README section or `Docs/officeimo.entities.md` user guide (API, catalog, limits), with the design doc updated where the implementation differs.
- PR open, ready for review, CI green, CodeRabbit review addressed.
