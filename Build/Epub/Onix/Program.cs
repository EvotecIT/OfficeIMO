using OfficeIMO.Epub;
using OfficeIMO.Workflows;
using System.Net;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml;
using System.Xml.Schema;

if (args.Length != 2) throw new ArgumentException("Supply the local reference XSD directory and a new task-owned output directory.");
string schemaDirectory = Path.GetFullPath(args[0]), outputDirectory = Path.GetFullPath(args[1]);
string[] schemaNames = ["ONIX_BookProduct_3.1_reference.xsd", "ONIX_BookProduct_CodeLists.xsd", "ONIX_XHTML_Subset.xsd"];
var schemaFiles = schemaNames.Select(name => Path.Combine(schemaDirectory, name)).ToArray();
var schemas = new XmlSchemaSet { XmlResolver = new LocalSchemaResolver(schemaFiles) };
using (var reader = XmlReader.Create(schemaFiles[0], new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
    schemas.Add(null, reader);
schemas.Compile();
if (Directory.Exists(outputDirectory) || File.Exists(outputDirectory)) throw new IOException("Output already exists; use a new directory.");
Directory.CreateDirectory(outputDirectory);
var evidence = new List<object>();
var messageProducts = new List<BookOnixExportResult>();
var timestamp = new DateTimeOffset(2026, 10, 5, 12, 0, 0, TimeSpan.Zero);
foreach (var profile in new[] { (Name: "early", Language: "en", Onix: "eng", Notification: BookOnixNotification.Early),
    (Name: "advance", Language: "pl", Onix: "pol", Notification: BookOnixNotification.Advance),
    (Name: "confirmed", Language: "fr", Onix: "fre", Notification: BookOnixNotification.Confirmed),
    (Name: "priced", Language: "pl", Onix: "pol", Notification: BookOnixNotification.Confirmed),
    (Name: "collateral-xhtml", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collateral", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collateral-unicode", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collateral-usage", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collateral-usage-quantities", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collateral-licenses", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collateral-license-transition", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "review-ratings", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "complexity", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "complexity-audience", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "adult-unrated", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "adult-general", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "adult-advice", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "audience-headings", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "audience-codes", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "audience-grades", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "audience", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "audience-months", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "audience-open", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collection", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collection-identifiers", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collection-brand-universe", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collection-hierarchy", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "collection-frequency", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "no-collection", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "edition", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "no-edition", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "alternative-titles", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "title-sorting", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "title-no-prefix", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "discoverability", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "discount-coded", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "discounted", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "taxed", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "withdrawn", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "accessibility-unknown", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "accessibility-provenance", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "accessibility-claims", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed) }) {
    var project = BookProject.Create(profile.Name == "title-sorting" ? "The history & future" : "Publishing & metadata — " + profile.Name, profile.Language);
    project.Publication.Identifier = "urn:officeimo:fixture:onix:" + profile.Name;
    project.Publication.AddIdentifier("digital-isbn", new EpubIdentifierMetadata { Value = "978-0-306-40615-7", Kind = EpubIdentifierKind.Isbn13 });
    if (profile.Name == "discoverability") project.Publication.AddTitle("selected-title", new() {
        Text = "Selected catalog title", Kind = EpubTitleKind.Main
    });
    BookOnixCommercialMetadata commercial = CommercialFixtures.Create(profile.Name);
    var options = new BookOnixExportOptions {
        SenderName = "Example Press", PublisherName = "Example Press", RecordReference = "fixture-" + profile.Name,
        SentAt = timestamp, Notification = profile.Notification, IdentifierId = "digital-isbn", LanguageCode = profile.Onix,
        TitleId = profile.Name == "discoverability" ? "selected-title" : null,
        TitleSorting = profile.Name == "title-sorting" ? new() { Prefix = "The " } : profile.Name == "title-no-prefix" ? new() : null,
        AlternativeTitles = profile.Name == "alternative-titles" ? AlternativeTitleFixtures.Create() : [],
        Audience = ComplexityFixtures.Create(profile.Name) ?? AdultAudienceFixtures.Create(profile.Name) ?? AudienceFixtures.Create(profile.Name),
        CollateralTexts = profile.Name.StartsWith("collateral-usage", StringComparison.Ordinal) ? UsageFixtures.Create(profile.Name == "collateral-usage-quantities") : profile.Name.StartsWith("collateral-license", StringComparison.Ordinal) ? LicenseFixtures.Create(profile.Name == "collateral-license-transition") : profile.Name == "review-ratings" ? ReviewRatingFixtures.Create() : profile.Name == "collateral-xhtml" ? CollateralXhtmlFixtures.Create() : CollateralFixtures.Create(profile.Name),
        Collections = profile.Name == "collection-identifiers" ? CollectionIdentifierFixtures.Create() :
            profile.Name == "collection-brand-universe" ? BrandUniverseFixtures.Create() :
            profile.Name == "collection" ? CollectionFixtures.Create() :
            profile.Name == "title-sorting" ? TitleSortingFixtures.Collections() : CollectionHierarchyFixtures.Create(profile.Name),
        NoCollection = profile.Name == "no-collection",
        Edition = profile.Name == "edition" ? new() { Number = 2, VersionNumber = "1.2",
            Types = [BookOnixEditionType.Revised, BookOnixEditionType.Annotated],
            Statements = [new("Second revised and annotated edition", "eng"), new("Drugie wydanie", "pol")] }
            : profile.Name == "no-edition" ? new() { NoEdition = true } : null,
        Subjects = profile.Name == "discoverability" ? SubjectFixtures.Create() : [],
        Subtitle = "Explicit publishing assertions", PublicationDate = new DateOnly(2026, 10, 5),
        NoContributors = profile.Name == "early",
        Contributors = profile.Name == "early" ? [] : [
            new("Alice Example", BookOnixContributorRole.Author), new("Jan Kowalski", BookOnixContributorRole.Translator),
            new("Example Studio", BookOnixContributorRole.Illustrator, true), new("Anne Editor", BookOnixContributorRole.Editor),
            new("Other Creator", BookOnixContributorRole.Other)],
        Commercial = commercial, Accessibility = AccessibilityFixtures.Create(profile.Name)
    };
    if (profile.Name == "audience-grades") {
        var gradeCodes = new List<object>();
        foreach (var system in Enum.GetValues<BookOnixGradeSystem>())
            foreach (var grade in Enum.GetValues<BookOnixGrade>()) {
                var gradeResult = project.ExportOnix(options with { Audience = new() { GradeRanges = [new(system, grade, grade)] } },
                    schemas, new EpubWriteOptions { ModifiedAt = timestamp });
                System.Xml.Linq.XNamespace ns = "http://ns.editeur.org/onix/3.1/reference";
                var range = System.Xml.Linq.XDocument.Load(new MemoryStream(gradeResult.Bytes)).Descendants(ns + "AudienceRange").Single();
                gradeCodes.Add(new { system = system.ToString(), grade = grade.ToString(),
                    qualifier = range.Element(ns + "AudienceRangeQualifier")!.Value,
                    precision = range.Element(ns + "AudienceRangePrecision")!.Value,
                    value = range.Element(ns + "AudienceRangeValue")!.Value });
            }
        File.WriteAllText(Path.Combine(outputDirectory, "audience-grade-codes.json"), JsonSerializer.Serialize(gradeCodes,
            new JsonSerializerOptions { WriteIndented = true }));
    }
    if (profile.Name == "edition") {
        // Exercise every supported list 21 mapping against the authoritative schema.
        foreach (var type in Enum.GetValues<BookOnixEditionType>())
            project.ExportOnix(options with { Edition = new() { Types = [type] } }, schemas,
                new EpubWriteOptions { ModifiedAt = timestamp });
    }
    if (profile.Name == "collateral-xhtml") {
        bool invalidNestingRejected = false;
        try {
            project.ExportOnix(options with { CollateralTexts = [new() {
                Type = BookOnixTextType.Description, Audiences = [BookOnixContentAudience.Unrestricted],
                Texts = [new("<p><div>Invalid block nesting</div></p>") { Format = BookOnixCollateralTextFormat.Xhtml }]
            }] }, schemas);
        } catch (InvalidDataException error) when (error.Message.StartsWith("ONIX schema validation failed:", StringComparison.Ordinal)) {
            invalidNestingRejected = true;
        }
        if (!invalidNestingRejected) throw new InvalidDataException("The supplied schema accepted invalid XHTML paragraph nesting.");
    }
    var result = project.ExportOnix(options, schemas, new EpubWriteOptions { ModifiedAt = timestamp });
    if (profile.Name == "collateral-usage") UsageFixtures.ProbeRecentCodes(project, options, schemas, outputDirectory, timestamp);
    if (profile.Name == "collateral-licenses") LicenseFixtures.ProbeDatedLicenses(project, options, schemas, outputDirectory, timestamp);
    if (profile.Name == "alternative-titles") AlternativeTitleFixtures.Verify(result, schemas);
    if (profile.Name == "collection-identifiers") CollectionIdentifierFixtures.Verify(result, schemas);
    if (profile.Name == "collection-brand-universe") BrandUniverseFixtures.Verify(result, schemas);
    if (profile.Name == "collection") CollectionFixtures.Verify(result, schemas);
    if (profile.Name is "title-sorting" or "title-no-prefix")
        TitleSortingFixtures.Verify(result, profile.Name == "title-no-prefix", schemas);
    if (profile.Name is "collection-hierarchy" or "collection-frequency")
        CollectionHierarchyFixtures.Verify(result, profile.Name == "collection-frequency", schemas);
    if ((profile.Name.StartsWith("collateral-usage", StringComparison.Ordinal) || profile.Name.StartsWith("collateral-license", StringComparison.Ordinal) || profile.Name == "review-ratings" || profile.Name == "collateral-xhtml" || profile.Name == "audience-codes" || profile.Name == "audience-headings" || profile.Name.StartsWith("adult-", StringComparison.Ordinal) || profile.Name.StartsWith("complexity", StringComparison.Ordinal)) &&
        !BookOnixMessage.Create([result], schemas).Bytes.SequenceEqual(result.Bytes))
        throw new InvalidDataException("Record composition changed retained content or whitespace.");
    File.WriteAllBytes(Path.Combine(outputDirectory, profile.Name + ".onix"), result.Bytes);
    if (profile.Name is "early" or "priced") {
        // Existing single-record fixtures deliberately share an ISBN. Create a distinct second edition
        // for the catalog rather than silently treating repeated product identities as separate books.
        var catalogResult = result;
        if (profile.Name == "priced") {
            var catalogProject = BookProject.Create("Catalog second edition", profile.Language);
            catalogProject.Publication.Identifier = "urn:officeimo:fixture:onix:catalog-priced";
            catalogProject.Publication.AddIdentifier("catalog-isbn", new EpubIdentifierMetadata { Value = "9781861972712", Kind = EpubIdentifierKind.Isbn13 });
            catalogResult = catalogProject.ExportOnix(options with { IdentifierId = "catalog-isbn" }, schemas, new EpubWriteOptions { ModifiedAt = timestamp });
            File.WriteAllBytes(Path.Combine(outputDirectory, "catalog-priced.epub"), catalogResult.Publication.Bytes);
        }
        messageProducts.Add(catalogResult);
    }
    File.WriteAllBytes(Path.Combine(outputDirectory, profile.Name + ".epub"), result.Publication.Bytes);
    bool invalidCodeRejected = false;
    try { project.ExportOnix(options with { LanguageCode = "zzz" }, schemas); }
    catch (InvalidDataException error) when (error.Message.StartsWith("ONIX schema validation failed:", StringComparison.Ordinal)) { invalidCodeRejected = true; }
    if (!invalidCodeRejected) throw new InvalidDataException("The supplied schema did not reject an invalid list 74 value.");
    bool invalidCountryRejected = false;
    try { project.ExportOnix(options with { Commercial = commercial with {
        SalesRights = [new(BookOnixSalesRightsKind.Exclusive, new() { Countries = ["QQ"] })], Supplies = []
    } }, schemas); }
    catch (InvalidDataException error) when (error.Message.StartsWith("ONIX schema validation failed:", StringComparison.Ordinal)) { invalidCountryRejected = true; }
    if (!invalidCountryRejected) throw new InvalidDataException("The supplied schema did not reject an invalid list 91 value.");
    bool? invalidCurrencyRejected = null;
    BookOnixSupply? pricedSupply = commercial.Supplies.FirstOrDefault(supply => supply.Prices.Count != 0);
    if (pricedSupply != null) {
        invalidCurrencyRejected = false;
        try { project.ExportOnix(options with { Commercial = commercial with {
            Supplies = [pricedSupply with { Prices = [pricedSupply.Prices[0] with { CurrencyCode = "ZZZ" }] }]
        } }, schemas); }
        catch (InvalidDataException error) when (error.Message.StartsWith("ONIX schema validation failed:", StringComparison.Ordinal)) { invalidCurrencyRejected = true; }
        if (invalidCurrencyRejected != true) throw new InvalidDataException("The supplied schema did not reject an invalid list 96 value.");
    }
    evidence.Add(new { profile = profile.Name, onixSha256 = Convert.ToHexString(SHA256.HashData(result.Bytes)),
        epubSha256 = result.PublicationSha256, schemaValidation = "passed", invalidLanguageCode = "rejected",
        invalidCountryRejected, invalidCurrencyRejected, accessibilityAssertionsSynthetic = options.Accessibility != null });
}
var catalog = BookOnixMessage.Create(messageProducts, schemas);
File.WriteAllBytes(Path.Combine(outputDirectory, "catalog.onix"), catalog.Bytes);
File.WriteAllText(Path.Combine(outputDirectory, "evidence.json"), JsonSerializer.Serialize(new {
    schemaFiles = schemaFiles.Select(path => new { name = Path.GetFileName(path), sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) }),
    records = evidence,
    catalog = new { onixSha256 = Convert.ToHexString(SHA256.HashData(catalog.Bytes)),
        products = catalog.Products.Select(product => new { product.OnixSha256, product.PublicationSha256 }), schemaValidation = "passed" },
    independentValidation = "not-performed", retailerAcceptance = "not-performed"
}, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Generated and schema-validated {evidence.Count} ONIX/EPUB pairs in {outputDirectory}.");

sealed class LocalSchemaResolver(string[] files) : XmlResolver {
    private readonly HashSet<string> _files = new(files, StringComparer.Ordinal);
    public override ICredentials? Credentials { set { } }
    public override object GetEntity(Uri absoluteUri, string? role, Type? ofObjectToReturn) {
        if (!absoluteUri.IsFile || !_files.Contains(Path.GetFullPath(absoluteUri.LocalPath)))
            throw new XmlException("Only the three explicitly supplied local schema files may be resolved.");
        return File.OpenRead(absoluteUri.LocalPath);
    }
}
