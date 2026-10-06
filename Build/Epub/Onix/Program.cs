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
var timestamp = new DateTimeOffset(2026, 10, 5, 12, 0, 0, TimeSpan.Zero);
foreach (var profile in new[] { (Name: "early", Language: "en", Onix: "eng", Notification: BookOnixNotification.Early),
    (Name: "advance", Language: "pl", Onix: "pol", Notification: BookOnixNotification.Advance),
    (Name: "confirmed", Language: "fr", Onix: "fre", Notification: BookOnixNotification.Confirmed),
    (Name: "priced", Language: "pl", Onix: "pol", Notification: BookOnixNotification.Confirmed),
    (Name: "withdrawn", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "accessibility-unknown", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "accessibility-provenance", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed),
    (Name: "accessibility-claims", Language: "en", Onix: "eng", Notification: BookOnixNotification.Confirmed) }) {
    var project = BookProject.Create("Publishing & metadata — " + profile.Name, profile.Language);
    project.Publication.Identifier = "urn:officeimo:fixture:onix:" + profile.Name;
    project.Publication.AddIdentifier("digital-isbn", new EpubIdentifierMetadata { Value = "978-0-306-40615-7", Kind = EpubIdentifierKind.Isbn13 });
    BookOnixCommercialMetadata commercial = CommercialFixtures.Create(profile.Name);
    var options = new BookOnixExportOptions {
        SenderName = "Example Press", PublisherName = "Example Press", RecordReference = "fixture-" + profile.Name,
        SentAt = timestamp, Notification = profile.Notification, IdentifierId = "digital-isbn", LanguageCode = profile.Onix,
        Subtitle = "Explicit publishing assertions", PublicationDate = new DateOnly(2026, 10, 5),
        NoContributors = profile.Name == "early",
        Contributors = profile.Name == "early" ? [] : [
            new("Alice Example", BookOnixContributorRole.Author), new("Jan Kowalski", BookOnixContributorRole.Translator),
            new("Example Studio", BookOnixContributorRole.Illustrator, true), new("Anne Editor", BookOnixContributorRole.Editor),
            new("Other Creator", BookOnixContributorRole.Other)],
        Commercial = commercial, Accessibility = AccessibilityFixtures.Create(profile.Name)
    };
    var result = project.ExportOnix(options, schemas, new EpubWriteOptions { ModifiedAt = timestamp });
    File.WriteAllBytes(Path.Combine(outputDirectory, profile.Name + ".onix"), result.Bytes);
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
File.WriteAllText(Path.Combine(outputDirectory, "evidence.json"), JsonSerializer.Serialize(new {
    schemaFiles = schemaFiles.Select(path => new { name = Path.GetFileName(path), sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) }),
    records = evidence, independentValidation = "not-performed", retailerAcceptance = "not-performed"
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
