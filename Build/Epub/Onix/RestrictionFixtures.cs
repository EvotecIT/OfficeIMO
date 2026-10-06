using OfficeIMO.Epub;
using OfficeIMO.Workflows;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Schema;

internal static class RestrictionFixtures {
    internal static void Write(BookProject project, BookOnixExportOptions options, XmlSchemaSet schemas, string output, DateTimeOffset timestamp) {
        var evidence = new List<object>(); byte[]? publication = null;
        foreach (var kind in Enum.GetValues<BookOnixSalesRestrictionKind>()) {
            bool needsOutlet = kind is BookOnixSalesRestrictionKind.RetailerExclusiveOrOwnBrand or BookOnixSalesRestrictionKind.RetailerExclusive or
                BookOnixSalesRestrictionKind.RetailerOwnBrand or BookOnixSalesRestrictionKind.RetailerException or
                BookOnixSalesRestrictionKind.SelectedSubscriptionServices or BookOnixSalesRestrictionKind.SubscriptionServiceExclusive;
            var restriction = new BookOnixSalesRestriction(kind) {
                Notes = [new("Publisher assertion & explanation", "eng"), new("Wyjaśnienie ograniczenia", "pol")],
                ValidFrom = new DateOnly(2026, 10, 1), ValidUntil = new DateOnly(2026, 10, 31),
                Outlets = needsOutlet ? [new() { Name = "Example & Books", NameLanguageCode = "eng", Identifiers = [
                    new(BookOnixSalesOutletScheme.Proprietary, "retailer-1") { SchemeName = "Publisher outlets" },
                    new(BookOnixSalesOutletScheme.Onix, "AMZ"), new(BookOnixSalesOutletScheme.Gln, "1234567890128"),
                    new(BookOnixSalesOutletScheme.San, "1234567")] }] : []
            };
            var result = project.ExportOnix(options with { Commercial = options.Commercial! with {
                SalesRights = options.Commercial.SalesRights.Select(right => right with { Restrictions = [restriction] }).ToArray(),
                Supplies = options.Commercial.Supplies.Select((supply, index) => supply with { MarketReference = "restricted-" + index, Restrictions = [restriction] }).ToArray()
            } }, schemas, new EpubWriteOptions { ModifiedAt = timestamp });
            publication ??= result.Publication.Bytes;
            if (!publication.SequenceEqual(result.Publication.Bytes)) throw new InvalidDataException("Restriction metadata altered EPUB bytes.");
            string name = "restriction-" + ((int)kind).ToString("D2") + ".onix";
            File.WriteAllBytes(Path.Combine(output, name), result.Bytes);
            if (!BookOnixMessage.Create([result], schemas).Bytes.SequenceEqual(result.Bytes)) throw new InvalidDataException("Restriction composition altered the full record.");
            if (kind == BookOnixSalesRestrictionKind.RetailerExclusive)
                File.WriteAllBytes(Path.Combine(output, "update-restriction.onix"), BookOnixMessage.CreateBlockUpdates([
                    new(result) { ReplaceMarketReferences = ["restricted-0"], ReplaceBlocks = [BookOnixBlock.PublishingDetail] }], schemas).Bytes);
            evidence.Add(new { file = name, kind = kind.ToString(), onixSha256 = result.OnixSha256, epubSha256 = result.PublicationSha256 });
        }
        var formatted = new BookOnixSalesRestriction(BookOnixSalesRestrictionKind.Unspecified) {
            Notes = [new("<p xmlns='http://www.w3.org/1999/xhtml'>Only <strong>approved</strong> outlets &amp; partners.</p>", "eng") {
                Format = BookOnixCollateralTextFormat.Xhtml
            }]
        };
        var formattedResult = project.ExportOnix(options with { Commercial = options.Commercial! with {
            SalesRights = options.Commercial.SalesRights.Select(right => right with { Restrictions = [formatted] }).ToArray(),
            Supplies = options.Commercial.Supplies.Select((supply, index) => supply with {
                MarketReference = "formatted-" + index, Restrictions = [formatted]
            }).ToArray()
        } }, schemas, new EpubWriteOptions { ModifiedAt = timestamp });
        if (!publication!.SequenceEqual(formattedResult.Publication.Bytes)) throw new InvalidDataException("Formatted restrictions altered EPUB bytes.");
        if (!BookOnixMessage.Create([formattedResult], schemas).Bytes.SequenceEqual(formattedResult.Bytes))
            throw new InvalidDataException("Formatted restriction composition changed the record.");
        bool invalidNestingRejected = false;
        try {
            project.ExportOnix(options with { Commercial = options.Commercial! with {
                SalesRights = options.Commercial.SalesRights.Select(right => right with { Restrictions = [formatted with {
                    Notes = [new("<p><div>Invalid nesting</div></p>") { Format = BookOnixCollateralTextFormat.Xhtml }]
                }] }).ToArray()
            } }, schemas);
        } catch (InvalidDataException error) when (error.Message.StartsWith("ONIX schema validation failed:", StringComparison.Ordinal)) {
            invalidNestingRejected = true;
        }
        if (!invalidNestingRejected) throw new InvalidDataException("Schema accepted invalid restriction-note XHTML nesting.");
        File.WriteAllBytes(Path.Combine(output, "restriction-xhtml.onix"), formattedResult.Bytes);
        var formattedUpdate = BookOnixMessage.CreateBlockUpdates([new(formattedResult) {
            ReplaceMarketReferences = ["formatted-0"], ReplaceBlocks = [BookOnixBlock.PublishingDetail]
        }], schemas);
        File.WriteAllBytes(Path.Combine(output, "update-restriction-xhtml.onix"), formattedUpdate.Bytes);
        File.WriteAllBytes(Path.Combine(output, "restrictions.epub"), publication!);
        File.WriteAllText(Path.Combine(output, "restrictions-evidence.json"), JsonSerializer.Serialize(new {
            records = evidence, epubSha256 = Convert.ToHexString(SHA256.HashData(publication!)),
            schemaValidation = "passed", recipientAcceptance = "not-performed", assertions = "synthetic publisher inputs"
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
