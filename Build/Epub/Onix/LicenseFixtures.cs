using OfficeIMO.Epub;
using OfficeIMO.Workflows;
using System.Text.Json;
using System.Xml.Schema;

internal static class LicenseFixtures {
    internal static IReadOnlyList<BookOnixCollateralText> Create(bool transition) => [new() {
        Type = BookOnixTextType.Excerpt, Audiences = [BookOnixContentAudience.EndCustomers],
        Texts = [new("Synthetic licensed excerpt.", "eng")], SourceLinks = ["https://example.org/excerpt"],
        PublishedOn = new(2026, 10, 5),
        Licenses = transition ? [
            new() { Names = [new("Previous terms")], ValidUntil = new(2026, 12, 31) },
            new() { Names = [new("Replacement terms")], ValidFrom = new(2027, 1, 1) }
        ] : [new() {
            Names = [new("Synthetic terms & permissions", "eng"), new("Przykładowe warunki", "pol")],
            Expressions = Enum.GetValues<BookOnixLicenseExpressionType>().Select(type =>
                new BookOnixLicenseExpression(type, "https://example.org/license?format=" + type + "&version=1")).ToArray(),
            ValidFrom = new(2026, 1, 1), ValidUntil = new(2026, 12, 31)
        }]
    }];

    // Record the supplied schema's actual behavior; never alter dates to satisfy an older schema.
    internal static void ProbeDatedLicenses(BookProject project, BookOnixExportOptions options,
        XmlSchemaSet schemas, string outputDirectory, DateTimeOffset timestamp) {
        byte[] before = project.ToProjectBytes();
        var item = Create(true)[0];
        var dated = item with { Licenses = [
            item.Licenses[0] with { ValidFrom = new(2026, 1, 1) },
            item.Licenses[1] with { ValidUntil = new(2027, 12, 31) }
        ] };
        string status, detail;
        try {
            var result = project.ExportOnix(options with { CollateralTexts = [dated] }, schemas,
                new EpubWriteOptions { ModifiedAt = timestamp });
            if (!BookOnixMessage.Create([result], schemas).Bytes.SequenceEqual(result.Bytes))
                throw new InvalidDataException("License composition changed the record.");
            File.WriteAllBytes(Path.Combine(outputDirectory, "dated-license-probe.onix"), result.Bytes);
            status = "passed"; detail = "Both licenses retain both date roles and pass the supplied schema.";
        } catch (InvalidDataException error) when (error.Message.StartsWith("ONIX schema validation failed:", StringComparison.Ordinal) &&
            error.Message.Contains("TextContent_EpubLicenseDateRole_must_be_unique", StringComparison.Ordinal)) {
            status = "schema-rejected"; detail = error.Message;
        }
        if (!project.ToProjectBytes().SequenceEqual(before)) throw new InvalidDataException("License probe changed the project.");
        File.WriteAllText(Path.Combine(outputDirectory, "dated-license-probe.json"), JsonSerializer.Serialize(new {
            status, detail, projectUnchanged = true, datesRewritten = false
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
