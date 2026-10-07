using OfficeIMO.Epub;
using OfficeIMO.Workflows;
using System.Text.Json;
using System.Xml.Schema;

internal static class UsageFixtures {
    internal static IReadOnlyList<BookOnixCollateralText> Create(bool quantities) => [new() {
        Type = BookOnixTextType.Excerpt, Audiences = [BookOnixContentAudience.EndCustomers], Texts = [new("Synthetic usage assertions.")],
        SourceLinks = ["https://example.org/excerpt"], Licenses = [new() { Names = [new("Publisher terms")] }],
        UsageConstraints = quantities ? Enum.GetValues<BookOnixUsageUnit>().Select(Constraint).ToArray() :
            Enum.GetValues<BookOnixUsageType>().Where(t => t is not (BookOnixUsageType.NoConstraints or BookOnixUsageType.PrivatePurchaseAi or BookOnixUsageType.PrivateReadingAi))
                .Select(t => new BookOnixUsageConstraint(t, t == BookOnixUsageType.TextAndDataMining ? BookOnixUsageStatus.Prohibited : BookOnixUsageStatus.Unlimited)).ToArray()
    }];

    private static BookOnixUsageConstraint Constraint(BookOnixUsageUnit unit) {
        var limits = new List<BookOnixUsageLimit>();
        switch (unit) {
            case BookOnixUsageUnit.MediaDuration:
                limits.Add(BookOnixUsageLimit.Time(unit, TimeSpan.FromSeconds(30))); break;
            case BookOnixUsageUnit.StartTime:
            case BookOnixUsageUnit.EndTime:
                limits.Add(BookOnixUsageLimit.Time(BookOnixUsageUnit.StartTime, TimeSpan.FromMilliseconds(1230)));
                limits.Add(BookOnixUsageLimit.Time(BookOnixUsageUnit.EndTime, TimeSpan.FromSeconds(90))); break;
            case BookOnixUsageUnit.ValidFrom:
            case BookOnixUsageUnit.ValidUntil:
                limits.Add(BookOnixUsageLimit.Date(BookOnixUsageUnit.ValidFrom, new(2026, 1, 1)));
                limits.Add(BookOnixUsageLimit.Date(BookOnixUsageUnit.ValidUntil, new(2026, 12, 31))); break;
            case BookOnixUsageUnit.StartPage:
            case BookOnixUsageUnit.EndPage:
                limits.Add(BookOnixUsageLimit.Number(BookOnixUsageUnit.StartPage, 2));
                limits.Add(BookOnixUsageLimit.Number(BookOnixUsageUnit.EndPage, 10)); break;
            case BookOnixUsageUnit.PercentagePerPeriod:
                limits.Add(BookOnixUsageLimit.Number(unit, 12.50m));
                limits.Add(BookOnixUsageLimit.Number(BookOnixUsageUnit.Days, 7)); break;
            default: limits.Add(BookOnixUsageLimit.Number(unit, 5)); break;
        }
        return new(BookOnixUsageType.Preview, BookOnixUsageStatus.Limited) { Limits = limits };
    }

    internal static void ProbeRecentCodes(BookProject project, BookOnixExportOptions options, XmlSchemaSet schemas,
        string outputDirectory, DateTimeOffset timestamp) {
        var evidence = new List<object>();
        foreach (var type in new[] { BookOnixUsageType.PrivatePurchaseAi, BookOnixUsageType.PrivateReadingAi }) {
            byte[] before = project.ToProjectBytes(); string status, detail;
            var item = Create(false)[0] with { UsageConstraints = [
                new(BookOnixUsageType.TextAndDataMining, BookOnixUsageStatus.Prohibited), new(type, BookOnixUsageStatus.Unlimited)] };
            try {
                var result = project.ExportOnix(options with { CollateralTexts = [item] }, schemas, new EpubWriteOptions { ModifiedAt = timestamp });
                File.WriteAllBytes(Path.Combine(outputDirectory, "usage-" + type + ".onix"), result.Bytes);
                status = "passed"; detail = "The supplied schema accepts this usage code.";
            } catch (InvalidDataException error) when (error.Message.StartsWith("ONIX schema validation failed:", StringComparison.Ordinal) &&
                error.Message.Contains("EpubUsageType", StringComparison.Ordinal)) {
                status = "schema-rejected"; detail = error.Message;
            }
            if (!project.ToProjectBytes().SequenceEqual(before)) throw new InvalidDataException("Usage probe changed the project.");
            evidence.Add(new { type = type.ToString(), status, detail, projectUnchanged = true });
        }
        File.WriteAllText(Path.Combine(outputDirectory, "usage-recent-codes.json"), JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
    }
}
