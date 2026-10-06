using OfficeIMO.Workflows;

internal static class ResourceFixtures {
    private static BookOnixResourceVersion Version(string link, string? format = null) => new() {
        Form = BookOnixResourceForm.Downloadable, Links = [new(link)], FileFormatCode = format,
        UpdatedOn = new(2026, 10, 6)
    };
    private static BookOnixSupportingResource Resource(BookOnixResourceContentType type, BookOnixResourceMode mode, params BookOnixResourceVersion[] versions) => new() {
        Type = type, Mode = mode, Audiences = [BookOnixContentAudience.Unrestricted], Versions = versions
    };

    internal static IReadOnlyList<BookOnixSupportingResource> Create(string profile) {
        if (profile == "resource-types") return Enum.GetValues<BookOnixResourceContentType>().Select(type =>
            Resource(type, BookOnixResourceMode.MultiMode, Version("https://example.org/resources/" + type))).ToArray();
        if (profile == "resource-formats") return new[] {
            "A103", "A104", "A105", "A106", "A107", "A108", "A111", "D101", "D102", "D103", "D104", "D105", "D106", "D107", "D108", "D109", "D401",
            "D501", "D502", "D503", "D504", "D505", "D506", "D507", "D508", "D509", "D510", "D511", "E101", "E105", "E107", "E112", "E113", "E115", "E116", "E139", "E140"
        }.Select(format => Resource(BookOnixResourceContentType.SampleContent, BookOnixResourceMode.MultiMode,
            Version("https://example.org/formats/" + format, format))).ToArray();
        if (profile != "supporting-resources") return [];
        return [
            Resource(BookOnixResourceContentType.FrontCover, BookOnixResourceMode.Image,
                Version("https://example.org/cover?edition=1&size=large", "D502") with {
                    ImageWidth = 1200, ImageHeight = 1800, FileName = "cover.jpg", ByteLength = 123456, Sha256 = new string('a', 64),
                    UsableFrom = new(2026, 1, 1), UsableUntil = new(2027, 12, 31),
                    UsageConstraints = [new(BookOnixUsageType.TextAndDataMining, BookOnixUsageStatus.Prohibited)],
                    Licenses = [new() { Names = [new("Promotional terms")], ValidUntil = new(2027, 12, 31) }]
                },
                Version("https://example.org/cover-small.jpg", "D502") with { Form = BookOnixResourceForm.Linkable, ImageWidth = 200, ImageHeight = 300 }) with {
                    Credits = [new("Publisher & artist", "eng"), new("Wydawca i artysta", "pol")],
                    Captions = [new("Front cover")], CopyrightHolders = [new("Synthetic publisher")],
                    AlternativeTexts = [new("A blue cover with the title in white", "eng"), new("Niebieska okładka z białym tytułem", "pol")]
                },
            Resource(BookOnixResourceContentType.ContributorReading, BookOnixResourceMode.Audio, Version("https://example.org/reading.mp3", "A103")) with { LengthMinutes = 2 },
            Resource(BookOnixResourceContentType.Trailer, BookOnixResourceMode.Video, Version("https://example.org/trailer.mp4", "D105")) with { LengthMinutes = 1 },
            Resource(BookOnixResourceContentType.SampleContent, BookOnixResourceMode.Text, Version("https://example.org/sample.pdf", "E107") with {
                Links = [new("https://example.org/en/sample.pdf", "eng"), new("https://example.org/pl/sample.pdf", "pol")]
            }),
            Resource(BookOnixResourceContentType.Widget, BookOnixResourceMode.Application, Version("https://example.org/widget") with { Form = BookOnixResourceForm.EmbeddableApplication }),
            Resource(BookOnixResourceContentType.PromotionalEventMaterial, BookOnixResourceMode.MultiMode, Version("https://example.org/promotion") with { Form = BookOnixResourceForm.Linkable })
        ];
    }
}
