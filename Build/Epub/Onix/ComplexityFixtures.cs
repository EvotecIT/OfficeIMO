using OfficeIMO.Workflows;

internal static class ComplexityFixtures {
    internal static BookOnixAudienceMetadata? Create(string profile) => profile switch {
        "complexity" => new() { Complexities = [
            new(BookOnixComplexityScheme.FryReadability, "15"), new(BookOnixComplexityScheme.IoeBookBand, "Pink A"),
            new(BookOnixComplexityScheme.FountasAndPinnell, "Z+"), new(BookOnixComplexityScheme.Lexile, "AD0L"),
            new(BookOnixComplexityScheme.Atos, "17.0"), new(BookOnixComplexityScheme.FleschKincaid, "-2.5"),
            new(BookOnixComplexityScheme.GuidedReading, "J"), new(BookOnixComplexityScheme.ReadingRecovery, "20"),
            new(BookOnixComplexityScheme.Lix, "42"), new(BookOnixComplexityScheme.LexileAudio, "600L"),
            new(BookOnixComplexityScheme.LexileSpanish, "880L")] },
        "complexity-audience" => new() {
            Categories = [new(BookOnixAudienceType.Children)], Descriptions = [new("Synthetic audience and complexity assertions", "eng")],
            Complexities = [new(BookOnixComplexityScheme.Atos, "0"), new(BookOnixComplexityScheme.Atos, "17"),
                new(BookOnixComplexityScheme.FleschKincaid, "25.5")] },
        _ => null
    };
}
