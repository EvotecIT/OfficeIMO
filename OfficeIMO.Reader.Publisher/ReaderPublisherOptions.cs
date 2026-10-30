using OfficeIMO.Publisher;

namespace OfficeIMO.Reader.Publisher;

/// <summary>Controls native Publisher recovery and its semantic Reader projection.</summary>
public sealed class ReaderPublisherOptions {
    /// <summary>Native recovery settings. Item and text ceilings also bound emitted Reader projection work.</summary>
    public PublisherReadOptions? ReadOptions { get; set; }
    /// <summary>Copies the document's current embedded image bytes into the rich result, including in-memory replacements. Default: false.</summary>
    public bool IncludeImagePayloads { get; set; }

    internal ReaderPublisherOptions Clone() => new() {
        ReadOptions = (ReadOptions ?? new PublisherReadOptions()).Clone(),
        IncludeImagePayloads = IncludeImagePayloads
    };
}
