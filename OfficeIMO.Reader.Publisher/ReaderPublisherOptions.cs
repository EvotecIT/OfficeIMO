using OfficeIMO.Publisher;

namespace OfficeIMO.Reader.Publisher;

/// <summary>Controls native Publisher recovery and its semantic Reader projection.</summary>
public sealed class ReaderPublisherOptions {
    /// <summary>Native recovery settings. Item and text ceilings also bound emitted Reader projection work.</summary>
    public PublisherReadOptions? ReadOptions { get; set; }
    /// <summary>Copies original embedded image bytes into the rich result. Default: false.</summary>
    public bool IncludeImagePayloads { get; set; }

    internal ReaderPublisherOptions Clone() => new() {
        ReadOptions = (ReadOptions ?? new PublisherReadOptions()).Clone(),
        IncludeImagePayloads = IncludeImagePayloads
    };
}
