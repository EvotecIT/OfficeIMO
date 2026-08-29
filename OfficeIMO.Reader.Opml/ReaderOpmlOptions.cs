using OfficeIMO.Opml;

namespace OfficeIMO.Reader.Opml;

/// <summary>Options for adapting native OPML documents to Reader chunks.</summary>
public sealed class ReaderOpmlOptions {
    /// <summary>Includes OPML validation diagnostics as chunk warnings.</summary>
    public bool IncludeDiagnostics { get; set; } = true;
    /// <summary>Native bounded read options.</summary>
    public OpmlReadOptions ReadOptions { get; set; } = new OpmlReadOptions();
}

internal static class ReaderOpmlOptionsCloner {
    internal static ReaderOpmlOptions Clone(ReaderOpmlOptions? options) {
        ReaderOpmlOptions source = options ?? new ReaderOpmlOptions();
        OpmlReadOptions read = source.ReadOptions ?? new OpmlReadOptions();
        return new ReaderOpmlOptions { IncludeDiagnostics = source.IncludeDiagnostics, ReadOptions = new OpmlReadOptions {
            MaxInputBytes = read.MaxInputBytes, MaxCharacters = read.MaxCharacters, MaxDepth = read.MaxDepth,
            MaxOutlines = read.MaxOutlines, MaxAttributes = read.MaxAttributes
        } };
    }
}
