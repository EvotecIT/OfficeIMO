using OfficeIMO.Opml;

namespace OfficeIMO.Reader.Opml;

/// <summary>Options for adapting native OPML documents to Reader chunks.</summary>
public sealed class ReaderOpmlOptions {
    /// <summary>Includes OPML validation diagnostics as chunk warnings.</summary>
    public bool IncludeDiagnostics { get; set; } = true;
    /// <summary>Native bounded read options.</summary>
    public OpmlLoadOptions ReadOptions { get; set; } = new OpmlLoadOptions();
    /// <summary>Bounded shared-model conversion options.</summary>
    public OpmlConversionOptions ConversionOptions { get; set; } = new OpmlConversionOptions();
}

internal static class ReaderOpmlOptionsCloner {
    internal static ReaderOpmlOptions Clone(ReaderOpmlOptions? options) {
        ReaderOpmlOptions source = options ?? new ReaderOpmlOptions();
        OpmlLoadOptions read = source.ReadOptions ?? new OpmlLoadOptions();
        OpmlConversionOptions conversion = source.ConversionOptions ?? new OpmlConversionOptions();
        return new ReaderOpmlOptions { IncludeDiagnostics = source.IncludeDiagnostics, ReadOptions = new OpmlLoadOptions {
            MaxInputBytes = read.MaxInputBytes, MaxCharacters = read.MaxCharacters, MaxDepth = read.MaxDepth,
            MaxElements = read.MaxElements, MaxOutlines = read.MaxOutlines, MaxAttributes = read.MaxAttributes
        }, ConversionOptions = new OpmlConversionOptions {
            MaxStructureDepth = conversion.MaxStructureDepth,
            MaxStructureNodes = conversion.MaxStructureNodes,
            MaxDetailedDiagnosticsPerCode = conversion.MaxDetailedDiagnosticsPerCode
        } };
    }
}
