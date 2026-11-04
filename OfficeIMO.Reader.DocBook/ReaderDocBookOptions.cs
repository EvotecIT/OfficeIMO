using OfficeIMO.DocBook;

namespace OfficeIMO.Reader.DocBook;

/// <summary>Options for adapting native DocBook documents to Reader chunks.</summary>
public sealed class ReaderDocBookOptions {
    /// <summary>Includes bounded profile diagnostics as chunk warnings.</summary>
    public bool IncludeDiagnostics { get; set; } = true;
    /// <summary>Native bounded read options.</summary>
    public DocBookLoadOptions ReadOptions { get; set; } = new DocBookLoadOptions();
    /// <summary>Bounded shared-model conversion options.</summary>
    public DocBookConversionOptions ConversionOptions { get; set; } = new DocBookConversionOptions();
}

internal static class ReaderDocBookOptionsCloner {
    internal static ReaderDocBookOptions Clone(ReaderDocBookOptions? options) {
        ReaderDocBookOptions source = options ?? new ReaderDocBookOptions();
        DocBookLoadOptions read = source.ReadOptions ?? new DocBookLoadOptions();
        DocBookConversionOptions conversion = source.ConversionOptions ?? new DocBookConversionOptions();
        return new ReaderDocBookOptions { IncludeDiagnostics = source.IncludeDiagnostics, ReadOptions = new DocBookLoadOptions {
            MaxInputBytes = read.MaxInputBytes, MaxCharacters = read.MaxCharacters, MaxDepth = read.MaxDepth,
            MaxElements = read.MaxElements, MaxAttributes = read.MaxAttributes, MaxCharactersFromEntities = read.MaxCharactersFromEntities
        }, ConversionOptions = new DocBookConversionOptions {
            MaxStructureDepth = conversion.MaxStructureDepth,
            MaxStructureNodes = conversion.MaxStructureNodes,
            MaxTotalTextCharacters = conversion.MaxTotalTextCharacters,
            MaxTableColumns = conversion.MaxTableColumns,
            MaxTableRows = conversion.MaxTableRows,
            MaxTableCells = conversion.MaxTableCells,
            MaxDetailedDiagnosticsPerCode = conversion.MaxDetailedDiagnosticsPerCode
        } };
    }
}
