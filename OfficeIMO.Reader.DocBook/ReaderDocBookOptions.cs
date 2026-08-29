using OfficeIMO.DocBook;

namespace OfficeIMO.Reader.DocBook;

/// <summary>Options for adapting native DocBook documents to Reader chunks.</summary>
public sealed class ReaderDocBookOptions {
    /// <summary>Includes bounded profile diagnostics as chunk warnings.</summary>
    public bool IncludeDiagnostics { get; set; } = true;
    /// <summary>Native bounded read options.</summary>
    public DocBookReadOptions ReadOptions { get; set; } = new DocBookReadOptions();
}

internal static class ReaderDocBookOptionsCloner {
    internal static ReaderDocBookOptions Clone(ReaderDocBookOptions? options) {
        ReaderDocBookOptions source = options ?? new ReaderDocBookOptions();
        DocBookReadOptions read = source.ReadOptions ?? new DocBookReadOptions();
        return new ReaderDocBookOptions { IncludeDiagnostics = source.IncludeDiagnostics, ReadOptions = new DocBookReadOptions {
            MaxInputBytes = read.MaxInputBytes, MaxCharacters = read.MaxCharacters, MaxDepth = read.MaxDepth,
            MaxElements = read.MaxElements, MaxAttributes = read.MaxAttributes, MaxCharactersFromEntities = read.MaxCharactersFromEntities
        } };
    }
}
