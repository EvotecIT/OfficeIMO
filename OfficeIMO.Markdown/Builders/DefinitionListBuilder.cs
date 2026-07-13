namespace OfficeIMO.Markdown;

/// <summary>
/// Builder for definition lists (term/definition).
/// </summary>
public sealed class DefinitionListBuilder {
    private readonly DefinitionListBlock _dl = new DefinitionListBlock();
    /// <summary>Adds a term/definition pair.</summary>
    public DefinitionListBuilder Item(string term, string definition) { _dl.AddEntry(term, definition); return this; }
    internal DefinitionListBlock Build() => _dl;
}
