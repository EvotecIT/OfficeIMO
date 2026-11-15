namespace OfficeIMO.Yaml;

/// <summary>
/// Variants of YAML that some producers write and that the specification does not allow. Each one is opt-in, through
/// <see cref="YamlReaderOptions.Dialect"/> or <see cref="YamlLoadOptions.Dialect"/>.
/// </summary>
public enum YamlDialect : byte {
    /// <summary>YAML as the specification defines it, with the leniencies <see cref="YamlReaderOptions.Strict"/> describes.</summary>
    Standard = 0,

    /// <summary>
    /// Unity's serialization files (scenes, prefabs, assets, <c>.meta</c> files). Every document uses the YAML 1.1 schema
    /// unless <see cref="YamlReaderOptions.Schema"/> is set; a document without <c>%TAG</c> directives uses the
    /// directives of the document before it, so <c>!u!</c> tags resolve after the first document; and the word
    /// <c>stripped</c> after a document's anchor on its <c>---</c> line marks the object as a prefab placeholder
    /// (<see cref="YamlReader.IsStripped"/>, <see cref="YamlDocumentPart.IsStripped"/>) instead of failing the read.
    /// </summary>
    Unity
}
