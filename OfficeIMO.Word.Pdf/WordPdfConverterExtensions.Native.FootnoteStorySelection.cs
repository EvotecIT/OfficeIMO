
namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Resolves the loaded stories actually used by this document's page-note layout.</summary>
    private static (OfficeXmlNode Separator, OfficeXmlNode ContinuationSeparator, OfficeXmlNode ContinuationNotice) ResolveNativePageFootnoteStories(
        WordDocument document) {
        OfficeXmlNode source = document.FootnotesNode();
        ISet<long> registered = GetNativeRegisteredNoteStories(document, "footnotePr", "footnote");
        OfficeXmlNode separator = ResolveNativeLoadedFootnoteStory(source, registered, continuation: false);
        // Native modern layout reuses the ordinary separator on continued pages.
        // An explicitly loaded empty story remains empty; only an absent loaded
        // story receives the detached implicit default.
        return (separator, UsesModernNativeWordLayout(document) ? separator :
            ResolveNativeLoadedFootnoteStory(source, registered, continuation: true),
            FindNativeLoadedFootnoteStory(source, registered, "continuationNotice"));
    }

    /// <summary>The special note story identifiers the settings part registers in <c>w:footnotePr</c> or <c>w:endnotePr</c>.</summary>
    private static ISet<long> GetNativeRegisteredNoteStories(WordDocument document, string propertiesName, string referenceName) {
        var registered = new HashSet<long>();
        foreach (OfficeXmlNode reference in document.SettingsNode().WChild(propertiesName).WElements(referenceName)) {
            string? value = reference.WAttr("id");
            if (value == null) continue;
            registered.Add(WordXml.ParseInt32(value) ?? throw new FormatException("The note story identifier is not an integer."));
        }
        return registered;
    }

    private static OfficeXmlNode ResolveNativeLoadedFootnoteStory(OfficeXmlNode source,
        ISet<long> registered, bool continuation) => ResolveNativeLoadedNoteStory(
            source.WElements("footnote"), registered, continuation, endnote: false);

    private static OfficeXmlNode ResolveNativeLoadedNoteStory(IEnumerable<OfficeXmlNode> source,
        ISet<long> registered, bool continuation, bool endnote) {
        string type = continuation ? "continuationSeparator" : "separator";
        OfficeXmlNode stored = FindNativeLoadedNoteStory(source, registered, type);
        if (stored.Exists) return stored;
        // Registration decides which stored definitions Word loads. Unregistered
        // definitions must not supply content, font dependencies or admission
        // restrictions. Construct the existing default without editing the part.
        return WordConverterNodes.CreateImplicitNoteSeparatorStory(endnote, continuation);
    }

    // Notices are optional in both compatibility profiles. A missing or
    // unregistered notice remains absent instead of acquiring an implicit story.
    private static OfficeXmlNode FindNativeLoadedFootnoteStory(OfficeXmlNode source,
        ISet<long> registered, string type) =>
        FindNativeLoadedNoteStory(source.WElements("footnote"), registered, type);

    private static OfficeXmlNode FindNativeLoadedNoteStory(IEnumerable<OfficeXmlNode> source,
        ISet<long> registered, string type) =>
        source.FirstOrDefault(note => string.Equals(note.WAttr("type"), type, StringComparison.Ordinal) &&
                WordXml.ParseInt64(note.WAttr("id")) is long id && registered.Contains(id));
}
