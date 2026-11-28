namespace OfficeIMO.Word.Pdf;


/// <summary>Source-only field metadata scoped to one PDF conversion operation.</summary>
internal sealed class NativeFieldResultDiagnosticIndex {
    internal NativeFieldResultDiagnosticIndex(WordDocument document) {
        FieldsByParagraph = WordFieldInventory.InspectWithSource(document)
            .Where(field => !field.HasStoredResult && field.SourceNode().Ancestor(WordXml.W, "p").Exists)
            .ToLookup(field => field.SourceNode().Ancestor(WordXml.W, "p"));
    }

    internal ILookup<OfficeXmlNode, WordFieldInventory.FieldInfoCandidate> FieldsByParagraph { get; }
    internal HashSet<OfficeXmlNode> DiagnosedSources { get; } = new();
}

public static partial class WordPdfConverterExtensions {
    private static void RecordNativeFieldResultDiagnostics(WordParagraph paragraph, WordPdfExportOptions options,
        string source, bool runningStory) {
        OfficeXmlNode paragraphNode = paragraph.Node();
        if (paragraphNode.IsNone) return;
        NativeFieldResultDiagnosticIndex index = options.FieldResultDiagnostics ??=
            new NativeFieldResultDiagnosticIndex(paragraph._document);
        foreach (WordFieldInventory.FieldInfoCandidate field in index.FieldsByParagraph[paragraphNode]) {
            if (!index.DiagnosedSources.Add(field.SourceNode()) || !IsNativeDisplayResultField(field.Info.FieldType) ||
                !WordConverterNodes.IsFieldVisibleBefore(field.SourceNode())) continue;
            // These running counters are calculated by the pagination owner even
            // when the source has no result separator or display cache.
            if (runningStory && TryGetNativeHeaderFooterFieldToken(field.Info.InstructionText, out _, out _)) continue;
            OfficeXmlNode formattingRun = field.SourceNode().Ancestor(WordXml.W, "r");
            if (formattingRun.IsNone) formattingRun = paragraphNode.CreateElement(WordXml.W, "r");
            if (IsNativeHiddenTextRun(paragraph._document.RunView(paragraphNode, formattingRun), paragraph)) continue;
            AddNativeExportWarning(options, "native-field-result-unavailable",
                source + " " + field.Info.PartUri + " field " + field.Info.Index.ToString(System.Globalization.CultureInfo.InvariantCulture),
                "The visible " + field.Info.FieldType + " field has no stored display result. PDF conversion does not evaluate this field here, so its display result is omitted; the Word source is unchanged.");
        }
    }

    // Keep non-displaying instructions (including ASK, SET, TC, XE and mail-merge
    // control fields) distinct from fields whose visible result needs a cache.
    private static bool IsNativeDisplayResultField(WordFieldType? type) => type is
        WordFieldType.AddressBlock or WordFieldType.Author or WordFieldType.AutoNum or
        WordFieldType.AutoNumLgl or WordFieldType.AutoNumOut or WordFieldType.AutoText or
        WordFieldType.AutoTextList or WordFieldType.Bibliography or WordFieldType.Citation or
        WordFieldType.Comments or WordFieldType.Compare or WordFieldType.CreateDate or
        WordFieldType.Database or WordFieldType.Date or WordFieldType.DocProperty or
        WordFieldType.DocVariable or WordFieldType.FileName or WordFieldType.FileSize or
        WordFieldType.GoToButton or WordFieldType.GreetingLine or WordFieldType.HyperlinkIf or
        WordFieldType.IncludeText or WordFieldType.Index or WordFieldType.Info or
        WordFieldType.Keywords or WordFieldType.LastSavedBy or WordFieldType.ListNum or
        WordFieldType.MacroButton or WordFieldType.MergeField or WordFieldType.MergeRec or
        WordFieldType.MergeSeq or WordFieldType.NoteRef or WordFieldType.NumChars or
        WordFieldType.NumPages or WordFieldType.NumWords or WordFieldType.Page or
        WordFieldType.PageRef or WordFieldType.PrintDate or WordFieldType.Quote or
        WordFieldType.Ref or WordFieldType.RevNum or WordFieldType.SaveDate or
        WordFieldType.Section or WordFieldType.SectionPages or WordFieldType.Seq or
        WordFieldType.StyleRef or WordFieldType.Subject or WordFieldType.Symbol or
        WordFieldType.Template or WordFieldType.Time or WordFieldType.Title or
        WordFieldType.TOA or WordFieldType.TOC or WordFieldType.UserAddress or
        WordFieldType.UserInitials or WordFieldType.UserName or WordFieldType.Formula or WordFieldType.EQ;
}
