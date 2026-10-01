namespace OfficeIMO.Rtf.Writing;

/// <summary>Document defaults shared by every semantic story and table writer.</summary>
internal sealed partial class RtfWriteContext {
    internal RtfWriteContext(int? defaultLanguageId, int unicodeSkipCount, int? defaultParagraphStyleId, bool preserveStyleInheritance = false,
        RtfDocument? formattingDocument = null, RtfParagraph? formattingParagraph = null) {
        DefaultLanguageId = defaultLanguageId;
        UnicodeSkipCount = unicodeSkipCount;
        DefaultParagraphStyleId = defaultParagraphStyleId;
        PreserveStyleInheritance = preserveStyleInheritance || defaultParagraphStyleId.HasValue;
        FormattingDocument = formattingDocument;
        FormattingParagraph = formattingParagraph;
    }

    internal int? DefaultLanguageId { get; }
    internal int UnicodeSkipCount { get; }
    internal int? DefaultParagraphStyleId { get; }
    internal bool PreserveStyleInheritance { get; }
    private RtfDocument? FormattingDocument { get; }
    private RtfParagraph? FormattingParagraph { get; }
    private MaterializationCounts? Counts { get; set; }

    internal RtfWriteContext ForParagraph(RtfParagraph paragraph) => FormattingDocument == null ? this :
        new RtfWriteContext(DefaultLanguageId, UnicodeSkipCount, DefaultParagraphStyleId, PreserveStyleInheritance, FormattingDocument, paragraph) { Counts = Counts };

    internal RtfParagraph ResolveParagraph(RtfParagraph paragraph) {
        if (FormattingDocument == null) return paragraph;
        RtfParagraph resolved = FormattingDocument.GetParagraphFormatting(paragraph);
        if (Counts != null && HasMaterializedParagraphFormatting(paragraph, resolved)) Counts.Paragraphs++;
        return resolved;
    }

    internal RtfRun ResolveRun(RtfRun run) {
        if (FormattingDocument == null || FormattingParagraph == null) return run;
        RtfRun resolved = FormattingDocument.GetRunFormatting(FormattingParagraph, run);
        if (Counts != null && HasMaterializedRunFormatting(run, resolved)) Counts.Runs++;
        return resolved;
    }

    internal RtfWriteContext ForCharacterScope(bool preserveStyleInheritance) =>
        PreserveStyleInheritance || !preserveStyleInheritance ? this :
        new RtfWriteContext(DefaultLanguageId, UnicodeSkipCount, DefaultParagraphStyleId, preserveStyleInheritance: true,
            formattingDocument: FormattingDocument, formattingParagraph: FormattingParagraph) { Counts = Counts };
}
