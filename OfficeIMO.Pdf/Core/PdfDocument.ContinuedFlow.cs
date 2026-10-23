namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Captures a static continued story in ordinary flow without creating a page-associated note reference.</summary>
    /// <remarks>Used by document adapters for endnote stories. Content uses the common resumable note-flow profile;
    /// decorations are static paragraph, panel, image or rule flow and cannot contain recursive references or tables.</remarks>
    internal void ContinuedFlow(Action<PdfContentBuilder> build, Action<PdfContentBuilder>? separator,
        Action<PdfContentBuilder>? continuationSeparator, Action<PdfContentBuilder>? continuationNotice,
        bool reserveFinalNotice) {
        Guard.NotNull(build, nameof(build));
        IReadOnlyList<IPdfBlock> body = BuildFlowBlocks(build);
        IReadOnlyList<IPdfBlock> Capture(Action<PdfContentBuilder>? action) =>
            action == null ? Array.Empty<IPdfBlock>() : BuildFlowBlocks(action);
        var first = Capture(separator);
        var later = Capture(continuationSeparator);
        var notice = Capture(continuationNotice);
        ValidateFootnoteBlocks(body, _options);
        ValidateFootnoteBlocks(first, _options, allowTables: false);
        ValidateFootnoteBlocks(later, _options, allowTables: false);
        ValidateFootnoteBlocks(notice, _options, allowTables: false);
        AddBlock(new ContinuedFlowBlock(body, first, later, notice, reserveFinalNotice));
    }
}
