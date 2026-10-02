using OfficeIMO.IWork;

namespace OfficeIMO.Reader.IWork;

internal sealed partial class IWorkReadProjection {
    internal void AddPages(IWorkPagesProjection source) {
        var page = NewPage(null, "Pages document", null);
        var drawableLookup = source.Drawables.Where(drawable => Identity(drawable) != null)
            .GroupBy(drawable => Identity(drawable)!.RecordIdentifier)
            .Where(group => group.Count() == 1).ToDictionary(group => group.Key, group => group.Single());
        var inlineDrawables = new HashSet<ulong>();
        foreach (IWorkTextParagraph paragraph in source.Body.Paragraphs) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!paragraph.Runs.Any(run => run.InlineObject != null)) { AddParagraph(page, paragraph, "body"); continue; }
            ReportParagraphDetails(page, paragraph);
            var pending = new List<IWorkTextRun>();
            bool firstSegment = true;
            foreach (IWorkTextRun run in paragraph.Runs) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (run.InlineObject == null) { pending.Add(run); continue; }
                Flush();
                ulong identifier = run.InlineObject.Drawable.RecordIdentifier;
                if (drawableLookup.TryGetValue(identifier, out IWorkPagesDrawable? drawable)) {
                    ReaderLocation? location = AddDrawable(drawable, inline: true);
                    if (run.Hyperlink != null) AddLink(page, run.Hyperlink, location ?? Location(page));
                    inlineDrawables.Add(identifier);
                } else {
                    AddRunLinks(page, new[] { run });
                    _diagnostics.Add(new OfficeDocumentDiagnostic {
                        Category = OfficeDocumentDiagnosticCategory.Content, Code = "IWORK_READER_INLINE_OBJECT_UNAVAILABLE",
                        Message = "A Pages inline attachment has no supported projected drawable.",
                        Source = "OfficeIMO.Reader.IWork", Location = Location(page)
                    });
                }
            }
            Flush();
            if (paragraph.Text.Length > 0 || paragraph.ListLevel >= 0) {
                _diagnostics.Add(new OfficeDocumentDiagnostic {
                    Category = OfficeDocumentDiagnosticCategory.Content, Code = "IWORK_READER_INLINE_PARAGRAPH_SPLIT",
                    Message = "A Pages paragraph containing inline objects is split into ordered text and drawable blocks; its inline paragraph layout is not preserved.",
                    Source = "OfficeIMO.Reader.IWork", Location = Location(page)
                });
            }
            void Flush() {
                if (pending.Count == 0) return;
                AddParagraph(page, new IWorkTextParagraph(pending, paragraph.Style,
                    firstSegment ? paragraph.ListIdentifier : null, firstSegment ? paragraph.ListLevel : -1,
                    firstSegment ? paragraph.ListLabel : null, IWorkParagraphBreakKind.None,
                    firstSegment ? paragraph.ListFontName : null), "body");
                firstSegment = false; pending.Clear();
            }
        }
        foreach (IWorkPagesDrawable drawable in source.Drawables) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (Identity(drawable) is { } identity && inlineDrawables.Contains(identity.RecordIdentifier)) continue;
            AddDrawable(drawable);
        }
        foreach (IWorkTextContent header in source.Sections.SelectMany(section => section.SelectedHeaderContents)) AddRichContent(page, header, "header");
        foreach (IWorkTextContent footer in source.Sections.SelectMany(section => section.SelectedFooterContents)) AddRichContent(page, footer, "footer");
        if (source.PageLayout is { } layout) {
            page.Width = layout.WidthPoints;
            page.Height = layout.HeightPoints;
        }
        AddDiagnostics(source.Diagnostics);

        ReaderLocation? AddDrawable(IWorkPagesDrawable drawable, bool inline = false) {
            switch (drawable.Kind) {
                case IWorkPagesDrawableKind.TextBox:
                    AddTextBox(page, drawable.TextBox!, "text-box");
                    break;
                case IWorkPagesDrawableKind.Image:
                    return AddImage(page, drawable.Image!, includeAnchorBlock: inline);
                case IWorkPagesDrawableKind.Table:
                    return AddTable(page, drawable.Table!);
            }
            return null;
        }
    }

    private static IWorkObjectIdentity? Identity(IWorkPagesDrawable drawable) =>
        drawable.Image?.SourceIdentity ?? drawable.Table?.SourceIdentity ?? drawable.TextBox?.SourceIdentity;

}
