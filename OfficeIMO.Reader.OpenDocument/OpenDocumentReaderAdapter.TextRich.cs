using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentReaderAdapter {
    private static void ProjectTextRich(OdtDocument document, string path, ReaderOptions options,
        OpenDocumentRichResources resources, OpenDocumentRichBlocks blocks, OfficeDocumentReadResult result, CancellationToken token) {
        ProjectStory(document.ContentBlocks, "body", null, false);
        foreach (OdtHeaderFooter story in document.HeaderFooters) {
            token.ThrowIfCancellationRequested();
            ProjectStory(story.ContentBlocks, story.Kind, story.MasterPageName, true);
        }
        void ProjectStory(IEnumerable<OdtContentBlock> source, string story, string? anchor, bool retainText) {
            int index = 0;
            foreach (OdtContentBlock block in source) {
                token.ThrowIfCancellationRequested();
                var location = new ReaderLocation { Path = path, SourceBlockIndex = index++,
                    SourceBlockKind = story == "body" ? block.Kind.ToString().ToLowerInvariant() : story,
                    BlockAnchor = anchor == null ? null : "styles.xml#" + anchor + "/" + story };
                if (block.Paragraph != null) ProjectParagraph(block.Paragraph, location, retainText);
                else if (block.Table != null) {
                    int maxRows = options.MaxTableRows > 0 ? options.MaxTableRows : 200;
                    if (retainText) {
                        ReaderChunk table = BuildTableChunk(block.Table, path, location.SourceBlockIndex!.Value, options,
                            new ProjectionBudget(Math.Max(1, blocks.RemainingCharacters)), token);
                        foreach (ReaderTable value in table.Tables!) value.Location = location;
                        result.Tables = result.Tables.Concat(table.Tables!).ToArray();
                        blocks.Add(table.Text, "table", location);
                        if (table.Warnings != null) foreach (string warning in table.Warnings)
                            resources.AddDiagnostic("ODF_STORY_TABLE_BOUNDED", warning, location);
                    }
                    foreach (OdtTableRow row in block.Table.Rows.Take(maxRows))
                        foreach (OdtTableCell cell in row.Cells.Take(MaximumTableColumns))
                            foreach (OdtParagraph paragraph in cell.Paragraphs) ProjectParagraph(paragraph, location, false);
                }
            }
        }
        void ProjectParagraph(OdtParagraph paragraph, ReaderLocation location, bool retainText) {
            if (retainText) blocks.Add(paragraph.Text, paragraph.IsHeading ? "heading" : "paragraph", location);
            foreach (OdtHyperlink link in paragraph.Hyperlinks) resources.AddLink(link.Href, link.Text, location);
            foreach (OdtImage image in paragraph.Images) resources.AddImage(image.Path, location);
            var pending = new Stack<OdtNote>(paragraph.Notes.Reverse());
            while (pending.Count > 0) {
                token.ThrowIfCancellationRequested();
                OdtNote note = pending.Pop();
                var noteLocation = new ReaderLocation { Path = path, SourceBlockKind = note.Kind?.ToString().ToLowerInvariant() ?? "note",
                    SourceBlockIndex = location.SourceBlockIndex, BlockAnchor = note.Id };
                if (!note.HasOnlyParagraphs) resources.AddDiagnostic("ODF_NOTE_PROFILE_PARTIAL", "The native note contains body elements outside its paragraph profile.", noteLocation);
                foreach (OdtParagraph child in note.Paragraphs) {
                    blocks.Add(child.Text, "note", noteLocation);
                    foreach (OdtHyperlink link in child.Hyperlinks) resources.AddLink(link.Href, link.Text, noteLocation);
                    foreach (OdtImage image in child.Images) resources.AddImage(image.Path, noteLocation);
                    foreach (OdtNote nested in child.Notes.Reverse()) pending.Push(nested);
                }
            }
        }
    }
}
