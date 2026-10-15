using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentReaderAdapter {
    private static void ProjectPresentationRich(OdpPresentation document, string path, ReaderOptions options,
        ReaderOpenDocumentOptions formatOptions, OpenDocumentRichResources resources, OpenDocumentRichBlocks blocks,
        OfficeDocumentReadResult result, CancellationToken token) {
        for (int index = 0; index < document.Slides.Count; index++) {
            token.ThrowIfCancellationRequested();
            OdpSlide slide = document.Slides[index];
            ProjectShapes(slide.Shapes, index + 1, null);
            if (formatOptions.IncludeSpeakerNotes && slide.SpeakerNotes != null)
                foreach (OdpParagraph paragraph in slide.SpeakerNotes.Paragraphs)
                    ProjectParagraph(paragraph, new ReaderLocation { Path = path, Slide = index + 1, SourceBlockKind = "speaker-notes" }, null);
        }
        // Keep each authored master once, independently of whether any slide references it.
        foreach (OdpMasterPage master in document.MasterPages) { token.ThrowIfCancellationRequested(); ProjectShapes(master.Shapes, null, master.Name); }

        void ProjectShapes(IEnumerable<OdpShape> source, int? slide, string? master) {
            var pending = new Stack<(OdpShape Shape, bool InGroup)>();
            foreach (OdpShape shape in source.Reverse()) pending.Push((shape, false));
            int shapeIndex = 0;
            while (pending.Count > 0) {
                token.ThrowIfCancellationRequested();
                var item = pending.Pop(); OdpShape shape = item.Shape;
                var location = new ReaderLocation { Path = path, Slide = slide, SourceBlockIndex = shapeIndex++,
                    SourceBlockKind = master == null ? "shape" : "master-shape", BlockAnchor = master == null ? "content.xml#" + (shape.XmlId ?? shape.Name) : "styles.xml#" + master + "/" + (shape.XmlId ?? shape.Name) };
                OfficeDocumentRegion? region = null;
                if (shape is OdpGroup group) {
                    foreach (OdpShape child in group.Shapes.Reverse()) pending.Push((child, true));
                } else if (!item.InGroup && shape is not OdpLine && string.IsNullOrWhiteSpace(shape.Transform)) {
                    try { region = OpenDocumentRichResources.Region(shape.Bounds); }
                    catch (NotSupportedException) { resources.AddDiagnostic("ODF_GEOMETRY_UNRESOLVED", "Native relative geometry was retained without assigning an absolute region.", location); }
                }
                if (item.InGroup || !string.IsNullOrWhiteSpace(shape.Transform))
                    resources.AddDiagnostic("ODF_GEOMETRY_TRANSFORM_NOT_APPLIED", "Shape geometry belongs to a group or authored transform; no computed absolute region was assigned.", location);
                foreach (OdpParagraph paragraph in shape.TextParagraphs) ProjectParagraph(paragraph, location, region);
                if (shape is OdpImage image) resources.AddImage(image.Path, location, region);
                if (shape is OdpTable table) {
                    int maximumRows = options.MaxTableRows > 0 ? options.MaxTableRows : 200;
                    foreach (OdpTableRow row in table.Rows.Take(maximumRows))
                        foreach (OdpTableCell cell in row.Cells.Take(MaximumTableColumns))
                            foreach (OdpParagraph paragraph in cell.Paragraphs) ProjectParagraph(paragraph, location, region);
                    if (master != null) {
                        ReaderTable value = BuildPresentationTable(table, path, 0, result.Tables.Count, options,
                            new ProjectionBudget(formatOptions.MaxExtractedCharacters), token, out bool truncatedRows, out bool truncatedColumns);
                        value.Location = location;
                        result.Tables = result.Tables.Concat(new[] { value }).ToArray();
                        if (truncatedRows || truncatedColumns) resources.AddDiagnostic("ODF_MASTER_TABLE_BOUNDED", "Master table rows or columns were omitted by the bounded table profile.", location);
                    }
                }
            }
        }
        void ProjectParagraph(OdpParagraph paragraph, ReaderLocation location, OfficeDocumentRegion? region) {
            blocks.Add(paragraph.Text, location.SourceBlockKind == "speaker-notes" ? "note" : "paragraph", location, region);
            var pending = new Stack<OdpInlineNode>(paragraph.InlineNodes.Reverse());
            while (pending.Count > 0) {
                token.ThrowIfCancellationRequested();
                OdpInlineNode node = pending.Pop();
                if (node.Hyperlink != null) resources.AddLink(node.Hyperlink.Href, node.Hyperlink.Text, location);
                foreach (OdpInlineNode child in node.Children.Reverse()) pending.Push(child);
            }
        }
    }
}
