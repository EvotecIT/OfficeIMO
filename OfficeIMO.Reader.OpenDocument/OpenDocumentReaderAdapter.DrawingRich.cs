using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentReaderAdapter {
    private static void ProjectDrawingRich(OdgDocument document, string path, OpenDocumentRichResources resources,
        OpenDocumentRichBlocks blocks, OfficeDocumentReadResult result, CancellationToken token) {
        for (int index = 0; index < document.Pages.Count; index++) {
            token.ThrowIfCancellationRequested();
            OdgPage source = document.Pages[index];
            ProjectShapes(source.Shapes, index + 1, null);
            OfficeDocumentPage? page = result.Pages.FirstOrDefault(page => page.Location.Page == index + 1);
            if (page != null) {
                try { page.Width = source.Width.ToPoints(); page.Height = source.Height.ToPoints(); }
                catch (Exception error) when (error is InvalidDataException || error is NotSupportedException) {
                    resources.AddDiagnostic("ODF_PAGE_DIMENSIONS_UNRESOLVED", error.Message, page.Location);
                }
            }
        }
        foreach (OdgMasterPage master in document.MasterPages) { token.ThrowIfCancellationRequested(); ProjectShapes(master.Shapes, null, master.Name); }

        void ProjectShapes(IEnumerable<OdgShape> source, int? page, string? master) {
            var pending = new Stack<(OdgShape Shape, bool InGroup)>();
            foreach (OdgShape shape in source.Reverse()) pending.Push((shape, false));
            int index = 0;
            while (pending.Count > 0) {
                token.ThrowIfCancellationRequested();
                var item = pending.Pop(); OdgShape shape = item.Shape;
                var location = new ReaderLocation { Path = path, Page = page, SourceBlockIndex = index++,
                    SourceBlockKind = master == null ? "shape" : "master-shape", BlockAnchor = master == null ? "content.xml#" + (shape.XmlId ?? shape.Name) : "styles.xml#" + master + "/" + (shape.XmlId ?? shape.Name) };
                if (shape.IsGroup) { foreach (OdgShape child in shape.Children.Reverse()) pending.Push((child, true)); continue; }
                OfficeDocumentRegion? region = null;
                if (!item.InGroup && shape.ElementName is not "line" and not "connector" && string.IsNullOrWhiteSpace(shape.Transform)) {
                    try { region = OpenDocumentRichResources.Region(shape.Bounds); }
                    catch (NotSupportedException) { resources.AddDiagnostic("ODF_GEOMETRY_UNRESOLVED", "Native relative geometry was retained without assigning an absolute region.", location); }
                }
                if (item.InGroup || !string.IsNullOrWhiteSpace(shape.Transform))
                    resources.AddDiagnostic("ODF_GEOMETRY_TRANSFORM_NOT_APPLIED", "Shape geometry belongs to a group or authored transform; no computed absolute region was assigned.", location);
                if (!string.IsNullOrEmpty(shape.Text)) blocks.Add(shape.Text, "paragraph", location, region);
                foreach (OfficeDocumentModelLink link in shape.InspectHyperlinks())
                    resources.AddLink(link.Uri ?? string.Empty, link.Text ?? string.Empty, location);
                if (shape.ImagePath != null) resources.AddImage(shape.ImagePath, location, region);
            }
        }
    }
}
