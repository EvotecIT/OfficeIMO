using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentInspectionProjection {
    private static void AddPackageFeatures(ReaderInspectionMetadataBuilder builder, OdfDocument document, string path, CancellationToken token) {
        OdfFeatureReport report = document.InspectFeatures(token);
        foreach (OdfFeatureFinding finding in report.Findings)
            builder.Add("odf.package-feature", finding.Name, finding.Count.ToString(CultureInfo.InvariantCulture), "integer",
                new ReaderLocation { Path = path, SourceBlockKind = "package-part", BlockAnchor = finding.PartPath }, finding.PartPath,
                ReaderInspectionMetadataBuilder.Attributes(("nativeSupport", finding.Support.ToString()), ("executed", "false")));
        foreach (OdfFeatureDiagnostic diagnostic in report.Diagnostics)
            builder.Diagnostic(diagnostic.Code, diagnostic.Message,
                new ReaderLocation { Path = path, SourceBlockKind = "package-part", BlockAnchor = diagnostic.PartPath });
        builder.Coverage("package-features", report.IsComplete ? "complete" : "partial",
            "Native loaded XML feature presence and owner support profile; preserved scripts, objects, media and signatures are inert and not trust-validated.", new ReaderLocation { Path = path });
    }

    private static void AddPresentationVisibility(ReaderInspectionMetadataBuilder builder, OdpPresentation document, string path, CancellationToken token) {
        for (int index = 0; index < document.Slides.Count; index++) {
            token.ThrowIfCancellationRequested(); OdpSlide slide = document.Slides[index];
            var location = new ReaderLocation { Path = path, Slide = index + 1, SourceBlockKind = "slide" };
            builder.Add("odf.presentation", "Slide master reference", slide.MasterPageName, "native-reference", location,
                attributes: ReaderInspectionMetadataBuilder.Attributes(("layoutName", slide.LayoutName)));
            if (slide.Hidden) builder.Add(OfficeDocumentInspectionCategories.Visibility, "Slide", "hidden", "visibility-state", location);
            AddShapes(slide.Shapes, index + 1, null);
        }
        foreach (OdpMasterPage master in document.MasterPages) AddShapes(master.Shapes, null, master.Name);
        void AddShapes(IEnumerable<OdpShape> shapes, int? slide, string? master) {
            var pending = new Stack<OdpShape>(shapes.Reverse());
            while (pending.Count > 0) {
                token.ThrowIfCancellationRequested(); OdpShape shape = pending.Pop();
                if (shape.Hidden) builder.Add(OfficeDocumentInspectionCategories.Visibility, "Shape", "hidden", "visibility-state",
                    new ReaderLocation { Path = path, Slide = slide, SourceBlockKind = master == null ? "shape" : "master-shape", BlockAnchor = master }, shape.XmlId ?? shape.Name,
                    ReaderInspectionMetadataBuilder.Attributes(("masterName", master), ("name", shape.Name), ("transform", shape.Transform)));
                if (shape is OdpGroup group) foreach (OdpShape child in group.Shapes.Reverse()) pending.Push(child);
            }
        }
    }

    private static void AddDrawingVisibility(ReaderInspectionMetadataBuilder builder, OdgDocument document, string path, CancellationToken token) {
        for (int index = 0; index < document.Pages.Count; index++) {
            token.ThrowIfCancellationRequested(); OdgPage page = document.Pages[index];
            builder.Add("odf.drawing", "Page master reference", page.MasterPageName, "native-reference",
                new ReaderLocation { Path = path, Page = index + 1, SourceBlockKind = "drawing-page" });
            AddLayers(page.EffectiveLayers, index + 1, null);
        }
        foreach (OdgMasterPage master in document.MasterPages) AddLayers(master.Layers, null, master.Name);
        void AddLayers(IEnumerable<OdgLayer> layers, int? page, string? master) {
            foreach (OdgLayer layer in layers) {
                token.ThrowIfCancellationRequested();
                builder.Add(OfficeDocumentInspectionCategories.Visibility, "Layer", layer.Display.ToString(), "visibility-state",
                    new ReaderLocation { Path = path, Page = page, SourceBlockKind = "layer", BlockAnchor = master }, layer.Name,
                    ReaderInspectionMetadataBuilder.Attributes(("visibleOnScreen", layer.IsVisible() ? "true" : "false"),
                        ("visibleInPrint", layer.IsVisible(true) ? "true" : "false"), ("masterName", master), ("protectedEditingHint", layer.IsProtected ? "true" : "false")));
            }
        }
    }
}
