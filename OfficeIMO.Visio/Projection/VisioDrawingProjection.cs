using System.Threading;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal sealed partial class VisioDrawingProjection {
    private readonly VisioDrawingOptions _options;
    private readonly VisioDrawingConversionReport _report;
    private readonly CancellationToken _token;
    private int _shapes, _points;
    private long _imagePixels, _imageBytes;

    internal VisioDrawingProjection(VisioDrawingOptions options, VisioDrawingConversionReport report, CancellationToken token) {
        _options = options; _report = report; _token = token;
    }

    internal OfficeDrawing Project(VisioPage page, int ordinal) {
        _token.ThrowIfCancellationRequested();
        VisioRenderProjection projection = VisioRenderProjection.Create(page, 72D);
        // Preview surfaces impose a minimum size. Diagram pages retain the actual physical dimensions.
        var drawing = new OfficeDrawing(page.Width * projection.GeometryDensity, page.Height * projection.GeometryDensity);
        drawing.Fonts.AddRange(_options.Fonts);
        drawing.TextShapingProvider = _options.TextShapingProvider;
        drawing.TextShapingLanguage = _options.TextShapingLanguage;
        string location = $"page:{ordinal}:{page.Name}";
        _report.Add("VISIO_DRAWING_PAGE", "Physical page dimensions and cached drawing scale are projected in points.", OfficeConversionLossKind.None, location);
        var imageDiagnostics = new List<OfficeImageExportDiagnostic>();
        foreach (VisioPage contentPage in VisioBackgroundComposition.Resolve(page, _token, imageDiagnostics, location, _options.MaximumPages)) {
            bool background = !ReferenceEquals(contentPage, page);
            string contentLocation = background ? location + ":background:" + contentPage.Id + ":" + contentPage.Name : location;
            if (background) _report.Add("VISIO_DRAWING_BACKGROUND", "Cached background content is composed at its own physical scale behind the page.", OfficeConversionLossKind.None, contentLocation);
            ProjectContent(contentPage, drawing, VisioRenderProjection.CreateForContent(contentPage, page, 72D), imageDiagnostics, contentLocation);
        }
        foreach (OfficeImageExportDiagnostic diagnostic in imageDiagnostics)
            _report.Add(diagnostic.Code, diagnostic.Message, diagnostic.LossKind, diagnostic.Source ?? location);
        _token.ThrowIfCancellationRequested();
        return drawing;
    }

    private void ProjectContent(VisioPage page, OfficeDrawing drawing, VisioRenderProjection projection,
        List<OfficeImageExportDiagnostic> imageDiagnostics, string location) {
        if (page.Comments.Count > 0 || page.PreservedPageContentElements.Count > 0 || page.PreservedShapesContainerElements.Count > 0)
            _report.Add("VISIO_DRAWING_PAGE_CONTENT", "Comments and opaque page content are retained in the source but are not projected.", OfficeConversionLossKind.Omission, location);
        var visibility = new VisioRenderLayerVisibility(page, _options.LayerMode);
        var styles = new VisioNativeTextStyleResolver(page.OwnerDocument, _token, imageDiagnostics, location);
        var pending = new Stack<VisioShape>();
        for (int index = page.Shapes.Count - 1; index >= 0; index--) pending.Push(page.Shapes[index]);
        while (pending.Count > 0) {
            VisioShape shape = pending.Pop(); ChargeShape();
            string shapeLocation = location + ":shape:" + shape.Id;
            if (visibility.IsVisible(shape)) ProjectShape(page, drawing, shape, projection, styles, imageDiagnostics, shapeLocation);
            else _report.Add("VISIO_DRAWING_LAYER_EXCLUDED", "The shape is excluded by the selected layer policy; its children are considered independently.", OfficeConversionLossKind.None, shapeLocation);
            for (int index = shape.Children.Count - 1; index >= 0; index--) pending.Push(shape.Children[index]);
        }
        foreach (VisioConnector connector in page.Connectors) {
            ChargeShape(); string connectorLocation = location + ":connector:" + connector.Id;
            if (visibility.IsVisible(connector)) ProjectConnector(page, drawing, connector, projection, styles, connectorLocation);
            else _report.Add("VISIO_DRAWING_LAYER_EXCLUDED", "The connector is excluded by the selected layer policy.", OfficeConversionLossKind.None, connectorLocation);
        }
    }

    private void ChargeShape() {
        _token.ThrowIfCancellationRequested();
        if (_shapes >= _options.MaximumShapes) throw new InvalidDataException("Visio shape and connector count exceeds MaximumShapes.");
        _shapes++;
    }

    private void ChargePoints(int count) {
        _token.ThrowIfCancellationRequested();
        if (count > _options.MaximumGeometryPoints - _points) throw new InvalidDataException("Visio projected geometry exceeds MaximumGeometryPoints.");
        _points += count;
    }

    private void ChargeImagePixels(OfficeRasterImage image) {
        _token.ThrowIfCancellationRequested();
        long count = (long)image.Width * image.Height;
        if (count > _options.MaximumTotalImagePixels - _imagePixels) throw new InvalidDataException("Visio decoded image instances exceed MaximumTotalImagePixels.");
        _imagePixels += count;
    }

    private void ChargeImageBytes(byte[] bytes) {
        _token.ThrowIfCancellationRequested();
        if (bytes.LongLength > _options.MaximumTotalImageBytes - _imageBytes) throw new InvalidDataException("Visio projected image payloads exceed MaximumTotalImageBytes.");
        _imageBytes += bytes.LongLength;
    }
}
