using System.Linq;
using System.Text;
using CodeGlyphX;
using CodeGlyphX.DataMatrix;
using CodeGlyphX.Rendering.Png;
using CodeGlyphX.Rendering.Svg;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.CodeGlyphX;
using Xunit;

namespace OfficeIMO.Drawing.CodeGlyphX.Tests;

public sealed class CodeGlyphDrawingIntegrationTests {
    [Fact]
    public void NeutralSvgBoundaryImportsQrWithoutTheAdapterApi() {
        QrCode qr = QrCode.Encode("https://evotec.xyz/codeglyphx");
        string svg = SvgQrRenderer.Render(qr.Modules, new QrSvgRenderOptions());

        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing, out int unsupported));
        Assert.NotNull(drawing);
        Assert.Equal(0, unsupported);
        Assert.NotEmpty(drawing!.Shapes);
    }

    [Fact]
    public void AdapterImportsStyledQrAsReusableShapes() {
        QrCode qr = QrCode.Encode("STYLED-QR-OFFICEIMO");
        var options = new QrSvgRenderOptions {
            ModuleShape = QrPngModuleShape.Circle,
            ModuleScale = 0.78,
            DarkColor = "#2457A6",
            LightColor = "#F7FAFF"
        };

        OfficeDrawing drawing = qr.ToOfficeDrawing(out int unsupported, options);

        Assert.Equal(0, unsupported);
        Assert.True(drawing.Shapes.Count > 50);
        Assert.Contains(drawing.Shapes, item => item.Shape.Kind == OfficeShapeKind.Ellipse);
    }

    [Fact]
    public void AdapterImportsDataMatrix() {
        BitMatrix modules = DataMatrixEncoder.Encode("OFFICEIMO-DATAMATRIX-1234567890");

        OfficeDrawing drawing = modules.ToOfficeDrawing(out int unsupported, new MatrixSvgRenderOptions());

        Assert.Equal(0, unsupported);
        Assert.NotEmpty(drawing.Shapes);
        Assert.True(drawing.Width > 0D);
        Assert.True(drawing.Height > 0D);
    }

    [Fact]
    public void AdapterImportsDataBarExpandedStackedOutput() {
        BitMatrix modules = MatrixBarcodeEncoder.Encode(BarcodeType.GS1DataBarExpandedStacked, "1234567890");

        OfficeDrawing drawing = modules.ToOfficeDrawing(out int unsupported, new MatrixSvgRenderOptions());

        Assert.Equal(0, unsupported);
        Assert.NotEmpty(drawing.Shapes);
        Assert.True(drawing.Width > drawing.Height);
    }

    [Fact]
    public void AdapterKeepsLinearBarcodeLabelAsSearchableText() {
        const string label = "ORDER-1234";
        Barcode1D barcode = BarcodeEncoder.Encode(BarcodeType.Code128, label);
        var options = new BarcodeSvgRenderOptions { LabelText = label };

        OfficeDrawing drawing = barcode.ToOfficeDrawing(out int unsupported, options);

        Assert.Equal(0, unsupported);
        Assert.NotEmpty(drawing.Shapes);
        OfficeDrawingText text = Assert.Single(drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(label, text.Text);
    }
}
