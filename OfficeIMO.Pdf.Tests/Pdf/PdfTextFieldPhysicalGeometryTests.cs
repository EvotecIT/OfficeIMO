using System;
using System.Linq;
using OfficeIMO;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextFieldPhysicalGeometryTests {
    [Fact]
    public void FlowAndTableCellDefaultsKeepTheirSavedWidgetDimensions() {
        using var document = PdfDocument.Create(builder => builder.Content(content => content
            .TextField("Flow", value: "Ada")
            .Table(new[] { new[] { new PdfTableCell("Owner", formFields: new[] { PdfTableCellFormField.TextField("Cell", "Grace") }) } },
                style: new PdfTableStyle { HeaderRowCount = 0 })));
        using var readback = PdfDocument.Load(document.ToBytes());
        var fields = readback.Reader.DocumentInfo().FormFields;
        AssertWidget(fields.Single(field => field.Name == "Flow"), 180, 22);
        AssertWidget(fields.Single(field => field.Name == "Cell"), 120, 18);
        Assert.Equal("Ada", fields.Single(field => field.Name == "Flow").Value);
        Assert.Equal("Grace", fields.Single(field => field.Name == "Cell").Value);
    }

    [Fact]
    public void MixedUnitsReachNativeWidgetsThroughFlowCanvasAndTableCell() {
        var cell = PdfTableCellFormField.TextField("Cell", "Grace", OfficeLength.Inches(1), OfficeLength.Millimeters(6.35), OfficeLength.Points(9));
        Assert.Equal(OfficeLength.Points(72), cell.Width);
        Assert.Equal(OfficeLength.Points(18), cell.Height);
        using var document = PdfDocument.Create(builder => builder.Content(content => content
            .TextField("Flow", OfficeLength.Centimeters(2.54), OfficeLength.Pixels(24), "Ada", fontSize: OfficeLength.Points(11),
                spacingBefore: OfficeLength.Millimeters(2.54), spacingAfter: OfficeLength.Points(3))
            .Canvas(canvas => canvas.TextField("Canvas", "Lin", OfficeLength.Inches(1), OfficeLength.Points(120),
                OfficeLength.Pixels(96), OfficeLength.Millimeters(6.35), OfficeLength.Points(9)))
            .Table(new[] { new[] { new PdfTableCell("", formFields: new[] { cell }) } }, style: new PdfTableStyle { HeaderRowCount = 0 })));
        using var readback = PdfDocument.Load(document.ToBytes());
        var fields = readback.Reader.DocumentInfo().FormFields;
        AssertWidget(fields.Single(field => field.Name == "Flow"), 72, 18);
        AssertWidget(fields.Single(field => field.Name == "Canvas"), 72, 18);
        AssertWidget(fields.Single(field => field.Name == "Cell"), 72, 18);
        Assert.Contains(" 11 Tf", fields.Single(field => field.Name == "Flow").DefaultAppearance);
        Assert.Contains(" 9 Tf", fields.Single(field => field.Name == "Canvas").DefaultAppearance);
    }

    [Fact]
    public void ExplicitZeroIsRejectedInsteadOfBeingReplacedWithAnOmittedDefault() {
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create(builder => builder.Content(content => content.TextField("Flow", width: OfficeLength.Zero))));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPageCanvas().TextField("Canvas", "", OfficeLength.Zero, OfficeLength.Zero, OfficeLength.Points(72), OfficeLength.Zero));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfTableCellFormField.TextField("Cell", fontSize: OfficeLength.Zero));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDocument.Create(builder => builder.Content(content => content.TextField("Flow", spacingBefore: OfficeLength.Points(-1)))));
    }

    [Fact]
    public void SharedCellModelKeepsExistingChoiceFactoryGeometryAndSnapshot() {
        var style = new PdfFormFieldStyle { IsRequired = true };
        var field = PdfTableCellFormField.ChoiceField("Choice", new[] { "EU", "US" }, "EU", 72, 18, 9, style: style);
        var cell = new PdfTableCell("", formFields: new[] { field });
        style.IsRequired = false;
        using var document = PdfDocument.Create(builder => builder.Content(content => content.Table(
            new[] { new[] { cell } }, style: new PdfTableStyle { HeaderRowCount = 0 })));
        using var readback = PdfDocument.Load(document.ToBytes());
        var saved = Assert.Single(readback.Reader.DocumentInfo().FormFields);
        AssertWidget(saved, 72, 18);
        Assert.Equal("EU", saved.Value);
        Assert.True(saved.IsRequired);
    }

    private static void AssertWidget(PdfFormField field, double width, double height) {
        var widget = Assert.Single(field.Widgets);
        Assert.Equal(width, widget.Width, 3);
        Assert.Equal(height, widget.Height, 3);
    }
}
