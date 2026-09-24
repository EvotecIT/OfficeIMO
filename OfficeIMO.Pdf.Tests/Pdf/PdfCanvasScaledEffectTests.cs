using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCanvasScaledEffectTests {
    [Fact]
    public void UniformScaleKeepsInteractiveFieldAtScaledPageCoordinates() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 300D,
                PageHeight = 200D,
                MarginLeft = 0D,
                MarginTop = 0D,
                MarginRight = 0D,
                MarginBottom = 0D
            })
            .Canvas(canvas => canvas.Effect(OfficeTransform.Scale(0.5D, 0.5D), 1D,
                content => content.TextField("query", "Hello", 20D, 120D, 100D, 24D, fontSize: 12D)))
            .ToBytes();

        PdfFormField field = Assert.Single(PdfInspector.Inspect(pdf).FormFields);
        PdfFormWidget widget = Assert.Single(field.Widgets);
        Assert.Equal("Hello", field.Value);
        Assert.Equal(10D, widget.X1, 2);
        Assert.Equal(60D, widget.X2, 2);
        Assert.Equal(128D, widget.Y1, 2);
        Assert.Equal(140D, widget.Y2, 2);
    }

    [Fact]
    public void UniformScaleMovesNamedDestinationWithContent() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 300D,
                PageHeight = 200D,
                MarginLeft = 0D,
                MarginTop = 0D,
                MarginRight = 0D,
                MarginBottom = 0D
            })
            .Canvas(canvas => canvas.Effect(OfficeTransform.Scale(0.5D, 0.5D), 1D,
                content => content.NamedDestination("target", 20D, 100D)))
            .ToBytes();

        PdfNamedDestination destination = Assert.Single(PdfInspector.Inspect(pdf).NamedDestinations);
        Assert.Equal("target", destination.Name);
        Assert.Equal(150D, destination.DestinationTop!.Value, 2);
    }

    [Fact]
    public void UniformScaleKeepsCheckboxChoiceAndRadioWidgets() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 300D,
                PageHeight = 200D,
                MarginLeft = 0D,
                MarginTop = 0D,
                MarginRight = 0D,
                MarginBottom = 0D
            })
            .Canvas(canvas => canvas.Effect(OfficeTransform.Scale(0.5D, 0.5D), 1D, content => content
                .CheckBox("accepted", true, 20D, 20D, 20D, 20D)
                .ChoiceField("color", new[] { "red", "blue" }, new[] { "blue" }, 20D, 60D, 100D, 24D)
                .RadioButton("size", "small", true, 20D, 100D, 20D, 20D)
                .RadioButton("size", "large", false, 60D, 100D, 20D, 20D)))
            .ToBytes();

        PdfFormField[] fields = PdfInspector.Inspect(pdf).FormFields.ToArray();
        Assert.Equal(3, fields.Length);
        Assert.Equal(10D, Assert.Single(fields, field => field.Name == "accepted").Widgets[0].X1, 2);
        Assert.Equal("blue", Assert.Single(fields, field => field.Name == "color").Value);
        Assert.Equal(2, Assert.Single(fields, field => field.Name == "size").Widgets.Count);
    }

    [Fact]
    public void UniformScalePreservesChoiceRowsAndScalesAppearanceContent() {
        string[] options = Enumerable.Range(0, 8).Select(index => "Option " + index).ToArray();
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 300D,
                PageHeight = 200D,
                MarginLeft = 0D,
                MarginTop = 0D,
                MarginRight = 0D,
                MarginBottom = 0D
            })
            .Canvas(canvas => canvas.Effect(OfficeTransform.Scale(0.5D, 0.5D), 1D,
                content => content.ChoiceField("options", options, new[] { "Option 5" }, 20D, 20D, 120D, 80D,
                    fontSize: 12D, isComboBox: false)))
            .ToBytes();

        string serialized = PdfEncoding.Latin1GetString(pdf);
        Assert.Contains("/TI 3", serialized, StringComparison.Ordinal);
        Assert.Contains("0.5 0 0 0.5 0 0 cm", serialized, StringComparison.Ordinal);
    }

    [Fact]
    public void NonUniformOrTranslucentEffectsStillRejectInteractiveFields() {
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().Canvas(canvas => canvas.Effect(
            OfficeTransform.Scale(0.5D, 0.75D), 1D,
            content => content.TextField("query", "Hello", 20D, 20D, 100D, 24D))));
        Assert.Throws<ArgumentException>(() => PdfDocument.Create().Canvas(canvas => canvas.Effect(
            OfficeTransform.Scale(0.5D, 0.5D), 0.5D,
            content => content.TextField("query", "Hello", 20D, 20D, 100D, 24D))));
    }
}
