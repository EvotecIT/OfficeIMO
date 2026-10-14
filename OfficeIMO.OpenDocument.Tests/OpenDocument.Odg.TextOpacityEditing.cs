using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgTextOpacityEditingTests {
    [Theory]
    [InlineData(0D, 0)]
    [InlineData(0.5D, 128)]
    [InlineData(1D, 255)]
    public void TypedEditingRetainsImportedTextAndUpdatesBothContainers(double opacity, byte alpha) {
        var doc = OdgDocument.LoadFlatXml(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Drawing", "libreoffice-transparent-text.fodg"));
        var imported = doc.Pages[0].Shapes[0].Paragraphs[0].Runs.Single();
        byte[] before = doc.GetPackageEntryBytes("content.xml")!;
        Assert.Equal(0.25D, imported.TextOpacity);
        Assert.Equal(before, doc.GetPackageEntryBytes("content.xml"));
        imported.TextOpacity = opacity;
        foreach (var read in RoundTrips(doc)) {
            var run = read.Pages[0].Shapes[0].Paragraphs[0].Runs.Single();
            Assert.Equal(opacity, run.TextOpacity);
            Assert.Equal("asdf", run.Text);
            Assert.Equal(OdfColor.Parse("#FF0000"), run.Color);
            var projected = read.Pages[0].ToDrawing();
            Assert.Equal(alpha, Text(projected.Value).Paragraphs.Single().Runs.Single().Color.A);
            Assert.DoesNotContain(projected.Report.Mappings, m => m.Feature.EndsWith(":text-opacity", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void LocalEditsDetachSharedStylesAndClearingExposesTheParent() {
        var doc = OdgDocument.Create(); var shape = Shape(doc);
        var parent = doc.Styles.CreateNamed("Body", OdfStyleFamily.Paragraph);
        parent.Color = OdfColor.Parse("#FF0000"); parent.FontFamily = "Arial";
        parent.FontSize = OdfLength.Points(14); parent.TextOpacity = 0.5D;
        var p = shape.AddParagraph(); p.StyleName = parent.Name;
        var shared = doc.Styles.CreateAutomatic(OdfStyleFamily.Text);
        shared.TextOpacity = 0.25D; shared.Bold = true;
        var a = p.AddRun("A"); a.StyleName = shared.Name;
        var b = p.AddRun("B"); b.StyleName = shared.Name;
        a.TextOpacity = 0D;
        Assert.NotEqual(a.StyleName, b.StyleName);
        Assert.Equal(0D, a.TextOpacity); Assert.Equal(0.25D, b.TextOpacity);
        shared.TextOpacity = 0.75D;
        Assert.Equal(0D, a.TextOpacity); Assert.Equal(0.75D, b.TextOpacity);
        a.TextOpacity = null;
        Assert.Equal(0.5D, a.TextOpacity); Assert.True(a.Bold);
        var nested = a.AddRun("D"); nested.TextOpacity = 0.25D;
        var link = p.AddHyperlink("C", "https://example.invalid/"); link.TextOpacity = 1D;
        foreach (var read in RoundTrips(doc)) {
            var actual = read.Pages[0].Shapes.Single().Paragraphs.Single();
            Assert.Equal(new double?[] { 0.5D, 0.25D, 0.75D }, actual.Runs.Select(r => r.TextOpacity));
            Assert.Equal(1D, actual.Hyperlinks.Single().TextOpacity);
            var projected = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            Assert.Equal(new byte[] { 128, 64, 191, 255 }, Text(projected.Value).Paragraphs.Single().Runs.Select(r => r.Color.A));
            Assert.Equal("ADBC", Text(projected.Value).PlainText);
        }
        Assert.Equal(0.5D, parent.TextOpacity); Assert.Equal(0.75D, shared.TextOpacity);
    }

    [Theory]
    [InlineData("50.0%", false)]
    [InlineData("25%", true)]
    [InlineData("101%", true)]
    public void ImportedAliasesAreReadWithoutMutationAndAnEditRepairsOnlyOpacity(string legacy, bool invalid) {
        var doc = OdgDocument.Create(); var shape = Shape(doc);
        var p = shape.AddParagraph("Aliases"); p.Color = OdfColor.Parse("#FF0000");
        var style = p.EnsureStyle(); style.Bold = true;
        style.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.LoExt + "opacity", "50%");
        style.SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.Draw + "opacity", legacy);
        XNamespace unknown = "urn:example:retained";
        style.SetProperty(OdfNamespaces.Style + "text-properties", unknown + "note", "Retained");
        byte[] before = doc.GetPackageEntryBytes("content.xml")!;
        if (invalid) {
            Assert.Throws<InvalidDataException>(() => style.TextOpacity);
            Assert.Throws<InvalidDataException>(() => p.TextOpacity);
            Assert.Throws<OdfConversionLossException>(() => doc.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        } else {
            Assert.Equal(0.5D, style.TextOpacity); Assert.Equal(0.5D, p.TextOpacity);
        }
        Assert.Equal(before, doc.GetPackageEntryBytes("content.xml"));
        p.TextOpacity = 0.2D;
        foreach (var read in RoundTrips(doc)) {
            var actual = read.Pages[0].Shapes.Single().Paragraphs.Single();
            Assert.Equal(0.2D, actual.TextOpacity); Assert.True(actual.Bold);
            var properties = actual.Styles.First().TextProperties!;
            Assert.Equal("20%", (string?)properties.Attribute(OdfNamespaces.LoExt + "opacity"));
            Assert.Null(properties.Attribute(OdfNamespaces.Draw + "opacity"));
            Assert.Equal("Retained", (string?)properties.Attribute(unknown + "note"));
            actual.TextOpacity = null;
            Assert.Null(actual.Styles.First().TextProperties!.Attribute(OdfNamespaces.LoExt + "opacity"));
            Assert.Null(actual.Styles.First().TextProperties!.Attribute(OdfNamespaces.Draw + "opacity"));
        }
    }

    [Fact]
    public void AValidLocalOverrideStopsAnInvalidAncestorUntilItIsCleared() {
        var doc = OdgDocument.Create(); var shape = Shape(doc);
        var p = shape.AddParagraph(); p.Color = OdfColor.Parse("#FF0000");
        p.EnsureStyle().SetProperty(OdfNamespaces.Style + "text-properties", OdfNamespaces.LoExt + "opacity", "101%");
        var run = p.AddRun("Nearest"); run.TextOpacity = 1D;
        Assert.Equal(1D, run.TextOpacity);
        doc.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        run.TextOpacity = null;
        Assert.Throws<InvalidDataException>(() => run.TextOpacity);
        Assert.Throws<OdfConversionLossException>(() => doc.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    [InlineData(-0.1D)]
    [InlineData(1.1D)]
    public void InvalidAssignmentsDoNotCreateStylesOrChangePackageXml(double opacity) {
        var doc = OdgDocument.Create(); var shape = Shape(doc); var p = shape.AddParagraph("Atomic");
        var named = doc.Styles.CreateNamed("Unchanged", OdfStyleFamily.Text);
        byte[] content = doc.GetPackageEntryBytes("content.xml")!, styles = doc.GetPackageEntryBytes("styles.xml")!;
        Assert.Throws<ArgumentOutOfRangeException>(() => p.TextOpacity = opacity);
        Assert.Throws<ArgumentOutOfRangeException>(() => named.TextOpacity = opacity);
        Assert.Throws<ArgumentOutOfRangeException>(() => shape.FillOpacity = opacity);
        Assert.Equal(content, doc.GetPackageEntryBytes("content.xml"));
        Assert.Equal(styles, doc.GetPackageEntryBytes("styles.xml"));
        Assert.Null(named.TextOpacity);
    }

    [Fact]
    public void PercentageSerializationKeepsSmallValidFractionsAndNormalizesNegativeZero() {
        var doc = OdgDocument.Create(); var shape = Shape(doc); var p = shape.AddParagraph("Small");
        p.Color = OdfColor.Parse("#FF0000"); p.TextOpacity = 1e-20D; shape.FillOpacity = 1e-20D;
        foreach (var read in RoundTrips(doc)) {
            var actual = read.Pages[0].Shapes.Single(); var text = actual.Paragraphs.Single();
            Assert.InRange(Math.Abs(text.TextOpacity!.Value / 1e-20D - 1D), 0D, 1e-15D);
            Assert.InRange(Math.Abs(actual.FillOpacity!.Value / 1e-20D - 1D), 0D, 1e-15D);
            read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            text.TextOpacity = -0D; actual.FillOpacity = -0D;
            Assert.Equal(0D, text.TextOpacity); Assert.Equal(0D, actual.FillOpacity);
        }
    }

    private static OdgShape Shape(OdgDocument doc) {
        var shape = doc.AddPage().Shapes.AddRectangle(OdfRect.FromCentimeters(1, 1, 12, 8));
        shape.FontFamily = "Arial"; shape.FontSize = OdfLength.Points(14);
        shape.EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "wrap-option", "no-wrap");
        return shape;
    }
    private static OfficeDrawingRichText Text(OfficeDrawing drawing) => Assert.Single(drawing.Elements.OfType<OfficeDrawingRichText>());
    private static OdgDocument[] RoundTrips(OdgDocument doc) {
        using var flat = new MemoryStream(); doc.SaveFlatXml(flat); flat.Position = 0;
        return new[] { doc, OdgDocument.Load(new MemoryStream(doc.ToBytes(new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.PreserveSource }))), OdgDocument.LoadFlatXml(flat) };
    }
}
