using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgDateTimeFieldTests {
    [Theory]
    [InlineData(OdfTextFieldKind.Date, "2024-02-29", null, "2024-02-29")]
    [InlineData(OdfTextFieldKind.Date, "2024-02-29T00:30:00+14:00", "P1D", "2024-03-01")]
    [InlineData(OdfTextFieldKind.Date, "2024-01-31", "P1M", "2024-02-29")]
    [InlineData(OdfTextFieldKind.Date, "2024-02-29", "P1Y", "2025-02-28")]
    [InlineData(OdfTextFieldKind.Date, "2024-02-29", "P1Y1M", "2025-03-29")]
    [InlineData(OdfTextFieldKind.Date, "2024-03-01", "-P1D", "2024-02-29")]
    [InlineData(OdfTextFieldKind.Time, "13:07:09+14:00", "PT90S", "13:08:09")]
    [InlineData(OdfTextFieldKind.Time, "13:07:09", "-PT90S", "13:06:09")]
    [InlineData(OdfTextFieldKind.Time, "13:07:09", "PT59.999S", "13:07:09")]
    [InlineData(OdfTextFieldKind.Time, "13:07:09", "PT59.9999999999999999999999999999999S", "13:07:09")]
    [InlineData(OdfTextFieldKind.Time, "23:59:09", "PT90S", "00:00:09")]
    [InlineData(OdfTextFieldKind.Time, "00:00:09", "-PT90S", "23:59:09")]
    [InlineData(OdfTextFieldKind.Time, "24:00:00", null, "00:00:00")]
    [InlineData(OdfTextFieldKind.Time, "2024-02-29T13:07:09Z", null, "13:07:09")]
    public void StoredCivilValuesAndNativeAdjustmentsSurviveBothContainersWithoutRefreshingXml(OdfTextFieldKind kind, string value, string? adjustment, string expected) {
        var document = Create(kind, out var field); field.DateTimeValueLexical = value; field.DateTimeAdjustmentLexical = adjustment;
        foreach (OdgDocument read in RoundTrips(document)) {
            string[] before = Parts(read);
            Assert.Equal("cache", Text(read)); Assert.Equal(expected, Text(read, Stored())); Assert.Equal(before, Parts(read));
            var saved = read.Pages[0].Shapes[0].Paragraphs[0].Fields.Single();
            Assert.Equal(value, saved.DateTimeValueLexical); Assert.Equal(adjustment, saved.DateTimeAdjustmentLexical);
        }
    }

    [Fact]
    public void RefreshUsesOneExplicitCivilTimestampAcrossMastersGroupsSpansAndFixedFields() {
        var document = Create(OdfTextFieldKind.Date, out var fixedField); fixedField.IsFixed = true; fixedField.DateTimeValueLexical = "2010-01-02";
        var first = document.Pages[0]; var second = document.AddPage(); second.MasterPageName = first.MasterPageName;
        var p = first.MasterShapes.AddGroup().Children.AddTextBox(Bounds(), "On ").Paragraphs[0];
        var span = p.AddRun(); span.Bold = true; var date = span.AddField(OdfTextFieldKind.Date, ""); date.DataStyleName = "D";
        p.AddText(" at "); var time = p.AddHyperlink("", "https://example.invalid/time").AddField(OdfTextFieldKind.Time, "");
        document.Styles.CreateTimeStyle("T"); time.DataStyleName = "T";
        var options = new OdfDateTimeFieldProjectionOptions { Mode = OdfDateTimeFieldProjectionMode.RefreshDynamic,
            RefreshTimestamp = new DateTimeOffset(2024, 2, 29, 23, 59, 58, TimeSpan.FromHours(14)) };
        foreach (var read in RoundTrips(document)) {
            string[] before = Parts(read);
            var result = read.ToDrawings(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported, false, 1000, default, options);
            foreach (var drawing in result.Value) {
                var texts = drawing.Elements.OfType<OfficeDrawingRichText>().ToArray();
                Assert.Contains(texts, text => text.PlainText == "On 2024-02-29 at 23:59:58");
                Assert.True(texts[0].Paragraphs[0].Runs.Single(run => run.Text == "2024-02-29").Bold);
            }
            Assert.Contains(result.Value[0].Elements.OfType<OfficeDrawingRichText>(), text => text.PlainText == "2010-01-02");
            Assert.Equal(before, Parts(read));
        }
        options.RefreshTimestamp = options.RefreshTimestamp.Value.AddDays(1);
        Assert.Contains("On 2024-03-01", Text(document, options));
    }

    [Theory]
    [InlineData("en-US", false, "February")]
    [InlineData("pl-PL", false, "luty")]
    [InlineData("pl-PL", true, "lutego")]
    public void DeclaredLocaleAndPossessiveMonthAreIndependentOfHostCulture(string culture, bool possessive, string expected) {
        var document = Create(OdfTextFieldKind.Date, out var field); field.DateTimeValueLexical = "2024-02-29";
        var style = document.Styles.FindDataStyle("D")!.Element; style.RemoveNodes();
        style.SetAttributeValue(OdfNamespaces.Number + "rfc-language-tag", culture);
        style.Add(new XElement(OdfNamespaces.Number + "month", new XAttribute(OdfNamespaces.Number + "style", "long"),
            new XAttribute(OdfNamespaces.Number + "textual", "true"), new XAttribute(OdfNamespaces.Number + "possessive-form", possessive ? "true" : "false")));
        CultureInfo previous = CultureInfo.CurrentCulture;
        try { CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("tr-TR"); Assert.Equal(expected, Text(document, Stored())); }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    [Fact]
    public void FractionalSecondsCarryAndElapsedLargestComponentsHaveExplicitSemantics() {
        var document = Create(OdfTextFieldKind.Time, out var field); field.DateTimeValueLexical = "23:59:59.9995";
        XElement style = document.Styles.FindDataStyle("T")!.Element;
        style.Element(OdfNamespaces.Number + "seconds")!.SetAttributeValue(OdfNamespaces.Number + "decimal-places", "3");
        Assert.Equal("00:00:00.000", Text(document, Stored()));
        field.DateTimeValueLexical = "13:07:09";
        style.RemoveNodes(); style.SetAttributeValue(OdfNamespaces.Number + "truncate-on-overflow", "false");
        style.Add(new XElement(OdfNamespaces.Number + "minutes", new XAttribute(OdfNamespaces.Number + "style", "long")),
            new XElement(OdfNamespaces.Number + "text", ":"), new XElement(OdfNamespaces.Number + "seconds"));
        Assert.Equal("787:9", Text(document, Stored()));
        style.SetAttributeValue(OdfNamespaces.Number + "truncate-on-overflow", "true"); Assert.Equal("07:9", Text(document, Stored()));
        style.RemoveNodes(); style.Add(new XElement(OdfNamespaces.Number + "hours"), new XElement(OdfNamespaces.Number + "text", " "), new XElement(OdfNamespaces.Number + "am-pm"));
        Assert.Equal("1 PM", Text(document, Stored()));
        field.DateTimeValueLexical = "00:00:00"; Assert.Equal("12 AM", Text(document, Stored()));
    }

    [Fact]
    public void OmittedMonthInflectionUsesDayContextWhileExplicitNominativeOverridesIt() {
        var document = Create(OdfTextFieldKind.Date, out var field); field.DateTimeValueLexical = "2024-02-29";
        var style = document.Styles.FindDataStyle("D")!.Element; style.RemoveNodes(); style.SetAttributeValue(OdfNamespaces.Number + "rfc-language-tag", "pl-PL");
        var month = new XElement(OdfNamespaces.Number + "month", new XAttribute(OdfNamespaces.Number + "textual", "true"), new XAttribute(OdfNamespaces.Number + "style", "long"));
        style.Add(new XElement(OdfNamespaces.Number + "day"), new XElement(OdfNamespaces.Number + "text", " "), month);
        Assert.Equal("29 lutego", Text(document, Stored()));
        month.SetAttributeValue(OdfNamespaces.Number + "possessive-form", "false"); Assert.Equal("29 luty", Text(document, Stored()));
        month.SetAttributeValue(OdfNamespaces.Number + "possessive-form", null); style.Elements().First().Remove(); style.Elements().First().Remove();
        Assert.Equal("luty", Text(document, Stored()));
    }

    [Fact]
    public void PartLocalStylesShadowCommonDefinitionsAcrossFlatRoundTripsAndPageImport() {
        var document = Create(OdfTextFieldKind.Date, out var local); local.DateTimeValueLexical = "2024-02-29";
        var master = document.Pages[0].MasterShapes.AddTextBox(Bounds(), "").Paragraphs[0].AddField(OdfTextFieldKind.Date, "");
        master.DataStyleName = "D"; master.DateTimeValueLexical = "2024-02-29";
        XElement shadow = new(OdfNamespaces.Number + "date-style", new XAttribute(OdfNamespaces.Style + "name", "D"),
            new XElement(OdfNamespaces.Number + "day"));
        document.GetXml("content.xml").Root!.Element(OdfNamespaces.Office + "automatic-styles")!.Add(shadow); document.MarkPartDirty("content.xml");
        foreach (var read in RoundTrips(document)) {
            var texts = read.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported, false, default, Stored()).Value.Elements.OfType<OfficeDrawingRichText>();
            Assert.Equal(new[] { "2024-02-29", "29" }, texts.Select(text => text.PlainText));
            var target = OdgDocument.Create(); target.AddPage(); target.ImportPage(read, 0);
            Assert.Equal(new[] { "2024-02-29", "29" }, target.Pages[1].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported, false, default, Stored()).Value.Elements.OfType<OfficeDrawingRichText>().Select(text => text.PlainText));
        }
        Assert.True(document.Styles.FindDataStyle("D")!.IsAutomatic); Assert.False(document.Styles.FindDataStyle("D", "styles.xml")!.IsAutomatic);
        Assert.Equal("yyyy-mm-dd", OdsDocument.Create().AddDateStyle("D").ToExcelNumberFormatCode());
    }

    [Theory]
    [InlineData("automatic-order", "true")]
    [InlineData("format-source", "language")]
    [InlineData("rfc-language-tag", "ar-SA")]
    [InlineData("transliteration-style", "long")]
    [InlineData("language", "")]
    public void UnsupportedStylesRetainCachesAndRejectStrictProjectionWithoutMutation(string name, string value) {
        var document = Create(OdfTextFieldKind.Date, out var field); field.DateTimeValueLexical = "2024-02-29";
        document.Styles.FindDataStyle("D")!.Element.SetAttributeValue(OdfNamespaces.Number + name, value);
        AssertUnsupported(document);
    }

    [Fact]
    public void MissingMalformedAndAmbiguousDefinitionsFailClosedAndProjectionBoundsExpansion() {
        var document = Create(OdfTextFieldKind.Date, out var field); AssertUnsupported(document);
        field.DateTimeValueLexical = "2024-02-29"; field.DataStyleName = "Missing"; AssertUnsupported(document);
        field.DataStyleName = "D";
        var definition = document.Styles.FindDataStyle("D")!.Element; definition.Parent!.Add(new XElement(definition)); AssertUnsupported(document);
        definition.Parent.Elements().Last().Remove();
        definition.RemoveNodes(); definition.Add(new XElement(OdfNamespaces.Number + "text", new string('x', 100001))); AssertUnsupported(document);
        definition.RemoveNodes(); definition.Add(new XElement(OdfNamespaces.Number + "year", new XAttribute(OdfNamespaces.Number + "calendar", "buddhist"))); AssertUnsupported(document);
        definition.RemoveNodes(); definition.Add(new XElement(OdfNamespaces.Number + "year"));
        var native = document.Pages[0].Shapes[0].Paragraphs[0].Element.Element(OdfNamespaces.Text + "date")!;
        native.SetAttributeValue(OdfNamespaces.Text + "date-value", "2024-02-30"); AssertUnsupported(document);
        Assert.Equal("2024-02-30", field.DateTimeValueLexical);
        var before = field.ToXml().ToString(); Assert.Throws<ArgumentException>(() => field.DateTimeValueLexical = "2024-02-31");
        Assert.Throws<ArgumentException>(() => field.DateTimeAdjustmentLexical = "PT"); Assert.Equal(before, field.ToXml().ToString());
    }

    [Fact]
    public void OptionsAreValidatedAndCancellationLeavesNativeFieldsUntouched() {
        var document = Create(OdfTextFieldKind.Date, out var field); field.DateTimeValueLexical = "2024-02-29";
        Assert.Throws<ArgumentException>(() => Text(document, new() { Mode = OdfDateTimeFieldProjectionMode.RefreshDynamic }));
        Assert.Throws<ArgumentException>(() => Text(document, new() { Mode = OdfDateTimeFieldProjectionMode.StoredValues, RefreshTimestamp = DateTimeOffset.MinValue }));
        Assert.Throws<ArgumentOutOfRangeException>(() => Text(document, new() { Mode = (OdfDateTimeFieldProjectionMode)99 }));
        Assert.Throws<NotSupportedException>(() => document.Pages[0].Shapes[0].Paragraphs[0].AddField(OdfTextFieldKind.PageCount).DataStyleName = "D");
        string[] before = Parts(document); using var source = new CancellationTokenSource(); source.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.ToDrawings(OdfConversionLossPolicy.ReportOnly, false, 1000, source.Token, Stored()));
        Assert.Equal(before, Parts(document));
    }

    [Fact]
    public void CacheFreeScalarFieldsCanRenderStoredValuesAndExcessPrecisionFailsClosed() {
        var document = Create(OdfTextFieldKind.Time, out var field); field.DisplayText = ""; field.DateTimeValueLexical = "13:07:09.1234567";
        var style = document.Styles.FindDataStyle("T")!.Element;
        style.Element(OdfNamespaces.Number + "seconds")!.SetAttributeValue(OdfNamespaces.Number + "decimal-places", "7");
        Assert.Equal("13:07:09.1234567", Text(document, Stored()));
        Assert.Throws<OdfConversionLossException>(() => Text(document));
        string before = field.ToXml().ToString(); Assert.Throws<ArgumentException>(() => field.DateTimeValueLexical = "13:07:09.12345678"); Assert.Equal(before, field.ToXml().ToString());
        field.DisplayText = "cache";
        style.Element(OdfNamespaces.Number + "seconds")!.SetAttributeValue(OdfNamespaces.Number + "decimal-places", "8");
        AssertUnsupported(document);
    }

    private static OdgDocument Create(OdfTextFieldKind kind, out OdfTextField field) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var p = page.Shapes.AddTextBox(Bounds(), "", "Clock").Paragraphs[0]; p.FontFamily = "Liberation Sans"; p.FontSize = OdfLength.Points(12);
        field = p.AddField(kind, "cache"); field.DataStyleName = kind == OdfTextFieldKind.Date ? "D" : "T";
        if (kind == OdfTextFieldKind.Date) document.Styles.CreateDateStyle("D"); else document.Styles.CreateTimeStyle("T");
        return document;
    }
    private static OdfRect Bounds() => OdfRect.FromCentimeters(1, 1, 15, 3);
    private static OdfDateTimeFieldProjectionOptions Stored() => new() { Mode = OdfDateTimeFieldProjectionMode.StoredValues };
    private static string Text(OdgDocument document, OdfDateTimeFieldProjectionOptions? options = null) => string.Join("\n", document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported, false, default, options).Value.Elements.OfType<OfficeDrawingRichText>().Select(text => text.PlainText));
    private static string[] Parts(OdgDocument document) => new[] { "content.xml", "styles.xml" }.Select(part => document.GetXml(part).ToString()).ToArray();
    private static void AssertUnsupported(OdgDocument document) {
        string[] before = Parts(document);
        var result = document.ToDrawings(OdfConversionLossPolicy.ReportOnly, false, 1000, default, Stored());
        Assert.Equal("cache", Assert.Single(result.Value[0].Elements.OfType<OfficeDrawingRichText>()).PlainText);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature.StartsWith("page:1:", StringComparison.Ordinal) && mapping.Feature.EndsWith(":field-date-time-format", StringComparison.Ordinal) && mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Throws<OdfConversionLossException>(() => Text(document, Stored())); Assert.Equal(before, Parts(document));
    }
    private static OdgDocument[] RoundTrips(OdgDocument document) {
        using var package = new MemoryStream(); document.Save(package); package.Position = 0;
        using var flat = new MemoryStream(); document.SaveFlatXml(flat); flat.Position = 0;
        return new[] { document, OdgDocument.Load(package), OdgDocument.LoadFlatXml(flat) };
    }
}
