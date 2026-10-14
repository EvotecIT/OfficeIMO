using System;
using System.IO;
using System.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class DrawPdfDateTimeFieldTests {
    [Fact]
    public void ExplicitRefreshReachesEveryPdfPageAndOptionsCloneOwnsIndependentSettings() {
        var source = Create();
        var options = new OdgToPdfOptions { DateTimeFields = new() {
            Mode = OdfDateTimeFieldProjectionMode.RefreshDynamic,
            RefreshTimestamp = new DateTimeOffset(2024, 2, 29, 23, 59, 58, TimeSpan.FromHours(14)) } };
        var copy = options.Clone(); options.DateTimeFields.RefreshTimestamp = options.DateTimeFields.RefreshTimestamp.Value.AddDays(1);
        var before = source.ToBytes(); var result = source.ToPdfDocumentResult(copy); var read = PdfReadDocument.Open(result.ToBytes());
        Assert.Equal(2, read.Pages.Count);
        foreach (var page in read.Pages) { Assert.Contains("2024-02-29", page.ExtractText()); Assert.DoesNotContain("stale", page.ExtractText()); }
        Assert.Equal(before, source.ToBytes());
        Assert.Contains(result.FidelityDiagnostics, diagnostic => diagnostic.Location!.StartsWith("page:2:", StringComparison.Ordinal) && diagnostic.Location.EndsWith(":field-date-time-value", StringComparison.Ordinal));
        Assert.Contains("stale", PdfReadDocument.Open(source.ToPdfBytes()).ExtractText());
    }

    [Fact]
    public void UnsupportedFieldFormattingLeavesStrictPdfDestinationUntouched() {
        var source = Create(); var field = source.Pages[0].MasterShapes[0].Paragraphs[0].Fields.Single(); field.DataStyleName = "Missing";
        var options = new OdgToPdfOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported,
            DateTimeFields = new() { Mode = OdfDateTimeFieldProjectionMode.RefreshDynamic, RefreshTimestamp = DateTimeOffset.MinValue } };
        byte[] original = { 1, 2, 3 }; using var output = new MemoryStream(original.ToArray(), writable: true);
        var result = source.SaveAsPdfResult(output, options);
        Assert.False(result.Succeeded); Assert.Equal(original, output.ToArray()); Assert.True(output.CanWrite);
        Assert.Contains(result.FidelityDiagnostics, diagnostic => diagnostic.Location!.EndsWith(":field-date-time-format", StringComparison.Ordinal));
        options.DateTimeFields.RefreshTimestamp = null;
        Assert.Throws<ArgumentException>(() => source.ToPdfDocumentResult(options));
    }

    private static OdgDocument Create() {
        var source = OdgDocument.Create(); var first = source.AddPage(); source.AddPage().MasterPageName = first.MasterPageName;
        source.Styles.CreateDateStyle("ISO"); var p = first.MasterShapes.AddTextBox(OdfRect.FromCentimeters(1, 1, 15, 2), "").Paragraphs[0];
        p.FontFamily = "Liberation Sans"; p.FontSize = OdfLength.Points(12);
        p.AddField(OdfTextFieldKind.Date, "stale").DataStyleName = "ISO";
        return source;
    }
}
