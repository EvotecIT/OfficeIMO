using System;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void ViewerPreferences_CanBeEmittedClonedInspectedAndPreservedOnExtraction() {
        PdfViewerPreferencesOptions preferencesOptions = new PdfViewerPreferencesOptions {
            DisplayDocTitle = true,
            HideToolbar = true,
            FitWindow = false,
            NonFullScreenPageMode = PdfNonFullScreenPageMode.UseOutlines,
            PrintScaling = PdfPrintScaling.None,
            Duplex = PdfDuplexMode.DuplexFlipLongEdge,
            ViewArea = PdfPageBoundaryBox.CropBox,
            PrintArea = PdfPageBoundaryBox.TrimBox,
            PickTrayByPdfSize = true,
            NumCopies = 2
        }.AddPrintPageRange(1, 1);
        PdfPrintPageRange mutableRange = new PdfPrintPageRange(2, 3);
        preferencesOptions.AddPrintPageRange(mutableRange);

        var options = new PdfOptions {
            ViewerPreferences = preferencesOptions
        };

        byte[] bytes = PdfDocument.Create(options)
            .Meta(title: "Viewer preference proof")
            .ViewerPreferences(preferences => {
                preferences.CenterWindow = true;
                preferences.HideMenubar = false;
                preferences.Direction = PdfViewerDirection.RightToLeft;
                preferences.ViewClip = PdfPageBoundaryBox.BleedBox;
                preferences.PrintClip = PdfPageBoundaryBox.ArtBox;
            })
            .Paragraph(p => p.Text("Viewer preferences proof."))
            .PageBreak()
            .Paragraph(p => p.Text("Second print page."))
            .PageBreak()
            .Paragraph(p => p.Text("Third print page."))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(bytes);
        PdfViewerPreferences viewerPreferences = Assert.IsType<PdfViewerPreferences>(info.ViewerPreferences);
        PdfViewerPreferencesOptions clone = options.Clone().ViewerPreferences!;

        Assert.Contains("/ViewerPreferences ", raw, StringComparison.Ordinal);
        Assert.Contains("/DisplayDocTitle true", raw, StringComparison.Ordinal);
        Assert.Contains("/HideToolbar true", raw, StringComparison.Ordinal);
        Assert.Contains("/FitWindow false", raw, StringComparison.Ordinal);
        Assert.Contains("/CenterWindow true", raw, StringComparison.Ordinal);
        Assert.Contains("/HideMenubar false", raw, StringComparison.Ordinal);
        Assert.Contains("/PickTrayByPDFSize true", raw, StringComparison.Ordinal);
        Assert.Contains("/NonFullScreenPageMode /UseOutlines", raw, StringComparison.Ordinal);
        Assert.Contains("/Direction /R2L", raw, StringComparison.Ordinal);
        Assert.Contains("/PrintScaling /None", raw, StringComparison.Ordinal);
        Assert.Contains("/Duplex /DuplexFlipLongEdge", raw, StringComparison.Ordinal);
        Assert.Contains("/ViewArea /CropBox", raw, StringComparison.Ordinal);
        Assert.Contains("/ViewClip /BleedBox", raw, StringComparison.Ordinal);
        Assert.Contains("/PrintArea /TrimBox", raw, StringComparison.Ordinal);
        Assert.Contains("/PrintClip /ArtBox", raw, StringComparison.Ordinal);
        Assert.Contains("/NumCopies 2", raw, StringComparison.Ordinal);
        Assert.Contains("/PrintPageRange [1 1 2 3]", raw, StringComparison.Ordinal);
        Assert.True(info.HasViewerPreferences);
        Assert.True(info.HasReadableViewerPreferences);
        Assert.True(preflight.Probe.HasViewerPreferences);
        Assert.True(preflight.CanRewrite);
        Assert.True(viewerPreferences.GetBoolean("DisplayDocTitle"));
        Assert.True(viewerPreferences.GetBoolean("HideToolbar"));
        Assert.False(viewerPreferences.GetBoolean("FitWindow"));
        Assert.True(viewerPreferences.GetBoolean("CenterWindow"));
        Assert.False(viewerPreferences.GetBoolean("HideMenubar"));
        Assert.True(viewerPreferences.GetBoolean("PickTrayByPDFSize"));
        Assert.Equal("UseOutlines", viewerPreferences.GetValue("NonFullScreenPageMode"));
        Assert.Equal("R2L", viewerPreferences.GetValue("Direction"));
        Assert.Equal("None", viewerPreferences.GetValue("PrintScaling"));
        Assert.Equal("DuplexFlipLongEdge", viewerPreferences.GetValue("Duplex"));
        Assert.Equal("CropBox", viewerPreferences.GetValue("ViewArea"));
        Assert.Equal("BleedBox", viewerPreferences.GetValue("ViewClip"));
        Assert.Equal("TrimBox", viewerPreferences.GetValue("PrintArea"));
        Assert.Equal("ArtBox", viewerPreferences.GetValue("PrintClip"));
        Assert.Equal("2", viewerPreferences.GetValue("NumCopies"));
        Assert.Equal("[1 1 2 3]", viewerPreferences.GetValue("PrintPageRange"));
        Assert.True(clone.DisplayDocTitle);
        Assert.True(clone.HideToolbar);
        Assert.False(clone.FitWindow);
        Assert.True(clone.PickTrayByPdfSize);
        Assert.Equal(PdfNonFullScreenPageMode.UseOutlines, clone.NonFullScreenPageMode);
        Assert.Equal(PdfPrintScaling.None, clone.PrintScaling);
        Assert.Equal(PdfDuplexMode.DuplexFlipLongEdge, clone.Duplex);
        Assert.Equal(PdfPageBoundaryBox.CropBox, clone.ViewArea);
        Assert.Equal(PdfPageBoundaryBox.TrimBox, clone.PrintArea);
        Assert.Equal(2, clone.NumCopies);
        Assert.Equal(2, clone.PrintPageRanges.Count);
        Assert.Null(clone.CenterWindow);
        Assert.Null(clone.Direction);
        Assert.Null(clone.ViewClip);
        Assert.Null(clone.PrintClip);

        byte[] extracted = PdfPageExtractor.ExtractPages(bytes, 1, 2);
        PdfViewerPreferences extractedPreferences = PdfInspector.Inspect(extracted).ViewerPreferences!;
        Assert.True(extractedPreferences.GetBoolean("DisplayDocTitle"));
        Assert.Equal("R2L", extractedPreferences.GetValue("Direction"));
        Assert.Equal("ArtBox", extractedPreferences.GetValue("PrintClip"));
        Assert.Equal("[1 1 2 3]", extractedPreferences.GetValue("PrintPageRange"));

        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfViewerPreferencesOptions { NonFullScreenPageMode = (PdfNonFullScreenPageMode)99 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfViewerPreferencesOptions { Direction = (PdfViewerDirection)99 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfViewerPreferencesOptions { PrintScaling = (PdfPrintScaling)99 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfViewerPreferencesOptions { Duplex = (PdfDuplexMode)99 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfViewerPreferencesOptions { ViewArea = (PdfPageBoundaryBox)99 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfViewerPreferencesOptions { ViewClip = (PdfPageBoundaryBox)99 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfViewerPreferencesOptions { PrintArea = (PdfPageBoundaryBox)99 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfViewerPreferencesOptions { PrintClip = (PdfPageBoundaryBox)99 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfViewerPreferencesOptions { NumCopies = 0 });
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPrintPageRange(0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new PdfPrintPageRange(3, 2));
        Assert.Throws<ArgumentNullException>(() => new PdfViewerPreferencesOptions().AddPrintPageRange(null!));
    }
}
