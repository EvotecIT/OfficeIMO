using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.TestAssets;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("document")]
    [InlineData("named")]
    [InlineData("standard")]
    [InlineData("substitution")]
    public void HtmlDefaultFont_ExplicitTargetBytesRemainAvailableToRetainedDrawing(string registration) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithVerticalMetrics('A', 1100, -200, 1050);
        var family = new PdfCore.PdfEmbeddedFontFamily("Default Fixture", data);
        var options = new HtmlToPdfOptions {
            Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        if (registration == "document") options.FontFamily = family;
        else if (registration == "named")
            options.PdfOptions.RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily("Arial", data));
        else if (registration == "standard")
            options.PdfOptions.RegisterFontFamily(PdfCore.PdfStandardFont.Helvetica, family);
        else {
            options.PdfOptions.RegisterNamedFontFamily(family);
            options.PdfOptions.RegisterFontFamilySubstitution("Arial", family.FamilyName);
        }
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<p style='margin:0;font:20px Arial'>AA</p>");

        HtmlRenderDocument raw = HtmlRenderEngine.Render(source, options);
        HtmlPdfRenderRequestResult converted = source.RenderToPdfResult(
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, options));
        HtmlRenderText rawText = Assert.Single(raw.Pages[0].Visuals.OfType<HtmlRenderText>());
        HtmlRenderText pdfText = Assert.Single(converted.RenderResult!.Document.Pages[0].Visuals.OfType<HtmlRenderText>());

        Assert.Contains(raw.Fonts.Faces, face => face.Data.SequenceEqual(data));
        Assert.Contains(raw.Pages[0].CreateDrawing().Fonts.Faces, face => face.Data.SequenceEqual(data));
        Assert.Contains(converted.RenderResult.Document.Fonts.Faces, face => face.Data.SequenceEqual(data));
        Assert.Equal(rawText.TextAdvanceWidth, pdfText.TextAdvanceWidth);
        Assert.Equal(rawText.Y, pdfText.Y, 6);
        Assert.Equal("AA", PdfCore.PdfReadDocument.Open(converted.ToBytes()).ExtractText().Trim());
    }

    [Fact]
    public void HtmlDefaultFont_SuppliedCssFaceKeepsOwnershipOverTargetDefault() {
        byte[] supplied = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A');
        byte[] target = ManagedTextShapingTestAssets.CreateFontWithVerticalMetrics('A', 1100, -200, 1050);
        var options = new HtmlToPdfOptions {
            FontFamily = new PdfCore.PdfEmbeddedFontFamily("Other Default", target),
            Margins = HtmlRenderMargins.All(0D), AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.Fonts.Add("Arial", supplied);

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse("<p style='margin:0;font:20px Arial'>A</p>"), options);

        OfficeFontFace face = Assert.Single(rendered.Pages[0].CreateDrawing().Fonts.Faces);
        Assert.Equal(supplied, face.Data);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == "InstalledFontResolved");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FontMetricsEstimated);
    }

    [Fact]
    public void HtmlDefaultFont_ExplicitRetentionHonoursDecodedResourceBudget() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A');
        var options = new HtmlToPdfOptions {
            FontFamily = new PdfCore.PdfEmbeddedFontFamily("Default Fixture", data),
            AllowSystemFontFallback = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            MaxResourceBytes = data.Length, MaxTotalResourceBytes = data.Length
        };

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(
            HtmlConversionDocument.Parse("<p style='margin:0;font:20px Arial'>A</p>"), options);

        Assert.Empty(rendered.Fonts.Faces);
        Assert.Contains(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded
                && diagnostic.LossKind != OfficeConversionLossKind.None);
        Assert.Throws<HtmlConversionException>(() => rendered.RequireNoLoss());
    }

    [Fact]
    public void HtmlDefaultFont_RegularOnlyStandardRegistrationKeepsSynthesizedItalicPaint() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A');
        var options = new HtmlToPdfOptions {
            AllowSystemFontFallback = false, Margins = HtmlRenderMargins.All(0D),
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };
        options.PdfOptions.RegisterFontFamily(PdfCore.PdfStandardFont.Helvetica,
            new PdfCore.PdfEmbeddedFontFamily("Default Fixture", data));
        HtmlRenderOptions referenceOptions = options.Clone();
        referenceOptions.Fonts.Add("Arial", data);
        HtmlConversionDocument source = HtmlConversionDocument.Parse("<p style='margin:0;font:italic 40px Arial'>AA</p>");

        HtmlRenderDocument actual = HtmlRenderEngine.Render(source, options);
        HtmlRenderDocument reference = HtmlRenderEngine.Render(source, referenceOptions);

        Assert.Equal(reference.Pages[0].CreateDrawing().ExportImage(OfficeImageExportFormat.Png).Bytes,
            actual.Pages[0].CreateDrawing().ExportImage(OfficeImageExportFormat.Png).Bytes);
    }
}
