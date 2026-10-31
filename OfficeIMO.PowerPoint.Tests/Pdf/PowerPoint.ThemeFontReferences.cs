using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointThemeFontReferenceTests {
    [Theory]
    [InlineData("+mj-lt", "Georgia")]
    [InlineData("+mn-lt", "Courier New")]
    public void PortablePdfReportsTheResolvedThemeFamilyInsteadOfItsReferenceToken(string reference, string family) {
        var theme = OfficeTheme.Modern.WithFonts(new OfficeThemeFonts(
            new OfficeThemeFont("Georgia", new[] { "serif" }),
            new OfficeThemeFont("Courier New", new[] { "monospace" })));
        using var document = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = theme });
        document.AddSlide().AddTextBox("ResolvedFont", OfficeLength.Points(50), OfficeLength.Points(50), OfficeLength.Points(600), OfficeLength.Points(60)).FontName = reference;
        var options = new PowerPointPdfExportOptions { ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic() };
        using var converted = document.ToPdf(options, out var report);
        Assert.Contains(report.Diagnostics, warning => warning.Code == "font-family-substitution" && warning.Message.IndexOf(family, StringComparison.Ordinal) >= 0);
        Assert.DoesNotContain(report.Diagnostics, warning => warning.Message.IndexOf(reference, StringComparison.Ordinal) >= 0);
    }

    [Theory]
    [InlineData("+mj-lt", "HeadingFixture")]
    [InlineData("+mn-lt", "BodyFixture")]
    public void PdfResolvesThemeReferencesBeforeRegisteringFontFamilies(string reference, string family) {
        var theme = OfficeTheme.Modern.WithFonts(new OfficeThemeFonts(
            new OfficeThemeFont("HeadingFixture", new[] { "serif" }),
            new OfficeThemeFont("BodyFixture", new[] { "sans-serif" })));
        using var document = PowerPointDocument.Create(new PowerPointCreateOptions { Theme = theme });
        var box = document.AddSlide().AddTextBox("ResolvedFont", OfficeLength.Points(50), OfficeLength.Points(50), OfficeLength.Points(600), OfficeLength.Points(60));
        box.FontName = reference;
        var options = new PowerPointPdfExportOptions {
            PdfOptions = new PdfCreateOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
                family, OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont()))
        };
        using var converted = document.ToPdf(options, out var report);
        byte[] bytes = converted.ToBytes(new PdfSaveOptions { AllowLoss = true });
        Assert.DoesNotContain(report.Diagnostics, warning => warning.Message.IndexOf(reference, StringComparison.Ordinal) >= 0);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var word = Assert.Single(pdf.GetPage(1).GetWords(), item => item.Text == "ResolvedFont");
        Assert.All(word.Letters, letter => Assert.Contains(family, letter.FontName, StringComparison.Ordinal));
    }
}
