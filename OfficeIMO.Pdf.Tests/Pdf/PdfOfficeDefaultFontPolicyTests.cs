using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfOfficeDefaultFontPolicyTests {
    [Theory]
    [InlineData("Calibri", "Calibri.ttf")]
    [InlineData("Calibri Light", "calibril.ttf")]
    [InlineData("Cambria", "Cambria.ttf")]
    [InlineData("Aptos", "Aptos.ttf")]
    public void InstalledOfficeFontsAreAvailableToBothPdfEmbeddingAndDrawing(string family, string fileName) {
        string path = Path.Combine("/Applications", "Microsoft Word.app", "Contents", "Resources", "DFonts", fileName);
        if (!File.Exists(path)) return;
        Assert.True(PdfEmbeddedFontFamily.TryFromSystem(family, out _), family);
        Assert.NotNull(OfficeTrueTypeFont.TryLoadFontFamily(family));
    }

    [Theory]
    [InlineData("Calibri", "Carlito", true)]
    [InlineData("Cambria", "Caladea", true)]
    [InlineData("Arial", "Liberation Sans", true)]
    [InlineData("Times New Roman", "Liberation Serif", true)]
    [InlineData("Courier New", "Liberation Mono", true)]
    [InlineData("Aptos", "Inter", false)]
    public void MissingOfficeFamilies_UseInstalledSubstitutesWithoutHidingTheirIdentity(string source, string firstSubstitute, bool compatible) {
        string[] substitutes = OfficeDocumentFontSubstitutions.Enumerate(source).ToArray();
        Assert.Equal(firstSubstitute, substitutes[0]);
        Assert.Equal(compatible, OfficeDocumentFontSubstitutions.IsIntendedCompatibleSubstitute(source, firstSubstitute));

        // Exercise the real host boundary without installing fonts or changing global state.
        // Exact selection must win on Office hosts; missing families use the first available alternative.
        string? expected = new[] { source }.Concat(substitutes)
            .FirstOrDefault(candidate => PdfEmbeddedFontFamily.TryFromSystem(candidate, out _));
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions().ReportDiagnosticsTo(report);
        Assert.Equal(expected != null, options.TryRegisterNamedOfficeFontFamily(source, out string? registered));
        Assert.Equal(expected, registered);
        if (expected == null) return;
        Assert.Contains(options.NamedFontFamilies.Keys, name => name == expected);
        if (expected == source) {
            Assert.Empty(options.FontFamilySubstitutions);
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "font-family-substitution");
        } else {
            PdfFontFamilySubstitution substitution = Assert.Single(options.FontFamilySubstitutions);
            Assert.Equal(source, substitution.SourceFontFamily);
            Assert.Equal(expected, substitution.TargetFontFamily);
            PdfConversionWarning warning = Assert.Single(report.Warnings, warning => warning.Code == "font-family-substitution");
            Assert.Equal(source, warning.Details["fontFamily"]);
            Assert.Equal(expected, warning.Details["resolvedFontFamily"]);
            Assert.Equal("FontUnavailable", warning.Details["substitutionReason"]);
            Assert.Equal("NotPerformed", warning.Details["fontMetricsVerification"]);
            Assert.Equal(substitution.Impact == PdfFontFamilySubstitutionImpact.Compatible
                ? "IntendedCompatibility" : "VisualSimilarity", warning.Details["compatibilityBasis"]);
            Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        }
    }

    [Fact]
    public void BalancedDefaults_AllowInstalledDocumentFontsWithoutGrantingFileOrNetworkAccess() {
        foreach (PdfResourcePolicy policy in new[] {
            PdfResourcePolicy.CreateDefault(), new WordPdfExportOptions().ResourcePolicy,
            new PowerPointPdfExportOptions().ResourcePolicy
        }) {
            Assert.True(policy.AllowSystemFontEmbedding);
            Assert.True(policy.AllowDocumentFontEmbedding);
            Assert.False(policy.AllowLocalFileAccess);
            Assert.False(policy.AllowRemoteResourceResolution);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordAndPowerPoint_DefaultsHonorInstalledArial_WhilePortableNeverLoadsHostFonts(bool portable) {
        PdfResourcePolicy policy = portable ? PdfResourcePolicy.CreatePortableDeterministic() : PdfResourcePolicy.CreateDefault();
        using WordDocument word = WordDocument.Create();
        word.Settings.FontFamily = "Arial";
        word.AddParagraph("Arial proof").FontFamily = "Arial";
        PdfDocumentConversionResult wordResult = word.ConvertToPdfWithEvidence(new WordPdfExportOptions { ResourcePolicy = policy });
        AssertFontSelection(wordResult.ToBytes(), wordResult.Report, portable);

        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        PowerPointTextBox text = presentation.AddSlide().AddTextBoxPoints("Arial proof", 20, 20, 300, 60);
        text.FontName = "Arial";
        text.FontSize = 16;
        PdfDocumentConversionResult slidesResult = presentation.ConvertToPdfWithEvidence(new PowerPointPdfExportOptions { ResourcePolicy = policy });
        AssertFontSelection(slidesResult.ToBytes(), slidesResult.Report, portable);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void EmbeddedTrueTypeCost_MatchesEmittedStream_AndUsesOnlyTheUsedGlyphSubset(bool named, bool compress) {
        byte[] font = File.ReadAllBytes(Assert.IsType<string>(PdfComplianceTestFonts.FindBundledTrueTypeFont()));
        var family = new PdfEmbeddedFontFamily("Cost Proof", font);
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { CompressEmbeddedFonts = compress }.ReportDiagnosticsTo(report);
        if (named) options.RegisterNamedFontFamily(family);
        else options.RegisterFontFamily(PdfStandardFont.Helvetica, family);
        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => {
            if (named) paragraph.FontFamily("Cost Proof");
            paragraph.Text("Subset proof");
        }).ToBytes();

        PdfConversionWarning cost = Assert.Single(report.Warnings, warning => warning.Code == "font-embedding-cost");
        Assert.Equal(PdfConversionWarningSeverity.Information, cost.Severity);
        Assert.Equal(OfficeConversionLossKind.None, cost.LossKind);
        Assert.Equal(font.Length, int.Parse(cost.Details["originalFontBytes"], CultureInfo.InvariantCulture));
        Assert.Equal("subset-truetype", cost.Details["embeddingMode"]);
        int subsetBytes = int.Parse(cost.Details["embeddedFontFileBytes"], CultureInfo.InvariantCulture);
        Assert.InRange(subsetBytes, 1, font.Length - 1);
        AssertFontStreamLength(bytes, cost);
        using var reopened = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal("Subset proof", reopened.GetPage(1).Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmbeddedCffCost_MatchesActualCompressedOrUncompressedPayload(bool compress) {
        byte[] font = File.ReadAllBytes(Assert.IsType<string>(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()));
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { CompressEmbeddedFonts = compress }.ReportDiagnosticsTo(report);
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("CFF Cost Proof", font));
        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.FontFamily("CFF Cost Proof").Text("CFF proof")).ToBytes();
        PdfConversionWarning cost = Assert.Single(report.Warnings, warning => warning.Code == "font-embedding-cost");
        Assert.Equal("OpenType/CFF", cost.Details["format"]);
        Assert.Equal(font.Length, int.Parse(cost.Details["originalFontBytes"], CultureInfo.InvariantCulture));
        AssertFontStreamLength(bytes, cost);
        using var reopened = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Equal("CFF proof", reopened.GetPage(1).Text);
    }

    private static void AssertFontSelection(byte[] bytes, PdfConversionReport report, bool portable) {
        string raw = Encoding.ASCII.GetString(bytes);
        if (portable) {
            Assert.DoesNotContain("/FontFile", raw, StringComparison.Ordinal);
            Assert.DoesNotContain(report.Warnings, warning => warning.Code == "font-embedding-cost");
        } else {
            string? expected = new[] { "Arial", "Liberation Sans" }
                .FirstOrDefault(candidate => PdfEmbeddedFontFamily.TryFromSystem(candidate, out _));
            if (expected != null) {
                Assert.Contains(report.Warnings, warning => warning.Code == "font-embedding-cost" &&
                    warning.Details["fontName"].Replace(" ", string.Empty).StartsWith(expected.Replace(" ", string.Empty), StringComparison.Ordinal));
                using var reopened = UglyToad.PdfPig.PdfDocument.Open(bytes);
                Assert.All(reopened.GetPage(1).Letters, letter =>
                    Assert.Contains(expected.Replace(" ", string.Empty), letter.FontName.Replace(" ", string.Empty), StringComparison.Ordinal));
            }
        }
    }

    private static void AssertFontStreamLength(byte[] bytes, PdfConversionWarning cost) {
        string raw = Encoding.ASCII.GetString(bytes);
        Match stream = Regex.Match(raw, @"(?:^|\n)" + cost.Details["fontFileObjectId"] + @" 0 obj\s*<<(?<dictionary>.*?)>>\s*stream", RegexOptions.Singleline);
        Assert.True(stream.Success);
        string dictionary = stream.Groups["dictionary"].Value;
        Assert.Equal(cost.Details["embeddedStreamBytes"], Regex.Match(dictionary, @"/Length\s+(\d+)").Groups[1].Value);
        bool compressed = cost.Details["compressed"] == "true";
        Assert.Equal(compressed, dictionary.Contains("/FlateDecode", StringComparison.Ordinal));
        if (!compressed) Assert.Equal(cost.Details["embeddedFontFileBytes"], cost.Details["embeddedStreamBytes"]);
    }
}
