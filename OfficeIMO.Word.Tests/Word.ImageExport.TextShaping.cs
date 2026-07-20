using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class WordImageExportTests {
    [Fact]
    public void WordRasterExport_ReachesTheSharedTextShapingProvider() {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.AddParagraph("A").SetFontFamily(ManagedTextShapingTestAssets.FamilyName);
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new WordImageExportOptions {
            TextShapingProvider = provider,
            TextShapingLanguage = "ar-SA"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));

        OfficeImageExportResult result = document.ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.Contains(provider.Requests, request =>
            request.Text == "A" &&
            request.Language == "ar-SA");
        Assert.DoesNotContain(result.Diagnostics, diagnostic =>
            diagnostic.Code == OfficeImageExportDiagnosticCodes.TextShapingFallback);
    }

    [Fact]
    public void WordRasterExport_StrictLossPolicyRejectsManagedComplexTextFallback() {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.AddParagraph("اب").SetFontFamily(ManagedTextShapingTestAssets.FamilyName);
        var options = new WordImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoLoss = true }
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(
                0x0020,
                0x0627,
                0x0628,
                0xFE8D,
                0xFE8F));

        OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(
            () => document.ExportImage(OfficeImageExportFormat.Png, options));

        OfficeImageExportDiagnostic diagnostic = Assert.Single(
            exception.Diagnostics,
            item => item.Code == OfficeImageExportDiagnosticCodes.TextShapingFallback);
        Assert.Equal(OfficeImageExportLossKind.Approximation, diagnostic.LossKind);
        Assert.Equal("Word document", diagnostic.Source);
    }

    [Fact]
    public void WordRasterExport_StrictLossPolicyRejectsUnshapedIndicText() {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.AddParagraph("कि").SetFontFamily(ManagedTextShapingTestAssets.FamilyName);
        var options = new WordImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoLoss = true }
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(0x0915, 0x093F));

        OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(
            () => document.ExportImage(OfficeImageExportFormat.Png, options));

        OfficeImageExportDiagnostic diagnostic = Assert.Single(
            exception.Diagnostics,
            item => item.Code == OfficeImageExportDiagnosticCodes.TextShapingFallback);
        Assert.Equal(OfficeImageExportLossKind.Approximation, diagnostic.LossKind);
        Assert.Contains("cannot provide complete", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void WordRasterExport_StrictLossPolicyRejectsUnshapedBugineseText() {
        using var stream = new MemoryStream();
        using WordDocument document = WordDocument.Create(stream);
        document.AddParagraph("\u1A00\u1A17").SetFontFamily(ManagedTextShapingTestAssets.FamilyName);
        var options = new WordImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoLoss = true }
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont(0x1A00, 0x1A17));

        OfficeImageExportPolicyException exception = Assert.Throws<OfficeImageExportPolicyException>(
            () => document.ExportImage(OfficeImageExportFormat.Png, options));

        OfficeImageExportDiagnostic diagnostic = Assert.Single(
            exception.Diagnostics,
            item => item.Code == OfficeImageExportDiagnosticCodes.TextShapingFallback);
        Assert.Equal(OfficeImageExportLossKind.Approximation, diagnostic.LossKind);
        Assert.Contains("cannot provide complete", diagnostic.Message, StringComparison.Ordinal);
    }
}
