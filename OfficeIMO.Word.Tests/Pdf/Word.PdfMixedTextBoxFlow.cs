using DocumentFormat.OpenXml;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig;
using W = DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body", false)]
    [InlineData("body", true)]
    [InlineData("header", false)]
    [InlineData("header", true)]
    [InlineData("footer", false)]
    [InlineData("footer", true)]
    public void NativeFlowPreservesTextAroundPeerAndNestedBoxes(string scope, bool nested) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordHeaderFooter? story = scope == "body" ? null : scope == "header"
            ? document.HeaderDefaultOrCreate : document.FooterDefaultOrCreate;
        WordParagraph paragraph = story == null ? document.AddParagraph("PREFIX") : story.AddParagraph("PREFIX");
        WordParagraph first = paragraph.AddTextBox("FIRSTBOX", WordImageTextWrapping.Square).Paragraphs.Single();
        if (nested) {
            first.AddTextBox("INNERBOX", WordImageTextWrapping.Square);
            first.AddText("AFTERINNER");
        }
        WordParagraph suffix = paragraph.AddText("SUFFIX");
        suffix.AddTextBox("SECONDBOX", WordImageTextWrapping.Square);
        paragraph.AddText("ENDTEXT");
        if (story != null) {
            // SECTIONPAGES deliberately selects full running-story flow even
            // when the mixed story cannot use plain numbered-story admission.
            story.AddParagraph()._paragraph.Append(new W.SimpleField(new W.Run(new W.Text("888"))) { Instruction = " SECTIONPAGES " });
            paragraph._paragraph.Append(new W.SimpleField(new W.Run(new W.Text("999"))) { Instruction = " PAGE " });
            document.AddParagraph("BODY");
        }
        string before = paragraph._paragraph.OuterXml;
        Assert.Empty(document.ValidateDocument());
        var result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        string visible = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        string[] tokens = nested
            ? new[] { "PREFIX", "FIRSTBOX", "INNERBOX", "AFTERINNER", "SUFFIX", "SECONDBOX", "ENDTEXT" }
            : new[] { "PREFIX", "FIRSTBOX", "SUFFIX", "SECONDBOX", "ENDTEXT" };
        int previous = -1;
        foreach (string token in tokens) {
            int position = visible.IndexOf(token, StringComparison.Ordinal);
            Assert.True(position > previous, $"Missing or reordered {token}: {visible}");
            Assert.Equal(1, visible.Split(new[] { token }, StringSplitOptions.None).Length - 1);
            previous = position;
        }
        if (story != null) {
            Assert.DoesNotContain("888", visible); Assert.DoesNotContain("999", visible);
            Assert.Equal(2, visible.Count(character => character == '1'));
        }
        Assert.Equal(before, paragraph._paragraph.OuterXml);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("header")]
    [InlineData("footer")]
    public void NativeMixedTextBoxFlowRetainsSiblingInlinePictures(string scope) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordHeaderFooter? story = scope == "body" ? null : scope == "header"
            ? document.HeaderDefaultOrCreate : document.FooterDefaultOrCreate;
        WordParagraph paragraph = story == null ? document.AddParagraph("PREFIX") : story.AddParagraph("PREFIX");
        paragraph.AddTextBox("BOX", WordImageTextWrapping.Square);
        using var image = new MemoryStream(OfficeIMO.Tests.Pdf.PdfPngTestImages.CreateRgbPng(2, 1));
        paragraph.AddImage(image, "inline.png", OfficeIMO.OfficeLength.Pixels(16D), OfficeIMO.OfficeLength.Pixels(32D));
        paragraph.AddText("SUFFIX");
        if (story != null) {
            story.AddParagraph().AddField(WordFieldType.SectionPages);
            document.AddParagraph("BODY");
        }
        Assert.Empty(document.ValidateDocument());
        using var pdf = PdfDocument.Open(document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Single(pdf.GetPages().SelectMany(page => page.GetImages()));
        Assert.Contains("PREFIX", pdf.GetPage(1).Text); Assert.Contains("SUFFIX", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("nested")]
    [InlineData("header")]
    [InlineData("footer")]
    public void NativeCellTextBoxesReportFlattenedLayoutBeforeDeferredPaint(string scope) {
        using WordDocument document = CreateCellTextBoxDiagnosticControl(scope, includeBox: true, hidden: false,
            out WordParagraph cell);
        string sourceXml = cell._paragraph.OuterXml;
        var result = document.ConvertToPdfWithEvidence(CellTextBoxDiagnosticOptions());
        var warning = Assert.Single(result.Report.Warnings,
            item => item.Code == "native-mixed-text-box-layout-approximated");
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.True(result.HasLoss);
        if (scope == "body")
            Assert.Equal(warning.Code, Assert.Single(result.Report.Warnings,
                item => item.LossKind != OfficeIMO.OfficeConversionLossKind.None).Code);
        // Running-story paint is deferred until ToBytes; the report must be truthful already.
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        using var pdf = PdfDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        int previous = -1;
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX" }) {
            int position = visible.IndexOf(token, StringComparison.Ordinal);
            Assert.True(position > previous, $"Missing or reordered {token}: {visible}");
            Assert.Equal(1, visible.Split(new[] { token }, StringSplitOptions.None).Length - 1);
            previous = position;
        }
        Assert.DoesNotContain("FALLBACK", visible, StringComparison.Ordinal);
        Assert.Single(result.Report.Warnings, item => item.Code == warning.Code);
        Assert.Equal(sourceXml, cell._paragraph.OuterXml);
    }

    [Theory]
    [InlineData("body", false)]
    [InlineData("body", true)]
    [InlineData("header", true)]
    public void NativeCellTextBoxWarningsIgnorePlainCellsAndHiddenHostingRuns(string scope, bool hiddenBox) {
        using WordDocument document = CreateCellTextBoxDiagnosticControl(scope,
            includeBox: hiddenBox, hidden: hiddenBox, out WordParagraph cell);
        string sourceXml = cell._paragraph.OuterXml;
        var result = document.ConvertToPdfWithEvidence(CellTextBoxDiagnosticOptions());
        Assert.DoesNotContain(result.Report.Warnings, item => item.Code == "native-mixed-text-box-layout-approximated");
        Assert.False(result.HasLoss);
        result.RequireNoLoss();
        using var pdf = PdfDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("PREFIXSUFFIX", visible, StringComparison.Ordinal);
        Assert.DoesNotContain("BOX", visible, StringComparison.Ordinal);
        Assert.DoesNotContain("FALLBACK", visible, StringComparison.Ordinal);
        Assert.Equal(sourceXml, cell._paragraph.OuterXml);
    }

    private static WordDocument CreateCellTextBoxDiagnosticControl(string scope, bool includeBox, bool hidden,
        out WordParagraph cell) {
        WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        document.AddParagraph("BODY").FontFamily = ManagedTextShapingTestAssets.FamilyName;
        WordTable table = scope switch {
            "header" => document.HeaderDefaultOrCreate.AddTable(1, 1, WordTableStyle.TableNormal),
            "footer" => document.FooterDefaultOrCreate.AddTable(1, 1, WordTableStyle.TableNormal),
            _ => document.AddTable(1, 1, WordTableStyle.TableNormal)
        };
        if (scope == "nested")
            table = table.Rows[0].Cells[0].AddTable(1, 1, WordTableStyle.TableNormal, removePrecedingParagraph: true);
        cell = table.Rows[0].Cells[0].Paragraphs[0];
        cell.Text = "PREFIX";
        cell.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        if (includeBox) {
            WordTextBox box = cell.AddTextBox("BOX", WordImageTextWrapping.Square);
            box.Paragraphs.Single().FontFamily = ManagedTextShapingTestAssets.FamilyName;
            foreach (W.Break lineBreak in box.Content!.Descendants<W.Break>().ToArray()) lineBreak.Remove();
            AlternateContent alternate = box.Anchor!.Ancestors<AlternateContent>().Single();
            alternate.Append(new AlternateContentFallback(new W.Picture(new V.Shape(new V.TextBox(
                new W.TextBoxContent(new W.Paragraph(new W.Run(new W.Text("FALLBACK")))))) {
                Id = "CellBoxFallback", Type = "#_x0000_t202", Style = "width:96pt;height:36pt"
            })));
            alternate.Remove();
            var host = new W.Run(alternate);
            if (hidden) host.RunProperties = new W.RunProperties(new W.Vanish());
            cell._paragraph.Append(host);
        }
        cell.AddText("SUFFIX").FontFamily = ManagedTextShapingTestAssets.FamilyName;
        return document;
    }

    private static WordPdfExportOptions CellTextBoxDiagnosticOptions() {
        var pdf = new PdfCore.PdfCreateOptions();
        pdf.RegisterNamedFontFamily(new PdfCore.PdfEmbeddedFontFamily(ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(Enumerable.Range(32, 95).ToArray())));
        return new WordPdfExportOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(), PdfOptions = pdf
        };
    }
}
