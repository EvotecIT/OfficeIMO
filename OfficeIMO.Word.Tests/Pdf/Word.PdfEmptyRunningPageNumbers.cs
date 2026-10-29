using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SaveAsPdf_ImplicitFooterPreservesSiblingPageAndSectionFields(int emptyVariant) {
        using WordDocument document = CreateEmptyRunningStoryDocument("absent", footer: true);
        WordHeaderFooter[] stories = { document.FooterFirstOrCreate, document.FooterEvenOrCreate, document.FooterDefaultOrCreate };
        foreach (WordHeaderFooter story in stories)
            foreach (var child in story.ChildElements.ToArray()) child.Remove();
        for (int index = 0; index < 3; index++) {
            if (index != emptyVariant) {
                WordParagraph field = stories[index].AddParagraph("FIELD" + index + ":");
                field.AddField(WordFieldType.Page);
                field.AddText(" S").AddField(WordFieldType.SectionPages);
            }
            WordParagraph body = document.AddParagraph("A");
            body.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(checked((long)(240) * 635)));
            body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            if (index < 2) body.AddBreak(WordBreakType.Page);
        }
        WordPdfExportOptions options = EmptyRunningOptions();
        options.IncludePageNumbers = true;
        options.PageNumberFormat = "NUMBER {page}/{pages}";
        using var pdf = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        for (int index = 0; index < 3; index++) {
            string text = pdf.GetPage(index + 1).Text;
            if (index != emptyVariant) {
                Assert.Contains($"FIELD{index}:{index + 1} S3", text);
                Assert.DoesNotContain("NUMBER", text);
            } else if (index == 2) Assert.Contains("NUMBER 3/3", text);
            else Assert.DoesNotContain("NUMBER", text);
        }
    }

    [Theory]
    [InlineData(0, "absent")]
    [InlineData(1, "absent")]
    [InlineData(2, "absent")]
    [InlineData(0, "implicit")]
    [InlineData(1, "implicit")]
    [InlineData(2, "implicit")]
    [InlineData(0, "explicit")]
    [InlineData(1, "explicit")]
    [InlineData(2, "explicit")]
    [InlineData(0, "hidden")]
    [InlineData(1, "hidden")]
    [InlineData(2, "hidden")]
    public void SaveAsPdf_EmptyRunningFooterSiblingPreservesAbsentAndEmptyNumberVariants(int variant, string kind) {
        using WordDocument source = CreateEmptyRunningStoryDocument("absent", footer: true);
        WordHeaderFooter[] stories = { source.FooterFirstOrCreate, source.FooterEvenOrCreate, source.FooterDefaultOrCreate };
        foreach (WordHeaderFooter story in stories)
            foreach (var child in story.ChildElements.ToArray()) child.Remove();
        // Creating first/even stories also creates header peers. This control
        // isolates footer admission from those implicit header line boxes.
        source.Sections[0]._sectionProperties.RemoveAllChildren<W.HeaderReference>();
        W.HeaderFooterValues footerKind = variant == 0 ? W.HeaderFooterValues.First
            : variant == 1 ? W.HeaderFooterValues.Even : W.HeaderFooterValues.Default;
        bool absent = kind == "absent";
        if (absent)
            foreach (var reference in source.Sections[0]._sectionProperties.Elements<W.FooterReference>()
                .Where(reference => reference.Type?.Value == footerKind).ToArray()) reference.Remove();
        else if (kind == "explicit" || kind == "hidden") {
            WordParagraph empty = stories[variant].AddParagraph();
            if (kind == "hidden") {
                empty._paragraph.ParagraphProperties = new W.ParagraphProperties(
                    new W.ParagraphMarkRunProperties(new W.Vanish()));
                empty._paragraph.Append(new W.Run(new W.RunProperties(new W.Vanish()), new W.Text("HIDDEN")));
            }
        }
        int targetLines = absent ? 60 : 40;
        for (int index = 0; index < 3; index++) {
            WordParagraph body = source.AddParagraph("A");
            body.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(checked((long)(240) * 635)));
            body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            body.KeepLinesTogether = false; body.AvoidWidowAndOrphan = false;
            for (int line = 1; line < (index == variant ? targetLines : 1); line++) body.AddBreak().AddText("A");
            if (index < 2) body.AddBreak(WordBreakType.Page);
        }
        // Reload reference changes so the importer observes the actual file,
        // including a missing default alongside present first/even siblings.
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        WordPdfExportOptions options = EmptyRunningOptions();
        options.IncludePageNumbers = true;
        options.PageNumberFormat = "NUMBER {page}/{pages}";
        for (int repeat = 0; repeat < 2; repeat++) {
            using var pdf = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
            Assert.Equal(3, pdf.NumberOfPages);
            for (int index = 0; index < 3; index++) {
                var page = pdf.GetPage(index + 1);
                Assert.Equal(index == variant ? targetLines : 1, page.Letters.Count(letter => letter.Value == "A"));
                // The existing flat export generates a default number even
                // without a default part, and leaves text-free first/even
                // variants blank. Neither case gains a body exclusion.
                if (index == 2) Assert.Contains("NUMBER 3/3", page.Text);
                else Assert.DoesNotContain("NUMBER", page.Text);
                Assert.DoesNotContain("HIDDEN", page.Text);
            }
        }
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    public void SaveAsPdf_EmptyRunningFooterPreservesGeneratedPageNumberPresentation(int emptyVariant, bool typography) {
        using WordDocument document = CreateEmptyRunningStoryDocument("absent", footer: true);
        WordHeaderFooter[] stories = { document.FooterFirstOrCreate, document.FooterEvenOrCreate, document.FooterDefaultOrCreate };
        foreach (var child in stories[emptyVariant].ChildElements.ToArray()) child.Remove();
        for (int index = 0; index < 3; index++) {
            if (index != emptyVariant) stories[index].AddParagraph("STORY" + index);
            WordParagraph body = document.AddParagraph("A");
            body.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(checked((long)(240) * 635)));
            body.LineSpacingBefore = body.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            if (index < 2) body.AddBreak(WordBreakType.Page);
        }
        WordPdfExportOptions options = EmptyRunningOptions();
        options.IncludePageNumbers = true;
        options.PageNumberFormat = "NUMBER {page}/{pages}";
        options.PdfOptions = new PdfCreateOptions { FooterOffsetY = typography ? 18D : 24D };
        if (typography) {
            options.PdfOptions.FooterFont = PdfStandardFont.Courier;
            options.PdfOptions.FooterFontSize = 24D;
            options.PdfOptions.FooterAlign = PdfAlign.Left;
            options.PdfOptions.FooterTextColor = global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0);
        }
        using var withImplicit = PdfPigDocument.Open(document.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        W.HeaderFooterValues kind = emptyVariant == 0 ? W.HeaderFooterValues.First
            : emptyVariant == 1 ? W.HeaderFooterValues.Even : W.HeaderFooterValues.Default;
        foreach (var reference in document.Sections[0]._sectionProperties.Elements<W.FooterReference>()
            .Where(reference => reference.Type?.Value == kind).ToArray()) reference.Remove();
        using WordDocument withoutDocument = WordDocument.Load(new MemoryStream(document.ToBytes()));
        using var withoutImplicit = PdfPigDocument.Open(withoutDocument.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, withImplicit.NumberOfPages);
        Assert.Equal(3, withoutImplicit.NumberOfPages);
        for (int index = 0; index < 3; index++) {
            // Removing an empty source part must not restyle or reposition a
            // generated number, nor alter either nonempty sibling variant.
            Assert.Equal(withoutImplicit.GetPage(index + 1).Letters.Select(letter =>
                    (letter.Value, letter.StartBaseLine, letter.FontSize, letter.FontName, letter.Color.ToRGBValues())),
                withImplicit.GetPage(index + 1).Letters.Select(letter =>
                    (letter.Value, letter.StartBaseLine, letter.FontSize, letter.FontName, letter.Color.ToRGBValues())));
            if (index == emptyVariant && emptyVariant < 2)
                Assert.DoesNotContain("NUMBER", withImplicit.GetPage(index + 1).Text);
            else Assert.Contains($"NUMBER {index + 1}/3", withImplicit.GetPage(index + 1).Text);
        }
    }
}
