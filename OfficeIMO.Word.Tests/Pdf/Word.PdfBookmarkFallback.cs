using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_WordAuthoredCachedTocKeepsTextWhenATextBoxTargetIsUnprojected() {
        string path = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "FieldEvaluation", "word-generated-toc-table-text-box.docx"));
        using WordDocument document = WordDocument.Load(path, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        string original = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        var options = FootnoteDiagnosticOptions();
        var callerPdfOptions = new PdfOptions { DefaultFont = PdfStandardFont.TimesRoman };
        options.PdfOptions = callerPdfOptions;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);

        PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "NativeBookmarkTargetUnprojected");
        Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.True(result.HasLoss);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        byte[] bytes = result.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("Word-generated table text-box TOC overview", text, StringComparison.Ordinal);
        Assert.Contains("Word-generated table text-box TOC detail", text, StringComparison.Ordinal);
        PdfDocumentInfo navigation = PdfInspector.Inspect(bytes);
        Assert.NotEmpty(navigation.LinkAnnotations.Where(link => link.DestinationName == "_Toc233703443"));
        Assert.Single(navigation.NamedDestinations, target => target.Name == "_Toc233703443");
        Assert.DoesNotContain(navigation.LinkAnnotations, link => link.DestinationName == "_Toc233703444");
        Assert.DoesNotContain(navigation.NamedDestinations, target => target.Name == "_Toc233703444");
        Assert.Equal(bytes, result.ToBytes());
        Assert.Equal(original, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Same(callerPdfOptions, options.PdfOptions);
        Assert.Equal(PdfStandardFont.TimesRoman, callerPdfOptions.DefaultFont);
    }

    [Theory]
    [InlineData("table")]
    [InlineData("textbox")]
    [InlineData("header")]
    public void SaveAsPdf_UnprojectedBookmarkLinksKeepStyledTabTextAndOtherNavigation(string owner) {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        if (owner == "table") source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].SetText("TARGET").AddBookmark("UnprojectedTarget");
        else if (owner == "textbox") source.AddParagraph(string.Empty).AddTextBox("TARGET", WordImageTextWrapping.Square).Paragraphs[0].AddBookmark("UnprojectedTarget");
        else source.HeaderDefaultOrCreate.AddParagraph("TARGET").AddBookmark("UnprojectedTarget");
        source.AddParagraph("VALIDTARGET").AddBookmark("ValidTarget");
        WordParagraph body = source.AddParagraph(string.Empty);
        body._paragraph.Append(BookmarkFallbackLink("UNSAFE", "BODYTAIL", "UnprojectedTarget"));
        WordParagraph cell = source.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        cell._paragraph.Append(BookmarkFallbackLink("CELL", "CELLTAIL", "UnprojectedTarget"));
        source.AddParagraph(string.Empty).AddHyperLink("VALIDLINK", "ValidTarget");
        source.AddParagraph(string.Empty).AddHyperLink("EXTERNAL", new Uri("https://example.test/bookmark-fallback"));
        Assert.Empty(source.ValidateDocument());
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string original = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml +
            string.Concat(document._wordprocessingDocument.MainDocumentPart.HeaderParts.Select(part => part.Header.OuterXml));
        var options = FootnoteDiagnosticOptions();
        options.FontFamily = "Courier";
        var callerPdfOptions = new PdfOptions { DefaultFont = PdfStandardFont.Courier };
        options.PdfOptions = callerPdfOptions;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
        PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code.StartsWith("NativeBookmarkTarget", StringComparison.Ordinal));
        Assert.Equal("NativeBookmarkTargetUnprojected", warning.Code);
        Assert.Equal(OfficeConversionLossKind.Omission, warning.LossKind);
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var bodyPage = pdf.GetPages().Single(page => page.Text.Contains("UNSAFE", StringComparison.Ordinal));
        var letters = FindRichNoteToken(bodyPage.Letters.ToArray(), "UNSAFE");
        Assert.All(letters, letter => {
            Assert.Contains("Courier", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Bold", letter.FontName, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(12D, letter.FontSize, 4);
        });
        var linkLetters = FindRichNoteToken(bodyPage.Letters.ToArray(), "BODYTAIL");
        Assert.True(linkLetters[0].StartBaseLine.X > letters[letters.Length - 1].EndBaseLine.X + 1D);
        Assert.Contains(bodyPage.Paths, path => path.IsStroked && path.GetBoundingRectangle() is { } rectangle &&
            rectangle.Left <= letters[0].StartBaseLine.X + .01D &&
            rectangle.Right >= letters[letters.Length - 1].EndBaseLine.X - .01D &&
            rectangle.Top <= letters[0].StartBaseLine.Y && rectangle.Bottom >= letters[0].StartBaseLine.Y - 5D &&
            rectangle.Height < .01D);
        Assert.Contains("CELL", string.Concat(pdf.GetPages().Select(page => page.Text)), StringComparison.Ordinal);
        PdfDocumentInfo navigation = PdfInspector.Inspect(result.ToBytes());
        Assert.DoesNotContain(navigation.LinkAnnotations, link => link.DestinationName == "UnprojectedTarget");
        Assert.Single(navigation.LinkAnnotations, link => link.DestinationName == "ValidTarget");
        Assert.Single(navigation.LinkAnnotations, link => link.Uri == "https://example.test/bookmark-fallback");
        Assert.Equal(2, navigation.LinkAnnotations.Count);
        Assert.Equal(original, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml +
            string.Concat(document._wordprocessingDocument.MainDocumentPart.HeaderParts.Select(part => part.Header.OuterXml)));
        Assert.Equal("Courier", options.FontFamily);
        Assert.Same(callerPdfOptions, options.PdfOptions);
        Assert.Equal(PdfStandardFont.Courier, callerPdfOptions.DefaultFont);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SaveAsPdf_UnprojectedBodyAliasAndMissingSourceHaveDistinctLossDiagnostics() {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        WordParagraph target = source.AddParagraph("TARGET");
        target._paragraph.InsertAt(new W.BookmarkStart { Id = "20", Name = "FirstTarget" }, 0);
        target._paragraph.Append(new W.BookmarkEnd { Id = "20" },
            new W.BookmarkStart { Id = "21", Name = "ExtraTarget" }, new W.BookmarkEnd { Id = "21" });
        source.AddParagraph(string.Empty).AddHyperLink("FIRSTLINK", "FirstTarget");
        source.AddParagraph(string.Empty).AddHyperLink("EXTRALINK", "ExtraTarget");
        source.AddParagraph(string.Empty).AddHyperLink("MISSINGLINK", "AbsentTarget");
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string original = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        Assert.Equal(OfficeConversionLossKind.Omission,
            Assert.Single(result.Warnings, item => item.Code == "NativeBookmarkTargetUnprojected").LossKind);
        Assert.Equal(OfficeConversionLossKind.Omission,
            Assert.Single(result.Warnings, item => item.Code == "NativeBookmarkTargetMissing").LossKind);
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        foreach (string token in new[] { "TARGET", "FIRSTLINK", "EXTRALINK", "MISSINGLINK" }) Assert.Contains(token, text, StringComparison.Ordinal);
        PdfDocumentInfo navigation = PdfInspector.Inspect(result.ToBytes());
        Assert.Single(navigation.LinkAnnotations);
        Assert.Single(navigation.NamedDestinations, item => item.Name == "FirstTarget");
        Assert.Equal("FirstTarget", navigation.LinkAnnotations[0].DestinationName);
        Assert.Equal(original, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_ReferencedNoteAndFlowBreakBookmarkTargetsRemainNavigable() {
        using WordDocument source = CreateFootnoteDiagnosticDocument();
        WordParagraph target = source.AddParagraph(string.Empty);
        target._paragraph.Append(new W.BookmarkStart { Id = "20", Name = "BeforeBreak" }, RichNoteRun("FIRSTTARGET"),
            new W.BookmarkEnd { Id = "20" }, new W.Run(new W.Break { Type = W.BreakValues.Page }),
            new W.BookmarkStart { Id = "21", Name = "AfterBreak" }, RichNoteRun("SECONDTARGET"), new W.BookmarkEnd { Id = "21" });
        OpenXmlCompositeElement note = AddRichNoteFixture(source, endnote: false);
        WordParagraph noteTarget = source.AddParagraph("NOTETARGET");
        noteTarget._paragraph.Remove();
        note.Append(noteTarget._paragraph);
        noteTarget.AddBookmark("SupportedNote");
        foreach (string name in new[] { "BeforeBreak", "AfterBreak", "SupportedNote" })
            source.AddParagraph(string.Empty).AddHyperLink("OPEN" + name.ToUpperInvariant(), name);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        var main = document._wordprocessingDocument.MainDocumentPart!;
        string original = main.Document.OuterXml + main.FootnotesPart!.Footnotes!.OuterXml;
        PdfDocumentConversionResult result = document.ToPdfDocumentResult(FootnoteDiagnosticOptions());
        Assert.DoesNotContain(result.Warnings, item => item.Code.StartsWith("NativeBookmarkTarget", StringComparison.Ordinal));
        byte[] bytes = result.ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        PdfDocumentInfo navigation = PdfInspector.Inspect(bytes);
        foreach ((string name, string token) in new[] { ("BeforeBreak", "FIRSTTARGET"), ("AfterBreak", "SECONDTARGET"), ("SupportedNote", "NOTETARGET") }) {
            PdfNamedDestination destination = Assert.Single(navigation.NamedDestinations, item => item.Name == name);
            Assert.Single(navigation.LinkAnnotations, link => link.DestinationName == name);
            var page = pdf.GetPages().Single(item => item.Text.Contains(token, StringComparison.Ordinal));
            Assert.Equal(page.Number, destination.PageNumber);
        }
        Assert.Equal(original, main.Document.OuterXml + main.FootnotesPart.Footnotes.OuterXml);
    }

    private static W.Hyperlink BookmarkFallbackLink(string first, string second, string name) => new(new W.Run(
        new W.RunProperties(new W.RunFonts { Ascii = "Courier", HighAnsi = "Courier" }, new W.Bold(), new W.FontSize { Val = "24" }),
        new W.Text(first), new W.TabChar(), new W.Text(second))) { Anchor = name };
}
