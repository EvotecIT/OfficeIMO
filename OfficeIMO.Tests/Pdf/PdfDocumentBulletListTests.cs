using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using UglyToad.PdfPig.Content;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocumentBulletListTests {
    [Fact]
    public void Bullets_WithNullItems_ThrowsArgumentNullException() {
        var doc = PdfDocument.Create();

        var exception = Assert.Throws<ArgumentNullException>(() => doc.Bullets(null!));

        Assert.Equal("items", exception.ParamName);
        Assert.Contains("Parameter 'items' cannot be null.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Numbered_WithNullItems_ThrowsArgumentNullException() {
        var doc = PdfDocument.Create();

        var exception = Assert.Throws<ArgumentNullException>(() => doc.Numbered(null!));

        Assert.Equal("items", exception.ParamName);
        Assert.Contains("Parameter 'items' cannot be null.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Numbered_WithInvalidStartNumber_ThrowsArgumentOutOfRangeException() {
        var doc = PdfDocument.Create();

        var exception = Assert.Throws<ArgumentOutOfRangeException>(() => doc.Numbered(new[] { "Invalid" }, startNumber: 0));

        Assert.Equal("startNumber", exception.ParamName);
        Assert.Contains("Numbered lists must start at 1 or greater.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ListBlocks_SnapshotInputItemsBeforeRendering() {
        var bulletItems = new System.Collections.Generic.List<string> { "Original bullet", "Stable bullet" };
        var numberedItems = new System.Collections.Generic.List<string> { "Original step", "Stable step" };

        var doc = PdfDocument.Create()
            .Bullets(bulletItems)
            .Numbered(numberedItems);

        bulletItems[0] = "Mutated bullet";
        bulletItems.Add("Late bullet");
        numberedItems[0] = "Mutated step";
        numberedItems.Add("Late step");

        using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
        string text = pdf.GetPage(1).Text;

        Assert.Contains("Original bullet", text, StringComparison.Ordinal);
        Assert.Contains("Stable bullet", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Mutated bullet", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Late bullet", text, StringComparison.Ordinal);
        Assert.Contains("Original step", text, StringComparison.Ordinal);
        Assert.Contains("Stable step", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Mutated step", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Late step", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ListBlocks_SnapshotInputStylesBeforeRendering() {
        var style = new PdfListStyle {
            FontSize = 13,
            LineHeight = 1.2,
            LeftIndent = 12,
            MarkerGap = 8,
            SpacingBefore = 4,
            SpacingAfter = 9,
            ItemSpacing = 5,
            Color = PdfColor.FromRgb(10, 20, 30)
        };

        var bulletBlock = new BulletListBlock(new[] { "A" }, PdfAlign.Left, null, style);
        var numberedBlock = new NumberedListBlock(new[] { "One" }, PdfAlign.Left, null, 1, style);

        style.FontSize = 20;
        style.LeftIndent = 0;
        style.Color = PdfColor.Black;

        Assert.Equal(13, bulletBlock.Style!.FontSize);
        Assert.Equal(1.2, bulletBlock.Style.LineHeight);
        Assert.Equal(12, bulletBlock.Style.LeftIndent);
        Assert.Equal(8, bulletBlock.Style.MarkerGap);
        Assert.Equal(4, bulletBlock.Style.SpacingBefore);
        Assert.Equal(9, bulletBlock.Style.SpacingAfter);
        Assert.Equal(5, bulletBlock.Style.ItemSpacing);
        Assert.Equal(PdfColor.FromRgb(10, 20, 30), bulletBlock.Style.Color);
        Assert.Equal(13, numberedBlock.Style!.FontSize);
    }

    [Fact]
    public void Options_SnapshotDefaultListStyle() {
        var style = new PdfListStyle {
            FontSize = 13,
            LeftIndent = 14,
            Color = PdfColor.FromRgb(10, 20, 30)
        };
        var options = new PdfOptions {
            DefaultListStyle = style
        };

        style.FontSize = 20;
        style.LeftIndent = 0;
        style.Color = PdfColor.Black;

        PdfListStyle readback = options.DefaultListStyle!;
        readback.FontSize = 8;

        PdfOptions clone = options.Clone();

        Assert.Equal(13, options.DefaultListStyle!.FontSize);
        Assert.Equal(14, options.DefaultListStyle.LeftIndent);
        Assert.Equal(PdfColor.FromRgb(10, 20, 30), options.DefaultListStyle.Color);
        Assert.Equal(13, clone.DefaultListStyle!.FontSize);
    }

    [Fact]
    public void ListBlockItemCollections_AreReadOnlySnapshots() {
        var bulletBlock = new BulletListBlock(new[] { "A", null!, "B" }, PdfAlign.Left, null);
        var numberedBlock = new NumberedListBlock(new[] { "One", null!, "Two" }, PdfAlign.Left, null, 1);

        Assert.Equal(new[] { "A", "B" }, bulletBlock.Items);
        Assert.Equal(new[] { "One", "Two" }, numberedBlock.Items);
        Assert.False(bulletBlock.Items is System.Collections.Generic.List<string>);
        Assert.False(numberedBlock.Items is System.Collections.Generic.List<string>);
    }

    [Fact]
    public void RichListBlocks_SnapshotInputRunsBeforeRendering() {
        var runs = new System.Collections.Generic.List<TextRun> {
            TextRun.Normal("Original"),
            TextRun.Bolded(" Bold")
        };
        var items = new System.Collections.Generic.List<PdfListItem> {
            new PdfListItem(runs)
        };

        var doc = PdfDocument.Create()
            .RichBullets(items)
            .RichNumbered(items);

        runs[0] = TextRun.Normal("Mutated");
        runs.Add(TextRun.Normal(" Late"));
        items[0] = PdfListItem.Plain("Late item");

        using var pdf = PdfPigDocument.Open(new MemoryStream(doc.ToBytes()));
        string text = pdf.GetPage(1).Text;

        Assert.Equal(2, CountOccurrences(text, "Original Bold"));
        Assert.DoesNotContain("Mutated", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Late", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RichBullets_CanAnchorBookmarksOnListItems() {
        byte[] bytes = PdfDocument.Create()
            .RichBullets(new[] {
                PdfListItem.Plain("Bookmarked bullet", "BulletAnchor"),
                PdfListItem.Rich(new[] { TextRun.Normal("Second bookmarked bullet") }, "SecondBulletAnchor")
            })
            .Paragraph(p => p.LinkToBookmark("Jump to second bullet", "SecondBulletAnchor", contents: "List bookmark jump"))
            .ToBytes();

        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes, new PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        var listItems = PdfTextExtractor.ExtractListItemsByPage(bytes)
            .SelectMany(page => page.ListItems)
            .ToList();

        Assert.Contains(logical.NamedDestinations, destination => destination.Name == "BulletAnchor");
        Assert.Contains(logical.NamedDestinations, destination => destination.Name == "SecondBulletAnchor");
        Assert.Contains(listItems, item => item.Text == "Bookmarked bullet");
        Assert.Contains(listItems, item => item.Text == "Second bookmarked bullet");
        Assert.Contains(logical.GetLinksByDestinationName("SecondBulletAnchor"), link => link.Contents == "List bookmark jump");
    }

    [Fact]
    public void RichListItems_CanRenderExplicitMarkers() {
        byte[] bytes = PdfDocument.Create()
            .RichNumbered(new[] {
                PdfListItem.Plain("Alpha item", marker: "a)"),
                PdfListItem.Plain("Roman item", marker: "iv)")
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string pageText = pdf.GetPage(1).Text;

        Assert.Contains("a)Alpha item", pageText, StringComparison.Ordinal);
        Assert.Contains("iv)Roman item", pageText, StringComparison.Ordinal);
        Assert.True(pageText.IndexOf("a)Alpha item", StringComparison.Ordinal) < pageText.IndexOf("iv)Roman item", StringComparison.Ordinal));
    }

    [Fact]
    public void RichBullets_RenderScopedRunStylesAndLinks() {
        const string linkUri = "https://evotec.xyz/pdf-rich-list";
        byte[] bytes = PdfDocument.Create()
            .RichBullets(new[] {
                new PdfListItem(new[] {
                    TextRun.Normal("Plain "),
                    TextRun.Bolded("Bold"),
                    TextRun.Normal(" "),
                    TextRun.Normal("Red", PdfColor.FromRgb(255, 0, 0)),
                    TextRun.Normal(" "),
                    TextRun.Normal("Marked", backgroundColor: PdfColor.FromRgb(255, 255, 0)),
                    TextRun.Normal(" "),
                    TextRun.Link("Linked", linkUri, contents: "Rich list metadata")
                })
            })
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int boldText = content.IndexOf("<426F6C64>", StringComparison.Ordinal);
        int redText = content.IndexOf("<526564>", StringComparison.Ordinal);
        int markedText = content.IndexOf("<4D61726B6564>", StringComparison.Ordinal);

        using (PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes))) {
            Assert.Contains("Plain Bold Red Marked Linked", pdf.GetPage(1).Text, StringComparison.Ordinal);
        }

        PdfLogicalDocument logical = PdfLogicalDocument.Load(bytes);
        PdfLogicalLinkAnnotation link = Assert.Single(logical.GetLinksByUri(linkUri));

        Assert.True(boldText >= 0, "Expected encoded 'Bold' text in the list content stream.");
        Assert.True(redText > boldText, "Expected encoded 'Red' text after the bold list run.");
        Assert.True(markedText > redText, "Expected encoded 'Marked' text after the colored list run.");
        Assert.True(content.LastIndexOf("/F2 11 Tf", boldText, StringComparison.Ordinal) >= 0, "Expected rich list bold text to use the bold PDF font resource.");
        Assert.True(content.LastIndexOf("1 0 0 rg", redText, StringComparison.Ordinal) >= 0, "Expected rich list run color to emit a red PDF fill color.");
        Assert.True(content.LastIndexOf("1 1 0 rg", markedText, StringComparison.Ordinal) >= 0, "Expected rich list highlight to emit a yellow PDF fill color.");
        Assert.Equal("Rich list metadata", link.Contents);
    }

    [Fact]
    public void Bullets_RenderGlyphsWithHangingIndent() {
        var doc = PdfDocument.Create();
        doc.Bullets(new[] {
            "Short bullet",
            "This bullet contains enough text to wrap across multiple lines so that we can validate hanging indentation in the generated PDF output."
        });

        var bytes = doc.ToBytes();
        Assert.NotEmpty(bytes);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var lineGroups = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .ToList();

        var bulletIndices = lineGroups
            .Select((group, index) => (group, index))
            .Where(tuple => tuple.group.Any(letter => letter.Value == "•"))
            .ToList();

        Assert.Equal(2, bulletIndices.Count);

        var firstLine = lineGroups[bulletIndices[0].index].OrderBy(letter => letter.StartBaseLine.X).ToList();
        Assert.Equal("•", firstLine[0].Value);
        double firstTextX = firstLine.First(letter => letter.Value != "•").StartBaseLine.X;
        double firstBulletX = firstLine[0].StartBaseLine.X;
        Assert.True(firstTextX - firstBulletX > 4);

        var secondLine = lineGroups[bulletIndices[1].index].OrderBy(letter => letter.StartBaseLine.X).ToList();
        Assert.Equal("•", secondLine[0].Value);
        double secondTextX = secondLine.First(letter => letter.Value != "•").StartBaseLine.X;
        double secondBulletX = secondLine[0].StartBaseLine.X;
        Assert.True(secondTextX - secondBulletX > 4);

        var wrapLine = lineGroups
            .Skip(bulletIndices[1].index + 1)
            .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
            .First();
        Assert.DoesNotContain(wrapLine, letter => letter.Value == "•");
        double wrapTextX = wrapLine[0].StartBaseLine.X;
        Assert.InRange(wrapTextX, secondTextX - 1, secondTextX + 1);
    }

    [Fact]
    public void Numbered_RenderMarkersWithHangingIndent() {
        var doc = PdfDocument.Create();
        doc.Numbered(new[] {
            "Short step",
            "This numbered item contains enough text to wrap across multiple lines so that we can validate hanging indentation in the generated PDF output."
        }, startNumber: 9);

        var bytes = doc.ToBytes();
        Assert.NotEmpty(bytes);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var lineGroups = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .ToList();

        var numberedIndices = lineGroups
            .Select((group, index) => (line: string.Concat(group.OrderBy(letter => letter.StartBaseLine.X).Select(letter => letter.Value)), index))
            .Where(tuple => tuple.line.StartsWith("9.", StringComparison.Ordinal) || tuple.line.StartsWith("10.", StringComparison.Ordinal))
            .ToList();

        Assert.Equal(2, numberedIndices.Count);

        var firstLine = lineGroups[numberedIndices[0].index].OrderBy(letter => letter.StartBaseLine.X).ToList();
        Assert.StartsWith("9.", string.Concat(firstLine.Select(letter => letter.Value)), StringComparison.Ordinal);
        double firstTextX = firstLine.First(letter => letter.Value == "S").StartBaseLine.X;
        double firstMarkerX = firstLine[0].StartBaseLine.X;
        Assert.True(firstTextX - firstMarkerX > 8);

        var secondLine = lineGroups[numberedIndices[1].index].OrderBy(letter => letter.StartBaseLine.X).ToList();
        Assert.StartsWith("10.", string.Concat(secondLine.Select(letter => letter.Value)), StringComparison.Ordinal);
        double secondTextX = secondLine.First(letter => letter.Value == "T").StartBaseLine.X;
        double secondMarkerX = secondLine[0].StartBaseLine.X;
        Assert.True(secondTextX - secondMarkerX > 8);

        var wrapLine = lineGroups
            .Skip(numberedIndices[1].index + 1)
            .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
            .First();
        string wrapText = string.Concat(wrapLine.Select(letter => letter.Value));
        Assert.False(wrapText.StartsWith("10.", StringComparison.Ordinal));
        double wrapTextX = wrapLine[0].StartBaseLine.X;
        Assert.InRange(wrapTextX, secondTextX - 1, secondTextX + 1);
    }

    [Fact]
    public void Bullets_RespectAlignmentOptions() {
        var options = new PdfOptions();
        var doc = PdfDocument.Create(options);
        doc.Bullets(new[] { "Left aligned" }, PdfAlign.Left);
        doc.Bullets(new[] { "Centered bullet" }, PdfAlign.Center);
        doc.Bullets(new[] { "Right aligned bullet" }, PdfAlign.Right);

        var bytes = doc.ToBytes();
        Assert.NotEmpty(bytes);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bulletLines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
            .Where(line => line.Any(letter => letter.Value == "•"))
            .ToList();

        Assert.Equal(3, bulletLines.Count);

        var leftLine = bulletLines[0];
        var centerLine = bulletLines[1];
        var rightLine = bulletLines[2];

        double leftBulletX = leftLine.First(letter => letter.Value == "•").StartBaseLine.X;
        double leftTextX = leftLine.First(letter => letter.Value != "•").StartBaseLine.X;
        double centerBulletX = centerLine.First(letter => letter.Value == "•").StartBaseLine.X;
        double centerTextX = centerLine.First(letter => letter.Value != "•").StartBaseLine.X;
        double rightBulletX = rightLine.First(letter => letter.Value == "•").StartBaseLine.X;
        double rightTextX = rightLine.First(letter => letter.Value != "•").StartBaseLine.X;

        Assert.True(centerBulletX > leftBulletX + 10);
        Assert.True(rightBulletX > centerBulletX + 10);
        double leftGap = leftTextX - leftBulletX;
        double centerGap = centerTextX - centerBulletX;
        double rightGap = rightTextX - rightBulletX;
        Assert.InRange(centerGap, leftGap - 1, leftGap + 1);
        Assert.InRange(rightGap, leftGap - 1, leftGap + 1);

        double contentRight = options.PageWidth - options.MarginRight;
        double rightmostTextX = rightLine.Where(letter => letter.Value != "•").Max(letter => letter.EndBaseLine.X);
        Assert.InRange(rightmostTextX, contentRight - 1.5, contentRight + 0.5);
    }

    [Fact]
    public void Bullets_ApplyCustomColorToGlyphs() {
        var doc = PdfDocument.Create();
        var color = new PdfColor(0.2, 0.4, 0.6);
        doc.Bullets(new[] { "Colored bullet" }, PdfAlign.Left, color);

        var bytes = doc.ToBytes();
        Assert.NotEmpty(bytes);

        var pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.2 0.4 0.6 rg", pdfContent);
    }

    [Fact]
    public void Numbered_ApplyCustomColorToMarkers() {
        var doc = PdfDocument.Create();
        var color = new PdfColor(0.2, 0.4, 0.6);
        doc.Numbered(new[] { "Colored marker" }, PdfAlign.Left, color);

        var bytes = doc.ToBytes();
        Assert.NotEmpty(bytes);

        var pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.2 0.4 0.6 rg", pdfContent);
        Assert.Contains("<312E> Tj", pdfContent);
    }

    [Fact]
    public void DefaultListStyle_AppliesFontColorIndentAndSpacingToFollowingLists() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = new PdfListStyle {
            FontSize = 13,
            LineHeight = 1,
            LeftIndent = 16,
            MarkerGap = 10,
            SpacingBefore = 8,
            SpacingAfter = 18,
            Color = PdfColor.FromRgb(10, 20, 30)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("BeforeList"), style: new PdfParagraphStyle {
                SpacingAfter = 0
            })
            .DefaultListStyle(style)
            .Bullets(new[] { "StyledBullet" })
            .Paragraph(p => p.Text("AfterList"), style: new PdfParagraphStyle {
                SpacingAfter = 0
            })
            .ToBytes();

        style.FontSize = 20;
        style.LeftIndent = 0;
        style.Color = PdfColor.Black;

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bulletLine = GetLineContaining(page, "StyledBullet");
        var afterLine = GetLineContaining(page, "AfterList");
        double bulletPointSize = bulletLine.First(letter => letter.Value == "S").PointSize;
        double bulletX = bulletLine.First(letter => letter.Value == "•").StartBaseLine.X;
        double textX = bulletLine.First(letter => letter.Value == "S").StartBaseLine.X;
        double afterGap = bulletLine.First().StartBaseLine.Y - afterLine.First().StartBaseLine.Y;
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(bulletPointSize, 12.5, 13.5);
        Assert.InRange(bulletX, options.MarginLeft + 15, options.MarginLeft + 17);
        Assert.InRange(textX - bulletX, 14, 16);
        Assert.InRange(afterGap, 28, 36);
        Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ListStyle_RejectsInvalidIntrinsicValues() {
        Assert.Equal("FontSize", Assert.Throws<ArgumentException>(() => new PdfListStyle { FontSize = 0 }).ParamName);
        Assert.Equal("LineHeight", Assert.Throws<ArgumentException>(() => new PdfListStyle { LineHeight = double.NaN }).ParamName);
        Assert.Equal("LeftIndent", Assert.Throws<ArgumentException>(() => new PdfListStyle { LeftIndent = -1 }).ParamName);
        Assert.Equal("MarkerGap", Assert.Throws<ArgumentException>(() => new PdfListStyle { MarkerGap = -1 }).ParamName);
        Assert.Equal("SpacingBefore", Assert.Throws<ArgumentException>(() => new PdfListStyle { SpacingBefore = -1 }).ParamName);
        Assert.Equal("SpacingAfter", Assert.Throws<ArgumentException>(() => new PdfListStyle { SpacingAfter = -1 }).ParamName);
        Assert.Equal("ItemSpacing", Assert.Throws<ArgumentException>(() => new PdfListStyle { ItemSpacing = -1 }).ParamName);
        Assert.Throws<ArgumentNullException>(() => PdfDocument.Create().DefaultListStyle(null!));
    }

    [Fact]
    public void Bullets_SplitLongItemsAcrossPagesWithoutCrossingBottomMargin() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        string longItem = string.Join(" ", Enumerable.Range(1, 150).Select(i => "bullet" + i.ToString("000")));

        byte[] bytes = PdfDocument.Create(options)
            .Bullets(new[] { longItem })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected one very long bullet item to continue onto another page.");
        AssertListTextStaysAboveBottomMargin(pdf, options);
        Assert.Contains("bullet001", pdf.GetPage(1).Text);
        Assert.Contains("bullet150", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Numbered_SplitLongItemsAcrossPagesWithoutCrossingBottomMargin() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        string longItem = string.Join(" ", Enumerable.Range(1, 150).Select(i => "step" + i.ToString("000")));

        byte[] bytes = PdfDocument.Create(options)
            .Numbered(new[] { longItem }, startNumber: 3)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected one very long numbered item to continue onto another page.");
        AssertListTextStaysAboveBottomMargin(pdf, options);
        Assert.Contains("step001", pdf.GetPage(1).Text);
        Assert.Contains("step150", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Bullets_WithIndentWiderThanContent_DoesNotShiftBulletPastMargin() {
        var options = new PdfOptions {
            PageWidth = 110,
            MarginLeft = 50,
            MarginRight = 50,
            DefaultFontSize = 11
        };

        var doc = PdfDocument.Create(options);
        doc.Bullets(new[] { "A" }, PdfAlign.Right);

        var bytes = doc.ToBytes();
        Assert.NotEmpty(bytes);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var bulletLine = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
            .First(line => line.Any(letter => letter.Value == "•"));

        double bulletX = bulletLine.First(letter => letter.Value == "•").StartBaseLine.X;
        double textX = bulletLine.First(letter => letter.Value != "•").StartBaseLine.X;

        Assert.InRange(bulletX, options.MarginLeft - 0.5, options.MarginLeft + 0.5);
        Assert.True(textX > bulletX);
    }

    private static void AssertListTextStaysAboveBottomMargin(PdfPigDocument pdf, PdfOptions options) {
        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom - 2, $"Expected list text to stay above the bottom margin on page {pageNumber}.");
        }
    }

    private static System.Collections.Generic.List<Letter> GetLineContaining(Page page, string text) {
        return page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
            .First(line => string.Concat(line.Select(letter => letter.Value)).Contains(text, StringComparison.Ordinal));
    }

    private static int CountOccurrences(string value, string search) {
        int count = 0;
        int index = 0;
        while ((index = value.IndexOf(search, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += search.Length;
        }

        return count;
    }
}
