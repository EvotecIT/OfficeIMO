using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocBulletListTests {
    [Fact]
    public void Bullets_WithNullItems_ThrowsArgumentNullException() {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentNullException>(() => doc.Bullets(null!));

        Assert.Equal("items", exception.ParamName);
        Assert.Contains("Parameter 'items' cannot be null.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Bullets_RenderGlyphsWithHangingIndent() {
        var doc = PdfDoc.Create();
        doc.Bullets(new[] {
            "Short bullet",
            "This bullet contains enough text to wrap across multiple lines so that we can validate hanging indentation in the generated PDF output."
        });

        var bytes = doc.ToBytes();
        Assert.NotEmpty(bytes);

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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
    public void Bullets_RespectAlignmentOptions() {
        var options = new PdfOptions();
        var doc = PdfDoc.Create(options);
        doc.Bullets(new[] { "Left aligned" }, PdfAlign.Left);
        doc.Bullets(new[] { "Centered bullet" }, PdfAlign.Center);
        doc.Bullets(new[] { "Right aligned bullet" }, PdfAlign.Right);

        var bytes = doc.ToBytes();
        Assert.NotEmpty(bytes);

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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
        var doc = PdfDoc.Create();
        var color = new PdfColor(0.2, 0.4, 0.6);
        doc.Bullets(new[] { "Colored bullet" }, PdfAlign.Left, color);

        var bytes = doc.ToBytes();
        Assert.NotEmpty(bytes);

        var pdfContent = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.2 0.4 0.6 rg", pdfContent);
    }

    [Fact]
    public void Bullets_WithIndentWiderThanContent_DoesNotShiftBulletPastMargin() {
        var options = new PdfOptions {
            PageWidth = 110,
            MarginLeft = 50,
            MarginRight = 50,
            DefaultFontSize = 11
        };

        var doc = PdfDoc.Create(options);
        doc.Bullets(new[] { "A" }, PdfAlign.Right);

        var bytes = doc.ToBytes();
        Assert.NotEmpty(bytes);

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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
}
