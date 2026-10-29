using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wpg = DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using WP14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RemainingGeometry_TextBoxPersistsExactFrameLengthsAndIndependentRelativeSizing(bool inline) {
        using var output = new MemoryStream();
        using (var document = WordDocument.Create()) {
            var box = document.AddTextBox("Exact frame");
            box.HorizontalPositionOffset = OfficeLength.Emus(-3);
            box.VerticalPositionOffset = OfficeLength.Emus(5);
            box.RelativeWidthPercentage = 37;
            box.RelativeHeightPercentage = 23;
            if (inline) box.WrapText = WordImageTextWrapping.InLineWithText;
            box.Width = OfficeLength.Emus(9526);
            box.Height = OfficeLength.Emus(19051);
            if (!inline) {
                Assert.Equal(37, box.RelativeWidthPercentage);
                Assert.Equal(23, box.RelativeHeightPercentage);
            }
            document.Save(output);
        }
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, false)) {
            var drawing = Assert.Single(package.MainDocumentPart!.Document!.Descendants<W.Drawing>());
            var extent = inline ? drawing.Inline!.Extent! : drawing.Anchor!.Extent!;
            Assert.Equal(9526L, extent.Cx!.Value);
            Assert.Equal(19051L, extent.Cy!.Value);
            if (!inline) {
                var anchor = drawing.Anchor!;
                Assert.Equal("-3", anchor.HorizontalPosition!.PositionOffset!.Text);
                Assert.Equal("5", anchor.VerticalPosition!.PositionOffset!.Text);
                Assert.Equal("37000", Assert.Single(anchor.Elements<WP14.RelativeWidth>()).PercentageWidth!.Text);
                Assert.Equal("23000", Assert.Single(anchor.Elements<WP14.RelativeHeight>()).PercentageHeight!.Text);
            }
        }
        output.Position = 0;
        using var loaded = WordDocument.Load(output);
        var saved = Assert.Single(loaded.TextBoxes);
        Assert.Equal(OfficeLength.Emus(9526), saved.Width);
        Assert.Equal(OfficeLength.Emus(19051), saved.Height);
        if (!inline) {
            Assert.Equal(OfficeLength.Emus(-3), saved.HorizontalPositionOffset);
            Assert.Equal(OfficeLength.Emus(5), saved.VerticalPositionOffset);
        }
        Assert.Empty(loaded.ValidateDocument());
    }

    [Fact]
    public void RemainingGeometry_TextBoxRejectsOffsetOverflowBeforeChangingAlignment() {
        using var document = WordDocument.Create();
        var box = document.AddTextBox("Aligned");
        box.HorizontalAlignment = WordTextBoxHorizontalAlignment.Right;
        Assert.Throws<OverflowException>(() => box.HorizontalPositionOffset = OfficeLength.Emus((long)int.MaxValue + 1));
        Assert.Equal(WordTextBoxHorizontalAlignment.Right, box.HorizontalAlignment);
        Assert.Null(box.HorizontalPositionOffset);
        box.HorizontalPositionOffset = null;
        Assert.Equal(WordTextBoxHorizontalAlignment.Right, box.HorizontalAlignment);
        box.VerticalPositionOffset = OfficeLength.Emus(int.MinValue);
        Assert.Throws<OverflowException>(() => box.VerticalPositionOffset = OfficeLength.Emus((long)int.MinValue - 1));
        Assert.Equal(OfficeLength.Emus(int.MinValue), box.VerticalPositionOffset);
    }

    [Fact]
    public void RemainingGeometry_ShapeGroupPersistsExactChildTransformsAndTypedLayout() {
        using var output = new MemoryStream();
        using (var document = WordDocument.Create()) {
            document.AddParagraph().AddShapeGroup(new[] {
                new WordShapeGroupItem(WordShapeType.Rectangle, OfficeLength.Emus(2), OfficeLength.Emus(3),
                    OfficeLength.Emus(9526), OfficeLength.Emus(19051)),
                new WordShapeGroupItem(WordShapeType.Ellipse, OfficeLength.Emus(10001), OfficeLength.Emus(5),
                    OfficeLength.Emus(12701), OfficeLength.Emus(7))
            }, OfficeLength.Emus(11), OfficeLength.Emus(13));
            document.Save(output);
        }
        output.Position = 0;
        using (var package = WordprocessingDocument.Open(output, false)) {
            var anchor = Assert.Single(package.MainDocumentPart!.Document!.Descendants<DW.Anchor>());
            Assert.Equal(22702L, anchor.Extent!.Cx!.Value);
            Assert.Equal(19054L, anchor.Extent.Cy!.Value);
            Assert.Equal("11", anchor.HorizontalPosition!.PositionOffset!.Text);
            Assert.Equal("13", anchor.VerticalPosition!.PositionOffset!.Text);
            var group = Assert.Single(anchor.Descendants<Wpg.WordprocessingGroup>());
            var child = group.Descendants<Wps.WordprocessingShape>().First();
            var transform = Assert.Single(child.Descendants<A.Transform2D>());
            Assert.Equal(2L, transform.Offset!.X!.Value);
            Assert.Equal(3L, transform.Offset.Y!.Value);
            Assert.Equal(9526L, transform.Extents!.Cx!.Value);
            Assert.Equal(19051L, transform.Extents.Cy!.Value);
        }
        output.Position = 0;
        using var loaded = WordDocument.Load(output);
        var saved = Assert.Single(loaded.Paragraphs).ShapeGroup!;
        Assert.True(saved.TryGetLayoutSnapshot(out var layout));
        Assert.Equal(OfficeLength.Emus(22702), layout!.Width);
        Assert.Equal(OfficeLength.Emus(19054), layout.Height);
        Assert.Equal(OfficeLength.Emus(11), layout.HorizontalOffset);
        Assert.Equal(OfficeLength.Emus(13), layout.VerticalOffset);
        Assert.Empty(loaded.ValidateDocument());
    }

    [Theory]
    [InlineData(WordLineSpacingRule.Auto, 360)]
    [InlineData(WordLineSpacingRule.Exact, 361)]
    [InlineData(WordLineSpacingRule.AtLeast, 362)]
    public void RemainingGeometry_MixedSpacingAndTableWidthUnitsRemainNative(WordLineSpacingRule rule, int value) {
        using var output = new MemoryStream();
        using (var document = WordDocument.Create()) {
            var paragraph = document.AddParagraph("Mixed quantity");
            paragraph.LineSpacing = (rule) switch { global::OfficeIMO.Word.WordLineSpacingRule.Exact => global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(checked((long)(value) * 635))), global::OfficeIMO.Word.WordLineSpacingRule.AtLeast => global::OfficeIMO.Word.WordLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Emus(checked((long)(value) * 635))), _ => global::OfficeIMO.Word.WordLineSpacing.Multiple((value) / 240D) };
            var percentage = document.AddTable(1, 1);
            percentage.SetWidthPercentage(37);
            var automatic = document.AddTable(1, 1);
            automatic.WidthType = WordTableWidthUnit.Auto;
            document.Save(output);
        }
        output.Position = 0;
        using var package = WordprocessingDocument.Open(output, false);
        var body = package.MainDocumentPart!.Document!.Body!;
        var spacing = body.Elements<Paragraph>().First().ParagraphProperties!.SpacingBetweenLines!;
        Assert.Equal(value.ToString(System.Globalization.CultureInfo.InvariantCulture), spacing.Line!.Value);
        Assert.Equal(rule.ToString().ToLowerInvariant(), spacing.LineRule!.InnerText!.ToLowerInvariant());
        var tables = body.Elements<Table>().ToArray();
        Assert.Equal(TableWidthUnitValues.Pct, tables[0].GetFirstChild<TableProperties>()!.TableWidth!.Type!.Value);
        Assert.Equal("1850", tables[0].GetFirstChild<TableProperties>()!.TableWidth!.Width!.Value);
        Assert.Equal(TableWidthUnitValues.Auto, tables[1].GetFirstChild<TableProperties>()!.TableWidth!.Type!.Value);
        Assert.Equal("0", tables[1].GetFirstChild<TableProperties>()!.TableWidth!.Width!.Value);
    }
}
