using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("auto", false)]
    [InlineData("pct", false)]
    [InlineData("auto", true)]
    [InlineData("pct", true)]
    public void SaveAsPdf_ExplicitNonAbsoluteCellSpacingClearsInheritedSpacing(string unit, bool derivedStyle) {
        using WordDocument document = WordDocument.Create();
        WordTable table = CreateAutomaticWidthControl(document, 2400, 2400);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Spacing0";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Spacing1";
        table._tableProperties!.TableCellSpacing = null;
        var styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new W.Style(new W.BasedOn { Val = "TableGrid" },
            new W.StyleTableProperties(new W.TableCellSpacing { Type = W.TableWidthUnitValues.Dxa, Width = "120" })) {
            StyleId = "SpacedParent", Type = W.StyleValues.Table
        });
        var spacing = new W.TableCellSpacing {
            Type = unit == "auto" ? W.TableWidthUnitValues.Auto : W.TableWidthUnitValues.Pct, Width = "120"
        };
        if (derivedStyle) {
            styles.Append(new W.Style(new W.BasedOn { Val = "SpacedParent" }, new W.StyleTableProperties(spacing)) {
                StyleId = "SpacedChild", Type = W.StyleValues.Table
            });
            table._tableProperties.TableStyle = new W.TableStyle { Val = "SpacedChild" };
        } else {
            table._tableProperties.TableStyle = new W.TableStyle { Val = "SpacedParent" };
            table._tableProperties.TableCellSpacing = spacing;
        }
        using WordDocument imported = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string sourceXml = imported._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        using PdfPigDocument pdf = PdfPigDocument.Open(imported.ToPdfBytes(new WordToPdfOptions {
            IncludePageNumbers = false, ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }));
        Assert.Equal(sourceXml, imported._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        var words = pdf.GetPage(1).GetWords();
        var left = Assert.Single(words, word => word.Text == "Spacing0");
        var right = Assert.Single(words, word => word.Text == "Spacing1");
        Assert.Equal(120D, right.BoundingBox.Left - left.BoundingBox.Left, 3);
    }
}
