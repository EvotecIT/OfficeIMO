using DocumentFormat.OpenXml;
using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(null, true)]
    [InlineData("on", true)]
    [InlineData("off", false)]
    public void TableCellBooleanGettersHonorExplicitValuesAcrossFormats(string? value, bool enabled) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        WordTableCell cell = table.Rows[0].Cells[0];
        W.TableCellProperties properties = cell._tableCell.GetFirstChild<W.TableCellProperties>()!;
        properties.AddChild(WithValue(new W.NoWrap(), value), true);
        properties.AddChild(WithValue(new W.TableCellFitText(), value), true);
        properties.AddChild(WithValue(new W.HideMark(), value), true);
        Assert.Equal(!enabled, cell.WrapText);
        Assert.Equal(enabled, cell.FitText);
        Assert.Equal(enabled, cell.HideMark);
        Assert.Empty(document.ValidateDocument());
        foreach (WordFileFormat format in new[] { WordFileFormat.Docx, WordFileFormat.Doc }) {
            using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(format)));
            WordTableCell read = restored.Tables[0].Rows[0].Cells[0];
            Assert.Equal(!enabled, read.WrapText);
            Assert.Equal(enabled, read.FitText);
            Assert.Equal(enabled, read.HideMark);
        }
    }

    [Theory]
    [InlineData("wrap")]
    [InlineData("rowBreak")]
    [InlineData("header")]
    public void TableBooleanSettersOverrideAnExistingExplicitOffValue(string setting) {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        WordTableRow row = table.Rows[0];
        WordTableCell cell = row.Cells[0];
        if (setting == "wrap") {
            cell._tableCell.GetFirstChild<W.TableCellProperties>()!.AddChild(WithValue(new W.NoWrap(), "off"), true);
            cell.WrapText = false;
        } else {
            row._tableRow.TableRowProperties ??= new W.TableRowProperties();
            if (setting == "header") {
                row._tableRow.TableRowProperties.AddChild(WithValue(new W.TableHeader(), "off"), true);
                row.RepeatHeaderRowAtTheTopOfEachPage = true;
            } else {
                row._tableRow.TableRowProperties.AddChild(WithValue(new W.CantSplit(), "off"), true);
                row.AllowRowToBreakAcrossPages = false;
            }
        }
        Assert.Equal(setting == "header", ReadTableBoolean(table, setting));
        Assert.Empty(document.ValidateDocument());
        foreach (WordFileFormat format in new[] { WordFileFormat.Docx, WordFileFormat.Doc }) {
            using WordDocument restored = WordDocument.Load(new MemoryStream(document.ToBytes(format)));
            Assert.Equal(setting == "header", ReadTableBoolean(restored.Tables[0], setting));
        }
    }

    private static T WithValue<T>(T element, string? value) where T : OpenXmlElement {
        if (value != null) element.SetAttribute(new OpenXmlAttribute("w", "val", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", value));
        return element;
    }

    private static bool ReadTableBoolean(WordTable table, string setting) => setting switch {
        "wrap" => table.Rows[0].Cells[0].WrapText,
        "header" => table.Rows[0].RepeatHeaderRowAtTheTopOfEachPage,
        "rowBreak" => table.Rows[0].AllowRowToBreakAcrossPages,
        _ => throw new ArgumentOutOfRangeException(nameof(setting))
    };
}
