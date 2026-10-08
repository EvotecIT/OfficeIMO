using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyDoc_MissingAutofitFlagPreservesFixedColumnDefault() {
        byte[] bytes = LegacyDocTestBuilder.CreateUnicodeDocWithTableCellWidths();
        using WordDocument imported = WordDocument.Load(new MemoryStream(bytes));
        WordTable table = Assert.Single(imported.Tables);
        Assert.Equal(WordTableLayoutMode.Fixed, table.LayoutMode);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_SaveMaterializesDocxAutofitDefault(bool emptyLayoutElement) {
        using WordDocument source = WordDocument.Create();
        WordTable table = source.AddTable(1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Automatic table";
        table._tableProperties!.TableLayout = emptyLayoutElement ? new W.TableLayout() : null;
        Assert.Equal(WordTableLayoutMode.AutoFit, table.LayoutMode);
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);
        Assert.True(ContainsBytePattern(ReadCompoundStream(bytes, "WordDocument"), 0x15, 0x36, 0x01),
            "DOCX's implicit AutoFit must be emitted explicitly because binary DOC defaults to fixed columns.");
        using WordDocument imported = WordDocument.Load(new MemoryStream(bytes));
        Assert.Equal(WordTableLayoutMode.AutoFit, Assert.Single(imported.Tables).LayoutMode);
    }
}
