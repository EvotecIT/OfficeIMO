using System;
using System.IO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body", "fit")]
    [InlineData("body", "noWrap")]
    [InlineData("body", "hide")]
    [InlineData("header", "fit")]
    [InlineData("header", "noWrap")]
    [InlineData("header", "hide")]
    [InlineData("footer", "fit")]
    [InlineData("footer", "noWrap")]
    [InlineData("footer", "hide")]
    [InlineData("body", "none")]
    [InlineData("header", "none")]
    [InlineData("footer", "none")]
    public void LegacyDoc_NativeCellTextLayoutDeclaresAFormatThatPreservesItsFlags(string story, string setting) {
        using WordDocument document = WordDocument.Create();
        WordTable table;
        if (story == "body") table = document.AddTable(1, 1);
        else {
            document.AddHeadersAndFooters();
            table = (story == "header" ? (WordHeaderFooter)document.Header.Default : document.Footer.Default).AddTable(1, 1);
        }
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.Paragraphs[0].Text = "Cell text";
        if (setting == "fit") cell.FitText = true;
        if (setting == "noWrap") cell.WrapText = false;
        if (setting == "hide") cell.HideMark = true;
        byte[] bytes = document.ToBytes(WordFileFormat.Doc);
        Assert.True(OfficeCompoundFileReader.TryRead(bytes, out OfficeCompoundFile? compound, out string? error), error);
        byte[] stream = compound!.Streams["WordDocument"];
        bool newerFormat = setting != "none";
        Assert.Equal(newerFormat ? 0x00D9 : 0x00C1, BitConverter.ToUInt16(stream, 2));
        Assert.Equal(newerFormat ? 0x00F0 : 0, BitConverter.ToUInt16(stream, 10) & 0x00F0);
        Assert.Equal(newerFormat ? 0x006C : 0x005D, BitConverter.ToUInt16(stream, 0x98));
        Assert.Equal(newerFormat ? 544 : 500, BitConverter.ToInt32(stream, 0x196));
        if (newerFormat) {
            Assert.Equal(2, BitConverter.ToUInt16(stream, 0x3FA));
            Assert.Equal(0x00D9, BitConverter.ToUInt16(stream, 0x3FC));
            Assert.True(BitConverter.ToInt32(stream, 0x38E) >= 36);
        }
        using WordDocument restored = WordDocument.Load(new MemoryStream(bytes));
        WordTable read = story == "body" ? restored.Tables[0]
            : (story == "header" ? (WordHeaderFooter)restored.Header.Default : restored.Footer.Default).Tables[0];
        Assert.Equal(setting == "fit", read.Rows[0].Cells[0].FitText);
        Assert.Equal(setting != "noWrap", read.Rows[0].Cells[0].WrapText);
        Assert.Equal(setting == "hide", read.Rows[0].Cells[0].HideMark);
    }
}
