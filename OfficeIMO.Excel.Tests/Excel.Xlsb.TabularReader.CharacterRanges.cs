using OfficeIMO.Excel.Xlsb.Biff12;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Fact]
    public void XlsbTabularReader_GetCharsCopiesUtf16RangesWithoutChangingNeighbors() {
        const string text = "Ażółw🐢<&>\r\nZ";
        using var worksheetPart = CreateTabularWorksheet((0, 0U));
        using var reader = CreateTabularReader(
            worksheetPart, new[] { text }, hasHeaderRow: false, new XlsbCellReadBudget(10));
        Assert.True(reader.Read());

        var actual = new char[text.Length];
        var buffer = new char[7];
        int offset = 0;
        while (offset < text.Length) {
            for (int index = 0; index < buffer.Length; index++) buffer[index] = '!';
            int count = (int)reader.GetChars(0, offset, buffer, 2, 3);
            Assert.Equal(Math.Min(3, text.Length - offset), count);
            Assert.Equal('!', buffer[0]);
            Assert.Equal('!', buffer[1]);
            for (int index = 2 + count; index < buffer.Length; index++) Assert.Equal('!', buffer[index]);
            Array.Copy(buffer, 2, actual, offset, count);
            offset += count;
        }

        Assert.Equal(text, new string(actual));
        Assert.Equal(0, reader.GetChars(0, text.Length, buffer, 2, 3));
        Assert.Equal(0, reader.GetChars(0, 1, buffer, 2, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetChars(0, -1, buffer, 0, 1));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetChars(0, text.Length + 1, null, 0, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetChars(0, long.MaxValue, buffer, 0, 1));
        Assert.ThrowsAny<ArgumentException>(() => reader.GetChars(0, 0, buffer, -1, 1));
        Assert.ThrowsAny<ArgumentException>(() => reader.GetChars(0, 0, buffer, 0, -1));
        Assert.ThrowsAny<ArgumentException>(() => reader.GetChars(0, 0, buffer, buffer.Length, 1));
    }
}
