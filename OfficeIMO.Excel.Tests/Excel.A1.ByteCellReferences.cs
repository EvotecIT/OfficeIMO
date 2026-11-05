using System.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelA1ByteCellReferenceTests {
        [Theory]
        [InlineData("A1")]
        [InlineData("ab12")]
        [InlineData(" AB12 ")]
        [InlineData("\tB7")]
        [InlineData("$A$1")]
        [InlineData("A01")]
        [InlineData("A0")]
        [InlineData("1A")]
        [InlineData("A")]
        [InlineData("")]
        [InlineData("AB12X")]
        [InlineData("A2147483648")]
        [InlineData("ZZZZZZZ1")]
        [InlineData("XFD1048576")]
        [InlineData("XFE1")]
        [InlineData("XFD1048577")]
        public void TryParseCellReferenceFast_SingleByteTextMatchesCharacterText(string reference) {
            byte[] bytes = reference.Select(character => checked((byte)character)).ToArray();

            bool expected = ExcelA1.TryParseCellReferenceFast(reference.AsSpan(), out int expectedRow, out int expectedColumn);
            bool actual = ExcelA1.TryParseCellReferenceFast(bytes, out int row, out int column);

            Assert.Equal((expected, expectedRow, expectedColumn), (actual, row, column));
        }
    }
}
