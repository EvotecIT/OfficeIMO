#if NET8_0_OR_GREATER
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.SharedSource.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void WriteRows_Utf8TextPreservesValuesAndStorage(bool sharedStrings) {
            string[] text = ["plain ASCII", "", " \tZażółć 🚀\r\n<&>\"' \u00a0", "\0 a\u0001b\t\n\r", new string('a', 30_000) + new string('λ', 2_000)];
            byte[][] input = text.Select(Encoding.UTF8.GetBytes).ToArray();
            using var output = new MemoryStream();
            var result = ExcelDocument.WriteRows(output, input, ["Text", "Number"],
                static (writer, value) => writer.WriteUtf8(value).Write(42),
                new ExcelTabularWriteOptions { UseSharedStrings = sharedStrings, RequireStreaming = true });
            Assert.Equal(text.Length, result.RowCount);
            using (var package = SpreadsheetDocument.Open(output, false)) {
                Assert.Empty(new OpenXmlValidator().Validate(package));
                Assert.Equal(sharedStrings, package.WorkbookPart!.SharedStringTablePart != null);
            }
            using var reader = ExcelDocument.OpenDataReader(output.ToArray(), new ExcelReadOptions { HasHeaderRow = true });
            for (int row = 0; row < text.Length; row++) {
                Assert.True(reader.Read());
                Assert.Equal(text[row].Replace("\0", "").Replace("\u0001", ""), reader.GetString(0));
                Assert.Equal(42, reader.GetInt32(1));
            }
            Assert.False(reader.Read());
        }

        [Theory]
        [InlineData("InvalidUtf8")]
        [InlineData("InvalidXmlScalar")]
        [InlineData("TooLong")]
        public void WriteRows_Utf8RejectsInvalidTextBeforeAdvancingCell(string kind) {
            byte[] invalid = kind switch {
                "InvalidUtf8" => [0xc3, 0x28],
                "InvalidXmlScalar" => Encoding.UTF8.GetBytes("\uFFFE"),
                _ => Encoding.UTF8.GetBytes(new string('a', 32_768))
            };
            using var output = new MemoryStream();
            ExcelDocument.WriteRows(output, new[] { invalid }, ["Text"], static (writer, value) => {
                Assert.Throws<ArgumentException>(() => writer.WriteUtf8(value));
                writer.WriteUtf8("valid"u8);
            });
            using var reader = ExcelDocument.OpenDataReader(output.ToArray(), new ExcelReadOptions { HasHeaderRow = true });
            Assert.True(reader.Read());
            Assert.Equal("valid", reader.GetString(0));
            Assert.False(reader.Read());
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void PooledUtf8TextWriter_PreservesMixedCharacterAndByteWriteOrder(bool flushBeforeBytes) {
            using var output = new MemoryStream();
            using (var writer = new PooledUtf8TextWriter(output, new UTF8Encoding(false), 16, leaveOpen: true)) {
                writer.Write("<prefix>\ud83d");
                if (flushBeforeBytes) writer.Flush();
                writer.WriteUtf8(" Zażółć 🚀 "u8);
                writer.Write(new string('a', 4096));
                writer.WriteUtf8(Encoding.UTF8.GetBytes(new string('λ', 20_000)));
                writer.Write("</suffix>");
            }
            byte[] expected = Encoding.UTF8.GetBytes("<prefix>\ufffd Zażółć 🚀 " + new string('a', 4096) + new string('λ', 20_000) + "</suffix>");
            Assert.Equal(expected, output.ToArray());
            Assert.True(output.CanWrite);
        }
    }
}
#endif
