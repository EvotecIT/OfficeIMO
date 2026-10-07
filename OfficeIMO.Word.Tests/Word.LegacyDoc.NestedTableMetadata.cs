using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LegacyDoc_NestedTableIncludesRequiredRevisionThreadingMetadata(bool withRevisions) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("BEFORE");
        if (withRevisions) {
            paragraph._paragraph.Append(new InsertedRun(new Run(new Text("FIRST"))) { Id = "1", Author = "First author" });
            paragraph._paragraph.Append(new InsertedRun(new Run(new Text("SECOND"))) { Id = "2", Author = "Second author" });
            paragraph._paragraph.Append(new InsertedRun(new Run(new Text("THIRD"))) { Id = "3", Author = "First author" });
        }
        WordTable outer = source.AddTable(1, 1, WordTableStyle.TableNormal);
        outer.Rows[0].Cells[0].AddTable(1, 1, WordTableStyle.TableNormal).Rows[0].Cells[0].Paragraphs[0].Text = "NESTED";
        byte[] bytes = source.ToBytes(WordFileFormat.Doc);

        for (int cycle = 0; cycle < 2; cycle++) {
            byte[] fib = ReadCompoundStream(bytes, "WordDocument");
            byte[] table = ReadCompoundStream(bytes, "1Table");
            Assert.Equal(0x00D9, BitConverter.ToUInt16(fib, 2));
            Assert.Equal(0, BitConverter.ToUInt16(fib, 0x3FE));

            int authorCount = 0;
            if (BitConverter.ToInt32(fib, 0x236) > 0) {
                int authorsOffset = BitConverter.ToInt32(fib, 0x232);
                authorCount = BitConverter.ToUInt16(table, authorsOffset + 2);
            }
            Assert.Equal(withRevisions ? 3 : 0, authorCount);

            int threadingOffset = BitConverter.ToInt32(fib, 0x38A);
            int threadingLength = BitConverter.ToInt32(fib, 0x38E);
            Assert.True(threadingOffset > 0);
            Assert.Equal(36 + 12 * authorCount, threadingLength);
            Assert.InRange(threadingOffset + threadingLength, 1, table.Length);
            int cursor = threadingOffset;
            int[] counts = { authorCount, authorCount, 0, 0, 0, 0 };
            int[] extraLengths = { 8, 0, 2, 0, 2, 0 };
            for (int sttb = 0; sttb < counts.Length; sttb++) {
                Assert.Equal(0xFFFF, BitConverter.ToUInt16(table, cursor));
                Assert.Equal(counts[sttb], BitConverter.ToUInt16(table, cursor + 2));
                Assert.Equal(extraLengths[sttb], BitConverter.ToUInt16(table, cursor + 4));
                cursor += 6;
                for (int author = 0; author < counts[sttb]; author++) {
                    Assert.Equal(0, BitConverter.ToUInt16(table, cursor));
                    cursor += 2 + extraLengths[sttb];
                }
            }
            Assert.Equal(threadingOffset + threadingLength, cursor);
            using WordDocument restored = WordDocument.Load(new MemoryStream(bytes));
            Assert.Equal("NESTED", Assert.Single(Assert.Single(restored.Tables).Rows[0].Cells[0].NestedTables).Rows[0].Cells[0].Paragraphs[0].Text);
            bytes = restored.ToBytes(WordFileFormat.Doc);
        }
    }
}
