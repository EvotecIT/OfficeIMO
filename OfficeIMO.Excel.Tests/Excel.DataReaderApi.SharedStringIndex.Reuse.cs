using System.Text;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OpenDataReader_IndexReuseKeepsOverlappingWorkbookFieldsAndDisposalIndependent(bool stored) {
        const int seedFillers = 12_288, firstFillers = 8_192, secondFillers = 9_216;
        string? seedPath = null, firstPath = null, secondPath = null;
        try {
            seedPath = CreateIndexedSharedStringWorkbook("<si><t>seed</t></si>",
                includeLastRow: true, stored: stored, lastIndex: seedFillers + 2, fillerCount: seedFillers);
            firstPath = CreateIndexedSharedStringWorkbook("<si><t xml:space=\"preserve\"> first tail </t></si>",
                includeLastRow: true, stored: stored, lastIndex: firstFillers + 2, fillerCount: firstFillers);
            secondPath = CreateIndexedSharedStringWorkbook("<si><t>second longer tail</t></si>",
                includeLastRow: true, stored: stored, lastIndex: secondFillers + 2, fillerCount: secondFillers);
            using ExcelWorkbookDataReader seed = ExcelDocument.OpenDataReader(seedPath);
            AssertIndexedWorkbookFields(seed, "seed");
            seed.Dispose();
            Assert.True(seed.IsClosed);
            File.Delete(seedPath);
            Assert.False(File.Exists(seedPath));

            using ExcelWorkbookDataReader first = ExcelDocument.OpenDataReader(firstPath);
            Assert.Equal("Header", first.GetName(0));
            Assert.True(first.Read());
            Assert.Equal("Ready", first.GetString(0));
            Assert.Equal("Ready", Assert.IsType<string>(first.GetValue(0)));
#if NET8_0_OR_GREATER
            Assert.True(first.TryGetUtf8Text(0, out ReadOnlySpan<byte> borrowedFirst));
#endif
            seed.Dispose();
            using (ExcelWorkbookDataReader second = ExcelDocument.OpenDataReader(secondPath)) {
                AssertIndexedWorkbookFields(second, "second longer tail");
                second.Dispose();
                second.Dispose();
                Assert.True(second.IsClosed);
            }
#if NET8_0_OR_GREATER
            Assert.True(borrowedFirst.SequenceEqual(Encoding.UTF8.GetBytes("Ready")));
#endif
            // The first tail has not been materialized when the overlapping owner closes.
            Assert.True(first.Read());
            Assert.Equal(" first tail ", first.GetString(0));
            Assert.Equal(" first tail ", Assert.IsType<string>(first.GetValue(0)));
#if NET8_0_OR_GREATER
            Assert.True(first.TryGetUtf8Text(0, out ReadOnlySpan<byte> firstTail));
            Assert.True(firstTail.SequenceEqual(Encoding.UTF8.GetBytes(" first tail ")));
#endif
            Assert.False(first.Read());
            first.Dispose();
            first.Dispose();
            Assert.True(first.IsClosed);
            File.Delete(firstPath);
            File.Delete(secondPath);
            Assert.False(File.Exists(firstPath));
            Assert.False(File.Exists(secondPath));
        } finally {
            if (seedPath != null) File.Delete(seedPath);
            if (firstPath != null) File.Delete(firstPath);
            if (secondPath != null) File.Delete(secondPath);
        }
    }

    private static void AssertIndexedWorkbookFields(ExcelWorkbookDataReader reader, string tail) {
        Assert.Equal("Header", reader.GetName(0));
        foreach (string expected in new[] { "Ready", tail }) {
            Assert.True(reader.Read());
            Assert.Equal(1, reader.FieldCount);
            Assert.Equal(expected, Assert.IsType<string>(reader.GetValue(0)));
            Assert.Equal(expected, reader.GetString(0));
#if NET8_0_OR_GREATER
            Assert.True(reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> text));
            Assert.True(text.SequenceEqual(Encoding.UTF8.GetBytes(expected)));
#endif
        }
        Assert.False(reader.Read());
        Assert.False(reader.NextResult());
    }
}
