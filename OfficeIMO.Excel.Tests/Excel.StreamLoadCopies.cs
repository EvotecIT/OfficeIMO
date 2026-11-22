using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
        private static byte[] CreateStreamLoadWorkbook(string value)
        {
            using var document = ExcelDocument.Create();
            document.AddSheet("Data").Cell(1, 1).SetValue(value);
            using var output = new MemoryStream();
            document.Save(output);
            return output.ToArray();
        }

        [Theory]
        [InlineData(DocumentAccessMode.ReadOnly)]
        [InlineData(DocumentAccessMode.ReadWrite)]
        public void StreamLoad_DoesNotShareTheCallersBuffer(DocumentAccessMode mode)
        {
            byte[] callerBuffer = CreateStreamLoadWorkbook("Detached");
            using var source = new MemoryStream(callerBuffer, writable: false);
            using var document = ExcelDocument.Load(source, new ExcelLoadOptions { AccessMode = mode });

            Array.Clear(callerBuffer, 0, callerBuffer.Length);

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string text));
            Assert.Equal("Detached", text);
        }

#if NET
        [Fact]
        [Trait("Category", "Performance")]
        public void StreamLoad_AllocatesOneCopyOfTheFile()
        {
            byte[] original = LargePartPackages.WithLargePart(CreateStreamLoadWorkbook("Large"), "xl");

            foreach (DocumentAccessMode mode in new[] { DocumentAccessMode.ReadOnly, DocumentAccessMode.ReadWrite })
            {
                using var source = new MemoryStream(original, writable: false);
                GC.Collect();
                long before = GC.GetAllocatedBytesForCurrentThread();
                using var document = ExcelDocument.Load(source, new ExcelLoadOptions { AccessMode = mode });
                long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
                Assert.True(allocated < original.Length * 3L / 2, $"{mode} load allocated {allocated} bytes for a {original.Length} byte file.");
            }
        }
#endif
    }
}
