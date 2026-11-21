using OfficeIMO.Core.Internal;
using OfficeIMO.Excel.LegacyXls.Write;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// The in-memory XLS writers hold the workbook stream and the finished compound file in single byte arrays.
    /// Reaching that limit must report the BIFF8 size limit before anything near 2 GiB is allocated, so these tests
    /// drive each buffer to its last position instead of filling it.
    /// </summary>
    public class ExcelLegacyXlsInMemorySizeLimitTests {
        [Fact]
        public void WorkbookStreamBufferReportsTheBiff8LimitAtTheEndOfAByteArray() {
            using var buffer = new LegacyXlsWriter.WorkbookStreamBuffer(int.MaxValue);
            buffer.Position = int.MaxValue;

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() => buffer.WriteByte(1));

            Assert.Contains("2 GiB BIFF8 limit", exception.Message, StringComparison.Ordinal);
            Assert.Equal(0, buffer.Length);
        }

        [Fact]
        public void WorkbookStreamBufferKeepsAConfiguredOutputLimit() {
            using var buffer = new LegacyXlsWriter.WorkbookStreamBuffer(16);

            InvalidDataException exception = Assert.Throws<InvalidDataException>(() => buffer.Write(new byte[17], 0, 17));

            Assert.True(OfficeOutputLimit.Is(exception));
        }

        [Fact]
        public void DirectTabularStreamReportsTheBiff8LimitAtTheEndOfAByteArray() {
            using var stream = new LegacyXlsWriter.DirectTabularWorkbookStream(64);
            stream.Position = int.MaxValue;

            Assert.Contains("2 GiB BIFF8 limit",
                Assert.Throws<NotSupportedException>(() => stream.WriteByte(1)).Message, StringComparison.Ordinal);
            Assert.Contains("2 GiB BIFF8 limit",
                Assert.Throws<NotSupportedException>(() => stream.Write(new byte[2], 0, 2)).Message, StringComparison.Ordinal);
        }

        [Fact]
        public void OutputLimitsFromXlsSubWritersReportTheBiff8LimitUnlessTheCallerSetALimit() {
            // Sub-writers such as the shared-string table raise the shared output-limit error when they run out of room.
            static byte[] SubWriterOutOfRoom() => throw OfficeOutputLimit.Create("Native XLS shared strings exceed the in-memory save limit.");

            NotSupportedException biff8 = Assert.Throws<NotSupportedException>(() =>
                LegacyXlsWriter.WithinWorkbookSizeLimit(int.MaxValue, SubWriterOutOfRoom));
            Assert.Contains("2 GiB BIFF8 limit", biff8.Message, StringComparison.Ordinal);
            Assert.True(OfficeOutputLimit.Is(biff8.InnerException!));

            InvalidDataException configured = Assert.Throws<InvalidDataException>(() =>
                LegacyXlsWriter.WithinWorkbookSizeLimit(1024 * 1024, SubWriterOutOfRoom));
            Assert.True(OfficeOutputLimit.Is(configured));
        }

        [Fact]
        public void InMemoryCompoundFileRejectsAWorkbookThatCannotFitOneByteArray() {
            // The stream is within the 2 GiB compound limit, but the file around it is not within one byte array.
            var streams = new[] {
                new OfficeCompoundStream("Workbook", int.MaxValue - 1024L,
                    () => throw new InvalidOperationException("The size check must run before any content is read."))
            };

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() =>
                LegacyXlsWriter.WriteInMemoryCompoundFile(streams));

            Assert.Contains("2 GiB BIFF8 limit", exception.Message, StringComparison.Ordinal);
        }
    }
}
