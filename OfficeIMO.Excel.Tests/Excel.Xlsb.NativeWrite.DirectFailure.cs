using OfficeIMO.Excel.Xlsb.Write;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false, false)]
        [InlineData(true, true)]
        public void Xlsb_DirectRecordWriter_DoesNotReplayFailedFlushOnDisposal(
            bool implicitFlush, bool rejectSecondWrite) {
            using var destination = new PartiallyFailingXlsbRecordStream(rejectSecondWrite);

            IOException failure = Assert.Throws<IOException>(() => {
                using var writer = new XlsbDirectRecordWriter(destination);
                if (implicitFlush) {
                    for (int index = 0; index < 2_000; index++) {
                        writer.WriteNumberCell(2, index, index);
                    }
                } else {
                    writer.WriteNumberCell(2, 0, 42);
                    writer.Flush();
                }
            });

            Assert.Same(destination.InitialFailure, failure);
            Assert.Equal(1, destination.WriteAttempts);
            Assert.Equal(8, destination.Length);
        }

        private sealed class PartiallyFailingXlsbRecordStream : MemoryStream {
            private readonly bool _rejectSecondWrite;

            internal PartiallyFailingXlsbRecordStream(bool rejectSecondWrite) {
                _rejectSecondWrite = rejectSecondWrite;
            }

            internal IOException InitialFailure { get; } = new IOException("Initial partial write failure.");
            internal int WriteAttempts { get; private set; }

            public override void Write(byte[] buffer, int offset, int count) {
                WriteAttempts++;
                if (WriteAttempts == 1) {
                    base.Write(buffer, offset, Math.Min(8, count));
                    throw InitialFailure;
                }
                if (_rejectSecondWrite) throw new IOException("Unexpected retry during disposal.");
                base.Write(buffer, offset, count);
            }
        }
    }
}
