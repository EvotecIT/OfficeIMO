using OfficeIMO.Core.Internal;
using OfficeIMO.Reader.Publisher;
using OfficeIMO.Reader.Visio;
using Xunit;

namespace OfficeIMO.Reader.Tests {
    public sealed class ReaderCompoundFormatDetectionTests {
        [Theory]
        [InlineData(true, false, ReaderInputKind.Publisher)]
        [InlineData(false, true, ReaderInputKind.Visio)]
        [InlineData(true, true, ReaderInputKind.Unknown)]
        public void NativeCompoundPayloadsRetainDistinctDetectionAndRejectAmbiguity(
            bool publisher, bool visio, ReaderInputKind expected) {
            var streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            var entries = new List<OfficeCompoundFileEntry>();
            if (publisher) {
                streams.Add("Contents", new byte[] { 1 });
                streams.Add("Quill/QuillSub/CONTENTS", new byte[] { 2 });
                entries.Add(new OfficeCompoundFileEntry("Contents", "Contents", 2, 1));
                entries.Add(new OfficeCompoundFileEntry("Quill", "Quill", 1, 0));
                entries.Add(new OfficeCompoundFileEntry("QuillSub", "Quill/QuillSub", 1, 0));
                entries.Add(new OfficeCompoundFileEntry("CONTENTS", "Quill/QuillSub/CONTENTS", 2, 1));
            }
            if (visio) {
                streams.Add("VisioDocument", new byte[] { 3 });
                entries.Add(new OfficeCompoundFileEntry("VisioDocument", "VisioDocument", 2, 1));
            }
            var source = new OfficeCompoundFile(streams, entries,
                new OfficeCompoundFileEntry("Root Entry", "Root Entry", 5, 0));
            byte[] bytes = OfficeCompoundFileWriter.Rewrite(source, new Dictionary<string, byte[]>());
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddPublisherHandler().AddVisioHandler().Build();

            ReaderDetectionResult detection = reader.Detect(bytes, "renamed.bin",
                new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent });

            Assert.Equal(expected, detection.ContentKind);
            Assert.Equal(expected, detection.Kind);
            if (expected != ReaderInputKind.Unknown) {
                Assert.Equal(ReaderDetectionConfidence.High, detection.ContentConfidence);
            }
        }
    }
}
