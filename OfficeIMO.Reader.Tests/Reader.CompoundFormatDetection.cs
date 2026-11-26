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
        public async Task NativeCompoundPayloadsRetainDistinctDetectionAndRejectAmbiguity(
            bool publisher, bool visio, ReaderInputKind expected) {
            var streams = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            var entries = new List<OfficeCompoundFileEntry>();
            byte[] marker = "%PDF-1.7"u8.ToArray();
            if (publisher) {
                streams.Add("Contents", marker);
                streams.Add("Quill/QuillSub/CONTENTS", new byte[] { 2 });
                entries.Add(new OfficeCompoundFileEntry("Contents", "Contents", 2, marker.Length));
                entries.Add(new OfficeCompoundFileEntry("Quill", "Quill", 1, 0));
                entries.Add(new OfficeCompoundFileEntry("QuillSub", "Quill/QuillSub", 1, 0));
                entries.Add(new OfficeCompoundFileEntry("CONTENTS", "Quill/QuillSub/CONTENTS", 2, 1));
            }
            if (visio) {
                streams.Add("VisioDocument", marker);
                entries.Add(new OfficeCompoundFileEntry("VisioDocument", "VisioDocument", 2, marker.Length));
            }
            var source = new OfficeCompoundFile(streams, entries,
                new OfficeCompoundFileEntry("Root Entry", "Root Entry", 5, 0));
            byte[] bytes = OfficeCompoundFileWriter.Rewrite(source, new Dictionary<string, byte[]>());
            Assert.True(bytes.AsSpan(0, Math.Min(bytes.Length, 1024)).IndexOf("%PDF-"u8) >= 0);
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddPublisherHandler().AddVisioHandler().Build();

            var options = new ReaderDetectionOptions { Mode = ReaderDetectionMode.PreferContent, InspectContainers = true };
            ReaderDetectionResult detection = reader.Detect(bytes, "renamed.bin", options);
            ReaderDetectionResult detectionAsync = await reader.DetectAsync(bytes, "renamed.bin", options);

            Assert.All(new[] { detection, detectionAsync }, result => {
                Assert.Equal(expected, result.ContentKind);
                Assert.Equal(expected, result.Kind);
                if (expected != ReaderInputKind.Unknown) {
                    Assert.Equal(ReaderDetectionConfidence.High, result.ContentConfidence);
                }
            });
        }
    }
}
