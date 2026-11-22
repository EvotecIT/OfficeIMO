using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioStreamLoadCopiesTests {
        private static byte[] CreateDiagram(string pageName) {
            VisioDocument document = VisioDocument.Create();
            document.AddPage(pageName);
            using var output = new MemoryStream();
            document.Save(output);
            return output.ToArray();
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task SeekableLoad_ReadsInPlaceAndDoesNotNeedTheStreamAfterLoading(bool async) {
            byte[] original = CreateDiagram("In place");
            VisioDocument document;
            using (var source = new MemoryStream(original, writable: false)) {
                document = async ? await VisioDocument.LoadAsync(source) : VisioDocument.Load(source);
                Assert.Equal(0, source.Position);
            }

            Assert.Equal("In place", Assert.Single(document.Pages).Name);
            using var saved = new MemoryStream();
            document.Save(saved);
            saved.Position = 0;
            Assert.Equal("In place", Assert.Single(VisioDocument.Load(saved).Pages).Name);
        }

        [Fact]
        public async Task SeekableLoadAsync_ReadsAsyncOnlyStream() {
            byte[] original = CreateDiagram("Slow storage");
            var gate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            using var source = new GatedReadStream(original, gate.Task);

            Task<VisioDocument> loading = VisioDocument.LoadAsync(source);
            bool returnedBeforeReading = !loading.IsCompleted;
            gate.SetResult(true);

            VisioDocument document = await loading;
            Assert.True(returnedBeforeReading);
            Assert.Equal("Slow storage", Assert.Single(document.Pages).Name);
        }

        [Fact]
        public void NonSeekableLoad_ReadsTheWholeDiagram() {
            byte[] original = CreateDiagram("Forward only");
            using var source = new NonSeekableReadStream(original);

            VisioDocument document = VisioDocument.Load(source);

            Assert.Equal("Forward only", Assert.Single(document.Pages).Name);
        }

#if NET
        [Fact]
        [Trait("Category", "Performance")]
        public void SeekableLoad_DoesNotCopyTheFile() {
            byte[] original = LargePartPackages.WithLargePart(CreateDiagram("Large"), "visio");
            using var source = new MemoryStream(original, writable: false);
            GC.Collect();
            long before = GC.GetAllocatedBytesForCurrentThread();
            VisioDocument document = VisioDocument.Load(source);
            long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

            Assert.Single(document.Pages);
            Assert.True(allocated < original.Length / 4, $"Load allocated {allocated} bytes for a {original.Length} byte file.");
        }
#endif
    }
}
