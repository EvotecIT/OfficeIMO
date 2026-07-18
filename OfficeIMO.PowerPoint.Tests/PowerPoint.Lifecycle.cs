using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLifecycleTests {
        [Fact]
        public void Create_Path_IsDetachedUntilExplicitSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTitle("Detached");
                    Assert.False(File.Exists(path));
                }
                Assert.False(File.Exists(path));

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTitle("Saved");
                    presentation.Save();
                }

                Assert.True(File.Exists(path));
                using PresentationDocument package = PresentationDocument.Open(path, false);
                Assert.Single(package.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>());
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Create_Detached_SaveWithoutDestinationFailsClearly() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide();
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => presentation.Save());
            Assert.Contains("no associated destination", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Load_Stream_PreservesCallerPositionAndDoesNotWriteByDefault() {
            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                presentation.AddSlide().AddTitle("Original");
                presentation.Save();
            }
            byte[] original = stream.ToArray();
            stream.Position = Math.Min(5, stream.Length);
            long originalPosition = stream.Position;

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(stream,
                       new PowerPointLoadOptions {
                           OpenSettings = new OpenSettings { AutoSave = true }
                       })) {
                presentation.Slides[0].AddTextBox("Unsaved");
                Assert.Equal(originalPosition, stream.Position);
            }

            Assert.Equal(originalPosition, stream.Position);
            Assert.Equal(original, stream.ToArray());
        }

        [Fact]
        public void Load_ReadOnlyRejectsSaveOnDispose() {
            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                presentation.AddSlide();
                presentation.Save();
            }

            Assert.Throws<ArgumentException>(() => PowerPointPresentation.Load(stream,
                new PowerPointLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly,
                    PersistenceMode = DocumentPersistenceMode.SaveOnDispose
                }));
        }

        [Fact]
        public async Task SaveAsync_PathAndStreamProduceEquivalentReloadablePackages() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                await using PowerPointPresentation presentation = PowerPointPresentation.Create();
                presentation.AddSlide().AddTitle("Async");
                byte[] bytes = presentation.ToBytes();
                Assert.NotEmpty(bytes);

                using var stream = new MemoryStream();
                await presentation.SaveAsync(stream);
                await presentation.SaveAsync(path);

                using PowerPointPresentation streamLoaded = PowerPointPresentation.Load(stream,
                    new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
                using PowerPointPresentation loaded = PowerPointPresentation.Load(path,
                    new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
                Assert.Equal("Async", streamLoaded.Slides[0].TextBoxes.First().Text);
                Assert.Equal("Async", loaded.Slides[0].TextBoxes.First().Text);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task SaveAsync_CanceledPathDoesNotCreateDestination() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                await using PowerPointPresentation presentation = PowerPointPresentation.Create();
                presentation.AddSlide();
                using var cancellation = new CancellationTokenSource();
                cancellation.Cancel();

                await Assert.ThrowsAsync<OperationCanceledException>(() =>
                    presentation.SaveAsync(path, cancellationToken: cancellation.Token));
                Assert.False(File.Exists(path));
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Load_Stream_HonorsCancellationDuringInputCopy() {
            byte[] bytes;
            using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                source.AddSlide().AddTitle("Cancellation boundary");
                bytes = source.ToBytes();
            }

            using var cancellation = new CancellationTokenSource();
            using var stream = new CancelOnBulkReadStream(bytes,
                cancellation.Cancel);

            Assert.ThrowsAny<OperationCanceledException>(() =>
                PowerPointPresentation.Load(stream,
                    new PowerPointLoadOptions(), cancellation.Token));
            Assert.True(stream.BulkReadObserved);
        }

        private sealed class CancelOnBulkReadStream : Stream {
            private readonly MemoryStream _inner;
            private readonly Action _cancel;

            internal CancelOnBulkReadStream(byte[] bytes, Action cancel) {
                _inner = new MemoryStream(bytes, writable: false);
                _cancel = cancel;
            }

            internal bool BulkReadObserved { get; private set; }
            public override bool CanRead => true;
            public override bool CanSeek => true;
            public override bool CanWrite => false;
            public override long Length => _inner.Length;
            public override long Position {
                get => _inner.Position;
                set => _inner.Position = value;
            }

            public override int Read(byte[] buffer, int offset, int count) {
                int read = _inner.Read(buffer, offset, count);
                if (!BulkReadObserved && count >= 81920) {
                    BulkReadObserved = true;
                    _cancel();
                }
                return read;
            }

            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) =>
                _inner.Seek(offset, origin);
            public override void SetLength(long value) =>
                throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) =>
                throw new NotSupportedException();

            protected override void Dispose(bool disposing) {
                if (disposing) _inner.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}
