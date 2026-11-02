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
                using (PowerPointDocument presentation = PowerPointDocument.Create(path)) {
                    presentation.AddSlide().AddTitle("Detached");
                    Assert.False(File.Exists(path));
                }
                Assert.False(File.Exists(path));

                using (PowerPointDocument presentation = PowerPointDocument.Create(path)) {
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
            using PowerPointDocument presentation = PowerPointDocument.Create();
            presentation.AddSlide();
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => presentation.Save());
            Assert.Contains("no associated destination", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Load_Stream_PreservesCallerPositionAndDoesNotWriteByDefault() {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                presentation.AddSlide().AddTitle("Original");
                presentation.Save();
            }
            byte[] original = stream.ToArray();
            stream.Position = Math.Min(5, stream.Length);
            long originalPosition = stream.Position;

            using (PowerPointDocument presentation = PowerPointDocument.Load(stream,
                       new PowerPointLoadOptions {
                           OpenSettings = new OfficeOpenXmlLoadSettings()
                       })) {
                presentation.Slides[0].AddTextBox("Unsaved");
                Assert.Equal(originalPosition, stream.Position);
            }

            Assert.Equal(originalPosition, stream.Position);
            Assert.Equal(original, stream.ToArray());
        }

        [Fact]
        public async Task ReadOnlySnapshotsDoNotEnableSavingToFilesOrStreams() {
            const string password = "synthetic-readonly-snapshot-password";
            using var stream = new MemoryStream();
            byte[] encrypted;
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                presentation.AddSlide();
                presentation.Save();
                encrypted = presentation.ToBytes(new PowerPointSaveOptions { Password = password });
            }

            Assert.Throws<ArgumentException>(() => PowerPointDocument.Load(stream,
                new PowerPointLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly,
                    PersistenceMode = DocumentPersistenceMode.SaveOnDispose
                }));

            using var encryptedSource = new MemoryStream(encrypted);
            byte[] original = encryptedSource.ToArray();
            using var readOnly = PowerPointDocument.Load(encryptedSource, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly, Password = password });
            using var captured = new MemoryStream(readOnly.ToBytes());
            using var reopened = PowerPointDocument.Load(captured, new PowerPointLoadOptions { Password = password });
            Assert.Single(reopened.Slides);
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".ppt");
            try {
                Assert.Throws<InvalidOperationException>(() => readOnly.Save(path));
                Assert.Throws<InvalidOperationException>(() => readOnly.SaveCopy(path));
                Assert.False(File.Exists(path));
                using var destination = new MemoryStream(new byte[] { 1, 2, 3 });
                Assert.Throws<InvalidOperationException>(() => readOnly.Save(destination));
                await Assert.ThrowsAsync<InvalidOperationException>(() => readOnly.SaveAsync(destination));
                Assert.Equal(new byte[] { 1, 2, 3 }, destination.ToArray());
                Assert.Equal(original, encryptedSource.ToArray());
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        [Fact]
        public async Task SaveAsync_PathAndStreamProduceEquivalentReloadablePackages() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                await using PowerPointDocument presentation = PowerPointDocument.Create();
                presentation.AddSlide().AddTitle("Async");
                byte[] bytes = presentation.ToBytes();
                Assert.NotEmpty(bytes);

                using var stream = new MemoryStream();
                await presentation.SaveAsync(stream);
                await presentation.SaveAsync(path);

                using PowerPointDocument streamLoaded = PowerPointDocument.Load(stream,
                    new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
                using PowerPointDocument loaded = PowerPointDocument.Load(path,
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
                await using PowerPointDocument presentation = PowerPointDocument.Create();
                presentation.AddSlide();
                using var cancellation = new CancellationTokenSource();
                cancellation.Cancel();

                await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                    presentation.SaveAsync(path, cancellationToken: cancellation.Token));
                Assert.False(File.Exists(path));
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Load_Stream_HonorsCancellationDuringInputCopy() {
            byte[] bytes;
            using (PowerPointDocument source = PowerPointDocument.Create()) {
                source.AddSlide().AddTitle("Cancellation boundary");
                bytes = source.ToBytes();
            }

            using var cancellation = new CancellationTokenSource();
            using var stream = new CancelOnBulkReadStream(bytes,
                cancellation.Cancel);

            Assert.ThrowsAny<OperationCanceledException>(() =>
                PowerPointDocument.Load(stream,
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
                // The format probe reads only the 8-byte package signature. Cancel on
                // the subsequent payload read without coupling the contract test to a
                // particular copy-buffer size or allocation strategy.
                if (!BulkReadObserved && count > 8) {
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
