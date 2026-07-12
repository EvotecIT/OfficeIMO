using System;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointStreamTests {
        [Fact]
        public void Create_ToStream_WithSaveOnDispose_WritesPackage() {
            using var stream = new MemoryStream();
            using (var presentation = PowerPointPresentation.Create(stream,
                       new PowerPointCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                presentation.AddSlide();
            }

            AssertValidPackage(stream, expectedSlides: 1);
        }

        [Fact]
        public void Create_ToStream_WithExplicitPersistence_DoesNotWriteOnDispose() {
            using var stream = new MemoryStream();

            using (var presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions())) {
                presentation.AddSlide();
            }

            Assert.Equal(0, stream.Length);
        }

        [Fact]
        public void Create_ToStream_WithExplicitPersistence_CanSaveToAssociatedStream() {
            using var stream = new MemoryStream();

            using (var presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions())) {
                presentation.AddSlide();
                presentation.Save();
            }

            AssertValidPackage(stream, expectedSlides: 1);
        }

        [Fact]
        public void Load_FromStream_WithSaveOnDispose_PersistsChanges() {
            using var stream = new MemoryStream();
            using (var presentation = PowerPointPresentation.Create(stream)) {
                presentation.AddSlide();
                presentation.Save();
            }

            using (var presentation = PowerPointPresentation.Load(stream, new PowerPointLoadOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose })) {
                presentation.AddSlide();
            }

            AssertValidPackage(stream, expectedSlides: 2);
        }

        [Fact]
        public void Load_FromNonSeekableReadStream_WorksWithExplicitPersistence() {
            byte[] bytes;
            using (var source = new MemoryStream()) {
                using (var presentation = PowerPointPresentation.Create(source)) {
                    presentation.AddSlide();
                    presentation.Save();
                }

                bytes = source.ToArray();
            }

            using var input = new NonSeekableReadStream(bytes);
            using var output = new MemoryStream();

            using (var presentation = PowerPointPresentation.Load(input)) {
                presentation.AddSlide();
                presentation.Save(output);
            }

            AssertValidPackage(output, expectedSlides: 2);
        }

        [Fact]
        public void Create_ToNonSeekableWritableStream_WithSaveOnDispose_Throws() {
            using var stream = new NonSeekableWriteStream();

            var exception = Assert.Throws<ArgumentException>(() => PowerPointPresentation.Create(stream,
                new PowerPointCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose }));
            Assert.Contains("support seeking", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Load_FromNonSeekableReadWriteStream_WithSaveOnDispose_Throws() {
            byte[] bytes;
            using (var source = new MemoryStream()) {
                using (var presentation = PowerPointPresentation.Create(source)) {
                    presentation.AddSlide();
                    presentation.Save();
                }

                bytes = source.ToArray();
            }

            using var stream = new NonSeekableReadWriteStream(bytes);

            var exception = Assert.Throws<ArgumentException>(() => PowerPointPresentation.Load(stream, new PowerPointLoadOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose }));
            Assert.Contains("support seeking", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task Load_FromNonSeekableReadWriteStream_DoesNotBecomePathlessSaveTarget() {
            byte[] bytes;
            using (var source = new MemoryStream()) {
                using var presentation = PowerPointPresentation.Create(source);
                presentation.AddSlide();
                presentation.Save();
                bytes = source.ToArray();
            }

            using var stream = new NonSeekableReadWriteStream(bytes);
            using PowerPointPresentation loaded = PowerPointPresentation.Load(stream);
            loaded.AddSlide();

            Assert.Throws<InvalidOperationException>(() => loaded.Save());
            await Assert.ThrowsAsync<InvalidOperationException>(() => loaded.SaveAsync());
            Assert.Equal(bytes, stream.ToArray());
        }

        [Fact]
        public void Create_ToNonSeekableAssociatedStream_WithExplicitPersistence_Throws() {
            using var stream = new NonSeekableWriteStream();

            ArgumentException exception = Assert.Throws<ArgumentException>(() =>
                PowerPointPresentation.Create(stream));

            Assert.Equal("stream", exception.ParamName);
            Assert.Contains("support seeking", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExportSlide_ToNonSeekableWritableStream_WritesStandalonePresentation() {
            using var source = PowerPointPresentation.Create();
            source.AddSlide().AddTitle("First");
            source.AddSlide().AddTitle("Exported");
            using var destination = new NonSeekableWriteStream();

            source.ExportSlide(1, destination);

            using var package = new MemoryStream(destination.ToArray(), writable: false);
            using PresentationDocument document = PresentationDocument.Open(package, false);
            Assert.Single(document.PresentationPart!.Presentation.SlideIdList!.ChildElements);
        }

        [Fact]
        public void Create_ToStream_WithSaveOnDispose_PropagatesPersistenceFailure() {
            using var stream = new FailingCopyBackStream();

            IOException exception = Assert.Throws<IOException>(() => {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(stream,
                    new PowerPointCreateOptions { PersistenceMode = OfficeIMO.Drawing.DocumentPersistenceMode.SaveOnDispose });
                presentation.AddSlide();
            });

            Assert.Contains("SetLength failed", exception.Message, StringComparison.Ordinal);
        }

        private static void AssertValidPackage(MemoryStream stream, int expectedSlides) {
            Assert.True(stream.Length > 0);
            stream.Position = 0;

            using var document = PresentationDocument.Open(stream, false);
            Assert.NotNull(document.PresentationPart);
            Assert.NotNull(document.PresentationPart!.Presentation);
            Assert.NotNull(document.PresentationPart.Presentation.SlideIdList);
            Assert.Equal(expectedSlides, document.PresentationPart.Presentation.SlideIdList!.ChildElements.Count);
        }

        private sealed class NonSeekableWriteStream : Stream {
            private readonly MemoryStream _inner = new();

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => _inner.Length;
            public override long Position {
                get => _inner.Position;
                set => throw new NotSupportedException();
            }

            public override void Flush() => _inner.Flush();
            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => _inner.SetLength(value);
            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

            public byte[] ToArray() => _inner.ToArray();

            protected override void Dispose(bool disposing) {
                if (disposing) {
                    _inner.Dispose();
                }

                base.Dispose(disposing);
            }
        }

        private sealed class NonSeekableReadWriteStream : Stream {
            private readonly MemoryStream _inner;

            public NonSeekableReadWriteStream(byte[] bytes) {
                _inner = new MemoryStream();
                _inner.Write(bytes, 0, bytes.Length);
                _inner.Position = 0;
            }

            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => throw new NotSupportedException();
            public override long Position {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override void Flush() => _inner.Flush();
            public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => _inner.SetLength(value);
            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

            public byte[] ToArray() => _inner.ToArray();

            protected override void Dispose(bool disposing) {
                if (disposing) {
                    _inner.Dispose();
                }

                base.Dispose(disposing);
            }
        }

        private sealed class FailingCopyBackStream : MemoryStream {
            public override void SetLength(long value) {
                throw new IOException("SetLength failed during copy-back.");
            }
        }
    }
}
