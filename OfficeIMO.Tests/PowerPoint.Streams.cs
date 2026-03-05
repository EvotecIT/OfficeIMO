using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointStreamTests {
        [Fact]
        public void Create_ToStream_WritesPackage() {
            using var stream = new MemoryStream();
            using (var presentation = PowerPointPresentation.Create(stream)) {
                presentation.AddSlide();
            }

            AssertValidPackage(stream, expectedSlides: 1);
        }

        [Fact]
        public void Create_ToStream_WithAutoSaveFalse_DoesNotWriteOnDispose() {
            using var stream = new MemoryStream();

            using (var presentation = PowerPointPresentation.Create(stream, autoSave: false)) {
                presentation.AddSlide();
            }

            Assert.Equal(0, stream.Length);
        }

        [Fact]
        public void Create_ToStream_WithAutoSaveFalse_CanBeSavedExplicitly() {
            using var stream = new MemoryStream();

            using (var presentation = PowerPointPresentation.Create(stream, autoSave: false)) {
                presentation.AddSlide();
                presentation.Save(stream);
            }

            AssertValidPackage(stream, expectedSlides: 1);
        }

        [Fact]
        public void Open_FromStream_WithAutoSaveTrue_PersistsChangesOnDispose() {
            using var stream = new MemoryStream();
            using (var presentation = PowerPointPresentation.Create(stream)) {
                presentation.AddSlide();
            }

            using (var presentation = PowerPointPresentation.Open(stream, readOnly: false, autoSave: true)) {
                presentation.AddSlide();
            }

            AssertValidPackage(stream, expectedSlides: 2);
        }

        [Fact]
        public void Open_FromNonSeekableReadStream_WorksWhenAutoSaveDisabled() {
            byte[] bytes;
            using (var source = new MemoryStream()) {
                using (var presentation = PowerPointPresentation.Create(source)) {
                    presentation.AddSlide();
                }

                bytes = source.ToArray();
            }

            using var input = new NonSeekableReadStream(bytes);
            using var output = new MemoryStream();

            using (var presentation = PowerPointPresentation.Open(input, readOnly: false, autoSave: false)) {
                presentation.AddSlide();
                presentation.Save(output);
            }

            AssertValidPackage(output, expectedSlides: 2);
        }

        [Fact]
        public void Create_ToNonSeekableWritableStream_WithAutoSaveEnabled_Throws() {
            using var stream = new NonSeekableWriteStream();

            var exception = Assert.Throws<ArgumentException>(() => PowerPointPresentation.Create(stream, autoSave: true));
            Assert.Contains("support seeking", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Open_FromNonSeekableReadWriteStream_WithAutoSaveEnabled_Throws() {
            byte[] bytes;
            using (var source = new MemoryStream()) {
                using (var presentation = PowerPointPresentation.Create(source)) {
                    presentation.AddSlide();
                }

                bytes = source.ToArray();
            }

            using var stream = new NonSeekableReadWriteStream(bytes);

            var exception = Assert.Throws<ArgumentException>(() => PowerPointPresentation.Open(stream, readOnly: false, autoSave: true));
            Assert.Contains("support seeking", exception.Message, StringComparison.OrdinalIgnoreCase);
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

            protected override void Dispose(bool disposing) {
                if (disposing) {
                    _inner.Dispose();
                }

                base.Dispose(disposing);
            }
        }
    }
}
