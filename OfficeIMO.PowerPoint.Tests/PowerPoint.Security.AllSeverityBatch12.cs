using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointAllSeverityBatch12SecurityTests {
    [Fact]
    public async Task LoadApisEnforceInputBudgetForSeekableAndNonSeekableStreams() {
        byte[] package;
        using (PowerPointPresentation source =
               PowerPointPresentation.Create()) {
            source.AddSlide().AddTextBox("safe");
            package = source.ToBytes();
        }
        var restrictive = new PowerPointLoadOptions {
            MaxInputBytes = package.Length - 1L
        };

        using var seekable = new MemoryStream(package, writable: false);
        Assert.Throws<InvalidDataException>(() =>
            PowerPointPresentation.Load(seekable, restrictive));

        using var nonSeekable = new Batch12NonSeekableStream(package);
        Assert.Throws<InvalidDataException>(() =>
            PowerPointPresentation.Load(nonSeekable, restrictive));

        using var asyncStream = new Batch12NonSeekableStream(package);
        await Assert.ThrowsAsync<InvalidDataException>(() =>
            PowerPointPresentation.LoadAsync(asyncStream, restrictive));

        using var validStream = new MemoryStream(package, writable: false);
        using PowerPointPresentation valid = PowerPointPresentation.Load(
            validStream,
            new PowerPointLoadOptions { MaxInputBytes = package.Length });
        Assert.Single(valid.Slides);
    }

    [Fact]
    public void DuplicateSlideRejectsCyclicPartRelationships() {
        using PowerPointPresentation presentation =
            PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        ExtendedPart first = slide.SlidePart.AddExtendedPart(
            "urn:officeimo:test:first",
            "application/xml",
            "rIdCycleFirst");
        ExtendedPart second = first.AddExtendedPart(
            "urn:officeimo:test:second",
            "application/xml",
            "rIdCycleSecond");
        WritePart(first, "<first />");
        WritePart(second, "<second />");
        second.AddPart(first, "rIdCycleBack");
        int originalSlideCount = presentation.Slides.Count;
        int originalTopLevelPartCount = presentation.OpenXmlDocument
            .PresentationPart!.Parts.Count();

        InvalidDataException exception = Assert.Throws<InvalidDataException>(
            () => presentation.DuplicateSlide(0));

        Assert.Contains("contain a cycle", exception.Message,
            StringComparison.Ordinal);
        Assert.Equal(originalSlideCount, presentation.Slides.Count);
        Assert.Equal(originalTopLevelPartCount, presentation.OpenXmlDocument
            .PresentationPart!.Parts.Count());
    }

    [Fact]
    public void ConversionPreservesConfiguredInputBudget() {
        string root = Path.Combine(Path.GetTempPath(),
            "officeimo-powerpoint-conversion-" + Guid.NewGuid().ToString("N"));
        string sourcePath = Path.Combine(root, "source.pptx");
        string destinationPath = Path.Combine(root, "destination.pptx");
        Directory.CreateDirectory(root);
        try {
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Create(sourcePath)) {
                presentation.AddSlide().AddTextBox("safe");
                presentation.Save();
            }
            long sourceLength = new FileInfo(sourcePath).Length;
            var options = new PowerPointPresentationConversionOptions {
                LoadOptions = new PowerPointLoadOptions {
                    MaxInputBytes = sourceLength - 1L
                }
            };

            Assert.Throws<InvalidDataException>(() =>
                PowerPointPresentation.AnalyzeConversion(
                    sourcePath, destinationPath, options));
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, true);
        }
    }

    private static void WritePart(OpenXmlPart part, string content) {
        using Stream stream = part.GetStream(FileMode.Create,
            FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(content);
    }

    private sealed class Batch12NonSeekableStream : Stream {
        private readonly MemoryStream _inner;

        internal Batch12NonSeekableStream(byte[] bytes) {
            _inner = new MemoryStream(bytes, writable: false);
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) =>
            _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) =>
            throw new NotSupportedException();
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
