using System.IO.Compression;
using OfficeIMO.PowerPoint;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointStreamLoadCopiesTests {
        private static byte[] CreatePresentation(string title, byte[]? picture = null) {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTitle(title);
            if (picture != null) {
                slide.AddImage(new MemoryStream(picture), OfficeIMO.Drawing.OfficeImageFormat.Png);
            }
            return presentation.ToBytes();
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task ReadOnlySeekableLoad_ReadsCallerStreamInPlaceAndLeavesItOpenAndUnchanged(bool async) {
            byte[] original = CreatePresentation("In place");
            using var source = new MemoryStream();
            source.Write(original, 0, original.Length);
            source.Position = 7;
            var options = new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly };

            PowerPointDocument presentation = async
                ? await PowerPointDocument.LoadAsync(source, options)
                : PowerPointDocument.Load(source, options);
            using (presentation) {
                Assert.Equal(7, source.Position);
                Assert.Equal("In place", presentation.Slides[0].TextBoxes.First().Text);
                Assert.Throws<InvalidOperationException>(() => presentation.Save(source));
                Assert.Throws<InvalidOperationException>(() => presentation.Save());
            }

            Assert.True(source.CanRead);
            Assert.Equal(original, source.ToArray());
        }

        [Fact]
        public void ReadOnlyLoad_ReportsAnUnreferencedSignaturePart() {
            byte[] original = WithUnreferencedSignaturePart(CreatePresentation("Signed elsewhere"));
            var options = new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly };
            using var seekable = new MemoryStream(original, writable: false);
            using var forwardOnly = new NonSeekableReadStream(original);
            using PowerPointDocument inPlace = PowerPointDocument.Load(seekable, options);
            using PowerPointDocument copied = PowerPointDocument.Load(forwardOnly, options);

            PowerPointSignatureReport inPlaceReport = inPlace.InspectSignatures();
            PowerPointSignatureReport copiedReport = copied.InspectSignatures();

            Assert.Equal(1, inPlaceReport.XmlSignaturePartCount);
            Assert.Equal(1, copiedReport.XmlSignaturePartCount);
            Assert.True(inPlaceReport.HasSignatureMetadata);
            Assert.True(copiedReport.HasSignatureMetadata);
            Assert.Equal(original, seekable.ToArray());
        }

        [Fact]
        public async Task ReadOnlySeekableLoadAsync_SnapshotsCustomStreamUsingAsyncReads() {
            byte[] original = CreatePresentation("Slow storage");
            var gate = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            using var source = new GatedReadStream(original, gate.Task, synchronousPrefixLength: 8);

            Task<PowerPointDocument> loading = PowerPointDocument.LoadAsync(source,
                new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            bool returnedBeforeReading = !loading.IsCompleted;
            gate.SetResult(true);

            using PowerPointDocument presentation = await loading;
            Assert.True(returnedBeforeReading);
            Assert.Equal("Slow storage", presentation.Slides[0].TextBoxes.First().Text);
        }

        private static byte[] WithUnreferencedSignaturePart(byte[] package) {
            using var output = new MemoryStream();
            output.Write(package, 0, package.Length);
            using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
                ZipArchiveEntry types = archive.GetEntry("[Content_Types].xml")!;
                string xml;
                using (var reader = new StreamReader(types.Open())) xml = reader.ReadToEnd();
                xml = xml.Replace("</Types>", "<Override PartName=\"/_xmlsignatures/sig1.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/></Types>");
                types.Delete();
                using (var writer = new StreamWriter(archive.CreateEntry("[Content_Types].xml").Open())) writer.Write(xml);
                using var signature = new StreamWriter(archive.CreateEntry("_xmlsignatures/sig1.xml").Open());
                signature.Write("<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\" />");
            }
            return output.ToArray();
        }

        [Fact]
        public void ReadOnlyNonSeekableLoad_DoesNotNeedTheStreamAfterLoading() {
            byte[] original = CreatePresentation("Forward only");
            PowerPointDocument presentation;
            using (var source = new NonSeekableReadStream(original)) {
                presentation = PowerPointDocument.Load(source, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            }

            using (presentation) {
                Assert.Equal("Forward only", presentation.Slides[0].TextBoxes.First().Text);
            }
        }

        [Fact]
        public void ReadWriteLoad_DoesNotShareTheCallersBuffer() {
            byte[] original = CreatePresentation("Detached");
            byte[] callerBuffer = (byte[])original.Clone();
            using var source = new MemoryStream(callerBuffer, writable: false);
            using PowerPointDocument presentation = PowerPointDocument.Load(source);

            Array.Clear(callerBuffer, 0, callerBuffer.Length);

            Assert.Equal("Detached", presentation.Slides[0].TextBoxes.First().Text);
            using var reloaded = PowerPointDocument.Load(new MemoryStream(presentation.ToBytes()),
                new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            Assert.Equal("Detached", reloaded.Slides[0].TextBoxes.First().Text);
        }

        [Fact]
        public void ReadWriteLoad_SavesEditsAndKeepsUntouchedPartsIntact() {
            byte[] picture = PdfPngTestImages.CreateRgbPng(0, 128, 255);
            byte[] original = CreatePresentation("Before", picture);
            using var source = new MemoryStream(original, writable: false);
            using PowerPointDocument presentation = PowerPointDocument.Load(source);
            presentation.Slides[0].TextBoxes.First().Text = "After";

            using var saved = new MemoryStream();
            presentation.Save(saved);
            saved.Position = 0;

            using var reloaded = PowerPointDocument.Load(saved, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
            Assert.Equal("After", reloaded.Slides[0].TextBoxes.First().Text);
            Assert.Equal(picture, reloaded.Slides[0].Images.Single().GetImageBytes());
        }

#if NET
        [Fact]
        [Trait("Category", "Performance")]
        public void Load_AllocatesNoMoreCopiesThanTheAccessModeNeeds() {
            byte[] original = LargePartPackages.WithLargePart(CreatePresentation("Large"), "ppt");

            long readOnly = MeasureLoad(original, DocumentAccessMode.ReadOnly);
            long readWrite = MeasureLoad(original, DocumentAccessMode.ReadWrite);

            Assert.True(readOnly < original.Length / 4, $"Read-only load allocated {readOnly} bytes for a {original.Length} byte file.");
            Assert.True(readWrite < original.Length * 3L / 2, $"Editable load allocated {readWrite} bytes for a {original.Length} byte file.");
        }

        private static long MeasureLoad(byte[] file, DocumentAccessMode mode) {
            using var source = new MemoryStream(file, writable: false);
            GC.Collect();
            long before = GC.GetAllocatedBytesForCurrentThread();
            using PowerPointDocument presentation = PowerPointDocument.Load(source, new PowerPointLoadOptions { AccessMode = mode });
            return GC.GetAllocatedBytesForCurrentThread() - before;
        }
#endif
    }
}
