using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using System.Threading.Tasks;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptTests {
        private static string FixturePath => Path.Combine(AppContext.BaseDirectory,
            "Documents", "LegacyPptCorpus", "BasicPowerPoint.ppt");

        [Fact]
        public void NeutralReader_DecodesRealBinaryPresentation() {
            LegacyPptPresentation legacy = LegacyPptPresentation.Load(FixturePath);

            LegacyPptSlide slide = Assert.Single(legacy.Slides);
            Assert.Equal(7680, legacy.SlideWidth);
            Assert.Equal(4320, legacy.SlideHeight);
            Assert.Equal(3, slide.Shapes.Count(shape => shape.Kind == LegacyPptShapeKind.TextBox));
            Assert.Contains(slide.Shapes, shape => shape.Text == "OfficeIMO PowerPoint Basics");
            Assert.Contains(slide.Shapes, shape => shape.PlaceholderKind == LegacyPptPlaceholderKind.Title);
            Assert.True(legacy.CreateImportReport().HasConversionLoss);
            Assert.Contains(legacy.Diagnostics, diagnostic => diagnostic.Code == "PPT-TEXT-FORMATTING-FLATTENED");
        }

        [Fact]
        public void NormalLoad_RoutesPptAndProjectsEditablePptxModel() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);

            Assert.Equal(PowerPointFileFormat.Ppt, presentation.SourceFormat);
            Assert.Equal(FixturePath, presentation.SourcePath);
            PowerPointSlide slide = Assert.Single(presentation.Slides);
            Assert.Equal(3, slide.TextBoxes.Count());
            Assert.Contains(slide.TextBoxes, textBox => textBox.Text == "OfficeIMO PowerPoint Basics");
            slide.AddTextBox("Edited after binary import");

            using var pptx = presentation.ToStream();
            using PowerPointPresentation reopened = PowerPointPresentation.Load(pptx);
            Assert.Contains(reopened.Slides[0].TextBoxes, textBox => textBox.Text == "Edited after binary import");
        }

        [Theory]
        [InlineData(".pot", PowerPointFileFormat.Pot)]
        [InlineData(".pps", PowerPointFileFormat.Pps)]
        public void NormalLoad_PreservesLegacyExtensionSemantics(string extension, PowerPointFileFormat format) {
            string copy = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
            try {
                File.Copy(FixturePath, copy);
                using PowerPointPresentation presentation = PowerPointPresentation.Load(copy);
                Assert.Equal(format, presentation.SourceFormat);
                Assert.Single(presentation.Slides);
            } finally {
                if (File.Exists(copy)) File.Delete(copy);
            }
        }

        [Fact]
        public void NormalLoad_RejectsOtherLegacyOfficeFamiliesClearly() {
            string doc = Path.Combine(AppContext.BaseDirectory, "Documents", "LegacyDocCorpus", "ComSimpleParagraphs.doc");
            string xls = Path.Combine(AppContext.BaseDirectory, "Documents", "LegacyXlsCorpus",
                "openpreserve-format-corpus", "valid.xls");

            InvalidDataException wordError = Assert.Throws<InvalidDataException>(() => PowerPointPresentation.Load(doc));
            InvalidDataException excelError = Assert.Throws<InvalidDataException>(() => PowerPointPresentation.Load(xls));

            Assert.Contains("legacy Word", wordError.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("legacy Excel", excelError.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void NormalLoad_PrefersZipContentOverMisleadingPptExtension() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".ppt");
            try {
                using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                    created.AddSlide().AddTextBox("Actually Open XML");
                    File.WriteAllBytes(path, created.ToBytes());
                }

                using PowerPointPresentation loaded = PowerPointPresentation.Load(path);
                Assert.Equal(PowerPointFileFormat.Pptx, loaded.SourceFormat);
                Assert.Contains(loaded.Slides[0].TextBoxes, textBox => textBox.Text == "Actually Open XML");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void NativeWriter_RoundTripsTextAndBasicShapesAcrossTwoSlides() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide first = presentation.AddSlide();
            first.AddTitle("Binary title", 540000, 360000, 7200000, 700000);
            first.AddTextBox("Line one\nLine two", 540000, 1200000, 6000000, 1800000);
            first.AddRectangle(600000, 3300000, 1200000, 700000);
            first.AddEllipse(2100000, 3300000, 1200000, 700000);
            first.AddShape(A.ShapeTypeValues.Line, 3600000, 3300000, 1600000, 400000);
            presentation.AddSlide().AddTextBox("Second slide", 700000, 700000, 5000000, 1000000);

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite);
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptPresentation binary = LegacyPptPresentation.Load(bytes);
            Assert.Equal(2, binary.Slides.Count);
            Assert.Contains(binary.Slides[0].Shapes, shape => shape.Text == "Binary title");
            Assert.Contains(binary.Slides[0].Shapes, shape => shape.Kind == LegacyPptShapeKind.Rectangle);
            Assert.Contains(binary.Slides[0].Shapes, shape => shape.Kind == LegacyPptShapeKind.Ellipse);
            Assert.Contains(binary.Slides[0].Shapes, shape => shape.Kind == LegacyPptShapeKind.Line);

            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation projected = PowerPointPresentation.Load(stream);
            Assert.Equal(PowerPointFileFormat.Ppt, projected.SourceFormat);
            Assert.Equal(2, projected.Slides.Count);
            Assert.Contains(projected.Slides[1].TextBoxes, textBox => textBox.Text == "Second slide");
        }

        [Fact]
        public void NativeWriter_BlocksKnownLossUnlessExplicitlyAllowed() {
            string image = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddPicture(image);

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.HasConversionLoss);
            Assert.Throws<NotSupportedException>(() => presentation.ToBytes(PowerPointFileFormat.Ppt));

            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt,
                new PowerPointSaveOptions { LossPolicy = PowerPointConversionLossPolicy.Allow });
            Assert.NotEmpty(bytes);
            Assert.Single(LegacyPptPresentation.Load(bytes).Slides);
        }

        [Fact]
        public void NativeWriter_PreflightReportsTransitionAndVisualStyleLoss() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.Transition = SlideTransition.Fade;
            slide.AddRectangle(100000, 100000, 1000000, 500000).Fill("FF0000");

            LegacyPptWritePreflightReport report = presentation.AnalyzeLegacyPptWrite();

            Assert.Contains(report.Findings, finding => finding.Code == "PPT-WRITE-TRANSITION");
            Assert.Contains(report.Findings, finding => finding.Code == "PPT-WRITE-SHAPE-STYLE");
        }

        [Fact]
        public void AssociatedLegacyStream_SavePreservesBinaryFormat() {
            byte[] source = File.ReadAllBytes(FixturePath);
            using var stream = new MemoryStream();
            stream.Write(source, 0, source.Length);
            stream.Position = 0;

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(stream)) {
                presentation.Slides[0].AddTextBox("Saved back as binary");
                Assert.Throws<NotSupportedException>(() => presentation.Save());
                presentation.Save(stream, presentation.SourceFormat,
                    new PowerPointSaveOptions { LossPolicy = PowerPointConversionLossPolicy.Allow });
            }

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(stream.ToArray());
            Assert.Contains(legacy.Slides[0].Shapes, shape => shape.Text == "Saved back as binary");
        }

        [Theory]
        [InlineData(".ppt")]
        [InlineData(".pot")]
        [InlineData(".pps")]
        public void SaveCopy_UsesBinaryWriterForLegacyExtensions(string extension) {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + extension);
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create();
                presentation.AddSlide().AddTextBox("Legacy extension");
                presentation.SaveCopy(path);

                LegacyPptPresentation binary = LegacyPptPresentation.Load(path);
                Assert.Single(binary.Slides);
                Assert.Contains(binary.Slides[0].Shapes, shape => shape.Text == "Legacy extension");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task NativeWriter_AsyncPathAndStreamOverloadsProduceBinaryFiles() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".ppt");
            try {
                await using PowerPointPresentation presentation = PowerPointPresentation.Create(path);
                presentation.AddSlide().AddTextBox("Async binary");

                await presentation.SaveAsync();
                Assert.Contains(LegacyPptPresentation.Load(path).Slides[0].Shapes,
                    shape => shape.Text == "Async binary");

                using var stream = new MemoryStream();
                await presentation.SaveAsync(stream, PowerPointFileFormat.Pps);
                Assert.Contains(LegacyPptPresentation.Load(stream.ToArray()).Slides[0].Shapes,
                    shape => shape.Text == "Async binary");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
