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
            Assert.NotEmpty(legacy.Package.UserEdits);
            Assert.NotEmpty(legacy.Package.PersistObjects);
            Assert.True(legacy.CreateImportReport().CompoundStreamCount >= 2);
        }

        [Fact]
        public void UnmodifiedBinarySave_PreservesTheOriginalPackageExactly() {
            byte[] source = File.ReadAllBytes(FixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            byte[] saved = presentation.ToBytes(PowerPointFileFormat.Ppt);

            Assert.True(preflight.CanWrite);
            Assert.Equal(source, saved);
        }

        [Fact]
        public void BinaryPreservationFingerprint_DetectsCorePropertyChanges() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);

            presentation.BuiltinDocumentProperties.Title = "Changed title";

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings,
                finding => finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }

        [Fact]
        public void ImportedBinaryGeometryEdit_AppendsIncrementalEditAndPreservesOpaqueStreams() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(FixturePath);
            LegacyPptShape originalTitle = original.Slides[0].Shapes.Single(shape =>
                shape.Text == "OfficeIMO PowerPoint Basics");
            IReadOnlyDictionary<string, byte[]> originalStreams = original.Package.CopyCompoundStreams();

            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            PowerPointTextBox title = presentation.Slides[0].TextBoxes.Single(textBox =>
                textBox.Text == "OfficeIMO PowerPoint Basics");
            title.Left += 15875;
            title.Width += 3175;

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite);
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptPresentation saved = LegacyPptPresentation.Load(bytes);
            LegacyPptShape savedTitle = saved.Slides[0].Shapes.Single(shape =>
                shape.Text == "OfficeIMO PowerPoint Basics");
            Assert.Equal(originalTitle.Bounds.Left + 10, savedTitle.Bounds.Left);
            Assert.Equal(originalTitle.Bounds.Width + 2, savedTitle.Bounds.Width);
            Assert.Equal(original.Package.UserEdits.Count + 1, saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0, original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));

            IReadOnlyDictionary<string, byte[]> savedStreams = saved.Package.CopyCompoundStreams();
            Assert.Equal(originalStreams.Keys.OrderBy(value => value), savedStreams.Keys.OrderBy(value => value));
            foreach (KeyValuePair<string, byte[]> stream in originalStreams) {
                if (stream.Key.Equals("PowerPoint Document", StringComparison.OrdinalIgnoreCase)
                    || stream.Key.Equals("Current User", StringComparison.OrdinalIgnoreCase)) continue;
                Assert.Equal(stream.Value, savedStreams[stream.Key]);
            }
        }

        [Fact]
        public void ImportedRichBinaryText_SameLengthEditPreservesFormattingRecords() {
            const string originalText = "OfficeIMO PowerPoint Basics";
            const string replacementText = "OfficeIMO BinaryDeck Basics";
            Assert.Equal(originalText.Length, replacementText.Length);

            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            PowerPointTextBox title = presentation.Slides[0].TextBoxes.Single(textBox =>
                textBox.Text == originalText);
            title.Text = replacementText;

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptPresentation saved = LegacyPptPresentation.Load(bytes);
            Assert.Contains(saved.Slides[0].Shapes, shape => shape.Text == replacementText);
            Assert.Contains(saved.Diagnostics, diagnostic =>
                diagnostic.Code == "PPT-TEXT-FORMATTING-FLATTENED");
        }

        [Fact]
        public void ImportedRichBinaryText_LengthChangingEditRemainsLossBlocked() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            PowerPointTextBox title = presentation.Slides[0].TextBoxes.Single(textBox =>
                textBox.Text == "OfficeIMO PowerPoint Basics");
            title.Text += "!";

            LegacyPptWritePreflightReport preflight = presentation.AnalyzeLegacyPptWrite();
            Assert.False(preflight.CanWrite);
            Assert.Throws<NotSupportedException>(() => presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedPlainBinaryText_ArbitraryLengthEditUsesIncrementalSave() {
            byte[] source;
            using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                created.AddSlide().AddTextBox("Short text", 100000, 100000, 3000000, 600000);
                source = created.ToBytes(PowerPointFileFormat.Ppt);
            }

            using var input = new MemoryStream(source);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            presentation.Slides[0].TextBoxes.Single().Text = "A substantially longer binary text value";

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            byte[] savedBytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            Assert.Contains(saved.Slides[0].Shapes,
                shape => shape.Text == "A substantially longer binary text value");
            Assert.Equal(2, saved.Package.UserEdits.Count);
        }

        [Fact]
        public void NormalLoad_RoutesPptAndProjectsEditablePptxModel() {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);

            Assert.Equal(PowerPointFileFormat.Ppt, presentation.SourceFormat);
            Assert.Equal(FixturePath, presentation.SourcePath);
            Assert.Single(presentation.LegacyPptProjectionMap!.Slides);
            Assert.Equal(3, presentation.LegacyPptProjectionMap.Slides[0].Shapes.Count);
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
        public void NativeWriter_RoundTripsHiddenSlideState() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Visible");
            PowerPointSlide hidden = presentation.AddSlide();
            hidden.AddTextBox("Hidden");
            hidden.Hidden = true;

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);

            LegacyPptPresentation legacy = LegacyPptPresentation.Load(bytes);
            Assert.False(legacy.Slides[0].Hidden);
            Assert.True(legacy.Slides[1].Hidden);
            using var stream = new MemoryStream(bytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(stream);
            Assert.False(reopened.Slides[0].Hidden);
            Assert.True(reopened.Slides[1].Hidden);
        }

        [Fact]
        public void ImportedBinaryHiddenState_TogglesThroughIncrementalEdits() {
            byte[] hiddenBytes;
            using (PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath)) {
                presentation.Slides[0].Hidden = true;
                Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
                hiddenBytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation hidden = LegacyPptPresentation.Load(hiddenBytes);
            Assert.True(hidden.Slides[0].Hidden);

            using var input = new MemoryStream(hiddenBytes);
            using PowerPointPresentation reopened = PowerPointPresentation.Load(input);
            reopened.Slides[0].Hidden = false;
            Assert.True(reopened.AnalyzeLegacyPptWrite().CanWrite);
            LegacyPptPresentation visible = LegacyPptPresentation.Load(
                reopened.ToBytes(PowerPointFileFormat.Ppt));
            Assert.False(visible.Slides[0].Hidden);
            Assert.Equal(hidden.Package.UserEdits.Count + 1, visible.Package.UserEdits.Count);
        }

        [Fact]
        public void ImportedBinarySlideOrder_ReordersPersistGroupsIncrementally() {
            byte[] sourceBytes;
            using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                created.AddSlide().AddTextBox("First");
                created.AddSlide().AddTextBox("Second");
                created.AddSlide().AddTextBox("Third");
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation source = LegacyPptPresentation.Load(sourceBytes);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            presentation.MoveSlide(0, 2);

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            Assert.Equal(new[] { "Second", "Third", "First" }, saved.Slides.Select(slide =>
                slide.Shapes.Single(shape => shape.Kind == LegacyPptShapeKind.TextBox).Text));
            Assert.Equal(source.Package.UserEdits.Count + 1, saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0, source.Package.DocumentStream.Length)
                .SequenceEqual(source.Package.DocumentStream));
        }

        [Fact]
        public void ImportedBinarySlideDeletion_RemovesPersistGroupIncrementally() {
            byte[] sourceBytes;
            using (PowerPointPresentation created = PowerPointPresentation.Create()) {
                created.AddSlide().AddTextBox("Keep first");
                created.AddSlide().AddTextBox("Delete middle");
                created.AddSlide().AddTextBox("Keep last");
                sourceBytes = created.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation source = LegacyPptPresentation.Load(sourceBytes);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            presentation.RemoveSlide(1);

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));
            Assert.Equal(new[] { "Keep first", "Keep last" }, saved.Slides.Select(slide =>
                slide.Shapes.Single(shape => shape.Kind == LegacyPptShapeKind.TextBox).Text));
            Assert.Equal(source.Package.UserEdits.Count + 1, saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0, source.Package.DocumentStream.Length)
                .SequenceEqual(source.Package.DocumentStream));
        }

        [Fact]
        public void ImportedBinarySlideAddition_AppendsPersistAndDrawingClusters() {
            LegacyPptPresentation source = LegacyPptPresentation.Load(FixturePath);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(FixturePath);
            PowerPointSlide added = presentation.AddSlide();
            added.AddTextBox("Incremental new slide", 200000, 200000, 4000000, 700000);
            added.AddRectangle(300000, 1200000, 1200000, 600000);
            added.Hidden = true;

            Assert.True(presentation.AnalyzeLegacyPptWrite().CanWrite);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                presentation.ToBytes(PowerPointFileFormat.Ppt));

            Assert.Equal(2, saved.Slides.Count);
            Assert.Contains(saved.Slides[0].Shapes, shape => shape.Text == "OfficeIMO PowerPoint Basics");
            Assert.Contains(saved.Slides[1].Shapes, shape => shape.Text == "Incremental new slide");
            Assert.Contains(saved.Slides[1].Shapes, shape => shape.Kind == LegacyPptShapeKind.Rectangle);
            Assert.True(saved.Slides[1].Hidden);
            Assert.Equal(source.Package.UserEdits.Count + 1, saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0, source.Package.DocumentStream.Length)
                .SequenceEqual(source.Package.DocumentStream));
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
