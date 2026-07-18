using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptInteropArtifactTests {
        private static readonly (string FileName, string Marker,
            PowerPointFileFormat Format)[] Variants = {
            ("libreoffice-ppt.ppt", "PPT", PowerPointFileFormat.Ppt),
            ("libreoffice-pot.pot", "POT", PowerPointFileFormat.Pot),
            ("libreoffice-pps.pps", "PPS", PowerPointFileFormat.Pps)
        };

        [Fact]
        [Trait("Category", "LegacyPptLibreOfficeArtifact")]
        public void EmitsRepresentativeBinaryPowerPointVariants() {
            string? requestedOutput = Environment.GetEnvironmentVariable(
                "OFFICEIMO_LEGACY_PPT_INTEROP_OUTPUT");
            bool keep = !string.IsNullOrWhiteSpace(requestedOutput);
            string output = keep
                ? Path.GetFullPath(requestedOutput!)
                : Path.Combine(Path.GetTempPath(), "OfficeIMO-LegacyPpt-Interop-"
                    + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(output);
            try {
                foreach ((string fileName, string marker,
                    PowerPointFileFormat format) in Variants) {
                    string path = Path.Combine(output, fileName);
                    WriteRepresentativePresentation(path, marker);

                    using LegacyPptLoadResult loaded = PowerPointPresentation
                        .LoadLegacyPptWithReport(path);
                    loaded.EnsureNoImportErrors();
                    Assert.False(loaded.HasConversionLoss,
                        string.Join(Environment.NewLine, loaded.Diagnostics));
                    loaded.EnsureNoConversionLoss();
                    Assert.Equal(format, loaded.Document.SourceFormat);
                    AssertRepresentativeSemantics(loaded.Document, marker);
                }
            } finally {
                if (!keep && Directory.Exists(output)) {
                    Directory.Delete(output, recursive: true);
                }
            }
        }

        [Fact]
        [Trait("Category", "LegacyPptLibreOfficeReopen")]
        public void ReopensLibreOfficeResavedVariantsWithExpectedSemantics() {
            string? requestedInput = Environment.GetEnvironmentVariable(
                "OFFICEIMO_LEGACY_PPT_INTEROP_INPUT");
            if (string.IsNullOrWhiteSpace(requestedInput)) return;

            string input = Path.GetFullPath(requestedInput!);
            string[] files = Directory.GetFiles(input, "*.ppt",
                SearchOption.AllDirectories);
            Assert.Equal(Variants.Length, files.Length);

            foreach ((string fileName, string marker,
                PowerPointFileFormat _) in Variants) {
                string expectedName = Path.GetFileNameWithoutExtension(fileName)
                    + ".ppt";
                string path = Assert.Single(files, candidate =>
                    Path.GetFileName(candidate).Equals(expectedName,
                        StringComparison.OrdinalIgnoreCase));

                using LegacyPptLoadResult loaded = PowerPointPresentation
                    .LoadLegacyPptWithReport(path);
                loaded.EnsureNoImportErrors();
                Assert.False(loaded.HasConversionLoss,
                    string.Join(Environment.NewLine, loaded.Diagnostics));
                loaded.EnsureNoConversionLoss();
                AssertRepresentativeSemantics(loaded.Document, marker);
            }
        }

        private static void WriteRepresentativePresentation(string path,
            string marker) {
            byte[] image = OfficePngWriter.Encode(new OfficeRasterImage(
                24, 16, OfficeColor.FromRgb(37, 99, 235)));
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create(path);
            presentation.SlideSize.SetPreset(
                PowerPointSlideSizePreset.Screen16x9);

            PowerPointSlide first = presentation.AddSlide();
            first.AddTitle("Binary interoperability " + marker);
            first.AddTextBox("OfficeIMO generated " + marker,
                800000, 1700000, 4200000, 700000);
            first.AddRectangle(800000, 2800000, 2200000, 900000,
                    "Interop rectangle")
                .Fill("D9EAF7")
                .Stroke("2563EB", 1.5D);
            using (var stream = new MemoryStream(image, writable: false)) {
                first.AddPicture(stream, ImagePartType.Png,
                    3600000, 2700000, 1800000, 1200000);
            }
            first.Notes.Text = "LibreOffice round-trip notes " + marker;
            first.Transition = SlideTransition.Fade;

            PowerPointSlide second = presentation.AddSlide();
            second.AddTextBox("Second slide " + marker,
                900000, 900000, 4200000, 800000);
            PowerPointTable table = second.AddTable(2, 2,
                900000, 2200000, 5000000, 1500000);
            table.HeaderRow = true;
            table.GetCell(0, 0).Text = "Region";
            table.GetCell(0, 1).Text = "Revenue";
            table.GetCell(1, 0).Text = marker;
            table.GetCell(1, 1).Text = "120";
            table.SetCellBorders(TableCellBorders.All, "2563EB",
                widthPoints: 1.25D);
            presentation.Save();
        }

        private static string ReadAllText(PowerPointPresentation presentation) =>
            string.Join("\n", presentation.Slides.SelectMany(slide =>
                slide.TextBoxes).Select(textBox => textBox.Text));

        private static void AssertRepresentativeSemantics(
            PowerPointPresentation presentation, string marker) {
            Assert.Equal(2, presentation.Slides.Count);
            Assert.InRange(presentation.SlideSize.AspectRatio, 1.77D, 1.79D);
            string text = ReadAllText(presentation);
            Assert.Contains("Binary interoperability " + marker, text);
            Assert.Contains("Second slide " + marker, text);
            PowerPointTable table = Assert.Single(
                presentation.Slides[1].Tables);
            Assert.Equal("Region", table.GetCell(0, 0).Text);
            Assert.Equal("Revenue", table.GetCell(0, 1).Text);
            Assert.Equal(marker, table.GetCell(1, 0).Text);
            Assert.Equal("120", table.GetCell(1, 1).Text);
            Assert.Contains(marker,
                presentation.Slides[0].GetSpeakerNotesText());
            Assert.True(presentation.Slides.Sum(slide =>
                slide.Pictures.Count()) >= 1);
            PowerPointAutoShape rectangle = Assert.Single(
                presentation.Slides[0].Shapes.OfType<PowerPointAutoShape>());
            Assert.Equal("D9EAF7", rectangle.FillColor, ignoreCase: true);
            Assert.Equal("2563EB", rectangle.OutlineColor, ignoreCase: true);
            Assert.Equal(SlideTransition.Fade,
                presentation.Slides[0].Transition);
            Assert.Empty(presentation.ValidateDocument());
        }
    }
}
