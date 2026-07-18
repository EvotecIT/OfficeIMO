using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void NormalAutoFit_IsExplicitlyBlockedAsNonRepresentable() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointTextBox textBox = presentation.AddSlide(
                    P.SlideLayoutValues.Blank)
                .AddTextBox("Shrink me");
            textBox.TextAutoFit = PowerPointTextAutoFit.Normal;

            LegacyPptWritePreflightReport preflight = presentation
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            LegacyPptWriteFinding finding = Assert.Single(
                preflight.Findings,
                item => item.Code == "PPT-WRITE-RICH-TEXT");
            Assert.Contains("no lossless classic binary PowerPoint mapping",
                finding.Description, StringComparison.Ordinal);
        }

        [Fact]
        public void ExplicitLineBreak_FreshEditAndRemove_RoundTripsNatively() {
            byte[] sourceBytes;
            using (PowerPointPresentation source =
                   PowerPointPresentation.Create()) {
                PowerPointTextBox textBox = source.AddSlide(
                        P.SlideLayoutValues.Blank)
                    .AddTextBox(string.Empty);
                Assert.IsType<P.Shape>(textBox.Element).TextBody =
                    new P.TextBody(new A.BodyProperties(),
                        new A.ListStyle(), new A.Paragraph(
                            new A.Run(new A.Text("Before")),
                            new A.Break {
                                RunProperties = new A.RunProperties {
                                    Bold = true,
                                    FontSize = 1800
                                }
                            },
                            new A.Run(new A.Text("After")),
                            new A.EndParagraphRunProperties()));

                LegacyPptWritePreflightReport preflight = source
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation sourceLegacy = LegacyPptPresentation.Load(
                sourceBytes);
            LegacyPptShape sourceShape = Assert.Single(
                Assert.Single(sourceLegacy.Slides).Shapes);
            Assert.Equal("Before\vAfter", sourceShape.Text);
            LegacyPptCharacterRun sourceBreak = Assert.Single(
                sourceShape.TextBody.CharacterRuns,
                run => run.Text == "\v");
            Assert.True(sourceBreak.Bold);
            Assert.Equal((short)18, sourceBreak.FontSizePoints);

            byte[] editedBytes;
            using (var input = new MemoryStream(sourceBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                A.Break lineBreak = Assert.Single(imported.Slides[0]
                    .SlidePart.Slide!.Descendants<A.Break>());
                Assert.True(lineBreak.RunProperties!.Bold!.Value);
                lineBreak.RunProperties.Bold = false;
                lineBreak.RunProperties.Italic = true;

                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                editedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation edited = LegacyPptPresentation.Load(
                editedBytes);
            LegacyPptCharacterRun editedBreak = Assert.Single(
                Assert.Single(Assert.Single(edited.Slides).Shapes)
                    .TextBody.CharacterRuns,
                run => run.Text == "\v");
            Assert.False(editedBreak.Bold);
            Assert.True(editedBreak.Italic);
            Assert.Equal(sourceLegacy.Package.UserEdits.Count + 1,
                edited.Package.UserEdits.Count);
            Assert.True(edited.Package.DocumentStream.AsSpan(0,
                    sourceLegacy.Package.DocumentStream.Length)
                .SequenceEqual(sourceLegacy.Package.DocumentStream));

            byte[] removedBytes;
            using (var input = new MemoryStream(editedBytes,
                       writable: false))
            using (PowerPointPresentation imported =
                   PowerPointPresentation.Load(input)) {
                Assert.Single(imported.Slides[0].SlidePart.Slide!
                    .Descendants<A.Break>()).Remove();
                Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
                removedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            }
            LegacyPptPresentation removed = LegacyPptPresentation.Load(
                removedBytes);
            Assert.Equal("BeforeAfter", Assert.Single(
                Assert.Single(removed.Slides).Shapes).Text);
            Assert.DoesNotContain(Assert.Single(Assert.Single(
                    removed.Slides).Shapes).TextBody.CharacterRuns,
                run => run.Text == "\v");
            Assert.Equal(edited.Package.UserEdits.Count + 1,
                removed.Package.UserEdits.Count);
            Assert.True(removed.Package.DocumentStream.AsSpan(0,
                    edited.Package.DocumentStream.Length)
                .SequenceEqual(edited.Package.DocumentStream));
        }

    }
}
