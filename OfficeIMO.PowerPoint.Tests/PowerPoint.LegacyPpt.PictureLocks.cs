using OfficeIMO.Drawing.Binary;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLegacyPptPictureLockTests {
        [Fact]
        public void ImportedPictureLockEdit_UsesIncrementalFoptRewrite() {
            string sourcePath = Path.Combine(AppContext.BaseDirectory,
                "Documents", "LegacyPptCorpus",
                "CroppedPicturePowerPoint.ppt");
            LegacyPptPresentation original = LegacyPptPresentation.Load(
                sourcePath);
            byte[] originalPictures = original.Package.CopyCompoundStreams()[
                "Pictures"];

            byte[] editedBytes;
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Load(sourcePath)) {
                foreach (PowerPointPicture model in presentation.Slides[0]
                             .Pictures) {
                    P.Picture picture = (P.Picture)model.Element;
                    P.NonVisualPictureDrawingProperties nonVisual = picture
                        .NonVisualPictureProperties!
                        .NonVisualPictureDrawingProperties!;
                    nonVisual.PreferRelativeResize = true;
                    A.PictureLocks locks = nonVisual
                        .GetFirstChild<A.PictureLocks>()!;
                    locks.NoSelection = true;
                    locks.NoChangeShapeType = true;
                }

                LegacyPptWritePreflightReport preflight = presentation
                    .AnalyzeLegacyPptWrite();
                Assert.True(preflight.CanWrite,
                    string.Join(Environment.NewLine, preflight.Findings));
                editedBytes = presentation.ToBytes(
                    PowerPointFileFormat.Ppt);
            }

            LegacyPptPresentation edited = LegacyPptPresentation.Load(
                editedBytes);
            LegacyPptShape[] editedPictures = Assert.Single(edited.Slides)
                .Shapes.Where(shape =>
                    shape.Kind == LegacyPptShapeKind.Picture).ToArray();
            Assert.Equal(2, editedPictures.Length);
            Assert.All(editedPictures, editedPicture => {
                OfficeArtShapeProtection protection =
                    OfficeArtShapeProtection.Decode(
                        editedPicture.Style.Properties);
                Assert.True(protection.LockAgainstSelect);
                Assert.True(editedPicture.Style.PreferRelativeResize);
                Assert.True(editedPicture.Style.LockShapeType);
            });
            Assert.Equal(originalPictures,
                edited.Package.CopyCompoundStreams()["Pictures"]);

            using var editedInput = new MemoryStream(editedBytes,
                writable: false);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(editedInput);
            Assert.All(reopened.Slides[0].Pictures, model => {
                P.NonVisualPictureDrawingProperties projected =
                    ((P.Picture)model.Element).NonVisualPictureProperties!
                    .NonVisualPictureDrawingProperties!;
                Assert.True(projected.PreferRelativeResize);
                Assert.True(projected.GetFirstChild<A.PictureLocks>()!
                    .NoSelection);
                Assert.True(projected.GetFirstChild<A.PictureLocks>()!
                    .NoChangeShapeType);
            });
            Assert.Empty(reopened.ValidateDocument());
        }
    }
}
