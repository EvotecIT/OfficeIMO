using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptMasterTests {
        [Fact]
        public void ImportedMainMasterShapeMove_AppendsPreservingIncrementalRecord() {
            LegacyPptPresentation original = LegacyPptPresentation.Load(FixturePath);
            LegacyPptMaster[] mainMasters = original.Masters
                .Where(master => master.IsMainMaster).ToArray();
            int masterIndex = Array.FindIndex(mainMasters,
                master => master.Shapes.Count > 0);
            Assert.True(masterIndex >= 0);
            LegacyPptMaster originalMaster = mainMasters[masterIndex];

            using PowerPointPresentation imported = PowerPointPresentation.Load(
                FixturePath);
            SlideMasterPart masterPart = imported.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.ElementAt(masterIndex);
            IReadOnlyList<PowerPointShape> projectedShapes = LegacyPptWriter
                .ReadMasterShapesForWrite(masterPart, out string? reason);
            Assert.Null(reason);
            Assert.Equal(originalMaster.Shapes.Count, projectedShapes.Count);
            PowerPointShape projectedShape = projectedShapes[0];
            long expectedLeft = projectedShape.Left + 15875L;
            projectedShape.Left = expectedLeft;

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();
            Assert.True(preflight.CanWrite,
                string.Join(Environment.NewLine, preflight.Findings));
            LegacyPptPresentation saved = LegacyPptPresentation.Load(
                imported.ToBytes(PowerPointFileFormat.Ppt));
            LegacyPptMaster savedMaster = Assert.Single(saved.Masters,
                master => master.MasterId == originalMaster.MasterId);

            Assert.Equal(originalMaster.Shapes[0].Bounds.Left + 10,
                savedMaster.Shapes[0].Bounds.Left);
            Assert.Equal(original.Package.UserEdits.Count + 1,
                saved.Package.UserEdits.Count);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            AssertUnrelatedMasterChildrenEqual(original, saved,
                originalMaster.PersistId);

            using var stream = new MemoryStream(saved.Package.CopyOriginalBytes());
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(stream);
            SlideMasterPart reopenedMaster = reopened.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.ElementAt(masterIndex);
            PowerPointShape reopenedShape = LegacyPptWriter
                .ReadMasterShapesForWrite(reopenedMaster, out _)[0];
            Assert.Equal(expectedLeft, reopenedShape.Left);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedMainMasterPlainTextEdit_AppendsPreservingIncrementalRecord() {
            byte[] sourceBytes = CreateBinaryWithEditableMasterText();
            LegacyPptPresentation original = LegacyPptPresentation.Load(sourceBytes);
            LegacyPptMaster originalMaster = Assert.Single(original.Masters);

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            SlideMasterPart masterPart = imported.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            PowerPointTextBox textBox = Assert.IsType<PowerPointTextBox>(
                Assert.Single(LegacyPptWriter.ReadMasterShapesForWrite(
                    masterPart, out _)));
            textBox.Text = "Edited label";

            Assert.True(imported.AnalyzeLegacyPptWrite().CanWrite);
            byte[] savedBytes = imported.ToBytes(PowerPointFileFormat.Ppt);
            LegacyPptPresentation saved = LegacyPptPresentation.Load(savedBytes);
            LegacyPptMaster savedMaster = Assert.Single(saved.Masters);

            Assert.Equal("Edited label", Assert.Single(savedMaster.Shapes).Text);
            Assert.True(saved.Package.DocumentStream.AsSpan(0,
                    original.Package.DocumentStream.Length)
                .SequenceEqual(original.Package.DocumentStream));
            AssertUnrelatedMasterChildrenEqual(original, saved,
                originalMaster.PersistId);

            using var reopenedInput = new MemoryStream(savedBytes);
            using PowerPointPresentation reopened =
                PowerPointPresentation.Load(reopenedInput);
            SlideMasterPart reopenedMaster = reopened.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            PowerPointTextBox reopenedText = Assert.IsType<PowerPointTextBox>(
                Assert.Single(LegacyPptWriter.ReadMasterShapesForWrite(
                    reopenedMaster, out _)));
            Assert.Equal("Edited label", reopenedText.Text);
            Assert.Empty(reopened.ValidateDocument());
        }

        [Fact]
        public void ImportedMainMasterUnsupportedStyleEdit_RemainsLossBlocked() {
            using PowerPointPresentation imported = PowerPointPresentation.Load(
                FixturePath);
            P.Shape shape = imported.OpenXmlDocument.PresentationPart!
                .SlideMasterParts.SelectMany(master => master.SlideMaster!
                    .CommonSlideData!.ShapeTree!.Descendants<P.Shape>())
                .First();
            shape.ShapeProperties!.Transform2D!.Rotation = 60000;

            LegacyPptWritePreflightReport preflight = imported
                .AnalyzeLegacyPptWrite();

            Assert.False(preflight.CanWrite);
            Assert.Contains(preflight.Findings, finding =>
                finding.Code == "PPT-WRITE-IMPORT-LOSS");
        }

        private static byte[] CreateBinaryWithEditableMasterText() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            SlideMasterPart masterPart = presentation.OpenXmlDocument
                .PresentationPart!.SlideMasterParts.Single();
            P.ShapeTree tree = masterPart.SlideMaster!.CommonSlideData!
                .ShapeTree!;
            tree.Append(new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties {
                        Id = 2U,
                        Name = "Editable master label"
                    },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 400000, Y = 500000 },
                        new A.Extents { Cx = 3000000, Cy = 600000 }),
                    new A.PresetGeometry(new A.AdjustValueList()) {
                        Preset = A.ShapeTypeValues.Rectangle
                    }),
                new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(new A.Text("Master label")),
                        new A.EndParagraphRunProperties()))));
            presentation.AddSlide(P.SlideLayoutValues.Blank);
            return presentation.ToBytes(PowerPointFileFormat.Ppt);
        }

        private static void AssertUnrelatedMasterChildrenEqual(
            LegacyPptPresentation original, LegacyPptPresentation saved,
            uint persistId) {
            IReadOnlyList<byte[]> originalChildren = ReadMasterChildrenExceptDrawing(
                original, persistId);
            IReadOnlyList<byte[]> savedChildren = ReadMasterChildrenExceptDrawing(
                saved, persistId);
            Assert.Equal(originalChildren.Count, savedChildren.Count);
            for (int index = 0; index < originalChildren.Count; index++) {
                Assert.True(originalChildren[index]
                    .SequenceEqual(savedChildren[index]));
            }
        }

        private static IReadOnlyList<byte[]> ReadMasterChildrenExceptDrawing(
            LegacyPptPresentation presentation, uint persistId) {
            LegacyPptPersistObject persistObject =
                presentation.Package.PersistObjects[persistId];
            LegacyPptRecord record = LegacyPptRecordReader.ReadSingle(
                persistObject.RecordBytes, 0, new LegacyPptImportOptions());
            return record.Children.Where(child => child.Type != 0x040C)
                .Select(child => child.CopyRecordBytes()).ToArray();
        }
    }
}
