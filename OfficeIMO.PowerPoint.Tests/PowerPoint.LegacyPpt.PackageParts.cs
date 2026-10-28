using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void NativeWriter_BlocksUnrepresentablePackagePartsAndRelationships() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Package preflight");
            PresentationPart presentationPart = presentation.OpenXmlDocument
                .PresentationPart!;

            CustomXmlPart customXml = presentationPart.AddCustomXmlPart(
                CustomXmlPartType.CustomXml);
            Feed(customXml, "<root><value>42</value></root>");

            ExtendedPart unknown = presentationPart.AddExtendedPart(
                "urn:officeimo:test:relationship",
                "application/vnd.officeimo.test+xml", "xml");
            Feed(unknown, "<test />");

            EmbeddedObjectPart embedded = slide.SlidePart
                .AddEmbeddedObjectPart(
                    "application/vnd.openxmlformats-officedocument.oleObject");
            Feed(embedded, new byte[] { 1, 2, 3, 4 });

            presentationPart.AddExternalRelationship(
                "urn:officeimo:test:external",
                new Uri("https://example.invalid/data"));

            LegacyPptWritePreflightReport report =
                presentation.AnalyzeLegacyPptWrite();

            Assert.Contains(report.Findings, finding =>
                finding.Feature == LegacyPptFeature.CustomXml
                && finding.Code == "ppt-write-custom-xml");
            Assert.Contains(report.Findings, finding =>
                finding.Code == "ppt-write-extended-part");
            Assert.Contains(report.Findings, finding =>
                finding.Code == "ppt-write-embedded-package");
            Assert.Contains(report.Findings, finding =>
                finding.Code == "ppt-write-external-relationship");
            Assert.Throws<NotSupportedException>(() =>
                presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));
        }

        [Fact]
        public void ImportedPresentation_CustomXmlBypassesNeitherPreservationPath() {
            byte[] sourceBytes;
            using (PowerPointDocument source =
                    PowerPointDocument.Create()) {
                source.AddSlide().AddTextBox("Imported package preflight");
                sourceBytes = source.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt });
            }

            using var input = new MemoryStream(sourceBytes);
            using PowerPointDocument imported =
                PowerPointDocument.Load(input);
            CustomXmlPart customXml = imported.OpenXmlDocument
                .PresentationPart!.AddCustomXmlPart(
                    CustomXmlPartType.CustomXml);
            Feed(customXml, "<imported />");

            LegacyPptWritePreflightReport report =
                imported.AnalyzeLegacyPptWrite();

            Assert.Contains(report.Findings, finding =>
                finding.Code == "ppt-write-custom-xml");
            Assert.Throws<NotSupportedException>(() =>
                imported.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));
        }

        [Fact]
        public void NativeWriter_ClassifiesAdvancedExtendedParts() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            presentation.AddSlide().AddTextBox("Extended part preflight");
            PresentationPart presentationPart = presentation.OpenXmlDocument
                .PresentationPart!;

            ExtendedPart activeX = presentationPart.AddExtendedPart(
                "http://schemas.microsoft.com/office/2006/relationships/activeXControl",
                "application/vnd.ms-office.activeX+xml", "xml");
            Feed(activeX, "<ax:ocx xmlns:ax=\"urn:active-x\" />");

            ExtendedPart webExtension = presentationPart.AddExtendedPart(
                "http://schemas.microsoft.com/office/2011/relationships/webextension",
                "application/vnd.ms-office.webextension+xml", "xml");
            Feed(webExtension, "<we:webextension xmlns:we=\"urn:web-extension\" />");

            ExtendedPart vba = presentationPart.AddExtendedPart(
                "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
                "application/vnd.ms-office.vbaProject", "bin");
            Feed(vba, new byte[] { 5, 6, 7, 8 });

            LegacyPptWritePreflightReport report =
                presentation.AnalyzeLegacyPptWrite();

            Assert.Contains(report.Findings, finding =>
                finding.Feature == LegacyPptFeature.ActiveX
                && finding.Code == "ppt-write-activex");
            Assert.Contains(report.Findings, finding =>
                finding.Feature == LegacyPptFeature.UnknownRecordsAndStreams
                && finding.Code == "ppt-write-web-extension");
            Assert.Contains(report.Findings, finding =>
                finding.Feature == LegacyPptFeature.VbaProjects
                && finding.Code == "ppt-write-vba-extended-part");
            Assert.DoesNotContain(report.Findings, finding =>
                finding.Code == "ppt-write-extended-part");
        }

        [Fact]
        public void NativeWriter_BlocksTypedPartsAndPackageRootRelationships() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            presentation.AddSlide().AddTextBox("Typed package preflight");

            RibbonExtensibilityPart ribbon = presentation.OpenXmlDocument
                .AddRibbonExtensibilityPart();
            Feed(ribbon,
                "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\" />");
            presentation.OpenXmlDocument.AddExternalRelationship(
                "urn:officeimo:test:package-external",
                new Uri("https://example.invalid/package-data"));

            LegacyPptWritePreflightReport report =
                presentation.AnalyzeLegacyPptWrite();

            Assert.Contains(report.Findings, finding =>
                finding.Code == "ppt-write-package-part");
            Assert.Contains(report.Findings, finding =>
                finding.Code == "ppt-write-external-relationship");
            Assert.Throws<NotSupportedException>(() =>
                presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));
        }

        [Fact]
        public void NativeWriter_RecognizesMappedOleObjectInsideGroup() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            byte[] storageBytes = CreateOleTestStorage("Grouped OLE");
            using var storage = new MemoryStream(storageBytes,
                writable: false);
            PowerPointOleObject ole = slide.AddOleObject(storage, "Package",
                OfficeIMO.OfficeLength.Emus(100000L), OfficeIMO.OfficeLength.Emus(100000L), OfficeIMO.OfficeLength.Emus(1000000L), OfficeIMO.OfficeLength.Emus(600000L));
            PowerPointAutoShape marker = slide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(1200000L), OfficeIMO.OfficeLength.Emus(100000L), OfficeIMO.OfficeLength.Emus(600000L), OfficeIMO.OfficeLength.Emus(600000L));
            slide.GroupShapes(new PowerPointShape[] { ole, marker },
                "Grouped OLE");

            LegacyPptWritePreflightReport report =
                presentation.AnalyzeLegacyPptWrite();

            Assert.DoesNotContain(report.Findings, finding =>
                finding.Code == "ppt-write-embedded-package");
            Assert.True(report.CanWrite,
                string.Join(Environment.NewLine, report.Findings));
            Assert.NotEmpty(presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));
        }

        [Fact]
        public void GroupProjection_MapsAnimationsPastOmittedOleChildren() {
            byte[] bytes;
            using (PowerPointDocument presentation =
                   PowerPointDocument.Create()) {
                PowerPointSlide slide = presentation.AddSlide();
                byte[] storageBytes = CreateOleTestStorage(
                    "Animated group OLE");
                using var storage = new MemoryStream(storageBytes,
                    writable: false);
                PowerPointOleObject ole = slide.AddOleObject(storage,
                    "Package", OfficeIMO.OfficeLength.Emus(100000L), OfficeIMO.OfficeLength.Emus(100000L), OfficeIMO.OfficeLength.Emus(1000000L), OfficeIMO.OfficeLength.Emus(600000L));
                PowerPointAutoShape marker = slide.AddRectangle(
                    OfficeIMO.OfficeLength.Emus(1200000L), OfficeIMO.OfficeLength.Emus(100000L), OfficeIMO.OfficeLength.Emus(600000L), OfficeIMO.OfficeLength.Emus(600000L));
                PowerPointGroupShape group = slide.GroupShapes(
                    new PowerPointShape[] { ole, marker },
                    "Animated group");
                slide.AddClassicAnimation(marker,
                    PowerPointClassicAnimationEffect.Fade);
                Assert.Contains(slide.GetGroupChildren(group),
                    child => child.Id == marker.Id);
                bytes = presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt });
            }

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointDocument projected =
                PowerPointDocument.Load(input);
            PowerPointSlide projectedSlide = Assert.Single(
                projected.Slides);
            PowerPointGroupShape projectedGroup = Assert.Single(
                projectedSlide.Shapes.OfType<PowerPointGroupShape>());
            PowerPointShape projectedMarker = Assert.Single(
                projectedSlide.GetGroupChildren(projectedGroup));
            PowerPointClassicAnimation animation = Assert.Single(
                projectedSlide.ClassicAnimations);
            Assert.Equal(projectedMarker.Id, animation.ShapeId);
            Assert.Empty(projected.ValidateDocument());
        }

        [Fact]
        public void OleCreationReservesBothShapeIdsBeforeAddingParts() {
            byte[] bytes;
            using (PowerPointDocument presentation =
                   PowerPointDocument.Create()) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape first = slide.AddRectangle(
                    OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape second = slide.AddRectangle(
                    OfficeIMO.OfficeLength.Emus(4000), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointGroupShape group = slide.GroupShapes(
                    new PowerPointShape[] { first, second });
                group.GroupShape.Descendants<
                        P.NonVisualDrawingProperties>().Last().Id =
                    uint.MaxValue - 1U;
                bytes = presentation.ToBytes();
            }

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointDocument loaded =
                PowerPointDocument.Load(input);
            PowerPointSlide loadedSlide = Assert.Single(loaded.Slides);
            int partCount = loadedSlide.SlidePart.Parts.Count();
            using var storage = new MemoryStream(
                CreateOleTestStorage("ID reservation"), writable: false);

            Assert.Throws<InvalidOperationException>(() => loadedSlide
                .AddOleObject(storage, "Package"));

            Assert.Equal(partCount, loadedSlide.SlidePart.Parts.Count());
            PowerPointAutoShape finalShape = loadedSlide.AddRectangle(
                OfficeIMO.OfficeLength.Emus(7000), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(1000));
            Assert.Equal(uint.MaxValue, finalShape.Id);
        }

        [Fact]
        public void NativeWriter_ReportsOpenXmlSignatureBeforeMutationPolicyRuns() {
            using PowerPointDocument presentation =
                PowerPointDocument.Create();
            presentation.AddSlide().AddTextBox("Signed package preflight");
            DigitalSignatureOriginPart origin = presentation.OpenXmlDocument
                .AddDigitalSignatureOriginPart();
            XmlSignaturePart signature = origin.AddNewPart<XmlSignaturePart>();
            Feed(signature,
                "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\" />");

            LegacyPptWritePreflightReport report =
                presentation.AnalyzeLegacyPptWrite();

            Assert.Contains(report.Findings, finding =>
                finding.Feature == LegacyPptFeature.DigitalSignatures
                && finding.Code == "ppt-write-openxml-signature");
            Assert.Throws<PowerPointSignedDocumentMutationException>(() =>
                presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt }));

            presentation.SignatureMutationPolicy =
                OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures;
            byte[] bytes = presentation.ToBytes(new PowerPointSaveOptions { Format = PowerPointFileFormat.Ppt });
            Assert.NotEmpty(bytes);
            Assert.False(presentation.InspectSignatures().HasSignatureMetadata);
        }

        private static void Feed(OpenXmlPart part, string xml) =>
            Feed(part, Encoding.UTF8.GetBytes(xml));

        private static void Feed(OpenXmlPart part, byte[] bytes) {
            using var stream = new MemoryStream(bytes, writable: false);
            part.FeedData(stream);
        }
    }
}
