using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Capabilities;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointLegacyPptTests {
        [Fact]
        public void NativeWriter_BlocksUnrepresentablePackagePartsAndRelationships() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
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
                && finding.Code == "PPT-WRITE-CUSTOM-XML");
            Assert.Contains(report.Findings, finding =>
                finding.Code == "PPT-WRITE-EXTENDED-PART");
            Assert.Contains(report.Findings, finding =>
                finding.Code == "PPT-WRITE-EMBEDDED-PACKAGE");
            Assert.Contains(report.Findings, finding =>
                finding.Code == "PPT-WRITE-EXTERNAL-RELATIONSHIP");
            Assert.Throws<NotSupportedException>(() =>
                presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void ImportedPresentation_CustomXmlBypassesNeitherPreservationPath() {
            byte[] sourceBytes;
            using (PowerPointPresentation source =
                    PowerPointPresentation.Create()) {
                source.AddSlide().AddTextBox("Imported package preflight");
                sourceBytes = source.ToBytes(PowerPointFileFormat.Ppt);
            }

            using var input = new MemoryStream(sourceBytes);
            using PowerPointPresentation imported =
                PowerPointPresentation.Load(input);
            CustomXmlPart customXml = imported.OpenXmlDocument
                .PresentationPart!.AddCustomXmlPart(
                    CustomXmlPartType.CustomXml);
            Feed(customXml, "<imported />");

            LegacyPptWritePreflightReport report =
                imported.AnalyzeLegacyPptWrite();

            Assert.Contains(report.Findings, finding =>
                finding.Code == "PPT-WRITE-CUSTOM-XML");
            Assert.Throws<NotSupportedException>(() =>
                imported.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_ClassifiesAdvancedExtendedParts() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
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
                && finding.Code == "PPT-WRITE-ACTIVEX");
            Assert.Contains(report.Findings, finding =>
                finding.Feature == LegacyPptFeature.UnknownRecordsAndStreams
                && finding.Code == "PPT-WRITE-WEB-EXTENSION");
            Assert.Contains(report.Findings, finding =>
                finding.Feature == LegacyPptFeature.VbaProjects
                && finding.Code == "PPT-WRITE-VBA-EXTENDED-PART");
            Assert.DoesNotContain(report.Findings, finding =>
                finding.Code == "PPT-WRITE-EXTENDED-PART");
        }

        [Fact]
        public void NativeWriter_BlocksTypedPartsAndPackageRootRelationships() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
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
                finding.Code == "PPT-WRITE-PACKAGE-PART");
            Assert.Contains(report.Findings, finding =>
                finding.Code == "PPT-WRITE-EXTERNAL-RELATIONSHIP");
            Assert.Throws<NotSupportedException>(() =>
                presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void NativeWriter_RecognizesMappedOleObjectInsideGroup() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            byte[] storageBytes = CreateOleTestStorage("Grouped OLE");
            using var storage = new MemoryStream(storageBytes,
                writable: false);
            PowerPointOleObject ole = slide.AddOleObject(storage, "Package",
                100000L, 100000L, 1000000L, 600000L);
            PowerPointAutoShape marker = slide.AddRectangle(
                1200000L, 100000L, 600000L, 600000L);
            slide.GroupShapes(new PowerPointShape[] { ole, marker },
                "Grouped OLE");

            LegacyPptWritePreflightReport report =
                presentation.AnalyzeLegacyPptWrite();

            Assert.DoesNotContain(report.Findings, finding =>
                finding.Code == "PPT-WRITE-EMBEDDED-PACKAGE");
            Assert.True(report.CanWrite,
                string.Join(Environment.NewLine, report.Findings));
            Assert.NotEmpty(presentation.ToBytes(PowerPointFileFormat.Ppt));
        }

        [Fact]
        public void GroupProjection_MapsAnimationsPastOmittedOleChildren() {
            byte[] bytes;
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = presentation.AddSlide();
                byte[] storageBytes = CreateOleTestStorage(
                    "Animated group OLE");
                using var storage = new MemoryStream(storageBytes,
                    writable: false);
                PowerPointOleObject ole = slide.AddOleObject(storage,
                    "Package", 100000L, 100000L, 1000000L, 600000L);
                PowerPointAutoShape marker = slide.AddRectangle(
                    1200000L, 100000L, 600000L, 600000L);
                PowerPointGroupShape group = slide.GroupShapes(
                    new PowerPointShape[] { ole, marker },
                    "Animated group");
                slide.AddClassicAnimation(marker,
                    PowerPointClassicAnimationEffect.Fade);
                Assert.Contains(slide.GetGroupChildren(group),
                    child => child.Id == marker.Id);
                bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
            }

            using var input = new MemoryStream(bytes, writable: false);
            using PowerPointPresentation projected =
                PowerPointPresentation.Load(input);
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
        public void NativeWriter_ReportsOpenXmlSignatureBeforeMutationPolicyRuns() {
            using PowerPointPresentation presentation =
                PowerPointPresentation.Create();
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
                && finding.Code == "PPT-WRITE-OPENXML-SIGNATURE");
            Assert.Throws<PowerPointSignedPresentationMutationException>(() =>
                presentation.ToBytes(PowerPointFileFormat.Ppt));

            presentation.SignatureMutationPolicy =
                PowerPointSignatureMutationPolicy.RemoveInvalidatedSignatures;
            byte[] bytes = presentation.ToBytes(PowerPointFileFormat.Ppt);
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
