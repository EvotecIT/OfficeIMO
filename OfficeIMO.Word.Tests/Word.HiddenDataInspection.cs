using System.IO.Packaging;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Provenance;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_RemoveHiddenData_ScrubsWordAuthoredReviewFixtureAndKeepsTrackedChanges() {
            string filePath = CopyFixtureDoc(Path.Combine("Word", "PremiumGaps", "ReviewRedline", "word-authored-comments-revisions.docx"));
            HashSet<string> errorsBefore;

            using (WordDocument document = WordDocument.Load(filePath)) {
                errorsBefore = new HashSet<string>(document.ValidateDocument().Select(error => error.Description));
                OfficeHiddenDataReport found = document.InspectHiddenData();
                Assert.All(found.Findings, finding => Assert.Equal(OfficeHiddenDataOutcome.Found, finding.Outcome));
                Assert.NotEmpty(found.GetFindings(OfficeHiddenDataKind.Comments));
                Assert.NotEmpty(found.GetFindings(OfficeHiddenDataKind.TrackedChanges));
                Assert.Contains(found.GetFindings(OfficeHiddenDataKind.PersonalIdentities), finding => finding.Value == "OfficeIMO Word Fixture");

                OfficeHiddenDataReport removed = document.RemoveHiddenData();
                Assert.All(removed.GetFindings(OfficeHiddenDataKind.Comments), finding => Assert.Equal(OfficeHiddenDataOutcome.Removed, finding.Outcome));
                Assert.All(removed.GetFindings(OfficeHiddenDataKind.TrackedChanges), finding => Assert.Equal(OfficeHiddenDataOutcome.Retained, finding.Outcome));
                Assert.Contains(removed.GetFindings(OfficeHiddenDataKind.PersonalIdentities), finding => finding.Outcome == OfficeHiddenDataOutcome.Anonymized);
                document.Save();
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, false)) {
                MainDocumentPart main = package.MainDocumentPart!;
                Assert.Null(main.WordprocessingCommentsPart);
                Assert.Empty(main.Document!.Body!.Descendants<DocumentFormat.OpenXml.Wordprocessing.CommentReference>());
                Assert.Null(main.WordprocessingPeoplePart);
            }

            using (WordDocument reopened = WordDocument.Load(filePath)) {
                Assert.Empty(reopened.ValidateDocument().Select(error => error.Description).Where(description => !errorsBefore.Contains(description)));
                OfficeHiddenDataReport after = reopened.InspectHiddenData();
                Assert.Empty(after.GetFindings(OfficeHiddenDataKind.Comments));
                Assert.Empty(after.GetFindings(OfficeHiddenDataKind.DocumentProperties));
                Assert.Empty(after.GetFindings(OfficeHiddenDataKind.RevisionIdentifiers));
                Assert.NotEmpty(after.GetFindings(OfficeHiddenDataKind.TrackedChanges));
                Assert.All(after.GetFindings(OfficeHiddenDataKind.PersonalIdentities), finding => Assert.Equal("Author", finding.Value));
            }
        }

        [Fact]
        public void Test_RemoveHiddenData_AllCategoriesSurviveSaveAndKeepImagePixels() {
            string filePath = Path.Combine(_directoryWithFiles, "HiddenData.AllCategories.docx");
            string imagePath = Path.Combine(_directoryWithFiles, "gps.jpg");
            byte[] source = File.ReadAllBytes(Path.Combine(_directoryWithImages, "Kulek.jpg"));
            OfficeImageMetadata gps = OfficeImageMetadata.Read(source);
            gps.SetExifValue(OfficeExifTag.GPSLatitudeRef, "N");
            gps.SetExifValue(OfficeExifTag.Artist, "Private Photographer");
            File.WriteAllBytes(imagePath, OfficeImageMetadata.Apply(source, gps));

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Properties.Creator = "Private Author";
                document.Properties.LastModifiedBy = "Private Editor";
                document.ApplicationProperties.Company = "Private Company";
                document.CustomProperties.Add("ClientName", new WordCustomProperty("Private Client"));
                document.SetDocumentVariable("CaseNumber", "42");
                WordParagraph paragraph = document.AddParagraph("Visible text. ");
                paragraph.AddText("Hidden draft note.").Hidden = true;
                document.AddParagraph().AddImage(imagePath);
                document.AddParagraph("Reviewed paragraph").AddComment("Private Reviewer", "PR", "Check this.");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                OfficeHiddenDataReport removed = document.RemoveHiddenData(new OfficeHiddenDataOptions {
                    Kinds = OfficeHiddenDataKind.All,
                    TrackedChanges = OfficeTrackedChangesAction.AcceptAll
                });
                Assert.Contains(removed.Findings, finding => finding.Kind == OfficeHiddenDataKind.HiddenContent && finding.Value == "Hidden draft note.");
                Assert.Contains(removed.Findings, finding => finding.Kind == OfficeHiddenDataKind.ImageMetadata && finding.Description.Contains("GPS"));
                Assert.DoesNotContain(removed.Findings, finding => finding.Outcome == OfficeHiddenDataOutcome.Retained);
                document.Save();
            }

            using (WordDocument reopened = WordDocument.Load(filePath)) {
                Assert.Empty(reopened.ValidateDocument());
                OfficeHiddenDataReport after = reopened.InspectHiddenData();
                Assert.False(after.HasHiddenData, string.Join(Environment.NewLine, after.Findings));
                Assert.Empty(reopened.CustomProperties);
                Assert.Empty(reopened.DocumentVariables);
                Assert.DoesNotContain("Hidden draft note.", string.Concat(reopened.Paragraphs.Select(paragraph => paragraph.Text)));

                byte[] scrubbed = reopened.Images.Single().ToBytes();
                OfficeImageMetadata metadata = OfficeImageMetadata.Read(scrubbed);
                Assert.DoesNotContain(metadata.ExifValues, value => value.Tag.Directory == OfficeExifDirectory.Gps || value.Tag.Equals(OfficeExifTag.Artist));
                Assert.Equal(JpegScanData(source), JpegScanData(scrubbed));
            }
        }

        [Fact]
        public void Test_RemoveHiddenData_KeepsDigitalSourceTypeUnlessProvenanceIsSelected() {
            string filePath = Path.Combine(_directoryWithFiles, "HiddenData.DigitalSourceType.docx");
            string imagePath = Path.Combine(_directoryWithFiles, "ai-gps.jpg");
            byte[] source = File.ReadAllBytes(Path.Combine(_directoryWithImages, "Kulek.jpg"));
            OfficeImageMetadata tagged = OfficeImageMetadata.Read(source);
            tagged.SetExifValue(OfficeExifTag.GPSLatitudeRef, "N");
            tagged.XmpProfile = Encoding.UTF8.GetBytes(
                "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
                "<rdf:Description rdf:about=\"\" xmlns:Iptc4xmpExt=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\" xmlns:xmp=\"http://ns.adobe.com/xap/1.0/\" xmp:CreatorTool=\"Private Tool\">" +
                "<Iptc4xmpExt:DigitalSourceType rdf:resource=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
                "</rdf:Description></rdf:RDF></x:xmpmeta>");
            File.WriteAllBytes(imagePath, OfficeImageMetadata.Apply(source, tagged));
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph().AddImage(imagePath);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                OfficeHiddenDataFinding declaration = Assert.Single(document.InspectHiddenData().GetFindings(OfficeHiddenDataKind.Provenance));
                Assert.Equal(OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia, declaration.Provenance!.DigitalSourceKind);

                OfficeHiddenDataReport removed = document.RemoveHiddenData();
                Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(removed.GetFindings(OfficeHiddenDataKind.ImageMetadata)).Outcome);
                Assert.Equal(OfficeHiddenDataOutcome.Retained, Assert.Single(removed.GetFindings(OfficeHiddenDataKind.Provenance)).Outcome);
                document.Save();
            }

            Assert.True(WordDocument.InspectProvenance(filePath).HasGenerativeAiDeclaration);
            using (WordDocument document = WordDocument.Load(filePath)) {
                OfficeHiddenDataReport kept = document.InspectHiddenData();
                Assert.Empty(kept.GetFindings(OfficeHiddenDataKind.ImageMetadata));
                Assert.DoesNotContain("Private Tool", Encoding.UTF8.GetString(document.Images.Single().ToBytes()));

                OfficeHiddenDataReport removed = document.RemoveHiddenData(new OfficeHiddenDataOptions { Kinds = OfficeHiddenDataKind.Provenance });
                Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(removed.GetFindings(OfficeHiddenDataKind.Provenance)).Outcome);
                document.Save();
            }

            Assert.Empty(WordDocument.InspectProvenance(filePath).Evidence);
        }

        [Fact]
        public void Test_RemoveHiddenData_LeavesContentCredentialsIntactUnlessProvenanceIsSelected() {
            string filePath = Path.Combine(_directoryWithFiles, "HiddenData.ContentCredentials.docx");
            string imagePath = Path.Combine(_directoryWithFiles, "c2pa-gps.jpg");
            byte[] source = File.ReadAllBytes(Path.Combine(_directoryWithImages, "Kulek.jpg"));
            OfficeImageMetadata gps = OfficeImageMetadata.Read(source);
            gps.SetExifValue(OfficeExifTag.GPSLatitudeRef, "N");
            byte[] located = OfficeImageMetadata.Apply(source, gps);
            byte[] credentialed = Join(located.Take(2).ToArray(), CreateJpegApp11(CreateManifestStore()), located.Skip(2).ToArray());
            File.WriteAllBytes(imagePath, credentialed);
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph().AddImage(imagePath);
                document.Save();
            }
            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite)) {
                PackagePart manifest = package.CreatePart(new Uri("/META-INF/content_credential.c2pa", UriKind.Relative), "application/c2pa");
                using Stream stream = manifest.GetStream();
                byte[] store = CreateManifestStore();
                stream.Write(store, 0, store.Length);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                OfficeHiddenDataReport removed = document.RemoveHiddenData();
                OfficeHiddenDataFinding metadata = Assert.Single(removed.GetFindings(OfficeHiddenDataKind.ImageMetadata));
                Assert.Equal(OfficeHiddenDataOutcome.Retained, metadata.Outcome);
                Assert.Equal(OfficeHiddenDataReason.ProtectsContentCredentials, metadata.Reason);
                OfficeHiddenDataFinding image = Assert.Single(removed.GetFindings(OfficeHiddenDataKind.Provenance), finding => finding.Location != "/META-INF/content_credential.c2pa");
                Assert.Equal(OfficeHiddenDataOutcome.Retained, image.Outcome);
                OfficeHiddenDataFinding package = Assert.Single(removed.GetFindings(OfficeHiddenDataKind.Provenance), finding => finding.Location == "/META-INF/content_credential.c2pa");
                Assert.Equal(OfficeHiddenDataOutcome.Removed, package.Outcome);
                Assert.Equal(OfficeHiddenDataScope.Package, package.Scope);
                Assert.Equal(OfficeHiddenDataScope.EmbeddedImage, image.Scope);
                Assert.NotNull(package.Note);
                Assert.Equal(credentialed, document.Images.Single().ToBytes());
                document.Save();
            }

            Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, Assert.Single(WordDocument.InspectProvenance(filePath).Evidence).Carrier);
            using (WordDocument document = WordDocument.Load(filePath)) {
                OfficeHiddenDataReport removed = document.RemoveHiddenData(new OfficeHiddenDataOptions {
                    Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.Provenance
                });
                Assert.All(removed.GetFindings(OfficeHiddenDataKind.Provenance), finding => Assert.Equal(OfficeHiddenDataOutcome.Removed, finding.Outcome));
                Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(removed.GetFindings(OfficeHiddenDataKind.ImageMetadata)).Outcome);
                document.Save();
            }

            Assert.Empty(WordDocument.InspectProvenance(filePath).Evidence);
            using (WordDocument reopened = WordDocument.Load(filePath)) {
                Assert.Empty(reopened.ValidateDocument());
                byte[] scrubbed = reopened.Images.Single().ToBytes();
                Assert.DoesNotContain(OfficeImageMetadata.Read(scrubbed).ExifValues, value => value.Tag.Directory == OfficeExifDirectory.Gps);
                Assert.Equal(JpegScanData(source), JpegScanData(scrubbed));
            }
        }

        [Fact]
        public void Test_RemoveHiddenData_ProvenanceCarriersCanBeKeptSeparately() {
            string filePath = Path.Combine(_directoryWithFiles, "HiddenData.ProvenanceCarriers.docx");
            string imagePath = Path.Combine(_directoryWithFiles, "c2pa-ai.jpg");
            byte[] source = File.ReadAllBytes(Path.Combine(_directoryWithImages, "Kulek.jpg"));
            OfficeImageMetadata tagged = OfficeImageMetadata.Read(source);
            tagged.XmpProfile = Encoding.UTF8.GetBytes(
                "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
                "<rdf:Description rdf:about=\"\" xmlns:Iptc4xmpExt=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
                "<Iptc4xmpExt:DigitalSourceType rdf:resource=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/>" +
                "</rdf:Description></rdf:RDF></x:xmpmeta>");
            byte[] declared = OfficeImageMetadata.Apply(source, tagged);
            File.WriteAllBytes(imagePath, Join(declared.Take(2).ToArray(), CreateJpegApp11(CreateManifestStore()), declared.Skip(2).ToArray()));
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph().AddImage(imagePath);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var options = new OfficeHiddenDataOptions { Kinds = OfficeHiddenDataKind.Provenance };
                options.Provenance.RemoveC2paManifests = false;
                OfficeHiddenDataReport removed = document.RemoveHiddenData(options);
                OfficeHiddenDataFinding credential = Assert.Single(removed.GetFindings(OfficeHiddenDataKind.Provenance),
                    finding => finding.Provenance!.Carrier == OfficeProvenanceCarrierKind.C2paManifest);
                Assert.Equal(OfficeHiddenDataOutcome.Retained, credential.Outcome);
                Assert.Equal(OfficeHiddenDataOutcome.Removed, Assert.Single(removed.GetFindings(OfficeHiddenDataKind.Provenance),
                    finding => finding.Provenance!.Carrier == OfficeProvenanceCarrierKind.IptcDigitalSourceType).Outcome);
                document.Save();
            }

            OfficeProvenanceEvidence kept = Assert.Single(WordDocument.InspectProvenance(filePath).Evidence);
            Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, kept.Carrier);
        }

        [Fact]
        public void Test_RemoveHiddenData_ReportsAndCleansPicturesTheDocumentDoesNotUse() {
            string filePath = Path.Combine(_directoryWithFiles, "HiddenData.UnusedPicture.docx");
            byte[] source = File.ReadAllBytes(Path.Combine(_directoryWithImages, "Kulek.jpg"));
            OfficeImageMetadata gps = OfficeImageMetadata.Read(source);
            gps.SetExifValue(OfficeExifTag.GPSLatitudeRef, "N");
            byte[] located = OfficeImageMetadata.Apply(source, gps);
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Text only");
                document.Save();
            }
            using (Package package = Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite)) {
                using Stream stream = package.CreatePart(new Uri("/word/media/leftover.jpeg", UriKind.Relative), "image/jpeg").GetStream();
                stream.Write(located, 0, located.Length);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                // An unsaved edit that adds a part with its own relationships, which .NET Framework writes only when the package is flushed.
                document.PageSetup.Header!.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));

                OfficeHiddenDataFinding found = Assert.Single(document.InspectHiddenData().GetFindings(OfficeHiddenDataKind.ImageMetadata),
                    finding => finding.Part == "/word/media/leftover.jpeg");
                Assert.Equal(OfficeHiddenDataScope.EmbeddedImage, found.Scope);
                Assert.Contains("not used", found.Description);

                OfficeHiddenDataFinding removed = Assert.Single(document.RemoveHiddenData().GetFindings(OfficeHiddenDataKind.ImageMetadata),
                    finding => finding.Part == "/word/media/leftover.jpeg");
                Assert.Equal(OfficeHiddenDataOutcome.Removed, removed.Outcome);
                document.Save();
            }

            using (WordDocument reopened = WordDocument.Load(filePath)) {
                Assert.Empty(reopened.InspectHiddenData().GetFindings(OfficeHiddenDataKind.ImageMetadata));
            }
        }

        private static byte[] CreateJpegApp11(byte[] manifest) {
            byte[] payload = Join(new byte[] { 0x4A, 0x50, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01 }, manifest);
            int length = payload.Length + 2;
            return Join(new byte[] { 0xFF, 0xEB, (byte)(length >> 8), (byte)length }, payload);
        }

        private static byte[] CreateManifestStore() {
            byte[] Box(string type, byte[] payload) {
                byte[] box = new byte[payload.Length + 8];
                box[0] = (byte)(box.Length >> 24);
                box[1] = (byte)(box.Length >> 16);
                box[2] = (byte)(box.Length >> 8);
                box[3] = (byte)box.Length;
                Encoding.ASCII.GetBytes(type).CopyTo(box, 4);
                payload.CopyTo(box, 8);
                return box;
            }
            byte[] Description(string code, string label) => Box("jumd", Join(Encoding.ASCII.GetBytes(code),
                new byte[] { 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71, 0x03 }, Encoding.ASCII.GetBytes(label + "\0")));
            byte[] Labelled(string code, string label) => Box("jumb", Join(Description(code, label), Box("cbor", new byte[] { 0xA0 })));
            byte[] assertions = Box("jumb", Join(Description("c2as", "c2pa.assertions"), Labelled("c2ac", "c2pa.test")));
            byte[] manifest = Box("jumb", Join(Description("c2ma", "m"), assertions, Labelled("c2cl", "c2pa.claim"), Labelled("c2cs", "c2pa.signature")));
            return Box("jumb", Join(Description("c2pa", "c2pa"), manifest));
        }

        private static byte[] Join(params byte[][] parts) => parts.SelectMany(part => part).ToArray();

        private static byte[] JpegScanData(byte[] jpeg) {
            for (int index = 2; index + 1 < jpeg.Length; index++) {
                if (jpeg[index] == 0xFF && jpeg[index + 1] == 0xDA) return jpeg.Skip(index).ToArray();
            }
            throw new InvalidDataException("No JPEG start-of-scan marker.");
        }
    }
}
