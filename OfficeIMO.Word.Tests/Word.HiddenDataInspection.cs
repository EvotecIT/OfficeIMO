using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
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

        private static byte[] JpegScanData(byte[] jpeg) {
            for (int index = 2; index + 1 < jpeg.Length; index++) {
                if (jpeg[index] == 0xFF && jpeg[index + 1] == 0xDA) return jpeg.Skip(index).ToArray();
            }
            throw new InvalidDataException("No JPEG start-of-scan marker.");
        }
    }
}
